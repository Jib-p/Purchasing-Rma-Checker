import pandas as pd
import os
from datetime import datetime, timedelta

# === CONFIG ===
TODAY = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
LOOKBACK_DAYS = 14
INVOICES_FILE = "Invoices.csv"
QC_CODES_FILE = "RMA_Codes.csv"

QC_PASS_CODES = {"M00", "LCD-L0"}
UNLOCKED_CARRIERS = {"Unlocked", "International (Unlocked)", "Wi-Fi Only", "WiFi", ""}
CONDITION_FLAGS = ["parts message", "front glass only", "back glass only"]

ICE_COLS = [
    "WID", "Vendor", "Vendor Invoice #", "Cost Amount", "Condition", "IMEI",
    "Brand", "Model", "Carrier", "Carrier GSX", "QC Error Code",
    "Physical QC Board", "Inventory",
]

# Vendor definitions: vendor pattern -> { ice files, output, skip rules }
VENDORS = {
    "hyla": {
        "name": "HYLA Mobile",
        "invoice_pattern": "hyla",
        "ice_files": [
            "260410ICEReport_jibrilpascua197484055269d90d3a5418c4.99316304.xlsx",
        ],
        "output": "HYLA_Return_Candidates.csv",
        "lookback_days": 7,
        "skip_carrier_check": False,
        "condition_flags": ["parts message", "front glass only", "back glass only"],
    },
    "att": {
        "name": "AT&T Mobility",
        "invoice_pattern": "at&t|at.t mobility",
        "ice_files": [
            "AT&T/260410ICEReport_jibrilpascua132994518369d912ed0d0852.73543945.xlsx",
        ],
        "output": "ATT_Return_Candidates.csv",
        "lookback_days": 60,
        "skip_carrier_check": True,
        "condition_flags": [
            "front glass only", "back glass only", "not cleared",
            "id locked", "ber/scrap",
        ],
    },
}


def get_recent_invoices(vendor_cfg):
    """Get invoice numbers for a vendor within its lookback window."""
    cutoff = TODAY - timedelta(days=vendor_cfg["lookback_days"])
    df = pd.read_csv(INVOICES_FILE, encoding="utf-8-sig", low_memory=False)
    df = df[df["Vendor"].str.contains(vendor_cfg["invoice_pattern"], case=False, na=False)]
    df["Date Received"] = pd.to_datetime(df["Date Received"], format="mixed", errors="coerce")
    df = df[(df["Date Received"] >= cutoff) & (df["Date Received"] <= TODAY)]
    return set(df["Invoice #"].astype(str).str.strip())


def load_qc_error_codes():
    """Load QC error code descriptions into a dict."""
    df = pd.read_csv(QC_CODES_FILE, encoding="utf-8-sig")
    df = df[df["Code"].notna() & (df["Code"].str.strip() != "-- UNKNOWN --")]
    return dict(zip(df["Code"].str.strip(), df["Description"].str.strip()))


def build_reasons(row, qc_lookup, vendor_cfg):
    """Build a list of return reasons for a single device row."""
    reasons = []
    flags = {"qc": 0, "carrier": 0, "condition": 0, "battery": 0}

    # QC error codes
    qc_raw = row.get("QC Error Code")
    qc_code = str(qc_raw).strip() if pd.notna(qc_raw) else ""
    if qc_code:
        alarming = []
        for code in (c.strip() for c in qc_code.split(",")):
            if code and code not in QC_PASS_CODES:
                desc = qc_lookup.get(code)
                alarming.append(f"{code} ({desc})" if desc else code)
        if alarming:
            reasons.append(f"QC: {'; '.join(alarming)}")
            flags["qc"] = 1

    # Carrier lock (skip for vendors where all devices share the same carrier)
    if not vendor_cfg["skip_carrier_check"]:
        carrier = str(row.get("Carrier") or "").strip()
        gsx = str(row.get("Carrier GSX") or "").strip()
        if carrier not in UNLOCKED_CARRIERS and gsx != "Unlocked":
            lock_note = f"Carrier Locked: {carrier}"
            if gsx:
                lock_note += f" (GSX: {gsx})"
            reasons.append(lock_note)
            flags["carrier"] = 1

    # Condition flags (vendor-specific keywords)
    cond = str(row.get("Condition") or "").strip()
    cond_flags = vendor_cfg["condition_flags"]
    if any(flag in cond.lower() for flag in cond_flags):
        reasons.append(f"Condition: {cond}")
        flags["condition"] = 1

    # Low battery
    if str(row.get("Physical QC Board") or "").strip() == "Battery 0-69%":
        reasons.append(f"Low Battery: {row['Physical QC Board']}")
        flags["battery"] = 1

    return reasons, flags


def process_vendor(vendor_key, vendor_cfg, qc_lookup):
    """Process a single vendor's ICE report(s) and return candidates."""
    name = vendor_cfg["name"]
    print(f"\n{'=' * 60}")
    print(f"Vendor: {name}")
    print(f"{'=' * 60}")

    # Step 1: Get recent invoices for this vendor
    cutoff = TODAY - timedelta(days=LOOKBACK_DAYS)
    print(f"Looking for invoices received {cutoff:%m/%d/%y} - {TODAY:%m/%d/%y}...")
    target_invoices = get_recent_invoices(vendor_cfg)
    if not target_invoices:
        print(f"  No recent invoices found for {name}. Skipping.")
        return None
    print(f"  Found {len(target_invoices)} invoices: {', '.join(sorted(target_invoices))}")

    # Step 2: Read ICE report(s) and filter to target invoices
    frames = []
    for ice_path in vendor_cfg["ice_files"]:
        if not os.path.exists(ice_path):
            print(f"  WARNING: ICE file not found: {ice_path}")
            continue
        print(f"  Scanning: {ice_path}...")
        ice = pd.read_excel(
            ice_path, usecols=ICE_COLS, dtype={"IMEI": str, "WID": str}, engine="openpyxl",
        )
        ice["Vendor Invoice #"] = ice["Vendor Invoice #"].astype(str).str.strip()
        ice = ice[ice["Vendor Invoice #"].isin(target_invoices)].copy()
        frames.append(ice)

    if not frames:
        print(f"  No ICE data loaded for {name}.")
        return None
    ice = pd.concat(frames, ignore_index=True)
    ice["IMEI"] = ice["IMEI"].fillna("").str.split(".").str[0]
    print(f"  Matched {len(ice)} devices across target invoices")

    # Step 3: Build reasons for each device
    results = ice.apply(lambda row: build_reasons(row, qc_lookup, vendor_cfg), axis=1)
    ice["Reason(s)"] = results.apply(lambda x: " | ".join(x[0]))
    flag_df = results.apply(lambda x: x[1]).apply(pd.Series)

    # Keep only rows with at least one reason
    candidates = ice[ice["Reason(s)"] != ""].copy()

    # Step 4: Sort by priority
    candidates["_qc"] = ~candidates["Reason(s)"].str.contains("QC:", na=False)
    candidates["_cond"] = ~candidates["Reason(s)"].str.contains("Condition:", na=False)
    candidates["_gsx"] = ~candidates["Reason(s)"].str.contains("GSX: Locked", na=False)
    candidates["_carrier"] = ~(
        candidates["Reason(s)"].str.contains("Carrier Locked:", na=False)
        & ~candidates["Reason(s)"].str.contains("GSX:", na=False)
    )
    candidates = candidates.sort_values(["_qc", "_cond", "_gsx", "_carrier"])
    candidates = candidates.drop(columns=["_qc", "_cond", "_gsx", "_carrier"])

    # Step 5: Rename and write output
    candidates = candidates.rename(columns={
        "Vendor Invoice #": "Invoice",
        "Cost Amount": "Cost",
    })
    out_cols = [
        "WID", "Vendor", "Invoice", "Brand", "Model", "Carrier", "Carrier GSX",
        "Condition", "QC Error Code", "Physical QC Board", "Inventory",
        "IMEI", "Cost", "Reason(s)",
    ]
    output_file = vendor_cfg["output"]
    candidates[out_cols].to_csv(output_file, index=False, encoding="utf-8-sig")

    # Step 6: Print summary
    stats = flag_df.loc[candidates.index].sum()
    gsx_locked = candidates["Reason(s)"].str.contains("GSX: Locked", na=False).sum()
    carrier_no_gsx = int(stats.get("carrier", 0)) - gsx_locked

    print(f"\n  RESULTS for {name}")
    print(f"  {'-' * 50}")
    print(f"  Total return candidates: {len(candidates)}")
    print(f"    - QC Error Code issues:     {int(stats.get('qc', 0))}")
    print(f"    - Condition flagged:         {int(stats.get('condition', 0))}")
    print(f"    - Low Battery (0-69%):       {int(stats.get('battery', 0))}")
    if not vendor_cfg["skip_carrier_check"]:
        print(f"    - Carrier Locked (GSX):      {gsx_locked}")
        print(f"    - Carrier Locked (no GSX):   {carrier_no_gsx}")
    print(f"  Output saved to: {output_file}")

    # Preview
    print(f"\n  --- Preview (first 20) ---")
    for _, c in candidates.head(20).iterrows():
        model = str(c.get("Model") or "")[:30]
        imei = str(c.get("IMEI") or "")
        reasons = str(c.get("Reason(s)") or "")[:60]
        print(f"    WID {c['WID']} | {c.get('Invoice', '')} | {model:30s} | IMEI: {imei} | {reasons}")
    if len(candidates) > 20:
        print(f"    ... and {len(candidates) - 20} more")

    return candidates


def main():
    print("=" * 60)
    print("Return Candidate Finder")
    print("=" * 60)

    # Load shared QC error code definitions
    qc_lookup = load_qc_error_codes()
    print(f"Loaded {len(qc_lookup)} QC error code definitions")

    # Process each vendor
    all_candidates = []
    for vendor_key, vendor_cfg in VENDORS.items():
        result = process_vendor(vendor_key, vendor_cfg, qc_lookup)
        if result is not None:
            all_candidates.append(result)

    # Final summary
    total = sum(len(df) for df in all_candidates)
    print(f"\n{'=' * 60}")
    print(f"GRAND TOTAL: {total} return candidates across {len(all_candidates)} vendor(s)")
    print(f"{'=' * 60}")



if __name__ == "__main__":
    main()
