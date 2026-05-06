# WID Checker — Project Documentation

> **Purpose:** Automates identification of mobile devices eligible for vendor RMA (Return Merchandise Authorization), turning a manual cross-reference task into a one-click report.

---

## Table of Contents

1. [Overview](#overview)
2. [Business Context](#business-context)
3. [Who Uses It](#who-uses-it)
4. [How It Works (End-to-End Flow)](#how-it-works-end-to-end-flow)
5. [Supported Vendors](#supported-vendors)
6. [Core Features](#core-features)
7. [Routes & Endpoints](#routes--endpoints)
8. [Data Sources](#data-sources)
9. [Database Schema](#database-schema)
10. [Technology Stack](#technology-stack)
11. [Environment & Deployment](#environment--deployment)
12. [Admin Features](#admin-features)
13. [Key Business Logic](#key-business-logic)
14. [Glossary](#glossary)

---

## Overview

**WID Checker** (WID = *Warehouse Inventory Device*) is an internal Flask web application built for **Mannapov LLC**. It automates the process of determining which used/refurbished mobile devices received from vendors are eligible to be returned under each vendor's RMA policy.

Before this tool: staff had to manually cross-reference QC error codes, carrier lock status, physical condition, and grade level against dozens of vendor-specific rules — a tedious, error-prone, time-consuming task.

After this tool: upload one Excel file, pick a vendor, and instantly get a categorized list of returnable devices with a shareable link for the returns team.

---

## Business Context

Mannapov LLC sources used phones from multiple wholesale and carrier channels. Each vendor has its own:

- **Return window** (e.g., 5 days, 30 days after invoice date)
- **Accepted defect categories** (camera fail, cracked screen, carrier locked, etc.)
- **Grade-specific rules** (a Grade A+ device has different return options than a Grade C device)
- **Threshold requirements** (some vendors require a minimum % of a lot to be defective before accepting a return)

WID Checker codifies all these rules and evaluates them automatically against the warehouse's ICE (Inventory Control Export) report.

---

## Who Uses It

| Role | Use Case |
|---|---|
| **Inventory / QC Coordinators** | Upload ICE reports, review return candidates |
| **RMA Specialists** | Generate CSVs + shareable links to dispute with vendors |
| **Account Managers** | Track disputed lots via invoice pages |
| **Warehouse Staff** | Use the phone scanner UI to photograph lots |
| **Admins** | Attach photos, manage invoice records |

---

## How It Works (End-to-End Flow)

```
 ┌──────────────────────┐
 │ 1. User uploads ICE  │  (Excel from warehouse system)
 │    report + selects  │
 │    vendor            │
 └──────────┬───────────┘
            │
            ▼
 ┌──────────────────────┐
 │ 2. App pulls recent  │  (From Dropbox IncomingInvoices.xlsx;
 │    vendor invoices   │   filtered by vendor's lookback window)
 └──────────┬───────────┘
            │
            ▼
 ┌──────────────────────┐
 │ 3. Filter ICE rows   │  Keeps only devices on recent invoices
 │    to matching       │
 │    invoices          │
 └──────────┬───────────┘
            │
            ▼
 ┌──────────────────────┐
 │ 4. Apply vendor RMA  │  Grade-specific rules from
 │    rules per device  │  Vendor RMA Guidelines.xlsx
 └──────────┬───────────┘
            │
            ▼
 ┌──────────────────────┐
 │ 5. Calculate return  │  Return % vs. vendor threshold
 │    thresholds        │
 └──────────┬───────────┘
            │
            ▼
 ┌──────────────────────┐
 │ 6. Persist to MySQL  │  Shareable token per invoice
 │    + generate CSV    │
 └──────────┬───────────┘
            │
            ▼
 ┌──────────────────────┐
 │ 7. User downloads    │  Plus public invoice pages
 │    CSV + shares      │  for collaboration
 │    links             │
 └──────────────────────┘
```

---

## Supported Vendors

| Vendor | Description | Lookback | Threshold |
|---|---|---|---|
| **HYLA TPS** | Device lifecycle platform, Tier 1 | 7 days | Per-grade rules |
| **HYLA DLS** | Device lifecycle platform, Tier 2 | 7 days | Per-grade rules |
| **AT&T Mobility** | Carrier buyback | 5 days | Unlocked must be >5% of lot |
| **Verizon Wireless** | Carrier buyback | 30 days | No threshold (accepts any %) |
| **Sprint** | Legacy carrier | 30 days | Carrier-locked only |
| **Superior / B Stock** | Apple-focused buyback | 14 days | Threshold-based |
| **Touchstone** | Functional failures only | 14 days | Partial credit model |
| **Clover** | *Disabled — rules pending* | — | — |

Vendor configuration lives in `app.py` under the `VENDORS` dict. Per-grade RMA rules are loaded at startup from `Vendor RMA Guidelines.xlsx`.

---

## Core Features

### 1. ICE Report Processing
Drop-in upload of the warehouse ICE Excel export. The app automatically matches to recent invoices, applies vendor rules, flags returnable devices, and exports a CSV.

### 2. QC Error Code Mapping
50+ QC error codes map to standard RMA defect categories (e.g., `M11` → Camera Fail, `LCD-L3` → Burns Level 3, `PHY-F01` → Cracked Back). Descriptions are shown inline in results.

### 3. Carrier Lock Detection (Priority Chain)
1. **ICE Detail Report** (if uploaded) — authoritative override
2. **Google/Pixel GSX column** — authoritative for Pixel devices
3. **Carrier column + GSX fallback** — for all other devices

### 4. Shareable Invoice Pages
Every processed invoice gets a unique URL-safe token (`/invoice/<token>`) that renders the full lot, return %, threshold status, and photo gallery. Safe to share externally.

### 5. Phone Scanner UI (`/scan`)
Real-time in-browser camera interface powered by **TensorFlow.js COCO-SSD**:
- Detects mobile phones in frame, draws an outline with confidence score
- Auto-captures when phone is centered + steady for 22 frames
- Stores captures in localStorage; auto-uploads when admin logs in
- Can be scoped to a specific invoice token

### 6. Admin Photo Management
Authenticated admins can attach, upload, or remove photos on any invoice page. Photos persist to `uploads/photos/`.

---

## Routes & Endpoints

| Route | Method | Auth | Purpose |
|---|---|---|---|
| `/` | GET | Public | Home — vendor selection + upload form |
| `/process` | POST | Public | Process ICE report, return candidates |
| `/download/<filename>` | GET | Public | Download CSV results |
| `/api/invoices/<vendor_key>` | GET | Public | JSON list of recent invoices (feeds UI dropdown) |
| `/scan` | GET | Public | Phone scanner UI |
| `/admin/login` | GET / POST | Public | Admin login |
| `/admin/logout` | POST | Admin | Logout |
| `/invoice/<token>` | GET | Public | View invoice lot + photos |
| `/invoice/<token>/photo` | POST | Admin | Add photo via URL |
| `/invoice/<token>/photo/<id>/delete` | POST | Admin | Remove photo |
| `/invoice/<token>/photo/upload` | POST | Admin | Upload photo from scanner |
| `/photos/<filename>` | GET | Public | Serve uploaded photos |

---

## Data Sources

| Source | Format | Location | Purpose |
|---|---|---|---|
| **IncomingInvoices.xlsx** | Excel | Dropbox | Master list of recent vendor invoices |
| **ICE Report** | Excel (user upload) | `/process` form | Warehouse inventory snapshot |
| **ICE Detail Report** | CSV (optional upload) | `/process` form | Carrier lock test results (override) |
| **Vendor RMA Guidelines.xlsx** | Excel | Project root | Per-vendor, per-grade RMA rules |
| **RMA_HYLA_REASONING.xlsx** | CSV | Project root | QC error code descriptions |

---

## Database Schema

MySQL 8+, accessed via SQLAlchemy ORM (see `db.py` and `setup_db.py`).

### `invoices`
| Column | Type | Notes |
|---|---|---|
| `invoice_number` | VARCHAR | Primary key |
| `vendor_key` | VARCHAR | e.g., `hyla`, `att`, `verizon` |
| `share_token` | VARCHAR | Secure URL-safe token |
| `return_count` | INT | # of returnable devices |
| `return_pct` | FLOAT | % of lot that is returnable |
| `invoice_qty` | INT | Qty1 from source invoice |
| `notes` | TEXT | Freeform |
| `created_at` / `updated_at` | TIMESTAMP | |

### `invoice_phones`
| Column | Type | Notes |
|---|---|---|
| `id` | INT | Primary key |
| `invoice_number` | VARCHAR | FK → `invoices` |
| `wid`, `imei`, `brand`, `model`, `carrier`, `grade`, `qc_error_code` | — | Device attributes |
| `reasons` | TEXT | Categorized return reasons |
| `scanned` | BOOL | Whether captured via scanner |

### `invoice_photos`
| Column | Type | Notes |
|---|---|---|
| `id` | INT | Primary key |
| `invoice_number` | VARCHAR | FK → `invoices` |
| `photo_url` | VARCHAR | Local path or external URL |
| `label` | VARCHAR | Optional caption |
| `created_at` | TIMESTAMP | |

---

## Technology Stack

| Layer | Technology |
|---|---|
| Backend | Flask (Python) |
| Database | MySQL 8+ via SQLAlchemy ORM |
| Data Processing | Pandas |
| Excel I/O | openpyxl, calamine |
| Frontend | Vanilla JavaScript, HTML5 Canvas |
| Object Detection | TensorFlow.js COCO-SSD (browser-side) |
| Auth | Flask sessions + HMAC password compare |
| File Handling | Werkzeug secure filenames, 50 MB upload limit |
| Config | python-dotenv |

See `requirements.txt` for the pinned dependency list.

---

## Environment & Deployment

The app reads configuration from `.env` (example in `.env.example`). Required variables include:

- `ADMIN_PASSWORD` — admin login password
- `SECRET_KEY` — Flask session signing key
- Database connection details (host, user, password, database name)
- Dropbox access token (for `IncomingInvoices.xlsx`)

Default dev server: `http://localhost:5000`. Runs via `python app.py`.

---

## Admin Features

- Password-protected login (`/admin/login`)
- Session-based auth enforced by `@require_admin` decorator
- Add, upload, and delete photos on any invoice page
- Scanner-driven photo auto-upload

---

## Key Business Logic

1. **Grade-based rules override static fallbacks.** HYLA TPS/DLS use per-grade rules from the Excel guidelines; other vendors use static returnable / not-sure sets.
2. **Threshold calculation.** `return_pct = (candidates_per_invoice / Qty1) × 100`. Compared against the vendor's threshold to decide if a return is viable.
3. **Priority sort in results.** Devices are ordered by severity: QC issues first, then condition, then carrier lock, then battery — so the most actionable items surface at the top.
4. **Detail Report override.** If an ICE Detail Report CSV is uploaded, its `Carrier Lock Status` column becomes the source of truth for all devices in the lot.
5. **Stable shareable tokens.** Each invoice's token is generated via `secrets.token_urlsafe` and persisted — links remain stable across re-uploads.

---

## Glossary

| Term | Meaning |
|---|---|
| **WID** | Warehouse Inventory Device — internal unique ID for each phone |
| **RMA** | Return Merchandise Authorization |
| **ICE Report** | Inventory Control Export — the warehouse's master inventory snapshot |
| **GSX** | Global Service Exchange — Apple's device lookup (used here for carrier-lock signal) |
| **QC Code** | Quality Control error code assigned during device testing |
| **Grade** | Condition tier (A+, B+, C, etc.) that drives vendor RMA eligibility |
| **Lookback Window** | How many days back the app searches for recent invoices per vendor |
| **Threshold** | Minimum return-percentage a lot must hit for some vendors to accept the return |
