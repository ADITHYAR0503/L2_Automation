# L2 Automation // PRESTO Operational Command Suite

An enterprise-grade internal operations automation platform architected for **PRESTO / Metrolinx L2 Operations**. Designed to streamline daily incident reporting, SLA calculations, MTTR analytics, and infrastructure health monitoring while maintaining strict human-in-the-loop validation.

> **System Architect & Creator:** Adithya  
> **Target Environment:** Windows / Python / Microsoft Outlook / SOTI MobiControl

---

## 🚀 Key Modules & Capabilities

1. **PCM Reports Engine**
   - Supports 7 operational report types (*Internal Report*, *Device Type Report*, *External Report*, *Daily ECLRT*, *Weekly 905 & TTC*, *Weekly RTL*, *Weekly ECLRT PCM*).
   - Computes P1, P2, and P3 SLA adherence buckets and Mean Time Taken to Resolve (MTTR) grouped by device type and transit agency.
   - **1-Click Outlook Draft Generator (`.eml`)**: Compiles live HTML tables and metrics into a valid RFC-822 email file (`internalreport.eml`) with pre-configured headers (`From: adithya.rajeev@accenture.com`, editable `To`, blank `CC`) ready for analyst review.

2. **PRESTO OPS L2 Alerts Feed**
   - Real-time notification stream consolidating critical infrastructure events: SQL Job Failures, ServiceNow P1/P2 Incidents, Multiple Station Outages, and Device AppCrash Monitoring.
   - Includes live floating toast alert pop-ups for immediate operational awareness.

3. **System Health (SOTI)**
   - Telemetry module mirroring SOTI MobiControl System Health (`System Settings > System Health > System Overview`).
   - Tracks Deployment Server hostnames (`PRX1DPDM11`–`14.prestoprod.net`), connected device counts, and queue lengths with threshold alerts.

---

## 📂 Project Directory Structure

```text
L2_Automation/
│
├── app/
│   ├── __init__.py
│   ├── main.py              # FastAPI server & browser launcher
│   ├── config_loader.py     # JSON configuration parser
│   ├── excel_processor.py   # Data cleanup & template generation engine
│   ├── validators.py        # Data quality checks & exception handler
│   └── logger.py            # Execution logging system
│
├── config/
│   └── reports/
│       └── internal_report.json
│
├── ui/
│   └── index.html           # Obsidian Black, Royal Purple & White Command Center UI
│
├── input/                   # Staged raw ServiceNow / SLA Excel exports
├── templates/               # Unmodified master report templates (.xlsx)
├── output/                  # Generated reports and .eml drafts
├── logs/                    # Daily application logs
│
├── requirements.txt         # Project dependencies
└── README.md                # Standard Operating Procedure (SOP)