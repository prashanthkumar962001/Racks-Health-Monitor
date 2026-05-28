# Racks-Health-Monitor
Built a Python-based tool to monitor network rack status, check device connectivity, and automatically generate daily Excel reports for easier network monitoring.
-----------------------------------------------------------------------------------------------------------------------------
#  Rack Health Monitor
---> A Python automation tool that eliminates repetitive daily network rack monitoring by automatically checking device connectivity and generating color-coded Excel reports.

---
## Overview

**Rack Health Monitor** was built to solve a real operational problem: manually checking rack status and preparing reports every day was time-consuming and error-prone. This tool automates the entire process — from logging into the monitoring portal to delivering a ready-to-share Excel report — saving time and improving accuracy in telecom network operations.

---

## - How It Works

When the script is executed, it performs the following steps automatically:

1. **Opens** the monitoring portal using Selenium
2. **Logs in** using configured credentials
3. **Extracts** rack details for each device:
   - Rack Name & Rack ID
   - Device IP Address
   - Status & Location
4. **Pings** each device IP to verify reachability
5. **Classifies** each device as:
   -  **Online** — device is reachable
   -  **Offline** — device is unreachable
   -  **Unknown** — status could not be determined
6. **Generates** a formatted Excel report with color-coded rows:
   -  Green → Online
   -  Red → Offline
   -  Yellow → Unknown
7. **Saves** the report automatically with the current date in the filename

---
##  Features

- Automated login and rack data collection via Selenium
- Real-time ICMP ping-based device reachability checks
- Automated Excel report generation with OpenPyXL
- Color-coded status rows for at-a-glance monitoring
- Daily report saved with date-stamped filename
- Reduces manual effort across entire monitoring workflow

---
##  Technologies Used

| Technology | Purpose |
|---|---|
| Python | Core scripting language |
| Selenium | Browser automation & portal login |
| BeautifulSoup | HTML parsing & data extraction |
| OpenPyXL | Excel report generation |
| Subprocess | ICMP ping execution |
| WebDriver Manager | Automated ChromeDriver management |

---

##  Getting Started

### Prerequisites

```bash
**pip install selenium beautifulsoup4 openpyxl webdriver-manager**
```
### Configuration

Update the credentials and portal URL in the config section of the script before running:

```python
PORTAL_URL = "https://monitoring-portal.com"
USERNAME = "your_username"
PASSWORD = "your_password"
```
### Run

```bash
python rack_health_monitor.py
```
The script will open a browser, collect data, run ping checks, and save the Excel report in the current directory.

---
##  Project Impact

-  **Reduced** daily manual monitoring effort significantly
-  **Improved** operational efficiency in network health tracking
-  **Increased** reporting accuracy by eliminating human error
-  **Simplified** report sharing with auto-dated Excel files

---
##  Use Cases

This tool is useful for:

- Telecom network monitoring teams
- Network Operations Centers (NOC)
- Infrastructure and data center health monitoring
- Any team running repetitive daily device status checks

---
## Author
Prashanth Kumar Sake
Network Engineer transitioning to Network Automation
