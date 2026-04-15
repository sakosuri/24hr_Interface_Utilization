# 24hr_Interface_Utilization
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Catalyst Center Interface Monitoring Automation
This Python-based automation tool interacts with Cisco Catalyst Center (formerly DNA Center) to audit network interface health and performance. It retrieves real-time operational status and historical performance trends (last 24 hours), exporting all data into a formatted Excel report.
🚀 Features
* Multi-Controller Support: Monitor devices across multiple Catalyst Center instances in a single run.
* Operational Audit: Captures Admin Status, Operational Status, Speed, and Duplex settings.
* Performance Analytics: Retrieves 24-hour Min/Max trends for:
    * Tx/Rx Rates (bps)
    * Tx/Rx Errors (%)
    * Tx/Rx Discards (%)
    * Tx/Rx Utilization (%)
* Automated Reporting: Generates a professional, color-coded Excel file with auto-adjusted column widths for easy reading.
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
📋 Prerequisites
1. Python 3.8+ installed.
2. Network Access: The machine running the script must have HTTPS (TCP 443) access to the Catalyst Center VIP.
3. API Credentials: A user account with at least OBSERVER or NETWORK-OPERATOR permissions on Catalyst Center.
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
⚙️ Configuration
The script uses a YAML file for configuration. Do not hardcode credentials in the script itself.
1. Create a file named config.yaml in the root directory.
2. Use the following structure:
dna_centers:
  - name: "Primary_DNAC"
    ip: "<catalyst-center-ip-address"
    username: "<catalyst-center-username>"
    password: "<catalyst-center-password>"

targets:
  - dna_center_name: "Primary_DNAC"
    devices:
      - device_ip: "<device-ip-address>"
        interfaces:
          - "GigabitEthernet1/0/1"
          - "GigabitEthernet1/0/2"
      - device_ip: "<device-ip-address>"
        interfaces:
          - "TenGigabitEthernet1/1/1"
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
🏃 Usage
Run the script from your terminal:
python interface_monitor.py
Script Workflow:
1. Authentication: Logs into each defined Catalyst Center to obtain an execution token.
2. Interface Discovery: Validates the existence of the specified interfaces and retrieves their unique IDs along with the operational status, speed and duplex settings.
3. Trend Collection: Queries the Trend Analytics API for the last 24 hours of performance data which retrieves Tx/Rx rates, Errors, Discards, Utilization.
4. Excel Generation: Saves a file named interface_report_YYYYMMDD_HHMMSS.xlsx.
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
📊 Output Example
The generated Excel report includes:
* Device & Interface Info: IP Address, Name, Speed, Duplex.
* Status: Admin vs. Operational state.
* Traffic Stats: Peak (Max) and Baseline (Min) throughput.
* Health Stats: Error and Discard percentages to identify faulty cables or congestion.

⚠️ Disclaimer
This script is provided as a sample and is not an officially supported Cisco product. Always test automation scripts in a lab environment before deploying them against production infrastructure.
