# KRA Auto-Checker v2.0

Automated system for verifying KRA (Kenya Revenue Authority) eTIMS transaction submissions and monitoring station connectivity across multiple fuel stations.

## 🎯 What's New in v2.0

- ✅ **Simplified Sheet Updates** - Only Report and Logs tabs (no more 0/1 in original sheet)
- ✅ **Station Health Monitoring** - Real-time connectivity and system status
- ✅ **Emoji Status Indicators** - 🟢🟡🔴 for easy visual scanning
- ✅ **Deployment Automation** - Auto-generate 94 deployment folders from spreadsheet
- ✅ **Better Error Handling** - Graceful handling of interruptions
- ✅ **Separate Executables** - Modular design for reliability

## ✨ Features

### KRA Checker
- ✅ Runs automatically at 7 PM daily
- ✅ Picks random transaction from 4-7 PM
- ✅ Checks KRA submission status
- ✅ Retries overnight (midnight, 2 AM, 4 AM)
- ✅ Updates Report tab in Google Sheets
- ✅ Detailed logging to Logs tab

### Heartbeat Monitor
- ✅ Runs every 30 minutes
- ✅ Checks internet connectivity
- ✅ Monitors SQL Server status
- ✅ Tracks disk space
- ✅ Gets local IP address
- ✅ Updates Station Status tab

### Deployment Generator
- ✅ Reads station data from Excel
- ✅ Creates folders organized by officer
- ✅ Pre-configures config.json for each station
- ✅ Generates deployment summary

## 📊 Google Sheets Structure

Your spreadsheet will have **3 tabs** (auto-created):

1. **Report** - Daily KRA check summaries
   - 🟢 SUCCESS - Transaction submitted
   - 🔴 NOT SUBMITTED - Not verified
   - 🟡 ERROR - Needs manual check

2. **Station Status** - Real-time station health
   - 🟢 Online - All systems OK
   - 🟡 Partial - Some issues
   - 🔴 Offline - Unreachable

3. **Logs** - Detailed debugging information
   - 🟢 INFO - Normal operations
   - 🟡 WARNING - Minor issues
   - 🔴 ERROR - Problems detected

## 🚀 Quick Start

### Prerequisites

- Windows PC with SQL Server (per station)
- Python 3.8+ (for building executables)
- Google account with Sheets access
- Excel file with station data

### One-Time Setup (30 minutes)

1. **Create Google Service Account**
   - Follow Google Cloud Console guide
   - Download `credentials.json`
   - Share spreadsheet with service account

2. **Build Executables**
   ```bash
   pip install -r requirements.txt
   
   # Build KRA Checker
   pyinstaller --onefile --hidden-import=googleapiclient --hidden-import=google.auth --name kra_checker kra_auto_checker.py
   
   # Build Heartbeat Monitor
   pyinstaller --onefile --hidden-import=googleapiclient --hidden-import=google.auth --name heartbeat_monitor heartbeat_monitor.py
   ```

3. **Create Base Template Folder**
   ```
   Base_Template/
   ├── kra_checker.exe          (from dist/)
   ├── heartbeat_monitor.exe    (from dist/)
   ├── credentials.json         (from Google Cloud)
   └── install.bat              (from repository)
   ```

4. **Generate Deployments**
   ```bash
   python deployment_generator.py --spreadsheet stations.xlsx
   ```

5. **Update Generated Configs**
   - Edit SQL password in each `config.json`
   - Update spreadsheet ID in each `config.json`

### Per-Station Deployment (5 minutes)

1. Connect to station via AnyDesk
2. Copy deployment folder to desktop
3. Run `install.bat` as Administrator
4. Verify in Task Scheduler
5. Done!

## 📁 Project Structure

```
kra-auto-checker/
├── kra_auto_checker.py          # Main KRA checker script
├── heartbeat_monitor.py         # Station health monitor
├── deployment_generator.py      # Auto-generate deployments
├── install.bat                  # Installation script
├── requirements.txt             # Python dependencies
├── config_template.json         # Configuration template
├── docs/
│   ├── INSTALLATION.md          # Detailed setup guide
│   └── DEPLOYMENT.md            # Deployment walkthrough
└── README.md                    # This file
```

## ⚙️ Configuration

### config.json

```json
{
    "station_name": "Rubis Burnt Forest",
    "anydesk_code": "1330302483",
    "sql_server": ".\\SQLEXPRESS",
    "sql_database": "ETIMS",
    "sql_username": "sa",
    "sql_password": "your_password",
    "spreadsheet_id": "your_spreadsheet_id",
    "service_account_file": "credentials.json",
    "max_retries": 3,
    "retry_delay": 120,
    "timeout": 15,
    "retry_file": "retry_transaction.json"
}
```

## 🔧 Requirements

### Python Packages

```
pyodbc>=4.0.39
google-api-python-client>=2.108.0
google-auth>=2.25.0
requests>=2.31.0
pyinstaller>=6.3.0
openpyxl>=3.1.0
```

### SQL Server

- Table: `ETPumpSales`
- Columns: `TransDateTime`, `QRLink`
- Authentication: SQL Server authentication

### Spreadsheet Columns

- **Station** (or Station Name)
- **AnyDesk Code** (or Anydesk)
- **Officer Assigned** (or Officer)

## 📖 Documentation

- **[Installation Guide](docs/INSTALLATION.md)** - Complete setup instructions
- **[Deployment Guide](docs/DEPLOYMENT.md)** - How to deploy to all stations
- **[Troubleshooting](#troubleshooting)** - Common issues and fixes

## 🐛 Troubleshooting

### "This operation is not supported for this document"
**Cause:** Spreadsheet is Excel file, not Google Sheets  
**Fix:** File → Save as Google Sheets, update spreadsheet_id in config

### Database connection error
**Cause:** SQL Server not running or wrong credentials  
**Fix:** Check SQL Server status and password in config.json

### Google Sheets authentication error
**Cause:** Missing or invalid credentials.json  
**Fix:** Verify file exists and service account has Editor access

### No transactions found
**Cause:** No sales during check period  
**Fix:** Normal if no sales; check if eTIMS services are running

### Station not updating in Status tab
**Cause:** Heartbeat task not running  
**Fix:** Check Task Scheduler for "Station Heartbeat" task

## 📊 Monitoring

### Daily Operations (5 minutes)

**Morning routine:**
1. Open Google Spreadsheet
2. Check **Report** tab for yesterday's KRA results
3. Check **Station Status** tab for offline stations
4. Review **Logs** tab for any errors
5. Take action only if needed

### Health Metrics

Monitor in Google Sheets:
- KRA submission success rate per station
- Station uptime percentage
- Retry frequency patterns
- Common error types

## 🔄 Updates

To update scripts on all stations:

1. Build new executables
2. Replace files in Base_Template
3. Re-run deployment_generator.py
4. Distribute updated folders to officers
5. Officers replace files on their stations

## 💰 Cost

**100% FREE!**
- Google Sheets API: Free tier (60 req/min)
- 94 stations × 50 req/day = 4,700 req/day
- Average: 3.3 req/min (well under limit)

## 🤝 Contributing

Internal tool for fuel station operations.

### Development Setup

```bash
git clone https://github.com/yourorg/kra-auto-checker.git
cd kra-auto-checker
pip install -r requirements.txt
```

### Testing

```bash
# Test on single station first
python kra_auto_checker.py

# Test heartbeat
python heartbeat_monitor.py

# Test deployment generator
python deployment_generator.py --spreadsheet test_stations.xlsx
```

## 📝 License

Internal use only - eTIMS integration for fuel stations.

## 👥 Authors

Created for streamlining KRA transaction verification and station monitoring across 94+ Rubis fuel stations in Kenya.

## 🙏 Acknowledgments

- Kenya Revenue Authority (KRA) for eTIMS system
- Google Cloud Platform for Sheets API
- All station operators and IT consultants

---

## 📞 Support

For deployment assistance or issues:
- Check Google Sheets tabs (Report, Status, Logs)
- Review documentation in `/docs`
- Test manually with executables

---

**Version:** 2.0.0  
**Last Updated:** January 2026  
**Stations Deployed:** 0/94 (ready to deploy)

## 🎯 Quick Commands

```bash
# Generate all deployments
python deployment_generator.py

# Generate for specific officer
python deployment_generator.py --officer "John Doe"

# Test KRA checker
C:\KRA_Checker\kra_checker.exe

# Test heartbeat
C:\KRA_Checker\heartbeat_monitor.exe

# Build executables
pyinstaller --onefile --hidden-import=googleapiclient --hidden-import=google.auth --name kra_checker kra_auto_checker.py
pyinstaller --onefile --hidden-import=googleapiclient --hidden-import=google.auth --name heartbeat_monitor heartbeat_monitor.py
```

## 📈 Deployment Progress

Track your deployment:
- [ ] Google Service Account created
- [ ] Executables built
- [ ] Base template created
- [ ] Deployments generated
- [ ] SQL passwords updated
- [ ] Spreadsheet IDs updated
- [ ] Pilot station tested (1/94)
- [ ] Remaining stations deployed (0/93)
- [ ] All stations verified (0/94)

---

**🚀 Ready to automate 94 stations!**