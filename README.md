Scripts Repository
A collection of practical automation, administration, and optimization scripts for various platforms and workflows.

📂 Repository Contents
This repository is organized into dedicated directories for different scripting purposes:

📊 googlesheet/
Scripts for Google Sheets automation and integration:

API interaction and data manipulation

Automated reporting and data synchronization

Spreadsheet management utilities

🖥️ powershell/
Windows PowerShell scripts for system administration:

Active Directory management

Windows Server automation

User and permission management

Network and security configuration

Bulk operations and reporting

🐍 python/win11_optimizer/
Python-based tools for Windows 11 optimization:

System performance tuning scripts

Privacy and security configuration

Feature management and customization

Automated setup and deployment

Registry optimization utilities

🎯 Purpose
This repository serves as a centralized collection of useful scripts for:

IT Administrators needing automation tools

Developers looking for deployment and configuration scripts

Power Users wanting to optimize their Windows 11 experience

Data Professionals requiring Google Sheets automation

⚙️ Usage Instructions
Each script category contains specialized tools:

For PowerShell Scripts:
powershell
cd powershell
# Run scripts with appropriate execution policy
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\script-name.ps1
For Python Scripts:
bash
cd python/win11_optimizer
# Install requirements if needed
pip install -r requirements.txt
python script-name.py
For Google Sheets Scripts:
Set up Google Cloud credentials

Enable Google Sheets API

Configure OAuth 2.0 authentication

Run scripts with appropriate permissions

🔒 Security Notes
⚠️ Important Safety Information:

Review scripts before execution, especially those modifying system settings

Run PowerShell scripts with appropriate privilege levels

Google Sheets scripts require proper API key management

Backup important data before running optimization scripts

Test in non-production environments first

📁 File Structure
text
scripts/
├── googlesheet/          # Google Sheets automation
├── powershell/           # Windows PowerShell administration
└── python/
    └── win11_optimizer/  # Windows 11 optimization tools
🛠️ Requirements
PowerShell Scripts: Windows PowerShell 5.1+ or PowerShell Core

Python Scripts: Python 3.8+ with appropriate packages

Google Sheets Scripts: Google Cloud Platform account with Sheets API enabled

Permissions: Appropriate system/API permissions for script functionality

🤝 Contribution
While this is primarily a personal collection, improvements and fixes are welcome:

Ensure scripts are well-documented

Include safety checks and error handling

Test thoroughly before submitting

Maintain clear directory structure

📄 License
Unless otherwise specified in individual script headers, scripts are provided for educational and utility purposes. Use at your own discretion and review all code before execution.

Note: This repository has undergone recent cleanup and reorganization. Some documentation files have been removed as part of maintenance.
