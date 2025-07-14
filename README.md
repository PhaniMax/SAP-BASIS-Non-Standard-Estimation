SAP BASIS Non-Standard Effort Estimation Tool:
This Python-based tool is designed to automate the effort estimation process for SAP BASIS non-standard tasks. It embeds an Excel template directly into the code using Base64 encoding, eliminating any external file dependencies. The tool calculates efforts based on complexity, user inputs, and exports a fully formatted Excel sheet.

📁 Files Included
Final_Sheet_25_June_2025.xlsx: Original Excel template used for estimation.

SAP BASIS NON-STANDARD ESTIMATION TOOL.py: Main script that:

Collects inputs
Calculates efforts and cost
Embeds the Excel template using base64
Outputs a dynamically filled Excel sheet

🔧 Features

No external file needed — Excel template is embedded in the script.
Complexity-based logic — handles Simple, Medium, Complex scenarios.
Rate card logic — dynamically selects resource role and cost rates.
Excel output — auto-filled, styled, and ready for delivery.
Offline ready — just run the Python file, no need for additional downloads.

🚀 How to Use

# Step 1: Make sure Python is installed
python sap_basis_estimator.py
Follow the prompts in the terminal to enter volumes and complexity.

The script generates a ready-to-send Excel file based on your inputs.

🧠 Behind the Scenes
The Excel file is embedded using Base64, decoded during runtime, and written back to disk.

openpyxl is used to modify the Excel contents dynamically.

Designed for offline use in secure environments where external templates can’t be shared.

👨‍💻 Developed by
Phanindra Mekala
SAP AO Automation Team
T-Systems | Internship Project
