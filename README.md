# Google-Sheets-to-Contacts-CRM-Automated-WhatsApp-Sync
A lightweight, automated CRM built on Google Sheets and Google Apps Script. This tool instantly categorizes pasted customer data, securely syncs leads to Google Contacts in the background, and generates bulk CSV files optimized for WhatsApp broadcasts. Built with rate-limit protections and custom UI menus for seamless manual control.
# 📱 Google Sheets WhatsApp CRM & Contact Sync

A robust, automated Customer Relationship Management (CRM) tool built directly into Google Sheets using Google Apps Script. 

This script transforms a standard spreadsheet into a powerful engine that categorizes raw customer data, syncs it to Google Contacts using the People API, and exports bulk contact lists for WhatsApp broadcast campaigns.

## 🚀 Features

*   **⚡ Auto-Cleaning & Categorization:** Automatically processes data pasted into the `customer` sheet, removes duplicate phone numbers, and assigns categories based on customizable keyword rules.
*   **🔄 Background Contact Syncing:** Seamlessly syncs new contacts and creates Contact Groups in Google Contacts. Built-in time-limit handlers prevent the script from timing out during large syncs.
*   **📤 WhatsApp CSV Exporter:** Generates ready-to-use `.csv` files for WhatsApp broadcasts, automatically chunked (max 200 contacts per file) and saved directly to your Google Drive.
*   **🖥️ Custom UI Menu:** Includes a custom spreadsheet menu (`🟢 WhatsApp CRM`) with a sleek HTML popup for manual syncing and exporting.
*   **🛡️ Rate-Limit Protection:** Includes automatic retry logic, error handling, and sleep intervals to prevent Google API blocks.

## 🛠️ Prerequisites

1. A Google Account with Google Sheets.
2. The **Google People API** enabled in your Apps Script project.

## ⚙️ Setup & Installation

### 1. Prepare Your Google Sheet
Create a new Google Sheet and ensure it has the following three tabs exactly as named:
*   `customer` (Where you will paste raw data)
*   `Processed_data` (Where the script moves and formats data)
*   `Config` (Where you define your category keywords and rules)

### 2. Add the Script
1. In your Google Sheet, navigate to **Extensions > Apps Script**.
2. Delete any existing code in the editor and paste the provided CRM script.
3. Save the project (Ctrl+S / Cmd+S).

### 3. Enable the People API
1. In the Apps Script editor, look at the left sidebar and find **Services**.
2. Click the **+** button to add a service.
3. Scroll down, select **Google People API**, and click **Add**.

### 4. Initialize the Automation
1. Refresh your Google Sheet. You should see a new menu at the top called **🟢 WhatsApp CRM**.
2. Click **🟢 WhatsApp CRM > ⚙️ Initialize Automation & Background Sync**.
3. Google will prompt you to authorize the script. Click **Continue**, select your account, click **Advanced**, and then click **Go to [Project Name]**.
4. Grant the necessary permissions for the script to manage your contacts and drive files.

## 📖 How to Use

### Configuring Rules (`Config` sheet)
Set up your categorization rules in the `Config` sheet. The script reads this sheet to determine how to group customers based on keywords found in the "Product_Name" column.
*   *Format:* `Category Name` | `Initial` | `Keywords (separated by |)`
*   *Example:* `Premium Users` | `PRM` | `gold|platinum|vip`

### Processing Data
Simply paste your raw data into the `customer` sheet (starting from row 2). The `onEdit` trigger will instantly catch the data, clean the phone numbers, categorize them based on your `Config` rules, and move them to the `Processed_data` sheet.

### Syncing to Google Contacts
*   **Automatic:** The background trigger runs every 5 minutes to sync new "Pending" contacts.
*   **Manual:** Use the custom menu **📱 2. Manual Sync to Google Contacts** to select specific categories and force a sync with a real-time progress UI.

### Exporting for WhatsApp
Use the custom menu **📥 1. Export Selected to WhatsApp CSVs**. Select your desired categories, and the script will generate optimized CSV files in a folder named `WhatsApp_Exports` in your Google Drive.

## 🔒 Security & Privacy

*   **No Hardcoded Credentials:** This script relies entirely on Google's native OAuth flow. No API keys, passwords, or emails are hardcoded into the script.
*   **Data Ownership:** All data remains securely within your personal Google Sheets, Google Drive, and Google Contacts ecosystem. No third-party servers are used.

## 📝 License
This project is open-source and available for personal or commercial use. Feel free to modify the categorization logic and HTML UI to fit your specific logistics or business needs.
