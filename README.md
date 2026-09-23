# 📊 QA Dispute Management System

A web-based application built with Google Apps Script (GAS) to manage Quality Assurance (QA) score disputes within a Customer Service / Operations team.

## ✨ Features
- **Custom Time Constraints:** Form automatically opens and closes based on team operational hours (Closes Tuesday 02:00 AM, Opens Friday 00:00 AM).
- **Google Sheets Integration:** Seamlessly reads and writes data to a Google Spreadsheet acting as a database.
- **Cross-Sheet Syncing:** Automatically updates dispute statuses by matching Ticket IDs across different data source sheets.
- **Real-Time Search & Filtering:** Users can track their dispute statuses filtered by `Email` or `QA Week` using Google Visualization API (GViz).
- **Modern UI/UX:** Responsive design powered by **TailwindCSS** and interactive alerts using **SweetAlert2**.

## 🛠️ Tech Stack
- **Backend:** Google Apps Script (JavaScript)
- **Frontend:** HTML5, CSS3 (Tailwind CSS via CDN), Vanilla JavaScript
- **Database:** Google Sheets
- **Libraries:** SweetAlert2

## 🚀 How it works
1. Employees log in with their Google Workspace email.
2. They submit a dispute form detailing why a QA deduction is incorrect, attaching relevant Ticket URLs.
3. The data is securely appended to a protected Google Sheet.
4. Supervisors and QA teams review the sheet. The script automatically synchronizes the final decision (Approved/Rejected) back to the user dashboard.

*(Note: Sensitive company data, URLs, and specific Spreadsheet IDs have been sanitized for this public repository.)*
