Build a production-ready ISP Billing & Customer Management ERP as a Google Apps Script Web App using Google Sheets as the sole database. No external hosting, no SQL database — everything runs on google.script.run + SpreadsheetApp.

## TECH STACK
- Backend: Google Apps Script (Code.gs, multiple .gs files split by module)
- Frontend: HTML + CSS + Vanilla JS (HtmlService templates, google.script.run calls)
- Database: Google Sheets (one sheet per entity)
- Auth: Email OTP-based login + self-service signup

## CORE MODULES TO BUILD

### 1. Authentication & Roles
- Login page with Email OTP verification (send OTP via MailApp/GmailApp)
- 4 roles: Super Admin, Admin, Accountant, Collector
- Roles sheet storing per-module permissions (view/add/edit/delete flags) checked on every server call
- Session handling via PropertiesService or a Sessions sheet with expiry

### 2. Customer 360 CRM
- Customers sheet: name, phone, email, address, area, package, status, join date
- Customer detail page with tabs: Profile, Invoices, Payments, Complaints, Equipment, Documents, Activity Timeline
- Full CRUD with search, filters, and pagination

### 3. Packages & Areas
- Packages sheet: name, speed, price, billing cycle
- Areas sheet: area name, assigned staff/collector
- Dropdowns in customer form pull live from these sheets

### 4. Billing Engine
- "Run Billing" button: select month + year, bulk-generate invoices for all active customers with active packages
- Invoice sheet: invoice number (auto-incrementing), customer ID, amount, due date, status (Unpaid/Paid/Overdue)
- Prevent duplicate billing for the same customer/month

### 5. Payment Recording
- Record Payment modal: auto-fetch invoice number on customer selection, capture method (cash/online/bank), collector name, date
- On save, flip invoice status Unpaid → Paid and log to Payments sheet
- Link payment to a Bank Account (from Bank Accounts sheet dropdown)

### 6. WhatsApp & Call Integration
- One-click WhatsApp button on each customer/lead row — opens wa.me link prefilled with a payment reminder or follow-up message
- One-click Call button — uses tel: link or a click-to-call API webhook
- Template messages stored in a Settings/Templates sheet for easy editing

### 7. Leads CRM
- Leads sheet: name, phone, source, medium, status (New/Contacted/Won/Lost)
- "Mark Won" button converts a lead into a Customer row automatically
- Lead source/medium tracked for marketing ROI reporting

### 8. Complaints / Ticketing System
- Complaints sheet: customer ID, issue type, description, status (Open/In Progress/Resolved)
- Auto-lock ticket editing once marked Resolved
- Assign complaints to staff with due-date tracking

### 9. Collections & Recovery Dashboard
- Collection Runs sheet: date, collector, opening balance, cash collected, online collected, expenses, bank deposit reconciliation
- Recovery Board: list of overdue invoices with days-overdue, filterable by area/collector
- CSV export button for recovery list and collection run summaries

### 10. Finance Module
- Expenses sheet with category-wise tracking
- Vendors sheet for expense payees
- Bank Accounts sheet, linked wherever a payment/expense needs a bank source
- Monthly Income vs Expense summary view

### 11. Dashboard & Reporting
- Home dashboard: total customers, monthly revenue, unpaid invoices count, overdue amount, new leads this month
- Charts (Chart.js or Google Charts) for revenue trend and collection performance
- Exportable reports (CSV/PDF) for invoices, payments, and recovery

### 12. Settings & Branding
- Editable app name, logo upload (store in Drive, link in Settings sheet)
- Multiple color theme presets + light/dark mode toggle
- Currency symbol configuration
- Activity Log sheet capturing every create/update/delete with user + timestamp

## BUILD SEQUENCE (feed as sequential prompts to avoid context overload)
1. Foundation: sheets setup() + CRUD scaffolding for Customers, Areas, Packages
2. Auth: OTP login + Roles/permissions system
3. Customer 360 detail page with all tabs
4. Billing Engine (Run Billing bulk generation)
5. Payment recording + Bank Accounts linkage
6. WhatsApp/Call quick actions
7. Leads CRM + Complaints ticketing
8. Finance: Expenses, Vendors, Income/Expense summary
9. Collections & Recovery dashboard with CSV export
10. Dashboard charts, Settings, theming, Activity Logs, final polish

## NON-NEGOTIABLE CONSTRAINTS
- Zero external hosting or paid database — Sheets is the only data store
- All server logic must go through google.script.run with success/failure handlers
- Every module must respect role-based permissions before executing CRUD actions
- UI must be responsive and usable on mobile browsers