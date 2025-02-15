# Bulk Email Generator

![Bulk Email Generator](https://img.shields.io/badge/Status-Working-brightgreen.svg) ![GitHub stars](https://img.shields.io/github/stars/your-repo.svg)

Automate bulk email creation with **Outlook** using **VBA** or **Python**! This tool allows users to generate email drafts dynamically based on a list of **RecipientCodes** and **attachments**.

---

## 📌 Features
✅ Generate bulk Outlook drafts dynamically  
✅ Append **RecipientCodes** to email **Subject**  
✅ Load **RecipientCodes** & **Attachments** from an Excel sheet  
✅ Works with both **VBA (Excel User Form)** & **Python (Flask App)**  
✅ Supports **PythonAnywhere Hosting**

---

## 🖥️ VBA Implementation

### 🔹 Setup
1. Open **Excel** and press `ALT + F11` to open **VBA Editor**.
2. Insert a new **User Form** and add:
   - `txtSubject` (Textbox for Email Subject)
   - `txtBody` (Textbox for Email Body, Multiline)
   - `lstClients` (ListBox for Client Codes & Attachments)
   - `btnLoadClients` (Button to Load Clients from Excel)
   - `btnGenerateEmails` (Button to Generate Outlook Drafts)
3. Copy and paste the **VBA Code** from [VBA Code File](emaildraft.vb).
4. Create an **Excel Sheet ("Data")** structured as:

| Recipient Codes | Attachment Path        |
|------------|----------------------|
| CODE001  | C:\path\to\file1.pdf |
| CODE002  | C:\path\to\file2.docx |

5. Run the macro **`OpenEmailForm`** to launch the form.

---

## 🐍 Python Implementation (Flask + Outlook)

### 🔹 Setup (Local)
1. Install dependencies:
   ```sh
   pip install flask pandas openpyxl pywin32
   ```
2. Create `clients.xlsx` file (same structure as above).
3. Copy and paste the **Python Code** from [app.py](app.py).
4. Run the app:
   ```sh
   python app.py
   ```
5. Open **http://127.0.0.1:5000/** in your browser.
6. Enter **Subject & Body**, then click **Generate Emails**.

### 🚀 Deploy on PythonAnywhere
PythonAnywhere doesn’t support Outlook directly. Use **Gmail API** or **SMTP**:
- ✅ [Gmail API Setup Guide](https://developers.google.com/gmail/api)
- ✅ [Flask-Mail SMTP Guide](https://pythonhosted.org/Flask-Mail/)

---

## 📜 License
This project is **open-source** under the [MIT License](LICENSE).

---

## 💙 Support & Contributions
Have suggestions or improvements? Feel free to **fork** & **contribute**! 🚀  
📩 Contact: [@wan_azhar on X](https://x.com/wan_azhar)

---

<p align="center">Made with ❤️ by <a href="https://github.com/wanazhar">wanazhar</a></p>
