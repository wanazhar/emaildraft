from flask import Flask, render_template, request, redirect, url_for
import win32com.client as win32
import pandas as pd

app = Flask(__name__)

# Load Excel Data
def load_client_data():
    df = pd.read_excel("clients.xlsx", sheet_name="Data")
    return df

# Function to Generate Draft Emails in Outlook
def create_outlook_drafts(subject, body, clients):
    outlook = win32.Dispatch("Outlook.Application")
    for _, row in clients.iterrows():
        client_code = row["Client Code"]
        attachment = row["Attachment Path"]

        mail = outlook.CreateItem(0)
        mail.To = "recipient@example.com"  # Change dynamically if needed
        mail.Subject = f"{subject} - {client_code}"
        mail.Body = body

        # Attach file if exists
        if pd.notna(attachment):
            mail.Attachments.Add(attachment)

        mail.Save()

# Flask Routes
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        subject = request.form["subject"]
        body = request.form["body"]
        clients = load_client_data()
        
        create_outlook_drafts(subject, body, clients)
        return redirect(url_for("success"))
    
    return render_template("index.html")

@app.route("/success")
def success():
    return "Emails have been saved as drafts!"

if __name__ == "__main__":
    app.run(debug=True)
