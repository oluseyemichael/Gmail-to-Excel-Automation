# 📩 Gmail to Excel Automation  

A Python script that reads unread emails from Gmail, extracts key details (date, sender, subject, body), and saves them into an Excel file.  

## 🚀 Features  
- Fetches **unread emails** from Gmail using the Gmail API  
- Extracts **Date, Sender, Subject, and a Preview** of the email body  
- Handles **both plain text and HTML emails**  
- Saves the extracted data to **emails_data.xlsx**  

## 🛠 Technologies Used  
- **Python**  
- **Google API Client** (for Gmail access)  
- **Pandas** (for Excel file handling)  
- **BeautifulSoup** (for parsing HTML emails)  
- **dotenv** (for secure credential storage)  

## 🔧 Setup Instructions  

1. **Clone the Repository**  
   ```bash
   git clone <your-repo-link>
   cd gmail-to-excel
   ```

2. **Create & Activate a Virtual Environment**  
   ```bash
   python -m venv venv
   source venv/bin/activate  # Mac/Linux
   venv\Scripts\activate  # Windows
   ```

3. **Install Dependencies**  
   ```bash
   pip install -r requirements.txt
   ```

4. **Enable Gmail API & Set Up Credentials**  
   - Go to [Google Cloud Console](https://console.cloud.google.com/)  
   - Enable **Gmail API**  
   - Create OAuth Credentials (**Desktop App**) and download `credentials.json`  
   - Move `credentials.json` to the project folder  

5. **Run the Script**  
   ```bash
   python script.py
   ```

## 📊 Output  
The script will generate `unread_mails.xlsx`, containing:  
| Date | Sender | Subject | Email Preview |  
|------|--------|---------|--------------|  
| 2025-01-31 14:42:03 | example@email.com | Meeting Reminder | Don't forget our meeting tomorrow... |  
