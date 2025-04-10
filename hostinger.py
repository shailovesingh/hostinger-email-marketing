import smtplib
import csv
import time
import random
import threading
import email.utils
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Set TESTING_MODE = True for testing (short delays) and False for production (1 and 2 days)
TESTING_MODE = False
followup_delay = 10 if TESTING_MODE else 86400  # 86400 seconds = 1 day

# List of sender accounts with their credentials and SMTP details
SENDER_ACCOUNTS = [
    {
        "email": "neal@onpagegenius.com",
        "password": "Fdg@2025#",
        "smtp_server": "smtp.hostinger.com",
        "smtp_port": 465
    },
    {
        "email": "nealnitesh@onpagegenius.com",
        "password": "Fdg@2025#",
        "smtp_server": "smtp.hostinger.com",
        "smtp_port": 465
    },
    {
        "email": "kevin@onpagegenius.com",
        "password": "Fdg@2025#",
        "smtp_server": "smtp.hostinger.com",
        "smtp_port": 465
    },
    {
        "email": "nealm@onpagegenius.com",
        "password": "Fdg@2025#",
        "smtp_server": "smtp.hostinger.com",
        "smtp_port": 465
    },
    {
        "email": "sales@onpagegenius.com",
        "password": "Fdg@2025#",
        "smtp_server": "smtp.hostinger.com",
        "smtp_port": 465
    },
    {
        "email": "sales@filldesigngroup.store",
        "password": "Fdg@2025#",
        "smtp_server": "smtp.hostinger.com",
        "smtp_port": 465
    },
    {
        "email": "kevin@filldesigngroup.store",
        "password": "Fdg@2025#",
        "smtp_server": "smtp.hostinger.com",
        "smtp_port": 465
    },
    {
        "email": "nealnitesh@filldesigngroup.store",
        "password": "Fdg@2025#",
        "smtp_server": "smtp.hostinger.com",
        "smtp_port": 465
    },
    {
        "email": "nealm@filldesigngroup.store",
        "password": "Fdg@2025#",
        "smtp_server": "smtp.hostinger.com",
        "smtp_port": 465
    },
    {
        "email": "neal@filldesigngroup.store",
        "password": "Fdg@2025#",
        "smtp_server": "smtp.hostinger.com",
        "smtp_port": 465
    },
    {
        "email": "paul@filldesignprojects.website",
        "password": "Fdg@2025#",
        "smtp_server": "smtp.hostinger.com",
        "smtp_port": 465
    },
    {
        "email": "kevin@filldesignprojects.website",
        "password": "Fdg@2025#",
        "smtp_server": "smtp.hostinger.com",
        "smtp_port": 465
    },
    {
        "email": "nealnitesh@filldesignprojects.website",
        "password": "Fdg@2025#",
        "smtp_server": "smtp.hostinger.com",
        "smtp_port": 465
    },
    {
        "email": "neal@filldesignprojects.website",
        "password": "Fdg@2025#",
        "smtp_server": "smtp.hostinger.com",
        "smtp_port": 465
    },
    {
        "email": "neal.nitesh@filldesigns.com",
        "password": "Fdg@9874#",
        "smtp_server": "smtp.hostinger.com",
        "smtp_port": 465
    },
    {
        "email": "neal@filldesigns.com",
        "password": "Fdg@9874#",
        "smtp_server": "smtp.hostinger.com",
        "smtp_port": 465
    },
    {
        "email": "support@filldesigns.com",
        "password": "Fdg@9874#",
        "smtp_server": "smtp.hostinger.com",
        "smtp_port": 465
    },
    {
        "email": "nealn@filldesigngroup.in",
        "password": "Fdg@9874#",
        "smtp_server": "smtp.hostinger.com",
        "smtp_port": 465
    },
    {
        "email": "neal.m@filldesigngroup.in",
        "password": "Fdg@9874#",
        "smtp_server": "smtp.hostinger.com",
        "smtp_port": 465
    }

]

def open_csv_with_fallback(path, encodings=('utf-8', 'cp1252')):
    """
    Try opening the file with each encoding in order.
    Returns an open file object.
    """
    last_exc = None
    for enc in encodings:
        try:
            f = open(path, newline='', encoding=enc)
            # Read a bit to force decoding errors early
            _ = f.read(1024)
            f.seek(0)
            print(f"Opened CSV using encoding: {enc}")
            return f
        except UnicodeDecodeError as e:
            last_exc = e
            print(f"Failed to decode with {enc}, trying next…")
    # If all encodings fail, raise the last UnicodeDecodeError
    raise last_exc

def get_random_sender():
    """Randomly selects a sender account from the list."""
    return random.choice(SENDER_ACCOUNTS)

def spin_email_template(person_name, company, is_followup=False, followup_number=None):
    # ... (copy your spin_email_template function exactly) ...
    greetings = [
        f"Hi {person_name},",
        f"Hello {person_name},",
        f"Dear {person_name},"
    ]
    sentence1_options = [
        "I see you booked your new domain, marking an important step toward establishing a strong online presence.",
        "I noticed you secured your new domain—an essential move toward building a reliable online identity.",
        "I noticed you secured your domain. This marks the beginning of your online journey."
    ]
    sentence2_options = [
        "In the past six months, we’ve worked with several businesses to build websites, improve their search performance, and refine their social media presence. Consider how a well-designed digital platform can support your goals.",
        "Over the past six months, we’ve assisted a number of companies with website design, search optimization, and social media strategy. Think about how a customized digital solution could benefit your business.",
        "Recently, we’ve helped several businesses develop websites, enhance their search performance, and improve their social media efforts. Imagine a digital solution that aligns with your business needs."
    ]
    sentence3_options = [
        "I’m contacting you personally to share how our services may be of benefit. Please take a moment to watch the brief video I recorded, which explains our approach.",
        "I’m contacting you directly to share more about our services. I’ve prepared a brief video introduction outlining our process.",
        "I’m reaching out personally to share how our services may help. I’ve recorded a short video to introduce myself and explain our approach."
    ]
    
    greeting = random.choice(greetings)
    sentence1 = random.choice(sentence1_options)
    sentence2 = random.choice(sentence2_options)
    sentence3 = random.choice(sentence3_options)
    
    extra_line = f"\nThis is follow-up #{followup_number}. Just checking in regarding my previous email." \
                 if is_followup and followup_number else ""
    
    loom_link = "https://www.loom.com/share/1915f664b7f145f193d7b0fd6873ecb1"
    text_body = f"""{greeting}

{sentence1}

{sentence2}

{sentence3}

{extra_line}

{loom_link}

Looking forward to hearing from you.

Best regards,  
Neal  
https://filldesigngroup.in/
"""
    html_body = f"""\
<html>
  <body>
    <p>{greeting}</p>
    <p>{sentence1}</p>
    <p>{sentence2}</p>
    <p>{sentence3}</p>
    {f"<p>{extra_line}</p>" if extra_line else ""}
    <div>
      <a href="{loom_link}">
        <img style="max-width:300px;" src="https://cdn.loom.com/sessions/thumbnails/1915f664b7f145f193d7b0fd6873ecb1-12ee91ac978e3ba5-full-play.gif" alt="Watch Video">
      </a>
    </div>
    <p>
      Looking forward to hearing from you.<br>
      Best regards,<br>
      Neal<br>
      <a href="https://filldesigngroup.in/">Fill Design Group</a>
    </p>
  </body>
</html>
"""
    return text_body, html_body

def choose_subject(company):
    subject_templates = [
        "Question for {Company}",
        "See this for {Company}",
        "Quick Question for {Company}"
    ]
    return random.choice(subject_templates).format(Company=company)

def check_reply(recipient_email):
    # Placeholder for reply checking
    return False

def send_followup(recipient_email, original_msg_id, person_name, company, followup_number,
                  sender, original_subject):
    msg = MIMEMultipart('alternative')
    msg['From'] = sender['email']
    msg['To'] = recipient_email
    msg['Subject'] = "Re: " + original_subject
    msg['In-Reply-To'] = original_msg_id
    msg['References'] = original_msg_id
    
    text_body, html_body = spin_email_template(person_name, company, True, followup_number)
    msg.attach(MIMEText(text_body, 'plain'))
    msg.attach(MIMEText(html_body, 'html'))
    
    try:
        with smtplib.SMTP_SSL(sender['smtp_server'], sender['smtp_port'], timeout=10) as server:
            server.login(sender['email'], sender['password'])
            server.sendmail(sender['email'], recipient_email, msg.as_string())
            print(f"Follow-up #{followup_number} sent to {recipient_email} from {sender['email']}")
    except Exception as e:
        print(f"Error sending follow-up #{followup_number} to {recipient_email}: {e}")

def followup_scheduler(recipient_email, original_msg_id, person_name, company,
                       sender, original_subject):
    print(f"Waiting {followup_delay} seconds for follow-up #1 for {recipient_email}")
    time.sleep(followup_delay)
    if not check_reply(recipient_email):
        send_followup(recipient_email, original_msg_id, person_name, company, 1, sender, original_subject)
    else:
        print(f"Reply received from {recipient_email}. No first follow-up sent.")
        return

    print(f"Waiting {followup_delay} seconds for follow-up #2 for {recipient_email}")
    time.sleep(followup_delay)
    if not check_reply(recipient_email):
        send_followup(recipient_email, original_msg_id, person_name, company, 2, sender, original_subject)
    else:
        print(f"Reply received from {recipient_email}. No second follow-up sent.")

def send_initial_email(row):
    company = row['company']
    person_name = row['name']
    recipient_email = row['email']
    subject = choose_subject(company)
    text_body, html_body = spin_email_template(person_name, company)
    
    sender = get_random_sender()
    msg = MIMEMultipart('alternative')
    msg['From'] = sender['email']
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg_id = email.utils.make_msgid()
    msg['Message-ID'] = msg_id
    
    msg.attach(MIMEText(text_body, 'plain'))
    msg.attach(MIMEText(html_body, 'html'))
    
    try:
        with smtplib.SMTP_SSL(sender['smtp_server'], sender['smtp_port'], timeout=10) as server:
            server.login(sender['email'], sender['password'])
            server.sendmail(sender['email'], recipient_email, msg.as_string())
            print(f"Initial email sent to {recipient_email} from {sender['email']}")
    except Exception as e:
        print(f"Error sending initial email to {recipient_email}: {e}")
        return None, None, None
    return msg_id, subject, sender

def send_emails(csv_path):
    try:
        csvfile = open_csv_with_fallback(csv_path)
        with csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                company = row['company']
                person_name = row['name']
                recipient_email = row['email']
                print(f"Processing: Company: {company}, Name: {person_name}, Email: {recipient_email}")
                
                msg_id, subject, sender = send_initial_email(row)
                if not msg_id:
                    continue
                
                threading.Thread(
                    target=followup_scheduler,
                    args=(recipient_email, msg_id, person_name, company, sender, subject)
                ).start()
                
                time.sleep(30)  # delay between initial emails
                
    except FileNotFoundError:
        print("Error: CSV file not found. Please check the file path.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    csv_path = "test Email.xlsx"
    send_emails(csv_path)
