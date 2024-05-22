import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def get_student_data(file_path):
    df = pd.read_excel(file_path, header=None)
    print("Loaded DataFrame:")
    print(df)
    return df


def compose_email(score):
    subject = "중간고사 시험 성적 안내"
    body = f"""

    """
    return subject, body


def send_email(to_email, subject, body):
    from_email = ""  # Replace with your email address
    password = ""  # Replace with your email password

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    try:
        server = smtplib.SMTP('smtp.naver.com', 587)
        server.starttls()
        server.login(from_email, password)
        server.sendmail(from_email, to_email, msg.as_string())
        server.quit()
        print(f"Email sent to {to_email}")
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")


def main():
    file_path = 'test.xlsx'

    df = get_student_data(file_path)

    for index, row in df.iterrows():
        email = row[1]
        score = row[2]
        subject, body = compose_email(score)
        send_email(email, subject, body)
        print(email, score)


main()