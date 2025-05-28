import json
import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from jira import JIRA

# Jira Configuration
JIRA_URL = "https://warthogs.atlassian.net"
JIRA_USER = "anton.vishnevskii@canonical.com"
JIRA_API_TOKEN = os.getenv("JIRA_API_TOKEN")
JIRA_JQL = ("\"Properties[Checkboxes]\" = \"Roadmap Item\" "
            "AND issuetype = Epic "
            "AND \"Roadmap State[Dropdown]\" IN (\"🟧 At Risk\",\"🟥 Excluded\") "
            "AND labels IN (25.10) "
            "AND updated >= -7d")
JIRA_API_ENDPOINT = f"{JIRA_URL}/rest/api/latest/search"

# Email Configuration
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_SENDER = "your-email@example.com"
EMAIL_PASSWORD = "your-email-password"
EMAIL_RECIPIENT = "recipient@example.com"

# Storage File
PREV_STATE_FILE = "previous_results.json"
ADD_STATE_FILE = "added_results.json"
REM_STATE_FILE = "removed_results.json"


def fetch_jira_issues():
    try:
        jira = JIRA(server=JIRA_URL, basic_auth=(JIRA_USER, JIRA_API_TOKEN))
        # customfield_10968 - id of the Roadmap State field. Change it for new env
        issues = jira.search_issues(JIRA_JQL, maxResults=100, fields=["key", "summary", "customfield_10968"])
        return {issue.key: "STATE: " + issue.fields.customfield_10968.value + " SUMMARY: " + issue.fields.summary for issue in issues}
    except Exception as e:
        print(f"Error fetching Jira issues: {e}")
        return {}


def load_previous_results():
    try:
        with open(PREV_STATE_FILE, "r") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"Error loading previous results: {e}")
        return {}


def save_results(results, filename):
    try:
        with open(filename, "w") as f:
            json.dump(results, f, indent=4)
    except Exception as e:
        print(f"Error saving current results: {e}")


def send_email(added_issues, removed_issues):
    try:
        subject = "Jira Issues State Change Alert"
        body = "The following issues have changed state:\n\n"

        if added_issues:
            body += "Added Issues:\n" + "\n".join(
                [f"{key}: {summary}" for key, summary in added_issues.items()]) + "\n\n"
        if removed_issues:
            body += "Removed Issues:\n" + "\n".join(
                [f"{key}: {summary}" for key, summary in removed_issues.items()]) + "\n\n"

        msg = MIMEMultipart()
        msg["From"] = EMAIL_SENDER
        msg["To"] = EMAIL_RECIPIENT
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain"))

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.sendmail(EMAIL_SENDER, EMAIL_RECIPIENT, msg.as_string())
    except Exception as e:
        print(f"Error sending email: {e}")


def main():
    current_results = fetch_jira_issues()
    previous_results = load_previous_results()

    added_issues = {k: v for k, v in current_results.items() if k not in previous_results}
    removed_issues = {k: v for k, v in previous_results.items() if k not in current_results}

    if added_issues or removed_issues:
        save_results(current_results, PREV_STATE_FILE)
        save_results(added_issues, ADD_STATE_FILE)
        save_results(removed_issues, REM_STATE_FILE)
        # send_email(added_issues, removed_issues)


if __name__ == "__main__":
    main()
