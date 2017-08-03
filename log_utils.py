import email_utils
from my_constants import MY_EMAIL
import time
import sys

def log_error(msg, service=None):
    time_str = time.strftime('%Y-%m-%d %H:%M')
    msg = '[Time: %s, MessageID: %s] %s\n' % (time_str, msg)

    # print error on stderr
    print(msg, file=sys.stderr)

    # put message in actual error log
    with open('error_log.txt', 'a') as error_log:
        error_log.write(msg)

    # write email to make error more visible
    # but only do this if service is available
    if service:
        subject = 'Error encountered in GMail App'
        message = email_utils.CreateMessage(MY_EMAIL, MY_EMAIL, subject, msg)
        #email_utils.SendMessage(service, 'me', message)