# Copyright (C) 2017 Jeff Everett - All Rights Reserved
# You may use, distribute and modify this code under the
# terms of the BSD-2-Clause license

import email_utils
from my_constants import MY_EMAIL
import time
import sys
import linecache

class DataException(Exception):
    pass

def format_exception():
    exc_type, exc_obj, tb = sys.exc_info()
    f = tb.tb_frame
    lineno = tb.tb_lineno
    filename = f.f_code.co_filename
    linecache.checkcache(filename)
    line = linecache.getline(filename, lineno, f.f_globals)
    return 'EXCEPTION IN ({}, LINE {} "{}"): {}'.format(filename, lineno, line.strip(), exc_obj)


def log_error(e, prefix, service=None):
    time_str = time.strftime('%Y-%m-%d %H:%M')
    msg = '[%s] %s %s' % (time_str, prefix, format_exception())

    # print error on stderr
    print(msg, file=sys.stderr)

    # put message in actual error log
    with open('error_log.txt', 'a') as error_log:
        error_log.write(msg)

    # write email to make error more visible
    # but only do this if service is available
    if service:
        subject = 'Error encountered in Gmail App'
        message = email_utils.CreateMessage(MY_EMAIL, MY_EMAIL, subject, msg)
        email_utils.SendMessage(service, 'me', message)