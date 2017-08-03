
import httplib2
import os

import zipfile
import base64
from time import sleep
import subprocess
import re
import statistics

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

import email_utils
from log_utils import log_error

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

from my_constants import (SCOPES, CLIENT_SECRET_FILE, APPLICATION_NAME, MY_EMAIL,
    PROCESS_SUBJECT, ORIG_BIN_DIR, EXT_BIN_DIR, DEBUGGING, CONFIG_FILE)

def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    credential_path = 'my_gmail_credentials.json'

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def unzip_all_in_dir(dirname):
    for filename in os.listdir(dirname):
        filename = os.path.join(dirname, filename)
        _,ext = os.path.splitext(filename.lower())
        if ext == '.zip':
            zip_arch = zipfile.ZipFile(filename)
            for member_file in zip_arch.namelist():
                if member_file.find('..') != -1:
                    raise Exception('Cannot extract zip files with ".." in the path.')
            
            zip_arch.extractall(dirname)
            zip_arch.close()

def file_is_rover(file):
    # call the file the rover if it contains 'rover' or the base if it contains 'base'
    # otherwise, look for an 'r' or 'b' in the filename and classify accordingly
    if file.find('rover') != -1:
        return True
    elif file.find('base') != -1:
        return False
    elif file.find('r') != -1:
        return True
    elif file.find('b') != -1:
        return False
    else:
        raise Exception('Cannot determine whether to classify %s as a rover or base.' % (filename,))

def get_binary_files(dirname):
    binary_exts = ['.ubx']
    rover_bin = None
    base_bin = None
    for filename in os.listdir(dirname):
        filename = os.path.join(dirname, filename)
        file,ext = os.path.splitext(filename.lower())
        if ext in binary_exts:
            is_rover = file_is_rover(file)
            if is_rover:
                rover_bin = filename
            else:
                base_bin = filename
    return (rover_bin, base_bin)

def get_text_files(dirname):
    # first get observation files
    obs_exts = ['.obs']
    obs_re = '^\.\d+o$'
    rover_obs = None
    base_obs = None

    for filename in os.listdir(dirname):
        filename = os.path.join(dirname, filename)
        file,ext = os.path.splitext(filename.lower())
        if ext in obs_exts or re.match(obs_re, ext):
            is_rover = file_is_rover(file)
            if is_rover:
                rover_obs = filename
            else:
                base_obs = filename

    if not rover_obs:
        raise Exception('Could not detect rover observation file in directory %s (even after running convbin if necessary).'
            % (dirname,))
    if not base_obs:
        raise Exception('Could not detect base observation file in directory %s (even after running convbin if necessary).'
            % (dirname,))

    # next find corresponding navigation files
    # take files with the same name as rover obs if they exist,
    # otherwise take whatever nav files can be found
    base_rover_name = os.path.splitext(os.path.basename(os.path.normpath(rover_obs)))[0]
    nav_re_strict = re.compile('^%s\..*nav$' % (base_rover_name,))
    nav_re_lenient = re.compile('^.*\..*nav$')
    nav_files = list(filter(nav_re_strict.match, os.listdir(dirname)))
    if len(nav_files) == 0:
        nav_files = list(filter(nav_re_lenient.match, os.listdir(dirname)))
        if len(nav_files) == 0:
            raise Exception('Could not find any navigation files (even after running convbin if necessary).')
    for i in range(len(nav_files)):
        nav_files[i] = os.path.join(dirname, nav_files[i])
    
    return (rover_obs, base_obs, nav_files)


def third_col_present(file):
    # determine if third col of obs file has a number
    with open(file) as obs_file:
        # skip first 100 lines to avoid header
        for line in obs_file.readlines()[100:]:
            if line[0] == '>' or len(line.strip()) < 19:
                continue
            if line[18] == ' ':
                return False
            else:
                return True

def parse_line(line):
    pos_eq = line.find('=')
    pos_hash = line.find('#')

    if pos_eq == -1:
        return None
    if pos_hash == -1:
        pos_hash = len(line)
    if pos_eq > pos_hash:
        return None

    name = line[:pos_eq].strip()
    value = line[pos_eq+1:pos_hash].strip()

    return (name, value)

def run_convbin(exe_dir, target_dir, binfile):
    rc = subprocess.call([os.path.join(exe_dir, 'convbin.exe'), '-od', '-os', '-oi', '-ot', '-v', '3.03', '-d', target_dir, binfile])
    if rc != 0:
        raise Exception('Non-zero exit code encountered while running convbin.exe.')

def run_rnx2rtkp(exe_dir, config, sln_file, rover_obs, base_obs, nav_files):
    rc = subprocess.call([os.path.join(exe_dir, 'rnx2rtkp.exe'), '-x', '3', '-y', '3', '-k', config, '-o', 
        sln_file, rover_obs, base_obs, ' '.join(nav_files)])
    if rc != 0:
        raise Exception('Non-zero exit code encountered while running rnx2rtkp.exe.')

def rtkplot_save_image(sln_file, plot_file):
    # note that the command-line arguments to rtkplot require absolute paths
    exe_file = os.path.join(EXT_BIN_DIR, 'rtkplot.exe')
    sln_file = os.path.abspath(sln_file)
    plot_file = os.path.abspath(plot_file)
    rc = subprocess.call([exe_file, '-s', plot_file, sln_file])
    if rc != 0:
        raise Exception('Non-zero exit code encountered while running rtkplot.exe.')

def process_message(service, msg_id, body, sender, thread_id, subject, general_msg_id):
    # create directory in which to work (message id should be unique)
    dirname = os.path.join('runs', msg_id)

    try:
        os.makedirs(dirname)
    except os.error as error:
        # exception will be raised if directory already exists
        # this may be the case if we are debugging, in which
        # case the exception may be ignored
        if not DEBUGGING:
            raise

    # fetch attachments
    email_utils.GetAttachments(service, 'me', msg_id, dirname)

    # detect if file is zipped, and if so, unzip it
    unzip_all_in_dir(dirname)

    # first check if there are rover and base binary files
    rover_bin, base_bin = get_binary_files(dirname)

    # convert binary files to text files if necessary
    orig_dir = os.path.join(dirname, 'orig')
    ext_dir = os.path.join(dirname, 'ext')
    if rover_bin and base_bin:
        run_convbin(EXT_BIN_DIR, orig_dir, rover_bin)
        run_convbin(EXT_BIN_DIR, ext_dir, rover_bin)
        run_convbin(EXT_BIN_DIR, orig_dir, base_bin)
        run_convbin(EXT_BIN_DIR, ext_dir, base_bin)

    # next get rover and base text files
    # the directories in which we look for them depend on whether we created them ourselves
    if rover_bin and base_bin:
        orig_rover_obs, orig_base_obs, orig_nav_files = get_text_files(orig_dir)
        ext_rover_obs, ext_base_obs, ext_nav_files = get_text_files(ext_dir)
    else:
        orig_rover_obs, orig_base_obs, orig_nav_files = get_text_file('')
        ext_rover_obs, ext_base_obs, ext_nav_files = orig_rover_obs, orig_base_obs, orig_nav_files

    # parse obs files to modfiy config file
    overwrites = {}
    # only do this if we generated them ourselves
    if rover_bin and base_bin:
        with open(ext_rover_obs) as obs_file:
            # first compute median delta to modify aroutcnt and arminfix
            times = []
            num_skipped = 0
            num_to_skip = 50
            num_read = 0
            num_to_read = 11
            for line in obs_file:
                if line[0] == '>':
                    if num_skipped < num_to_skip:
                        num_skipped += 1
                    else:
                        # time always starts and ends at same spot
                        times.append(float(line[19:29]))
                        num_read += 1
                        if num_read == num_to_read:
                            break
            if len(times):
                deltas = [times[i] - times[i-1] for i in range(1, len(times))]
                median_delta = statistics.median(deltas)
                overwrites['pos2-aroutcnt'] = 20/median_delta
                overwrites['pos2-arminfix'] = 20/median_delta

        # then use presence or absence of 3rd column to choose cont. or f.-a.-h. AR
        present_in_rover = third_col_present(ext_rover_obs)
        present_in_base = third_col_present(ext_base_obs)
        if present_in_rover and present_in_base:
            overwrites['pos2-armode'] = 'continuous'
            overwrites['pos2-gloarmode'] = 'on'

    # parse email body to modify config file
    for line in body.splitlines():
        nv_tuple = parse_line(line)
        if nv_tuple:
            overwrites[nv_tuple[0]] = nv_tuple[1]

    # perform actual overwrites to obtain modified config file
    config = CONFIG_FILE
    if len(overwrites):
        config = os.path.join(dirname, CONFIG_FILE)
        with open(CONFIG_FILE, 'r') as config_template, open(config, 'w') as my_config:
            for line in config_template:
                nv_tuple = parse_line(line)
                if nv_tuple and nv_tuple[0] in overwrites:
                    line = '%s=%s\n' % (nv_tuple[0], overwrites[nv_tuple[0]])
                my_config.write(line) 


    # do processing on these files
    orig_sln = os.path.join(orig_dir, 'out_orig.pos')
    ext_sln = os.path.join(ext_dir, 'out_ext.pos')
    run_rnx2rtkp(ORIG_BIN_DIR, config, orig_sln, orig_rover_obs, orig_base_obs, orig_nav_files)
    run_rnx2rtkp(EXT_BIN_DIR, config, ext_sln, ext_rover_obs, ext_base_obs, ext_nav_files)

    # graph these files
    orig_plot = os.path.join(orig_dir, 'plot_orig.jpg')
    ext_plot = os.path.join(ext_dir, 'plot_ext.jpg')
    rtkplot_save_image(orig_sln, orig_plot)
    rtkplot_save_image(ext_sln, ext_plot)

    # also graph the obs files located in the extended directory
    obs_rover_plot = os.path.join(dirname, 'plot_obs_rover.jpg')
    obs_base_plot = os.path.join(dirname, 'plot_obs_base.jpg')
    rtkplot_save_image(ext_rover_obs, obs_rover_plot)
    rtkplot_save_image(ext_base_obs, obs_base_plot)

    # send reply message
    html = """
    <html>
      <body>
        <div style="text-align:center;">
          <div style="margin-bottom:50px">
            <p>Original binaries plot:</p>
            <img src="cid:plot_orig.jpg">
          </div>
          <div style="margin-bottom:50px">
            <p>Extended binaries plot:</p>
            <img src="cid:plot_ext.jpg">
          </div>
          <div style="display:inline-block; margin-right:10px;">
            <p>Rover obs file:</p>
            <img src="cid:plot_obs_rover.jpg">
          </div>
          <div style="display:inline-block; margin-left:10px;">
            <p>Base obs file:</p>
            <img src="cid:plot_obs_base.jpg">
          </div>
      </body>
    </html>"""
    attachments = [
        {'path': orig_plot, 'disposition': 'inline'},
        {'path': ext_plot, 'disposition': 'inline'},
        {'path': obs_rover_plot, 'disposition': 'inline'},
        {'path': obs_base_plot, 'disposition': 'inline'},
        {'path': orig_sln, 'disposition': 'attachment'},
        {'path': ext_sln, 'disposition': 'attachment'}
    ]
    # because the reply will always be following an original message, "References" and "In-Reply-To" should be the same
    message = email_utils.CreateMessageWithAttachments(MY_EMAIL, sender, subject, html, True,
        attachments, thread_id, general_msg_id, general_msg_id)
    email_utils.SendMessage(service, 'me', message)


def process_messages(service):
    """Continuously loop, reading unread messages and processing
    them if necessary.
    """
    while (True):
        messages = service.users().messages().list(userId='me', maxResults=10000, q='is:unread').execute()

        if not messages or messages['resultSizeEstimate'] == 0:
            print('No messages to process. Sleeping for 10 seconds')
            sleep(10)
        else:
            num_processed = 0

            for message in messages['messages']:
                try:
                    contents = service.users().messages().get(userId='me', id=message['id']).execute()

                    # only process emails with subject="Process Request"
                    should_process_message = False
                    subject = ''
                    for header in contents['payload']['headers']:
                        if header['name'] == 'Subject':
                            if header['value'].lower() == PROCESS_SUBJECT:
                                should_process_message = True
                                # record exact subject (considering caps) for reply
                                subject = header['value']
                                break

                    # do actual processing if necessary
                    if should_process_message:
                        # determine message body
                        body = email_utils.GetMessageBody(contents)
                        if not body:
                            raise Exception('Could not determine message body.')

                        # determine sender and universal id
                        sender = None
                        general_msg_id = None
                        for header in contents['payload']['headers']:
                            if header['name'] == 'From':
                                sender = header['value']
                            if header['name'] == 'Message-ID':
                                general_msg_id = header['value']
                        if sender and general_msg_id:
                            process_message(service, message['id'], body, sender, message['threadId'], subject, general_msg_id)
                        else:
                            raise Exception('Could not determine one of: sender or general message id.')
                        num_processed += 1
                except Exception as e:
                    log_error(e, 'Error while processing message %s:' % (message['id'],), service)
                finally:
                    # mark message as read
                    if not DEBUGGING:
                        service.users().messages().modify(userId='me', id=message['id'],
                            body={'removeLabelIds': ['UNREAD'], 'addLabelIds': []}).execute()


            print('Processed %d messages in this run. Sleeping for 10 seconds' % (num_processed,))
            sleep(10)

def main():
    try:
        credentials = get_credentials()
        http = credentials.authorize(httplib2.Http())
        service = discovery.build('gmail', 'v1', http=http)
        process_messages(service)
    except Exception as e:
        log_error(e, 'Error outside of specific message processing:')

if __name__ == '__main__':
    main()