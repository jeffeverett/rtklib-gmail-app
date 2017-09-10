# GMail_RTKLIB

# Copyright (C) 2017 Jeff Everett - All Rights Reserved
# You may use, distribute and modify this code under the
# terms of the BSD-2-Clause license


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
from log_utils import log_error, DataException

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

from my_constants import (SCOPES, CLIENT_SECRET_FILE, APPLICATION_NAME, MY_EMAIL,
    PROCESS_SUBJECT, ORIG_BIN_DIR, DEMO5_BIN_DIR, DEBUGGING, ORIG_CONFIG_FILE,DEMO5_CONFIG_FILE)

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
                    raise DataException('Cannot extract zip files with ".." in the path.')
            
            zip_arch.extractall(dirname)
            zip_arch.close()

def file_is_rover(file):
    # call the file the rover if it contains 'rov' or the base if it contains 'base'
    # otherwise, look for an 'r' or 'b' in the filename and classify accordingly
    file,_ = os.path.splitext(os.path.basename(file))
    if file.find('rov') != -1:
        return True
    elif file.find('base') != -1:
        return False
    elif file.find('r') != -1:
        return True
    elif file.find('b') != -1:
        return False
    else:
        raise DataException('Cannot determine whether to classify %s as a rover or base.' % (file,))

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

def get_obs_file(dirname, is_rover):
    # first get observation files
    obs_exts = ['.obs']
    obs_re = '^\.\d+o$'

    for filename in os.listdir(dirname):
        filename = os.path.join(dirname, filename)
        file,ext = os.path.splitext(filename.lower())
        if ext in obs_exts or re.match(obs_re, ext):
            if file_is_rover(filename) == is_rover:
                return filename

def get_nav_files(dirname, rover_obs):
    # take files with the same name as rover obs if they exist,
    # otherwise take whatever nav files can be found
    base_rover_name = os.path.splitext(os.path.basename(os.path.normpath(rover_obs)))[0]
    nav_re_strict = re.compile('^%s\..*nav$' % (base_rover_name,))
    # nav_re_lenient = re.compile('^.*\..*nav$')
    nav_files = list(filter(nav_re_strict.match, os.listdir(dirname)))
    if len(nav_files) == 0:
        # nav_files = list(filter(nav_re_lenient.match, os.listdir(dirname)))
        nav_exts=['.nav','.17g']
        nav_re = '^\.\d+n$'
        for filename in os.listdir(dirname):
            filename = os.path.join(dirname, filename)
            file,ext = os.path.splitext(filename.lower())
            if ext in nav_exts or re.match(nav_re, ext):
                nav_files.append(filename)
    else:
        for i in range(len(nav_files)):
            nav_files[i] = os.path.join(dirname, nav_files[i])
    
    return nav_files


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
    rc = subprocess.call([os.path.join(exe_dir, 'convbin.exe'), '-od', '-os', '-oi', '-ot', '-ro', '-TRK_MEAS=2', '-v', '3.03', '-d', target_dir, binfile])
    if rc != 0:
        raise DataException('Error encountered while running convbin.exe.')

def run_rnx2rtkp(exe_dir, config, sln_file, rover_obs, base_obs, nav_files):
    rc = subprocess.call([os.path.join(exe_dir, 'rnx2rtkp.exe'), '-k', config, '-o', 
        sln_file, rover_obs, base_obs, ' '.join(nav_files)])
    if rc != 0:
        raise DataException('Error encountered while running rnx2rtkp.exe.')

def rtkplot_save_image(sln_file, plot_file):
    # note that the command-line arguments to rtkplot require absolute paths
    exe_file = os.path.join(DEMO5_BIN_DIR, 'rtkplot.exe')
    sln_file = os.path.abspath(sln_file)
    plot_file = os.path.abspath(plot_file)
    rc = subprocess.call([exe_file, '-s', plot_file, sln_file])
    if rc != 0:
        raise DataException('Error encountered while running rtkplot.exe.')

def process_message(service, msg_id, body, sender, thread_id, subject, general_msg_id):
    # create directory in which to work (message id should be unique)
    dirname = os.path.join('runs', msg_id)

    try:
        os.makedirs(dirname)
    except os.error as error:
        # exception will be raised if directory already exists
        # this may be the case if we are debugging, in which
        # case the exception may be ignored
        if 0: #not DEBUGGING:
            raise

    # fetch attachments
    email_utils.GetAttachments(service, 'me', msg_id, dirname)

    # detect if file is zipped, and if so, unzip it
    unzip_all_in_dir(dirname)

    # first check if there are rover and base binary files
    rover_bin, base_bin = get_binary_files(dirname)

    # convert binary files to text files if necessary
    orig_dir = os.path.join(dirname, 'orig')
    demo5_dir = os.path.join(dirname, 'demo5')
    if rover_bin:
        run_convbin(ORIG_BIN_DIR, orig_dir, rover_bin)
        run_convbin(DEMO5_BIN_DIR, demo5_dir, rover_bin)
    if base_bin:
        run_convbin(ORIG_BIN_DIR, orig_dir, base_bin)
        run_convbin(DEMO5_BIN_DIR, demo5_dir, base_bin)

    # next get rover and base text files
    # the directories in which we look for them depend on whether we created them ourselves
    if rover_bin:
        orig_rover_obs = get_obs_file(orig_dir, True)
        demo5_rover_obs = get_obs_file(demo5_dir, True)
    else:
        orig_rover_obs = demo5_rover_obs = get_obs_file(dirname, True)
    if not orig_rover_obs or not demo5_rover_obs:
        raise DataException('Could not detect rover observation file (even after running convbin if necessary).')
        
    if base_bin:
        orig_base_obs = get_obs_file(orig_dir, False)
        demo5_base_obs = get_obs_file(demo5_dir, False)
    else:
        orig_base_obs = demo5_base_obs = get_obs_file(dirname, False)
    if not orig_base_obs or not demo5_base_obs:
        raise DataException('Could not detect base observation file in directory %s (even after running convbin if necessary).')
    
    # next find navigation files
    # again, directories in which to look are dependent on previously run commands
    if rover_bin:
        orig_nav_files = get_nav_files(orig_dir, orig_rover_obs)
        demo5_nav_files = get_nav_files(demo5_dir, demo5_rover_obs)
    else:
        orig_nav_files = demo5_nav_files = get_nav_files(dirname, orig_rover_obs)
    if len(orig_nav_files) == 0 or len(demo5_nav_files) == 0:
        raise DataException('Could not find any navigation files (even after running convbin if necessary).')

    # parse obs files to modfiy config file
    overwrites = {}
    #if rover_bin and base_bin:
    with open(demo5_rover_obs) as obs_file:
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
            overwrites['pos2-aroutcnt'] = round(20/median_delta)
            overwrites['pos2-arminfix'] = round(20/median_delta) 
    # then use presence or absence of 3rd column (=M8T) to choose cont. or f.-a.-h. AR
    present_in_rover = third_col_present(demo5_rover_obs)
    present_in_base = third_col_present(demo5_base_obs)
    if present_in_rover and present_in_base:
        overwrites['pos2-armode'] = 'continuous'
        overwrites['pos2-gloarmode'] = 'on'

    # parse email body to modify config file
    for line in body.splitlines():
        nv_tuple = parse_line(line)
        if nv_tuple:
            overwrites[nv_tuple[0]] = nv_tuple[1]

    orig_config = os.path.join(dirname, ORIG_CONFIG_FILE)
    with open(ORIG_CONFIG_FILE, 'r') as config_template, open(orig_config, 'w') as my_config:
        for line in config_template:
            nv_tuple = parse_line(line)
            if nv_tuple and nv_tuple[0] in overwrites:
                line = '%s=%s\n' % (nv_tuple[0], overwrites[nv_tuple[0]])
            my_config.write(line) 
                
    demo5_config = os.path.join(dirname, DEMO5_CONFIG_FILE)
    with open(DEMO5_CONFIG_FILE, 'r') as config_template, open(demo5_config, 'w') as my_config:
        for line in config_template:
            nv_tuple = parse_line(line)
            if nv_tuple and nv_tuple[0] in overwrites:
                line = '%s=%s\n' % (nv_tuple[0], overwrites[nv_tuple[0]])
            my_config.write(line) 


    # do processing on these files
    orig_sln = os.path.join(orig_dir, 'out_orig.pos')
    demo5_sln = os.path.join(demo5_dir, 'out_demo5.pos')
    run_rnx2rtkp(ORIG_BIN_DIR, orig_config, orig_sln, orig_rover_obs, orig_base_obs, orig_nav_files)
    run_rnx2rtkp(DEMO5_BIN_DIR, demo5_config, demo5_sln, demo5_rover_obs, demo5_base_obs, demo5_nav_files)

    # graph these files
    orig_plot = os.path.join(orig_dir, 'plot_orig.jpg')
    demo5_plot = os.path.join(demo5_dir, 'plot_demo5.jpg')
    rtkplot_save_image(orig_sln, orig_plot)
    rtkplot_save_image(demo5_sln, demo5_plot)

    # also graph the obs files located in the extended directory
    obs_rover_plot = os.path.join(dirname, 'plot_obs_rover.jpg')
    obs_base_plot = os.path.join(dirname, 'plot_obs_base.jpg')
    rtkplot_save_image(demo5_rover_obs, obs_rover_plot)
    rtkplot_save_image(demo5_base_obs, obs_base_plot)

    # send reply message
    
    html = """
    <html>
      <head>
        <meta http-equiv="content-type" content="text/html;
          charset=windows-1252">
        <title>result</title>
      </head>
      <body>
        <div style="text-align:center;">
          <div style="margin-bottom:50px">
            <p align="left"><b>RTKLIB Demonstration Results</b>:<br>
            </p>
            <div align="left">Before looking at the solution, it's always a good
              idea to take a quick look at the base and rover observations.&nbsp; More often than not, 
              the reason for a poor solution can be fairly obvious in the observation plots.&nbsp; Things to
              look for are:<br>
              <ul>
                <li>&nbsp;All expected satellite constellations are present
                  for both base and rover</li>
                <li>Observation times coincide between base and rover</li>
                <li>There are no large gaps in the observations</li>
                <li>All observation lines are yellow or green (gray means missing navigation data)</li>
                <li>The number of cycle slips, particularly for the rover,
                  are not excessive<br>
                </li>
              </ul>
            </div>
            <p><br>
              <br>
            </p>
            <table width="900" height="37" cellspacing="2" cellpadding="2"
              border="0">
              <tbody>
                <tr>
                  <td align="center">BaseObservations </td>
                  <td align="center">RoverObservations</td>
                </tr>
                <tr>
                  <td valign="top" align="center"><img src="cid:plot_obs_base.jpg"
                      alt="base obs" width="95%" vspace="0" hspace="0"
                      border="0" align="top"></td>
                  <td valign="top" align="center"><img src="cid:plot_obs_rover.jpg"
                      alt="rov obs" width="95%"></td>
                </tr>
              </tbody>
            </table>
            <p align="center"> </p>
            <br>
            <br>
            <div align="left">Here is the position solution computed with the demo5 B28b
              version of RTKLIB.&nbsp; Yellow represents a float solution and green is a fixed solution.&nbsp; The solution file is also attached so
              you will want to download it and open it with RTKPLOT to take
              a closer look.&nbsp; The configuration file used for this solution is also
              attached and you may want to download it as well to verify
              the solution was run as you intended. You can re-run the solution with a modified configuration 
              by re-submitting the raw data
              with the modified lines from the config file cut and pasted into the body of the email.<br>
              <br>
              If both data sets were recognized as coming from u-blox M8T receivers then the solution was run with continuous
              ambiguity resolution with GLONASS AR enabled.  Otherwise it was run with fix-and-hold ambiguity resolution with
              relatively low tracking gain and GLONASS AR also set to fix-and-hold.  You can confirm how it was run by loooking
              at the ambiguity resolution settings in the attached config file or the header in the solution file.<br>
            </div>
            <br>
            <div align="left">
              <div align="center"><img src="cid:plot_demo5.jpg" alt="demo5 sol"
                  width="80%"><br>
              </div>
              <br>
              <br>
              <br>
              Just for reference and comparison, here is the same solution computed using convbin and rnx2rtkp from the B28
              version of the 2.4.3 RTKLIB code with a few adjustments to the
              config file appropriate for this code.&nbsp; This solution
              file and the config file are also attached.  Note that the chances of false fixes in this solution will
              be higher than in the demo5 solution because this code does not have some of the additional
              features designed to reduce the chances of fix-and-hold locking to a false fix&nbsp; <br>
            </div>
          </div>
          <div style="margin-bottom:50px" align="center"> <img
              src="cid:plot_orig.jpg" alt="2.4.3 sol" width="80%"><br>
            <br>
            <div align="left"><br>
            </div>
          </div>
        </div>
      </body>
</html>"""

    
    
    attachments = [
        {'path': orig_plot, 'disposition': 'inline'},
        {'path': demo5_plot, 'disposition': 'inline'},
        {'path': obs_rover_plot, 'disposition': 'inline'},
        {'path': obs_base_plot, 'disposition': 'inline'},
        {'path': orig_sln, 'disposition': 'attachment'},
        {'path': demo5_sln, 'disposition': 'attachment'},
        {'path': orig_config, 'disposition': 'attachment'},
        {'path': demo5_config, 'disposition': 'attachment'}
    ]
    # because the reply will always be following an original message, "References" and "In-Reply-To" should be the same
    print('Generate reply:')
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('gmail', 'v1', http=http)
    
    message = email_utils.CreateMessageWithAttachments(MY_EMAIL, sender, "Re:"+subject, html, True,
        attachments, thread_id, general_msg_id, general_msg_id)
    print('Send Reply:')
    email_utils.SendMessage(service, 'me', message)

            
def process_messages(service):
    """Continuously loop, reading unread messages and processing
    them if necessary.
    """
    while (True):
        messages = service.users().messages().list(userId='me', maxResults=10000, q='is:unread').execute()

        if not messages or messages['resultSizeEstimate'] == 0:
            print('No messages to process. Sleeping for 10 seconds.')
            sleep(10)
        else:
            num_processed = 0

            print('%d unread messages...' % (len(messages['messages'],)))
            for message in messages['messages']:
                try:
                    contents = service.users().messages().get(userId='me', id=message['id']).execute()

                    # only process emails with subject="Process Request"
                    should_process_message = False
                    subject = ''
                    for header in contents['payload']['headers']:
                        if header['name'] == 'Subject':
                            #print('subject= %s' % header['value'])
                            if header['value'].lower().find(PROCESS_SUBJECT) != -1:
                                should_process_message = True
                                # record exact subject (considering caps) for reply
                                subject = header['value']
                                break

                    # do actual processing if necessary
                    if should_process_message:
                        # determine message body
                        body = email_utils.GetMessageBody(contents)
                        #if not body:
                        #    raise Exception('Error reading body of email.')

                        # determine sender and universal id
                        sender = None
                        general_msg_id = None
                        for header in contents['payload']['headers']:
                            if header['name'] == 'From':
                                sender = header['value']
                            if header['name'] == 'Message-ID':
                                general_msg_id = header['value']
                        if sender and general_msg_id:
                            print('Processing message %s...' % (message['id'],))
                            process_message(service, message['id'], body, sender, message['threadId'], subject, general_msg_id)
                        else:
                            if not sender:
                                raise DataException('Could not determine sender.')
                            else:
                                raise DataException('Could not determine message ID.')
                        num_processed += 1
                except Exception as e:
                    reply_email_successful = True
                    try:
                        reply_text = 'RTKLIB was unable to process the data.  Please check that you followed all of the guidelines for submitting data.  At this point the process is still immature so it is quite possible the problem is on this end. '
                        if e is DataException:
                            reply_text += '\nNote: the specific error that triggered this response is "%s".' % (str(e),)
                        reply_message = email_utils.CreateMessageWithAttachments(MY_EMAIL, sender, subject, reply_text, False,
                            None, message['threadId'], general_msg_id, general_msg_id)
                        email_utils.SendMessage(service, 'me', reply_message)
                    except Exception as ee:
                        print('Failed to send reply to user. Error: %s.' % (ee,))
                        reply_email_successful = False
                    
                    text = 'Error while processing message %s:' % (message['id'],)
                    if not reply_email_successful:
                        text += '\nNote: reply email not successfully sent to data sender.'
                    log_error(e, text, service)
                finally:
                    # mark message as read
                    if not DEBUGGING:
                        service.users().messages().modify(userId='me', id=message['id'],
                            body={'removeLabelIds': ['UNREAD'], 'addLabelIds': []}).execute()


            print('Processed %d messages in this run. Sleeping for 10 seconds' % (num_processed,))
            sleep(10)

def authorize_and_process():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('gmail', 'v1', http=http)
    process_messages(service)

def run_continuously():
    try:
        authorize_and_process()
    except Exception as e:
        log_error(e, 'Error in authorization or message listing:')
        print('Sleeping for 10 seconds.')
        sleep(10)
        run_continuously()
        

if __name__ == '__main__':
    run_continuously()