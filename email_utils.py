"""Send an email message from the user's account.
"""

import base64
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import mimetypes
import os

from apiclient import errors

from log_utils import log_error


def SendMessage(service, user_id, message):
    """Send an email message.

    Args:
      service: Authorized Gmail API service instance.
      user_id: User's email address. The special value "me"
      can be used to indicate the authenticated user.
      message: Message to be sent.

    Returns:
      Sent Message.
    """
    message['raw'] = message['raw'].decode()
    message = (service.users().messages().send(userId=user_id, body=message).execute())
    return message


def CreateMessage(sender, to, subject, message_text):
    """Create a message for an email.

    Args:
      sender: Email address of the sender.
      to: Email address of the receiver.
      subject: The subject of the email message.
      message_text: The text of the email message.

    Returns:
      An object containing a base64url encoded email object.
    """
    message = MIMEText(message_text)
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject
    return {'raw': base64.urlsafe_b64encode(message.as_bytes())}


def CreateMessageWithAttachments(sender, to, subject, message_text, is_html, attachments,
    thread_id=None, in_reply_to=None, references=None):
    """Create a message for an email.

    Args:
      sender: Email address of the sender.
      to: Email address of the receiver.
      subject: The subject of the email message.
      message_text: The text of the email message.
      attachments: List of 'attachment' dictionaries, which should
        have 'path' and 'disposition' keys

    Returns:
      An object containing a base64url encoded email object.
    """
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject

    if is_html:
        msg = MIMEText(message_text, 'html')
    else:
        msg = MIMEText(message_text)

    message.attach(msg)

    for attachment in attachments:
        filename = os.path.basename(attachment['path'])
        content_type, encoding = mimetypes.guess_type(attachment['path'])

        if content_type is None or encoding is not None:
            content_type = 'application/octet-stream'
        main_type, sub_type = content_type.split('/', 1)
        if main_type == 'text':
            fp = open(attachment['path'], 'rb')
            msg = MIMEText(fp.read(), _subtype=sub_type)
            fp.close()
        elif main_type == 'image':
            fp = open(attachment['path'], 'rb')
            msg = MIMEImage(fp.read(), _subtype=sub_type)
            fp.close()
        elif main_type == 'audio':
            fp = open(attachment['path'], 'rb')
            msg = MIMEAudio(fp.read(), _subtype=sub_type)
            fp.close()
        else:
            fp = open(attachment['path'], 'rb')
            msg = MIMEBase(main_type, sub_type)
            msg.set_payload(fp.read())
            fp.close()

        if attachment['disposition'] == 'inline':
            msg.add_header('Content-Id', '<%s>' % (filename,))
            msg.add_header('Content-Disposition', 'inline', filename=filename)
        else:
            msg.add_header('Content-Disposition', 'attachment', filename=filename)
        message.attach(msg)

    # if thread id is set, message is a reply
    if thread_id:
        message['references'] = references
        message['in-reply-to'] = in_reply_to
        return {'raw': base64.urlsafe_b64encode(message.as_bytes()), 'threadId': thread_id}

    return {'raw': base64.urlsafe_b64encode(message.as_bytes())}


def GetAttachments(service, user_id, msg_id, dirname):
    """Get and store attachment from Message with given id.

    Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    msg_id: ID of Message containing attachment.
    prefix: prefix which is added to the attachment filename on saving
    """
    message = service.users().messages().get(userId=user_id, id=msg_id).execute()

    for part in message['payload']['parts']:
        if part['filename']:
            if 'data' in part['body']:
                data=part['body']['data']
            else:
                att_id=part['body']['attachmentId']
                att=service.users().messages().attachments().get(userId=user_id, messageId=msg_id,id=att_id).execute()
                data=att['data']
            file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
            path = os.path.join(dirname, part['filename'])

            with open(path, 'wb') as f:
                f.write(file_data)
    
def GetMessageBody(contents):
    # assumes plaintext message body
    for part in contents['payload']['parts']:
        if part['mimeType'] == 'text/plain':
            body = part['body']['data']
            return base64.urlsafe_b64decode(body.encode('UTF-8')).decode('UTF-8')
        elif 'parts' in part:
            # go two levels if necessary
            for sub_part in part['parts']:
                if sub_part['mimeType'] == 'text/plain':
                    body = sub_part['body']['data']
                    return base64.urlsafe_b64decode(body.encode('UTF-8')).decode('UTF-8')