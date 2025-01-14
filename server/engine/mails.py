"""
The module provides a high-level interface for managing emails
through Exchange Web Services (EWS) for a specific account that
exists on an Exchange server. Most of the procedures depend on
the 'exchangelib' package, which must be installed before using
the module.
"""

import os
import re
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from os.path import basename, isfile, join, splitext
from smtplib import SMTP
from typing import Union

import exchangelib as xlib
from exchangelib import Account, Message

# type aliases
FilePath = str

# custom message classes
class SmtpMessage(MIMEMultipart):
	"""Wraps MIMEMultipart objects
	which are sent via an SMTP server.
	"""

# custom exceptions and warnings
class UndeliveredError(Exception):
	"""Raised on message delivery failure."""

class CredentialsParameterMissingError(Exception):
	"""An authorization parameter is required
	but not found in the source file.
	"""

class CredentialsNotFoundError(Exception):
	"""File with credentials for an account 
	is requested but doesn't exist.
	"""

def _validate_emails(addr: Union[str,list]) -> list:
	"""Checks if email addresses comply to the company's naming standards."""

	mails = []
	validated = []

	if isinstance(addr, str):
		mails = [addr]
	elif isinstance(addr, list):
		mails = addr
	else:
		raise TypeError(f"Argument 'addr' has invalid type: {type(addr)}")

	for mail in mails:

		stripped = mail.strip()
		validated.append(stripped)

		# check if email is Ledvance-specific
		if re.search(r"\w+\.\w+@ledvance.com", stripped) is None:
			raise ValueError(f"Invalid email address format: '{stripped}'!")

	return validated

def _attach_data(email: SmtpMessage, payload: bytes, name: str):
	"""Attaches data to a message."""

	# The content type "application/octet-stream" means
	# that a MIME attachment is a binary file
	part = MIMEBase("application", "octet-stream")
	part.set_payload(payload)
	encoders.encode_base64(part)

	# Add header
	part.add_header(
		"Content-Disposition",
		f"attachment; filename = {name}"
	)

	# Add attachment to the message
	# and convert it to a string
	email.attach(part)

	return email

def _attach_file(email: SmtpMessage, file: FilePath, name: str) -> SmtpMessage:
	"""Attaches file to a message."""

	if not isfile(file):
		raise FileNotFoundError(f"Attachment not found at the path specified: '{file}'")

	with open(file, "rb") as stream:
		payload = stream.read()

	# The content type "application/octet-stream" means
	# that a MIME attachment is a binary file
	part = MIMEBase("application", "octet-stream")
	part.set_payload(payload)
	encoders.encode_base64(part)

	# Add header
	part.add_header(
		"Content-Disposition",
		f"attachment; filename = {name}"
	)

	# Add attachment to the message
	# and convert it to a string
	email.attach(part)

	return email

def _get_credentials(acc_name: str) -> xlib.OAuth2Credentials:
	"""Models an authorization for an account."""

	cred_dir = join(os.environ["APPDATA"], "bia")
	cred_path = join(cred_dir, f"{acc_name.lower()}.token.email.dat")

	if not isfile(cred_path):
		raise CredentialsNotFoundError(
			"File with credentials for the specified account "
			f"'{acc_name}' not found at path: '{cred_path}'")

	with open(cred_path, encoding = "utf-8") as stream:
		lines = stream.readlines()

	identity = xlib.Identity(primary_smtp_address = acc_name)

	params = {
		"client_id": None,
		"client_secret": None,
		"tenant_id": None,
		"identity": identity
	}

	for line in lines:

		if ":" not in line:
			continue

		tokens = line.split(":")
		param_name = tokens[0].strip()
		param_value = tokens[1].strip()

		if param_name == "Client ID":
			key = "client_id"
		elif param_name == "Client Secret":
			key = "client_secret"
		elif param_name == "Tenant ID":
			key = "tenant_id"
		else:
			raise ValueError(f"Unrecognized parameter '{param_name}'!")

		params[key] = param_value

	# verify loaded parameters
	if params["client_id"] is None:
		raise CredentialsParameterMissingError(
			"Parameter 'client_id' not found in the source file!")

	if params["client_secret"] is None:
		raise CredentialsParameterMissingError(
			"Parameter 'client_secret' not foundin the source file!")

	if params["tenant_id"] is None:
		raise CredentialsParameterMissingError(
			"Parameter 'tenant_id' not foundin the source file!")

	# params OK, create credentials
	creds = xlib.OAuth2Credentials(
		params["client_id"],
		params["client_secret"],
		params["tenant_id"],
		params["identity"]
	)

	return creds

def _compile_email(subj, from_addr, recips, body) -> SmtpMessage:
	"""Compiles the email object."""

	email = SmtpMessage()
	email["Subject"] = subj
	email["From"] = from_addr
	email["To"] = ";".join(recips)
	email.attach(MIMEText(body, "html"))

	return email

def _compile_attachment_name(name: str, file: FilePath) -> str:
	"""Compiles attachment name from the file name specified
	by the user and the file name in the file path."""

	ext = splitext(file)[1]

	if name.lower().endswith(ext.lower()):
		filename = name
	else:
		filename = "".join([name, ext])

	return filename

def create_smtp_message(
		sender: str, recipient: Union[str, list],
		subject: str, body: str,
		attachment: Union[FilePath, list, dict] = None, # type: ignore
	) -> SmtpMessage:
	"""Creates an SMTP-compatible message.

	Parameters:
	-----------
	sender:
		Email address of the sender.

	recipient:
		Email address or addresses of the recipient.

	subject:
		Message subject.

	body:
		Message body in HTML format.

	attachment:

		- `None` (default)
		- a valid file path
		- a `list` of file paths, 
		- a `dict` of file names and file paths.
		- a `dict` of file names and `byte-like` objects.

		By default, the message will be created without any attachment.
		If a file path or a `list` of file paths is passed, then that file or files
		will be attached to the message. If a `dict` of file names and file paths
		is used, then the files will be attached to the message and the file
		names will be used as attachment names. Attachment type is inferred
		form the file file type. An invalid file path raises `FileNotFoundError` exception.

		If a `bytes-like` object is used, then its contents will be attached to the message,
		then the files will be attached to the message and the file names will be used as
		attachment names.

	attachment_name:
		Name of the attachment (default: "attachment").

		The parameter is ignored if the `att` argument is `bytes-like`.

	Returns:
	--------
	The constructed message.
	"""

	if not isinstance(recipient, str) and len(recipient) == 0:
		raise ValueError("No message recipients provided in 'recipient' argument!")

	recips = _validate_emails(recipient)
	email = _compile_email(subject, sender, recips, body)

	if attachment is None:
		return email

	if isinstance(attachment, dict):
		for key, val in attachment.items():
			if isinstance(val, FilePath):
				name = _compile_attachment_name(key, val)
				email = _attach_file(email, val, name)
			elif isinstance(val, bytes):
				email = _attach_data(email, val, key)
			else:
				raise TypeError(f"Unsupported attachment type: {type(attachment)}")
	elif isinstance(attachment, list):
		for att in attachment:
			if not isfile(att):
				raise FileNotFoundError(f"Attachment not found at the path specified: '{att}'")
			email = _attach_file(email, att, basename(att))
	elif isinstance(attachment, FilePath):
		email = _attach_file(email, attachment, basename(attachment))

	return email

def send_smtp_message(
		msg: SmtpMessage,
		host: str, port: int,
		timeout: int = 30,
		debug: int = 0
	) -> None:
	"""Sends an SMTP message.

	If the message is not  delivered to all the specified 
	recipients, then an `UndeliveredError` exception is raised.

	Parameters:
	-----------
	msg:
		Message to send.

	host:
		Name of the SMTP host server used for message sending.

	port:
		Number o the SMTP server port.

	timeout:
		Number of seconds to wait for the message to be sent (default: 30).
		Exceeding this limit will raise an `TimeoutError` exception.

	debug:
		Whether debug messages for connection and for all messages
		sent to and received from the server should be captured:
		- 0: "off" (default)
		- 1: "verbose"
		- 2: "timestamped"
	"""

	try:
		with SMTP(host, port, timeout = timeout) as smtp_conn:
			smtp_conn.set_debuglevel(debug)
			send_errs = smtp_conn.sendmail(msg["From"], msg["To"].split(";"), msg.as_string())
	except TimeoutError as exc:
		raise TimeoutError(
			"Attempt to connect to the SMTP servr timed out! Possible reasons: "
			"Slow internet connection or an incorrect port number used.") from exc

	if len(send_errs) != 0:
		failed_recips = ";".join(send_errs.keys())
		raise UndeliveredError(f"Message undelivered to: {failed_recips}")

def get_account(mailbox: str, name: str, x_server: str) -> Account:
	"""Models an MS Exchange server user account.

	Parameters:
	-----------
	mailbox:
		Name of the shared mailbox.

	name:
		Name of the account.

	x_server:
		Name of the MS Exchange server.

	Raises:
	-------
	`CredentialsNotFoundError`:
		When the file with the account credentials
		parameters is not found at the path specified.

	`CredentialsParameterMissingError`:
		When a credential parameter is not found in the
		content of the file where credentials are stored.

	Returns:
	--------
	The user account object.
	"""

	credentials = _get_credentials(name)
	build = xlib.Build(major_version = 15, minor_version = 20)

	cfg = xlib.Configuration(
		credentials,
		server = x_server,
		auth_type = xlib.OAUTH2,
		version = xlib.Version(build)
	)

	acc = Account(
		mailbox,
		config = cfg,
		access_type = xlib.IMPERSONATION
	)

	return acc

def get_messages(acc: Account, email_id: str) -> list:
	"""Fetches messages with a specific message ID.

	The message ID corresponds to the `Message.message_id` property value.

	Parameters:
	-----------
	acc:
		Account to access the inbox where the messages are stored.

	email_id:
		The ID string of the message to fetch.

	Returns:
	--------
	A list of `exchangelib:Message` objects
	that represent the messages found.
	"""

	# sanitize input
	if not email_id.startswith("<"):
		email_id = f"<{email_id}"

	if not email_id.endswith(">"):
		email_id = f"{email_id}>"

	# process
	emails = acc.inbox.walk().filter(message_id = email_id).only(
		"subject", "text_body", "headers", "sender",
		"attachments", "datetime_received", "message_id"
	)

	if emails.count() == 0:
		return []

	return list(emails)

def get_attachments(msg: Message, ext: str = ".*") -> list:
	"""Fetches message attachments and their names.

	Parameters:
	-----------
	msg:
		Message from which attachments are fetched.

	ext:
		File extension, that filters the attachment file types to fetch.

		By default, any file attachments are fetched. If an extension
		(e. g. ".pdf") is used, then only attachments with that file type
		are fetched.

	Returns:
	--------
	A `list` of `dict` objects, each containing attachment parameters: 
	- "name" (`str`): Name of the attachment.
	- "data" (`bytes`): Attachment binary data.
	"""

	atts = []

	for att in msg.attachments:
		if ext is not None and att.name.lower().endswith(ext):
			atts.append({"name": att.name, "content": att.content})

	return atts
