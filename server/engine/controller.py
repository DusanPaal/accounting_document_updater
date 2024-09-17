# pylint: disable = C0103, E0401, E0611, W0703, W1203

"""The controller.py represents the middle layer of the application design \n
and mediates communication between the top layer (app.py) and the \n
highly specialized modules situated on the bottom layer of the design \n
(fb03.py, qm02.py, mails.py, report.py sap.py).
"""

import logging
import os
from datetime import datetime as dt
from datetime import timedelta
from glob import glob
from logging import Logger, config
from os.path import basename, isfile, join
from typing import Union

import pandas as pd
import yaml
from pandas import DataFrame
from win32com.client import CDispatch

from engine import fb03, mails, qm02, report, sap, utils, va03

ORDER_NUMBERING_PREFIX = "501"
SECTION_LINE_LENGTH = 35
MAX_RECOVERY_ATTEMPTS = 3

log = logging.getLogger("master")


# ====================================
# initialization of the logging system
# ====================================

def _compile_log_path(log_dir: str) -> str:
	"""Compiles the path to the log file
	by generating a log file name and then
	concatenating it to the specified log
	directory path."""

	date_tag = dt.now().strftime("%Y-%m-%d")
	nth = 0

	while True:
		nth += 1
		nth_file = str(nth).zfill(3)
		log_name = f"{date_tag}_{nth_file}.log"
		log_path = join(log_dir, log_name)

		if not isfile(log_path):
			break

	return log_path

def _read_log_config(cfg_path: str) -> dict:
	"""Reads logging configuration parameters from a yaml file."""

	# Load the logging configuration from an external file
	# and configure the logging using the loaded parameters.

	if not isfile(cfg_path):
		raise FileNotFoundError(
			f"The logging configuration file not found: '{cfg_path}'")

	with open(cfg_path, encoding = "utf-8") as stream:
		content = stream.read()

	return yaml.safe_load(content)

def _update_log_filehandler(log_path: str, logger: Logger) -> None:
	"""Changes the log path of a logger file handler."""

	prev_file_handler = logger.handlers.pop(1)
	new_file_handler = logging.FileHandler(log_path)
	new_file_handler.setFormatter(prev_file_handler.formatter)
	logger.addHandler(new_file_handler)

def _print_log_header(logger: Logger, header: list, terminate: str = "\n") -> None:
	"""Prints header to a log file."""

	for nth, line in enumerate(header, start = 1):
		if nth == len(header):
			line = f"{line}{terminate}"
		logger.info(line)

def _remove_old_logs(logger: Logger, log_dir: str, n_days: int) -> None:
	"""Removes old logs older than the specified number of days."""

	old_logs = glob(join(log_dir, "*.log"))
	n_days = max(1, n_days)
	curr_date = dt.now().date()

	for log_file in old_logs:
		log_name = basename(log_file)
		date_token = log_name.split("_")[0]
		log_date = dt.strptime(date_token, "%Y-%m-%d").date()
		thresh_date = curr_date - timedelta(days = n_days)

		if log_date < thresh_date:
			try:
				logger.info(f"Removing obsolete log file: '{log_file}' ...")
				os.remove(log_file)
			except PermissionError as exc:
				logger.error(str(exc))

def configure_logger(log_dir: str, cfg_path: str, *header: str) -> None:
	"""Configures application logging system.

	Parameters:
	-----------
	log_dir:
		Path t the directory to store the log file.

	cfg_path:
		Path to a yaml/yml file that contains
		application configuration parameters.

	header:
		A sequence of lines to print into the log header.
	"""

	log_path = _compile_log_path(log_dir)
	log_cfg = _read_log_config(cfg_path)
	config.dictConfig(log_cfg)
	logger = logging.getLogger("master")
	_update_log_filehandler(log_path, logger)
	if header is not None:
		_print_log_header(logger, list(header))
	_remove_old_logs(logger, log_dir, log_cfg.get("retain_logs_days", 1))

# ====================================
# 		application configuration
# ====================================

def load_app_config(cfg_path: str) -> dict:
	"""Reads application configuration
	parameters from a file.

	Parameters:
	-----------
	cfg_path:
		Path to a yaml/yml file that contains
		application configuration parameters.

	Returns:
	--------
	Application configuration parameters.
	"""

	log.info("Loading application configuration ...")

	if not cfg_path.endswith((".yaml", ".yml")):
		raise ValueError("The configuration file not a YAML/YML type!")

	with open(cfg_path, encoding = "utf-8") as stream:
		content = stream.read()

	cfg = yaml.safe_load(content)
	log.info("Configuration loaded.")

	return cfg


# ====================================
# 		Fetching of user input
# ====================================

def fetch_user_input(msg_cfg: dict, email_id: str) -> dict:
	"""Fetches the data provided by the user.

	Params:
	-------
	msg_cfg:
		Application 'messages' configuration parameters.

	email_id:
		The string ID of the message.

	Returns:
	--------
	Names of the processing parameters and their values:
		- "error_message: `str`, default: ""
			A detailed error message if the user input is invalid.
		- "email": `str`, default: ""
			Email address of the sender.
		- "data": `pandas.DataFrame`, default: None
		 	Converted attachment data.
	"""

	log.info("Retrieving user message ...")

	# init parameter with default values
	params = {
		"error_message": "",
		"email": "",
		"data": None,
	}

	user_req = msg_cfg["requests"]
	acc = mails.get_account(user_req["mailbox"], user_req["account"], user_req["server"])
	msg = mails.get_messages(acc, email_id)[0]

	if msg is None:
		raise RuntimeError(f"Could not find message with the specified ID: '{email_id}'")

	params.update({"email": msg.sender.email_address})
	log.info("User message retrieved.")

	log.info("Reading the contents of attached data ...")
	attachments = mails.get_attachments(msg, ".xlsm")

	if len(attachments) == 0:
		params.update({"error_message": "The user email contains no attachment!"})
		return params

	if len(attachments) > 1:
		params.update({"error_message": "The user email contains more than one attachment!"})
		return params

	log.info("Data reading completed.")

	# Converts user input data extracted from
	# the message attachment into a DataFrame.
	log.info("Converting attachment data ...")
	data = pd.read_excel(attachments[0]["content"], dtype = "string")

	# replace spaces in column names with underscore
	# to avoid ambiguity in space count.
	data.columns = data.columns.str.strip()
	data.columns = data.columns.str.replace(r"\s+", "_", regex = True)

	data["Company_Code"] = data["Company_Code"].str.zfill(4)
	data["Document_Number"] = pd.to_numeric(data["Document_Number"]).astype("UInt64")
	data["Notification"] = pd.to_numeric(data["Notification"]).astype("UInt64")
	data["Document_Year"] = pd.to_numeric(data["Document_Year"]).astype("UInt16")
	data["Case_ID"] = pd.to_numeric(data["Case_ID"]).astype("UInt64")

	log.info("Data conversion completed.")

	# Check if mandatory cols are correctly filled: valid rows
	# are those where all cells are filled or neither is filled.
	# Incomplete rows will be deleted without any warning. It's
	# down to the user to fill data correctly and this is also
	# clearly stated in the User guide.
	log.info("Validating data ...")
	cols_to_check = data[["Company_Code", "Document_Number", "Document_Year"]]
	invalid_recs = ~(cols_to_check.isna().all(axis = 1) | cols_to_check.notna().all(axis = 1))
	n_badrows = invalid_recs.sum()

	if n_badrows != 0:
		log.error(f"Found {n_badrows} invalid data rows.")
		badrows =  data[invalid_recs].to_string(
			header = True, index = False, justify  = "center")
		log.debug("".join(["Contents of the bad rows:\n", badrows]))
		# remove the bad rows and any remaining empty rows from the data
		data.drop(index = data[invalid_recs].index, inplace = True)
		empty_rows = data.isna().all(axis = 1)
		data.drop(index = data[empty_rows].index, inplace = True)

	log.info("Data validation completed.")

	# add a new field to data where processing results
	# will be recorded; default vals are empty strings.
	data = data.assign(Message = "", Credit_Note = pd.NA)
	params.update({"data": data})

	return params


# ====================================
# 		Management of SAP connection
# ====================================

def connect_to_sap(system: str) -> CDispatch:
	"""Creates connection to the SAP GUI scripting engine.

	Parameters:
	-----------
	system:
		The SAP system to use for connecting to the scripting engine.

	Returns:
	--------
	An SAP `GuiSession` object that represents active user session.
	"""

	log.info("Connecting to SAP ...")
	sess = sap.connect(system)
	log.info("Connection created.")

	return sess

def disconnect_from_sap(sess: CDispatch) -> None:
	"""Closes connection to the SAP GUI scripting engine.

	Parameters:
	-----------
	sess:
		An SAP `GuiSession` object (wrapped in the `win32:CDispatch` class)
		that represents an active user SAP GUI session.
	"""

	log.info("Disconnecting from SAP ...")
	sap.disconnect(sess)
	log.info("Connection to SAP closed.")


# ====================================
# 		Document processing
# ====================================

def assign_credit_note_numbers(sess: CDispatch, in_data: DataFrame) -> DataFrame:
	"""Retrieves credit note numbers for the order numbers
	contained in data input using VA03 transaction.

	Parameters:
	-----------
	sess:
		An SAP `GuiSession` object.

	data:
		Original user data containing order numbers
		for which credit note numbers will be retrieved.

	Returns:
	--------
	Original user data with added 'Credit_Note' field containing
	credit note numbers for each order where the credit note has already
	been created. If no credit note is found for a given order, then
	the default 'Credit_Note' field value is left in the input data.
	"""

	if in_data["Document_Number"].isna().all():
		return in_data

	result = in_data.copy()
	doc_nums = result["Document_Number"].astype("string")
	result = result.assign(Is_Order = doc_nums.str.startswith(ORDER_NUMBERING_PREFIX))

	n_orders = result["Is_Order"].sum()
	log.info(f"Total orders to process: {n_orders}")

	if n_orders == 0:
		return result

	log.info("=== Assinging credit note numbers to orders START ===")
	result.assign(Credit_Note = pd.NA)
	orders = result["Document_Number"][result["Is_Order"]]

	log.info("Starting VA03 ...")
	va03.start(sess)
	log.info("The VA03 has been started.")

	for nth, order in enumerate(orders, start = 1):

		if nth == 1:
			utils.print_section_break(log, SECTION_LINE_LENGTH)

		mask = result["Document_Number"] == order
		log.info(f"Retrieving credit note number for order ({nth} of {n_orders}): {order} ...")
		n_attempts = 0

		while n_attempts < MAX_RECOVERY_ATTEMPTS:
			try:
				credit_note = va03.get_creditnote_number(str(order))
			except va03.SapConnectionLostError as exc:
				log.error(exc)
				log.error(f"Attempt # {n_attempts + 1} to handle the error ...")
				log.error("Reconnecting to SAP ...")
				sess = connect_to_sap(sap.system_code)
				log.error("Connection restored.")
				log.error("Restarting VA03 ...")
				va03.start(sess)
				log.error("The transaction has been restarted.")
				n_attempts += 1
			except va03.DocumentProcessingError as exc:
				log.error(str(exc))
				result.loc[mask, "Message"] = str(exc)
				break
			except va03.DocumentProcessingWarning as wng:
				log.warning(wng)
				result.loc[mask, "Message"] = str(wng)
				break
			except va03.CreditNoteNotFoundWarning as wng:
				log.warning(wng)
				result.loc[mask, "Message"] = str(wng)
				result.loc[mask, "Credit_Note"] = None
			else:
				log.info(f"Credit note number retrieved: {credit_note}")
				result.loc[mask, "Credit_Note"] = credit_note
				n_attempts = 0
				break

		if n_attempts != 0:
			# critical error, do not return an error message but raise an exception
			raise RuntimeError("Attempts to handle the va03.SapConnectionLostError exception failed!")

		utils.print_section_break(log, SECTION_LINE_LENGTH)

	log.info("Closing VA03 ...")
	va03.close()
	log.info("The VA03 has been closed.")

	if result["Credit_Note"].isna().all():
		raise RuntimeError("None of the processed orders was found in VA03!")

	result["Credit_Note"] = result["Credit_Note"].astype("UInt64")

	log.info("=== Assinging credit note numbers to orders END ===\n")

	return result

def update_accounting_documents(sess: CDispatch, in_data: DataFrame) -> DataFrame:
	"""Updates document (credit note/invoice) text on case ID in FB03.

	Parameters:
	-----------
	sess:
		An SAP `GuiSession` object.

	in_data:
		Original user data containing numbers of documents to update.

	Returns:
	--------
	Original user data updated on 'Message' strings
	that inform the user of the document processing result.
	"""

	n_docs = in_data["Document_Number"].notna().sum()
	log.info(f"Total documents to update: {n_docs}")

	if n_docs == 0:
		return in_data

	log.info("=== Updating of accountig documents START ===")

	log.info("Starting FB03 ...")
	fb03.start(sess)
	log.info("The FB03 has been started.")

	result = in_data.copy()

	for nth, row in enumerate(in_data.itertuples(), start = 1):

		if nth == 1:
			utils.print_section_break(log, SECTION_LINE_LENGTH)

		idx = row.Index

		# insert case ID into the document text if not yet contained
		if not row.Is_Order:
			doc_num = row.Document_Number
		elif not pd.isna(row.Credit_Note):
			doc_num = int(row.Credit_Note)
		else:
			msg = "" if pd.isna(result.loc[idx, "Message"]) else result.loc[idx, "Message"]
			result.loc[idx, "Message"] = msg
			result.loc[idx, "Message"] += " Document skipped."
			result.loc[idx, "Message"] = result.loc[idx, "Message"].strip()
			continue

		doc_year = int(row.Document_Year)
		cocd = row.Company_Code
		case_id = int(row.Case_ID)

		log.info(f"Updating document ({nth} of {n_docs}): {doc_num} ...")
		n_attempts = 0

		while n_attempts < MAX_RECOVERY_ATTEMPTS:
			try:
				fb03.append_case_id(doc_num, doc_year, cocd, case_id)
			except fb03.SapConnectionLostError as exc:
				log.error(exc)
				log.error(f"Attempt # {n_attempts + 1} to handle the error ...")
				log.error("Reconnecting to SAP ...")
				sess = connect_to_sap(sap.system_code)
				log.error("Connection restored.")
				log.error("Restarting FB03 ...")
				fb03.start(sess)
				log.error("The transaction has been restarted.")
				n_attempts += 1
			except fb03.CaseIdContainedWarning as wng:
				result.loc[idx, "Message"] = str(wng)
				log.warning(wng)
				break
			except fb03.DocumentProcessingError as exc:
				result.loc[idx, "Message"] = str(exc)
				log.error(exc)
				break
			else:
				result.loc[idx, "Message"] = "Document updated."
				n_attempts = 0
				log.info("Document updated.")
				break

		if n_attempts != 0:
			# critical error, do not return an error message but raise an exception
			raise RuntimeError("Attempts to handle the fb03.SapConnectionLostError exception failed!")

		utils.print_section_break(log, SECTION_LINE_LENGTH)

	log.info("Closing FB03 ...")
	fb03.close()
	log.info("The FB03 has been closed.")

	log.info("=== Updating of accountig documents END ===\n")

	return result


# ====================================
# 		Notification closing
# ====================================

def close_service_notifications(sess: CDispatch, in_data: DataFrame) -> DataFrame:
	"""Closes service notifications in QM02.

	Parameters:
	-----------
	sess:
		An SAP `GuiSession` object.

	in_data:
		Original user data that contains
		numbers of notifications to close.

	Returns:
	--------
	Original user data updated on 'Message' strings
	that inform the user of the document processing result.
	"""

	result = in_data.copy()
	n_notifs = result["Notification"].notna().sum()

	log.info(f"Total notifications to close: {n_notifs}.")

	if n_notifs == 0:
		return in_data

	log.info("=== Closing of service notifiacitons START ===")

	log.info("Starting QM02 ...")
	qm02.start(sess)
	log.info("The QM02 has been started.")

	for nth, row in enumerate(in_data.itertuples(), start = 1):

		if nth == 1:
			utils.print_section_break(log, SECTION_LINE_LENGTH)

		idx = row.Index
		notif_id = int(row.Notification)

		log.info(f"Completing notification ({nth} of {n_notifs}): {notif_id} ...")

		n_attempts = 0

		while n_attempts < MAX_RECOVERY_ATTEMPTS:
			try:
				qm02.complete_notification(notif_id)
			except qm02.SapConnectionLostError as exc:
				log.error(exc)
				log.error(f"Attempt # {n_attempts + 1} to handle the error ...")
				log.error("Reconnecting to SAP ...")
				sess = connect_to_sap(sap.system_code)
				log.error("Connection restored.")
				log.error("Restarting FB03 ...")
				fb03.start(sess)
				log.error("The transaction has been restarted.")
				n_attempts += 1
			except qm02.NotificationCompletionError as exc:
				log.error(exc)
				msg = str(exc)
				break
			except qm02.NotificationCompletionWarning as wng:
				log.warning(wng)
				msg = str(wng)
				break
			except Exception as exc:
				log.exception(exc)
				msg = str(exc)
				break
			else:
				msg = "Notification completed."
				log.info(msg)
				n_attempts = 0
				break

		if n_attempts != 0:
			# critical error, do not return an error message but raise an exception
			raise RuntimeError("Attempts to handle the qm02.SapConnectionLostError exception failed!")

		result.loc[idx, "Message"] = " ".join([result.loc[idx, "Message"],  msg])
		utils.print_section_break(log, SECTION_LINE_LENGTH)

	log.info("Closing QM02 ...")
	qm02.close()
	log.info("The QM02 has been closed.")

	log.info("=== Closing of service notifiacitons END ===\n")

	return result


# ====================================
# 	Reporting of processing output
# ====================================

def create_report(temp_dir: str, report_cfg: dict, data: DataFrame) -> str:
	"""Creates user report from the processing result.

	Parameters:
	-----------
	temp_dir:
		Path to the directory where temporary files are stored.

	data_cfg:
		Application 'data' configuration parameters.

	data:
		The processing result from which report will be generated.

	Returns:
	--------
	Path to the report file.
	"""

	log.info("Creating user report ...")
	report_name =report_cfg["file_name"]
	report_path = join(temp_dir, f"{report_name}.xlsx")
	report.generate_excel_report(data, report_path, report_cfg["datasheet_name"])
	log.info("Report successfully created.")

	return report_path

def send_notification(
		msg_cfg: dict,
		user_mail: str,
		template_dir: str,
		attachment: Union[dict, str] = None, # type: ignore
		error_msg: str = ""
	) -> None:
	"""Sends a notification with processing result to the user.

	Parameters:
	-----------
	msg_cfg:
		Application 'messages' configuration parameters.

	user_mail:
		Email address of the user who requested processing.

	template_dir:
		Path to the application directory
		that contains notification templates.

	attachment:
		Attachment name and data or a file path.

	error_msg:
		Error message that will be included in the user notification.
		By default, no erro message is included.
	"""

	log.info("Sending notification to user ...")

	notif_cfg = msg_cfg["notifications"]

	if not notif_cfg["send"]:
		log.warning(
			"Sending of notifications to users "
			"is disabled in 'app_config.yaml'.")
		return

	if error_msg != "":
		templ_name = "template_error.html"
	else:
		templ_name = "template_completed.html"

	templ_path = join(template_dir, templ_name)

	with open(templ_path, encoding = "utf-8") as stream:
		html_body = stream.read()

	if error_msg != "":
		html_body = html_body.replace("$error_msg$", error_msg)

	if attachment is None:
		msg = mails.create_smtp_message(
			notif_cfg["sender"], user_mail,
			notif_cfg["subject"], html_body
		)
	elif isinstance(attachment, dict):
		msg = mails.create_smtp_message(
			notif_cfg["sender"], user_mail,
			notif_cfg["subject"], html_body,
			{attachment["name"]: attachment["content"]}
		)
	elif isinstance(attachment, str):
		msg = mails.create_smtp_message(
			notif_cfg["sender"], user_mail,
			notif_cfg["subject"], html_body,
			attachment
		)
	else:
		raise ValueError(f"Unsupported data type: '{type(attachment)}'!")

	try:
		mails.send_smtp_message(msg, notif_cfg["host"], notif_cfg["port"])
	except Exception as exc:
		log.error(exc)
		return

	log.info("Notification sent.")


# ====================================
# 			Data cleanup
# ====================================

def delete_temp_files(temp_dir: str) -> None:
	"""Removes all temporary files.

	Parameters:
	-----------
	temp_dir:
		Path to the directory where temporary files are stored.
	"""

	file_paths = glob(join(temp_dir, "*.*"))

	if len(file_paths) == 0:
		return

	log.info("Removing temporary files ...")

	for file_path in file_paths:
		try:
			os.remove(file_path)
		except Exception as exc:
			log.exception(exc)

	log.info("Files successfully removed.")
