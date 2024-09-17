# pylint: disable = C0103, C0301, W0703, W1203

"""
Description:
------------
The 'Accounting document updater' application automates updating of document parameters
in the FB03 transaction. User sends a data file containing a list of document numbers
to process along with mandatory document information such as document year, company code
(required to find the document in FB03), case ID (assigned in UDM_Dispute) an dnotification
number. The application then checks the text of each docuemnt for the presence of the
corresponding case ID If the text does not contain the case ID, the ID value is appended
to the text of the document. If any notification number (ID) is provided, then the
notification will be completed in QM02. Finally, a summary file is generated, attached to
an email and sent to the user.

Version history:
----------------
1.0.20201026 - Initial version.
1.1.20220613 - Added notification closing in QM02. Credit notes can now be identified in VA03
			   based on order numbers provided by the user and subsequently updated on case IDs.
1.1.20220616 - Updated storing of VA03 error messages into user data in 'get_credit_notes()'.
1.1.20221013 - Updated docstrings.
			 - Added new control handlers for signals received from service procedures.
1.1.20231106 - Updated dotrings. Major code refactoring.
"""

from os.path import join
from datetime import datetime as dt
import argparse
import logging
import sys
from engine import controller

log = logging.getLogger("master")

def main(args: dict) -> int:
	"""Program entry point.

	args:
	-----
	email_id:
		Identification string of the user message.
		The value of `Message.message_id` property.

	Returns:
	--------
	An integer representing the program's completion state:
	- 0: Program successfully completes.
	- 1: Program fails during while configuring the loging system.
	- 2: Program fails during the initialization phase.
	- 3: Program fails during the processing phase.
	- 4: Program fails during the reporting phase.
	"""

	app_dir = sys.path[0]
	log_dir = join(app_dir, "logs")
	temp_dir = join(app_dir, "temp")
	template_dir = join(app_dir, "notification")
	app_cfg_path = join(app_dir, "app_config.yaml")
	log_cfg_path = join(app_dir, "log_config.yaml")
	curr_date = dt.now().strftime("%d-%b-%Y")

	try:
		controller.configure_logger(
			log_dir, log_cfg_path,
			"Application name: Accounting Document Updater",
			"Application version: 1.1.20230118",
			f"Log date: {curr_date}")
	except Exception as exc:
		print(exc)
		print("CRITICAL: Unhandled exception while trying to configuring the logging system!")
		return 1

	try:
		log.info("=== Initialization START ===")
		cfg = controller.load_app_config(app_cfg_path)
		sess = controller.connect_to_sap(cfg["sap"]["system"])
		log.info("=== Initialization END ===\n")
	except Exception as exc:
		log.exception(exc)
		log.critical("Unhandled exception while trying to initialize the application!")
		return 2

	try:
		log.info("=== Fetching user input START ===")
		user_input = controller.fetch_user_input(cfg["messages"], args["email_id"])
		log.info("=== Fetching user input END ===\n")
	except Exception as exc:
		log.exception(exc)
		log.critical("=== Fetching user input FAILURE ===")
		log.info("=== Cleanup START ===")
		controller.disconnect_from_sap(sess)
		controller.delete_temp_files(temp_dir)
		log.info("=== Cleanup END ===\n")
		return 2

	if user_input["error_message"] != "":
		log.error(user_input["error_message"])
		controller.send_notification(
			cfg["messages"], user_input["email"], template_dir,
			error_msg = user_input["error_message"])
		log.info("=== Cleanup START ===")
		controller.disconnect_from_sap(sess)
		controller.delete_temp_files(temp_dir)
		log.info("=== Cleanup END ===\n")
		return 0

	try:
		log.info("=== Processing START ===")
		in_data = controller.assign_credit_note_numbers(sess, user_input["data"])
		update_output = controller.update_accounting_documents(sess, in_data)
		closing_output = controller.close_service_notifications(sess, update_output)
		log.info("=== Processing END ===\n")
	except Exception as exc:
		log.exception(exc)
		log.critical("=== Processing FAILURE ===\n")
		log.info("=== Cleanup START ===")
		controller.disconnect_from_sap(sess)
		controller.delete_temp_files(temp_dir)
		log.info("=== Cleanup END ===\n")
		return 3

	try:
		log.info("=== Reporting START ===")
		report_path = controller.create_report(temp_dir, cfg["report"], closing_output)
		controller.send_notification(cfg["messages"], user_input["email"], template_dir, report_path)
		log.info("=== Reporting END ===\n")
	except Exception as exc:
		log.exception(exc)
		return 4
	finally:
		log.info("=== Cleanup START ===")
		controller.delete_temp_files(temp_dir)
		controller.disconnect_from_sap(sess)
		log.info("=== Cleanup END ===\n")

	return 0

if __name__ == "__main__":

	parser = argparse.ArgumentParser()
	parser.add_argument("-e", "--email_id", required = True, help = "Sender message id.")
	parser.add_argument("-d", "--debug", required = False, default = False, help = "Enables debug logging level.")
	exit_code = main(vars(parser.parse_args()))

	log.info(f"=== System shutdown with return code: {exit_code} ===")
	logging.shutdown()
	sys.exit(exit_code)
