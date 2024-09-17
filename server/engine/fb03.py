# pylint: disable = C0103, E1101, W0603

"""
Description:
------------
The 'fb03.py' module provides the following exclusive procedures:
	- append_case_id():
		For appending case ID numbers to an existing 'Text' string
		of accounting documents.
	- remove_attachments():
		For removing files attached to an accounting document.

How to use:
-----------
The FB03 must be started by calling the `start()` procedure.

Attempt to use an exclusive procedure when FB03 has not been 
started results in the `UninitializedModuleError` exception.

After using the module, the transaction should be closed,
and the resources released by calling the `close()` procedure.

Version history:
----------------
1.0.20231102 - Initial version.
"""

import re
import pywintypes
from win32com.client import CDispatch

MAX_TEXT_FIELD_CHARS = 50

_sess = None
_main_wnd = None
_stat_bar = None

# keyboard to SAP virtual keys mapping
_virtual_keys = {
	"Enter":  0,
	"F2":     2,
	"F3":     3,
	"CtrlS":  11,
	"F12":    12,
	"CtrlF1": 25
}


class CaseIdContainedWarning(Warning):
	"""When document text already contains the case ID."""

class DocumentProcessingError(Exception):
	"""When document updating terminates with an error."""

class DocumentNotFoundError(Exception):
	"""When no docuemnt is found using the given search criteria."""

class SapConnectionLostError(Exception):
	"""When a connection to SAP is lost."""

class UninitializedModuleError(Exception):
	"""Attempting to use a procedure
	before starting the transaction.
	"""


def _press_key(name: str) -> None:
	"""Simulates pressing a keyboard button."""
	_main_wnd.SendVKey(_virtual_keys[name])

def _is_error_message() -> bool:
	"""Checks whether a status bar message indicates an error."""
	return _stat_bar.messageType == "E"

def _is_warning_message() -> bool:
	"""Checks whether a status bar message indicates a warning."""
	return _stat_bar.messageType == "W"

def _is_alert_message() -> bool:
	"""Checks whether a status bar message
	indicates an error or a warning.
	"""
	return _is_error_message() or _is_warning_message()

def _check_prerequisities() -> None:
	"""Checks whether prerequisities
	to run public procedures are met.
	"""

	if _sess is None:
		raise UninitializedModuleError(
			"Uninitialized module! Use the `start()` "
			"procedure to run the transaction first!")

def _is_popup_dialog() -> bool:
	"""Checks whether a window	is a pop-up dialog window."""
	return _sess.ActiveWindow.type == "GuiModalWindow"

def _is_cleared_document() -> bool:
	"""Checks whether a document has already been cleared."""
	return _main_wnd.findAllByName("BSEG-AUGBL", "GuiTextField").Count == 1

def _clear_search_criteria() -> None:
	"""Clears the contents of the search fields on the initial screen."""
	_set_search_criteria("", "", "")

def _close_popup_dialog(confirm: bool, max_attempts: int = 3) -> None:
	"""Confirms or declines a pop-up dialog."""

	nth = 0

	dialog_titles = (
		"Information",
		"Status check error",
		"Document lines: Display messages"
	)

	while _sess.ActiveWindow.text in (dialog_titles) and nth < max_attempts:
		if confirm:
			_press_key("Enter")
		else:
			_press_key("F12")

		nth += 1

	if _sess.ActiveWindow.type == "GuiModalWindow":
		dialog_title = _sess.ActiveWindow.text
		raise RuntimeError(f"Could not close the dialog window: '{dialog_title}'!")

	btn_caption = "Yes" if confirm else "No"

	for child in _sess.ActiveWindow.children:
		for grandchild in child.children:
			if grandchild.Type != "GuiButton":
				continue
			if btn_caption != grandchild.text.strip():
				continue
			grandchild.Press()
			return

def _save_changes() -> None:
	"""Saves changes made to the document parameters."""

	_press_key("CtrlS")

	msg = _stat_bar.text

	if _is_alert_message():
		# handle alert message

		if "Net due date on * is in the past" in msg:
			_press_key("Enter")
		else:
			raise DocumentProcessingError("Unable to save document changes!")

def _remove_duplicates(text: str, case_id: str) -> str:
	"""Removes duplicated case ID numbers from document text."""

	# temp fix for broken text in credit notes
	matches = re.findall(case_id, text)

	if len(matches) >= 2 and matches[0] == matches[1]:
		text = re.sub(case_id, "", text)
		text = text.strip()

	return text

def _delete_attachments() -> None:
	"""Deletes all document attachments."""

	# open attachment list
	_main_wnd.findById("titl/shellcont/shell").pressContextButton("%GOS_TOOLBOX")
	_main_wnd.findById("titl/shellcont/shell").selectContextMenuItem("%GOS_VIEW_ATTA")
	table = _sess.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell")

	# remove all attached files
	while table.rowcount != 0:
		table.selectedRows = "0"
		table.pressToolbarButton("%ATTA_DELETE")

	# confirm the empty attachment list
	_press_key("Enter")

def _clear_search_criteria() -> None:
	"""Clears all search criteria fields in the transaction initial mask."""
	_set_search_criteria("", "", "")

def _set_search_criteria(cocd: str, doc_num: int, doc_year: int) -> None:
	"""Enters search criteria into the transaction initial mask."""

	_set_company_code(cocd)
	_set_document_number(str(doc_num))
	_set_document_year(str(doc_year))
	_press_key("Enter")

	if "does not exist " in _stat_bar.text:
		raise DocumentProcessingError(_stat_bar.text)

def _set_company_code(val: str) -> None:
	"""Enters company code in the search field."""
	_main_wnd.findByName("RF05L-BUKRS", "GuiCTextField").text = val

def _set_document_number(val: str) -> None:
	"""Enters document number in the search field."""
	_main_wnd.findAllByName("RF05L-BELNR", "GuiTxtField")(1).text = val

def _set_document_year(val: str) -> None:
	"""Enters document year in the search field."""
	_main_wnd.findAllByName("RF05L-GJAHR", "GuiTxtField")(1).text = val

def _clear_search_criteria() -> None:
	"""Clears the contents of the search fields on the initial screen."""

	_set_company_code("")
	_set_document_number("")
	_set_document_year("")

def _to_initial_screen() -> None:
	"""Navigates back to the initial screen of the transaction."""

	_press_key("F12")
	_press_key("F12")

def _open_document_parameters() -> None:
	"""Opens the mask with document parameters."""

	_main_wnd.findAllByName("RF05L-ANZDT", "GuiTextField").ElementAt(0).SetFocus()
	_press_key("F2")

def _append_case_id(text: str, case_id: int) -> str:
	"""Appends case ID to a text."""

	# check if the document text contains the case id
	if str(case_id) in text:
		_to_initial_screen()
		raise CaseIdContainedWarning(
			"Document text already contains the case ID!")

	compiled_text = _remove_duplicates(text, str(case_id))
	compiled_text = f"{compiled_text.strip()} D {case_id}"

	# check if the new text length is within limit
	if not len(compiled_text) <= MAX_TEXT_FIELD_CHARS:
		raise DocumentProcessingError(
			f"'Text' value exceeds {MAX_TEXT_FIELD_CHARS} chars length!")

	return compiled_text

def _get_document_text() -> str:
	"""Retriefes document 'Text' value."""
	return _main_wnd.findByName("BSEG-SGTXT", "GuiCTextField").text

def _set_document_text(val: str) -> None:
	"""Enters a text value into the 'Text' field."""

	# Some documents that have already been deleted cannot be edited
	# and will throw an erroran attempts is made to write the the new
	# text into the 'Text' field. If this happens, throw an explicit
	# exception.
	try:
		_main_wnd.findByName("BSEG-SGTXT", "GuiCTextField").text = val
	except AttributeError as exc:
		_to_initial_screen()
		if _is_cleared_document():
			raise DocumentProcessingError(
				"The document is already cleared "
				"and the text cannot be changed!"
			) from exc

		raise DocumentProcessingError(str(exc)) from exc

def start(sess: CDispatch) -> None:
	"""Starts the FB03 transaction.

	If the FB03 has already been started, then
	the running transaction will be restarted.

	Parameters:
	-----------
	sess:
		An SAP `GuiSession` object.
	"""

	global _sess
	global _main_wnd
	global _stat_bar

	_sess = sess
	_main_wnd = _sess.findById("wnd[0]")
	_stat_bar = _main_wnd.findById("sbar")

	_sess.StartTransaction("FB03")

	# ensure that the mask contains no
	# initial values in the search fields
	_clear_search_criteria()

def close() -> None:
	"""Closes a running FB03 transaction.

	Attempt to close the transaction that has not been
	started by the `start()` procedure is ignored.
	"""

	global _sess
	global _main_wnd
	global _stat_bar

	if _sess is None:
		return

	_sess.EndTransaction()

	if _is_popup_dialog():
		_close_popup_dialog(confirm = True)

	_sess = None
	_main_wnd = None
	_stat_bar = None

def append_case_id(
		document: int, fiscal_year: int,
		company_code: str, case_id: int
	) -> None:
	"""Appends case ID to an existing document text.

	The case is appended to the original text so that
	in a specific format: "$original_text$ D $case_id$".

	If the document text already contains the case ID,
	then a CaseIdContainedWarning exception is raised.

	When the connection to SAP is lost due to an error,
	then a SapConnectionLostError exception is raised.

	A DocumentProcessingError exception is raised, when:
	- attempting to modify a cleared document that is no longer editable
	- attempting to edit a document result in an SAP error
	- attempting to save changes to a document generates an SAP error
	- text length limit is exceeded by inserting a case ID
	- a dialog window appears, for which no appropriate handler exists

	Example:
	--------
	>>> original text: "RET711884319"
	>>> append_case_id(544411698, 2024, "0010", 400081469)
	>>> updated text: "RET711884319 D 400081469"

	Parameters:
	-----------
	document:
		Number of the accounting document to modify.

	fiscal_year:
		Fiscal year in which the document was created.

	company_code:
		Company code for which the document was created.

	case_id:
		Identification number under which the case is stored in DMS.
	"""

	_check_prerequisities()

	try:

		# enter document parameters into the transaction's search
		# fields, then confirm teh values to open the document
		_set_search_criteria(company_code, document, fiscal_year)

		# possible alert messages
		if _is_alert_message():
			raise DocumentProcessingError(_stat_bar.text)

		# open the line item - first line
		_open_document_parameters()

		# get document text field value
		old_text = _get_document_text()

		# append case ID to the document text
		new_text = _append_case_id(old_text, case_id)

		if _is_popup_dialog():
			_close_popup_dialog(confirm = False)
			raise DocumentProcessingError("Unable to edit the document!")

		# toggle edit mode, set the new text value, and save the changes.
		_press_key("CtrlF1")
		_set_document_text(new_text)
		_save_changes()

		# ensure the contents in the search fields
		# is cleared before next call of the procedure
		_clear_search_criteria()

	except pywintypes.com_error:
		try:
			_sess.IsActive
		except AttributeError as exc:
			raise SapConnectionLostError("Connection to SAP lost!") from exc
		raise

def remove_attachments(document: int, fiscal_year: int, company_code: str) -> None:
	"""Removes all files attached to an accounting document.

	If no document  is found using the specified searching criteria,
	then a `DocumentNotFoundError` exception is raised.

	When the connection to SAP is lost due to an error,
	then a `SapConnectionLostError` exception is raised.

	Parameters:
	-----------
	document:
		Number of the accounting document.

	fiscal_year:
		Fiscal year in which the document was created.

	company_code:
		Company code for which the document was created.
	"""

	_check_prerequisities()

	try:
		_set_search_criteria(company_code, document, fiscal_year)
		_delete_attachments()

		# press back
		_press_key("F3")

		# ensure the contents in the search fields
		# is cleared before next call of the procedure
		_clear_search_criteria()

	except pywintypes.com_error:
		try:
			_sess.IsActive
		except AttributeError as exc:
			raise SapConnectionLostError("Connection to SAP lost!") from exc
		raise
