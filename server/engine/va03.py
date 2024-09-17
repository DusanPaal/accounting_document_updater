# pylint: disable = C0103, E1101, W0603, W0703

"""
Description:
------------
The 'va03.py' module automates the retrieval of credit note
numbers for credit note order requests using the SAP 
transaction VA03.

The module provides the following exclusive procedures:
	- get_creditnote_number():
		For retrieving a credit note number for an order.

How to use:
-----------
The VA03 must be started by calling the `start()` procedure.

Attempt to use an exclusive procedure when VA03 has not been 
started results in the `UninitializedModuleError` exception.

After using the module, the transaction should be closed,
and the resources released by calling the `close()` procedure.

Version history:
----------------
1.0.20231122 - Initial version.
"""

from typing import Union
import pywintypes
from win32com.client import CDispatch

_sess = None
_main_wnd = None
_stat_bar = None

_virtual_keys = {
	"Enter":        0,
	"F2":           2,
	"F3":           3,
	"F5":           5,
	"F12":          12,
	"ShiftF3":      15
}


class CreditNoteNotFoundWarning(Warning):
	"""Raised if a credit note does not exist
	yet the system a specific order."""

class DocumentProcessingWarning(Warning):
	"""Raised when a warning message appears
	during document processing.
	"""

class DocumentProcessingError(Exception):
	"""Raised when document processing
	terminates due to an error.
	"""

class UninitializedModuleError(Exception):
	"""Raised when attempting to use a procedure
	before starting the transaction.
	"""

class SapConnectionLostError(Exception):
	"""When a connection to SAP is lost."""

def _press_key(name: str) -> None:
	"""Simulates pressing a keyboard button."""
	_main_wnd.SendVKey(_virtual_keys[name])

def _check_prerequisities() -> None:
	"""Checks whether prerequisities
	to run public procedures are met.
	"""

	if _sess is None:
		raise UninitializedModuleError(
			"Uninitialized module! Use the `start()` "
			"procedure to run the transaction first!")

def _is_popup_dialog() -> bool:
	"""Checks whether a window is a pop-up dialog window."""
	return _sess.ActiveWindow.type == "GuiModalWindow"

def _is_error_message() -> bool:
	"""Checks whether a status bar message indicates an error."""
	return _stat_bar.messageType == "E"

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

def _get_dialog_message() -> str:
	"""Reads message text from a pop-up dialog window."""

	msg = ""
	idx = 1
	msg_coll = _sess.findById("wnd[1]/usr").findAllByName(f"MESSTXT{idx}", "GuiTextField")

	while len(msg_coll) > 0:
		msg = " ".join([msg, msg_coll[0].text])
		idx += 1
		msg_coll = _sess.findById("wnd[1]/usr").findAllByName(f"MESSTXT{idx}", "GuiTextField")

	return msg.strip()

def _get_reference_number(tree: CDispatch, node: CDispatch) -> int:
	"""Returns reference number for an accounting document."""

	credit_note_digits_count = 9
	tree.selectItem(node, "&Hierarchy")
	tree.doubleClickItem(node, "&Hierarchy")
	val = _main_wnd.findById("shellcont/shell").GetCellValue(0, "DOCNUV")
	num = int(val[-credit_note_digits_count:])

	return num

def _get_node_value(
		substr: str,
		tree: CDispatch = None,
		node: CDispatch = None
	) -> Union[None, int]:
	"""Searches SAP tree for a node containing a specific substring. 
	If such node is identified, then the node is double-clicked, 
	showing a table containing accounting details. From there,
	the credit note number is fetched, converted and returned.
	"""

	if tree is None:
		tree = _main_wnd.findById("usr/shell/shellcont[1]/shell[1]")
		nodes = tree.GetNodesCol()
		node = next(iter(nodes))

	text = tree.getNodeTextByKey(node)

	if substr in text:
		return _get_reference_number(tree, node)

	children = tree.GetSubNodesCol(node)

	if children is not None:

		res = _get_node_value(substr, tree, next(iter(children)))

		if res is not None:
			return res

	try:
		sibling = tree.GetNextNodeKey(node)
	except Exception:
		return None

	return _get_node_value(substr, tree, sibling)

def _set_order_number(val: str) -> None:
	"""Enters order number in the search field."""
	_main_wnd.findByName("VBAK-VBELN", "GuiCTextField").text = val

def _set_purchase_order_number(val: str) -> None:
	"""Enters purchase order number in the search field."""
	_main_wnd.findByName("RV45S-BSTNK", "GuiTextField").text = val

def _set_sold_to_number(val: str) -> None:
	"""Enters sold to party account number in the search field."""
	_main_wnd.findByName("RV45S-KUNNR", "GuiCTextField").text = val

def _set_delivery_number(val: str) -> None:
	"""Enters delivery number in the search field."""
	_main_wnd.findByName("RV45Z-LFNKD", "GuiCTextField").text = val

def _set_billing_document_number(val: str) -> None:
	"""Enters billing document number in the search field."""
	_main_wnd.findByName("RV45S-FAKKD", "GuiCTextField").text = val

def _set_wbs_element(val: str) -> None:
	"""Enters WBS element value in the search field."""
	_main_wnd.findByName("RV45S-PSPID", "GuiCTextField").text = val

def _clear_search_criteria() -> None:
	"""Clears the contents of the search fields on the initial screen."""

	_set_order_number("")
	_set_purchase_order_number("")
	_set_sold_to_number("")
	_set_delivery_number("")
	_set_billing_document_number("")
	_set_wbs_element("")

def _to_initial_screen() -> None:
	"""Navigates back to the initial screen of the transaction."""

	_press_key("ShiftF3")
	_press_key("F12")

def _open_document_parameters() -> None:
	"""Opens the mask with document parameters."""

	_main_wnd.findByName("KUWEV-KUNNR", "GuiCTextfield").setFocus()
	_press_key("F2")
	_press_key("F5")

def _handle_popup_dialog() -> None:
	"""Handles a pop-up dialog window based on the contained message."""

	msg = _get_dialog_message()
	_close_popup_dialog(confirm = True)

	if "information in the customer comment text" in msg:
		pass
	elif "Order is blocked" in msg:
		_press_key("F3")
		raise DocumentProcessingWarning(msg)
	else:
		_press_key("F3")
		raise DocumentProcessingError(msg)

def start(sess: CDispatch) -> None:
	"""Starts the FB03 transaction.

	If the FB03 has already been started,
	then the running transaction will be restarted.

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

	_sess.StartTransaction("VA03")
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

def get_creditnote_number(order: int) -> int:
	"""Retrieves credit note number related to an order.
	
	If a credit note does not exist yet in the system,
	then a `CreditNoteNotFoundWarning` warning is raised.

	If the order is blocked, then a `DocumentProcessingWarning` 
	warning is raised.
	
	If no document is found for a particular order number,
	or an unhandled error occurs, then a `DocumentProcessingError` 
	exception is raised.

	When the connection to SAP is lost due to an error, 
	then a `SapConnectionLostError` exception is raised.

	Parameters:
	-----------
	order:
		The number of the credited order.

	Returns:
	--------
	The number of the credit note created for an order.
	"""

	# check if prerequisities for procedure use are met
	_check_prerequisities()

	try:

		# enter the order number into the search mask and confirm
		_set_order_number(str(order))
		_press_key("Enter")

		# handle any pop-up dialog window that appears
		# after the order number has been entered
		if _is_popup_dialog():
			_handle_popup_dialog()

		if "is not in the database or has been archived" in _stat_bar.Text:
			raise DocumentProcessingError(f"Order {order} not found!")

		# open the mask with accounting parameters of the document
		# by clicking on the first row on the list of loaded postings
		_open_document_parameters()

		# handle any pop-up dialog window that
		# appears after the mask has been opened
		if _is_popup_dialog():
			_handle_popup_dialog()

		if _is_error_message():
			raise DocumentProcessingError(_stat_bar.Text)

		num = _get_node_value("Accounting document")
		_to_initial_screen()

		if num is None:
			raise CreditNoteNotFoundWarning(
				"The credit note does not exist yet in the system.")

		# ensure the contents in the search fields
		# is cleared before next call of the procedure
		_clear_search_criteria()

	except pywintypes.com_error:
		try:
			_sess.IsActive
		except AttributeError as exc:
			raise SapConnectionLostError("Connection to SAP lost!") from exc
		raise

	return num
