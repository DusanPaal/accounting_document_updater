# pylint: disable = C0103, E1101, W0603

"""
Description:
------------
The 'qm02.py' module provides the following exclusive procedures:
	- complete_notification():
		For completion of a service notification.

How to use:
-----------
The QM02 must be started by calling the `start()` procedure.

Attempt to use an exclusive procedure when QM02 has not been 
started results in the `UninitializedModuleError` exception.

After using the module, the transaction should be closed,
and the resources released by calling the `close()` procedure.

Version history:
----------------
1.0.20231122 - Initial version.
"""

import logging
from time import sleep
from typing import Union
import pywintypes
from win32com.client import CDispatch

_sess = None
_main_wnd = None
_stat_bar = None

log = logging.getLogger("master")

# keyboard to SAP virtual keys mapping
_virtual_keys = {
	"Enter":        0,
	"F2":           2,
	"F8":           8,
	"F9":           9,
	"CtrlS":        11,
	"F12":          12,
	"ShiftF3":      15,
	"ShiftF4":      16,
	"ShiftF12":     24,
	"CtrlF1":       25
}


class NotificationCompletionWarning(Warning):
	"""When a SAP warning appears during
	the process of document updating.
	"""

class NotificationCompletionError(Exception):
	"""When attempting to complete a
	notification results into an error.
	"""

class UninitializedModuleError(Exception):
	"""Attempting to use a procedure
	before starting the transaction.
	"""

class SapConnectionLostError(Exception):
	"""When a connection to SAP is lost."""

def _is_error_message() -> bool:
	"""Checks whether a status bar message indicates an error."""
	return _stat_bar.messageType == "E"

def _press_key(name: str) -> None:
	"""Simulates pressing a keyboard button."""
	_main_wnd.SendVKey(_virtual_keys[name])

def _to_initial_screen() -> None:
	"""Navigates back to the initial screen of the transaction."""

	_press_key("ShiftF3")
	_press_key("F12")

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

def _get_task_viewer() -> CDispatch:
	"""Returns a GuiTable object containing notification tasks."""
	return _main_wnd.FindByName("SAPLIQS0MASSNAHMEN_VIEWER", "GuiTableControl")

def _select_task(row_idx: int) -> None:
	"""Selects a task contained in QM02 task 
	table based on their task row index.
	"""

	# set scrollbar to tha appropriate position
	task_viewer = _get_task_viewer()
	task_viewer.VerticalScrollbar.position = row_idx

	# mark the task row as selected
	task_viewer = _get_task_viewer()
	task_viewer.getAbsoluteRow(row_idx).selected = True

def _activate_tab(name: str) -> None:
	"""Activates a specific tab."""

	if name == "Data":
		obj_name = "TAB_GROUP_10"
	else:
		assert False, "Unrecognized tab name!"

	_main_wnd.FindByName(obj_name, "GuiTabStrip").children[3].select()

def _complete_task(task_num: int) -> None:
	"""Completes a task."""

	_select_task(task_num - 1)
	_main_wnd.findByName("FC_ERLEDIGT", "GuiButton").press()

	if _sess.activewindow.text == "Status check error":
		_close_popup_dialog(confirm = False)
		return

	while _is_popup_dialog():
		_close_popup_dialog(confirm = True)

def _set_notification_id(val: str) -> None:
	"""Enters the notification ID into the 'Notification' field
	located on the QM02 initial window and confirms the value.
	"""
	_main_wnd.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").text = val

def _clear_notification_search_field() -> None:
	"""Clears the value in the notification
	field on the initial screen."""
	_set_notification_id("")

def _get_task_param(task_viewer: CDispatch, row_idx: int, param: str) -> Union[int,str]:
	"""Returns parameter value of a task defined by a row index."""

	if param == "number":
		num = task_viewer.GetCell(row_idx, 0).text
		param = int(num)
	elif param == "completion_date":
		param = task_viewer.GetCell(row_idx, 16).text
	else:
		assert False, f"Unrecognized parameter: {param}!"

	return param

def _get_available_tasks(task_viewer: CDispatch) -> dict:
	"""Returns a map of ID numbers of active  
	notification tasks and their completion dates.
	"""

	tasks = {}
	row_idx = 0

	# Never use the VisibleRowCount property since task
	# count may be higher than those visible in the grid!
	last_row_idx = task_viewer.RowCount - 1

	while row_idx < last_row_idx:

		visible_row_idx = row_idx % task_viewer.visibleRowCount

		# move down the scrollbar so that next positions appear in the table
		if visible_row_idx == 0 and row_idx > 0:
			task_viewer.VerticalScrollbar.position = row_idx
			task_viewer = _get_task_viewer()

		# program reaches the end of task list (the row contains no ID)
		if task_viewer.GetCell(visible_row_idx, 0).text == "":
			break

		# get task parameters
		task_num = _get_task_param(task_viewer, visible_row_idx, "number")
		compl_date = _get_task_param(task_viewer, visible_row_idx, "completion_date")

		# store task parameters
		tasks[task_num] = {
			"completion_date": compl_date,
			"row_idx": row_idx
		}

		# done analyzing a task, go to next row
		row_idx += 1

	return tasks

def _reopen_blocked(max_attempts: int, wait_secs: int) -> None:
	"""Opens a notification that hass been
	temporarily blocked by the user account."""

	n_attempts = 0

	while n_attempts <= max_attempts:
		n_attempts += 1
		log.debug(f"Attempt # {n_attempts} to reopen the notification ...")
		sleep(wait_secs)
		_press_key("Enter")

		if not _is_error_message():
			break # no longer blocked

	# still blocked after all the attempts - raise exception
	if _is_error_message():
		raise NotificationCompletionError(_stat_bar.text)

def _search_notification(val: int) -> None:
	"""Searches and opens a service notification."""

	_set_notification_id(str(val))
	_press_key("Enter")

	if _is_popup_dialog():
		_close_popup_dialog(confirm = True)

	msg = _stat_bar.text

	if "does not exist" in msg:
		raise NotificationCompletionError(msg)

	if "can only be displayed" in msg:
		_press_key("F12")
		raise NotificationCompletionWarning(msg)

	if "blocked" in msg:
		_reopen_blocked(max_attempts = 3, wait_secs = 2)
	elif _is_popup_dialog():
		_close_popup_dialog(confirm = True)
		_to_initial_screen()
		raise NotificationCompletionError(msg)

def _complete(task_viewer: CDispatch) -> CDispatch:
	"""Releases active tasks and complete the notification."""

	tasks = _get_available_tasks(task_viewer)

	# release active task(s) where no competion date exists
	for task_num, param in tasks.items():
		if param["completion_date"] != "":
			continue
		_complete_task(task_num)

	# no open taks left, return the refrence to the completion button
	_main_wnd.findById("tbar[1]/btn[20]").press()

	if _sess.ActiveWindow.text == "Status check error":
		_close_popup_dialog(confirm = False)
		_press_key("F12")

		if _is_popup_dialog():
			_close_popup_dialog(False)

		raise NotificationCompletionError("Status check error!")

def start(sess: CDispatch) -> None:
	"""Starts the QM02 transaction.

	If the QM02 has already been started, then
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

	_sess.StartTransaction("QM02")

	# ensure that the mask contains no initial
	# value in the notification search field
	_clear_notification_search_field()

def close() -> None:
	"""Closes a running QM02 transaction.

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

def complete_notification(notif_id: int) -> None:
	"""Completes an opened service notification.

	If the notification is already completed,
	then a `NotificationCompletionWarning` warning is raised.

	If the attempt to complete a notification fails,
	then a `NotificationCompletionError` exception is raised.

	When the connection to SAP is lost due to an error,
	then a `SapConnectionLostError` exception is raised.

	Parameters:
	-----------
	notif_id:
		Identification number of a service notification.
	"""

	try:

		_check_prerequisities()
		_search_notification(notif_id)
		_activate_tab(name = "Data")
		task_viewer = _get_task_viewer()
		_complete(task_viewer)

		# confirm completion
		_press_key("Enter")

		# ensure the contents in the search fields
		# is cleared before next call of the procedure
		_clear_notification_search_field()

	except pywintypes.com_error:
		try:
			_sess.IsActive
		except AttributeError as exc:
			raise SapConnectionLostError("Connection to SAP lost!") from exc
		raise
