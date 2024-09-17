# pylint: disable = E0110, E1101

"""
Description:
------------
The 'report.py' module generates user reports as 
Excel files from the document processing output.

Version history:
----------------
1.0.20231103 - Initial version.
"""

from os.path import exists, dirname
from pandas import DataFrame, Series, ExcelWriter

FilePath = str

class FolderNotFoundError(Exception):
	"""When a directory is requested but doesn't exist."""

def _calculate_max_column_width(column_data_values: Series, column_name: str) -> int:
	"""Returns excel column width calculated as the maximum count of
	characters contained in the column name and column data strings.

	Parameters:
	-----------
	column_data_values:
		A Series of field values for which the width will be calculated.

	column_name:
		Name of the column for which the width will be calculated.

	Returns:
	--------
	The width of a column in an Excel sheet.
	"""

	column_lengths = column_data_values.astype("string")
	column_lengths = column_lengths.dropna().str.len().to_list()
	column_lengths.append(len(column_name))
	column_width = max(column_lengths)

	if column_name != "Message":
		# add some points to fields other
		# than 'Message' for a better readbility
		column_width += 2

	return column_width

def generate_excel_report(data: DataFrame, file: FilePath, sheet_name: str) -> None:
	"""Creates a user Excel report from the processing result.

	Parameters:
	-----------
	data:
		Data to print to the report sheet.

	file:
		Path to the .xlsx report file to create.
		
		If the destination folder doesn't exist, then 
		an `FolderNotFoundError` exception is raised.

	sheet_name:
		Name of the sheet with printed data.
	"""

	if not exists(dirname(file)):
		raise FolderNotFoundError(
			"The report directory not found at the "
			f"path specified: '{dirname(file)}'")

	# The 'Is_Order' column is  ahelper column that indicates
	# whether an accounting document is a credit note order or not.
	# Helper columns are generally excluded from user reports.
	data.drop("Is_Order", axis = 1, inplace = True, errors = "ignore")

	# reorder fields so that they appear in a meaningful order for the user
	field_order = (
		"Company_Code", "Document_Number", "Document_Year",
		"Case_ID", "Notification", "Credit_Note", "Message"
	)

	data = data.reindex(field_order, axis = 1)

	with ExcelWriter(file, engine = "xlsxwriter") as wrtr:

		# remove underscores from column names, then write teh data to Excel.
		# Once the data is written, replace the spaces back with underscores
		# for a more convenient data manipulation across code.
		data.columns = data.columns.str.replace("_", " ")
		data.to_excel(wrtr, index = False, sheet_name = sheet_name)
		data.columns = data.columns.str.replace("_", " ")

		report = wrtr.book
		data_sht = wrtr.sheets[sheet_name]
		general_fmt = report.add_format({"align": "center"})

		for idx, col in enumerate(data.columns):
			col_width = _calculate_max_column_width(data[col], col)
			data_sht.set_column(idx, idx, col_width, general_fmt)
