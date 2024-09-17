"""Contains utility functions for common tasks across applications."""

from logging import Logger


def print_section_break(
		log: Logger, n_chars: int = 20, tag:
		str = "", char: str = "-", end: str = "",
		sides: str = "both") -> None:
	"""Print a log section line.

	Parameters:
	-----------
	log:
		Logger used to print the section break.

	n_chars:
		Number of characters used for the indentation from both sides (default: 20).

	tag:
		A text to insert before the counter of the section line (default: "").

	char:
		The character used to create the indentation sequence (default: "-").

	end:
		Ending character of the line (default: "").

	sides:
		Sides to indent:
		- "both": Both sides are indented (default behavior).
		- "left": Only the left side is indented.
		- "right": Only the right side is only indented.
	"""

	indentation = char * n_chars

	if sides.lower() == "both":
		log.info("".join([indentation, tag, indentation, end]))
	elif sides.lower() == "left":
		log.info("".join([indentation, tag, end]))
	elif sides.lower() == "right":
		log.info("".join([tag, indentation, end]))
	else:
		raise ValueError(f"Unrecognozed value for argument 'sides': '{sides}'")
