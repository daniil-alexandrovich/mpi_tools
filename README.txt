-------------------------
- PORTFOLIO UPDATE TOOL -
-------------------------


-------------------
- Getting Started -
-------------------

This package is compatible with Python 3.2+. Each of the external libraries used can be downloaded with pip:

    pip install openpyxl
    pip install pandas


----------------------------
- Running Script and Tests -
----------------------------

This is a command line tool that takes the following parameters:

	portfolio_updater.py input_file input_sheet output_file output_sheet [existing_file existing_sheet]

...which are defined as follows:

	* input_file: Filename or path of portfolio containing data to be
	   added to a Stylus portfolio.
	* input_sheet: Worksheet name within this file containing appropriate
	   data

	* output_file: Filename or path of Excel sheet to which the Stylus-
  	   formatted portfolio will be written.
	* output_sheet: Worksheet name within this file to which the formatted
   	   data is written. Must exist, for now.

	* existing_file: Filename or path of an optional Excel file to merge data
	   from input_file into.
	* existing_sheet: Sheet containing relevant data in this file.


Example tests can be run as follows:

	portfolio_updater.py in.xlsx in out.xlsx format
	* Generates a Stylus portfolio from in.xlsx

	portfolio_updater.py in.xlsx in out.xlsx join stylusformat.xlsx portfolio
	* Adds data from in.xlsx to existing portfolio data in stylusformat.xlsx

Attached files are as follows:
     - stylusformat.xlsx > Advanced: a Stylus-formatted Advanced Portfolio
     - in.xlsx > in: Additional data to be added to this portfolio. This
       contains new assets as well as an added date for an existing asset.
     - out.xlsx > out: Blank worksheet for output.

---------
- TO DO -
---------

TODO: ADD OVERWRITE SHEET FUNCTIONALITY
TODO: Make the main function way more Pythonic
TODO: Convert to named variables where it would help readability
