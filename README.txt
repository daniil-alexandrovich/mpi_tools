Uses pandas and openpyxl libraries.

Run portfolio_updater.py to test script.

This tool has the capacity to generate and manipulate a portfolio as well as
write it out to an Excel Sheet, properly formatted. The test script writes two
portfolios to sheets in out.xlsx:

	* join: Joins unformatted data to a previously existing Stylus portfolio
	* format: Formats unformatted data as a standalone Stylus portfolio

Eventually, this tool may be used to create and edit portfolios from scratch.

Attached files are as follows:
     - stylusformat.xlsx > Advanced: a Stylus-formatted Advanced Portfolio
     - in.xlsx > in: Additional data to be added to this portfolio. This
       contains new assets as well as an added date for an existing asset.
     - out.xlsx > out: Blank worksheet for output.

