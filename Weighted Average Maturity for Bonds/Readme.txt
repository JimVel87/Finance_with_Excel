This short project aimed to create an Excel formula to calculate the Weighted Average Maturity (WAM) of a bond, measured in years.

This solution avoids the use of VBA macros or the creation of a cashflow vector in other cells/sheets.

It assumes that the worksheet will have the first 7 columns (in blue) populated in the form of a portfolio consisting of bonds. The data in these columns are hardcoded to simulate input from a system, database or another file.

In order for the formula to work, the worksheet needs to include 3 columns (in pink) with information about the bond. The values in these cells are generated through formulas, which can either be already there and just be refreshed, or they can be populated through a macro. Of course, they could also come from the system, database or original file together with the columns in blue, but this solution assumes access to the minimum information possible.

The solution formula itself is in the last column (in yellow). It is an array formula, so it needs to be entered with CTRL+SHIFT+ENTER in order to work. The only hardcoded value in it is the assumption of 365 days in a year, which can easily be changed.