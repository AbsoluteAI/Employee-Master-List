# Employee-Master-List
This is a macro-enabled Excel workbook designed to streamline employee tracking and earnings management. The project combines VBA macros and built-in Excel formulas to provide an automated and interactive solution for monitoring employee details and financial data. 

## Key Features
* Automated Earnings Calculation: Calculates each employee’s earnings based on hours worked, pay rates, and total weeks, updating dynamically as data changes.
* Comprehensive Employee Tracking: Maintains a detailed log of employee information, including hours, pay, and allocations.
Projection Calculator: Allows users to estimate allocations and earnings by factoring in the number of employees, hours per week, total weeks, and pay rates.
* Interactive Scenario Planning: Users can modify hours per week to see updated projections, helping managers understand employee contributions and schedule impacts over time.
* Dynamic Summaries & Reporting: Summarizes key metrics across all employees, enabling quick insight into payroll costs and workforce allocation.
* VBA Automation: Reduces manual effort by automatically updating calculations, generating summaries, and maintaining consistency across large datasets.

This tool is ideal for HR professionals, payroll administrators, and managers who need a flexible and interactive way to track employee performance, calculate earnings, and plan workforce allocations efficiently.

## Getting Started

### Dependencies
* Windows 11
* Microsoft Excel 2016

### Downloading

* To download the macro-enabled Excel workbook follow the guide below:
  * [Github Docs](https://docs.github.com/en/get-started/start-your-journey/downloading-files-from-github) 
* Using a macro-enabled workbook requires a user to trust the program. The steps for this can be found at the following link:
  * [Microsoft Support](https://support.microsoft.com/en-us/office/enable-or-disable-macros-in-microsoft-365-files-12b036fd-d140-4e74-b45e-16fed1a7e5c6)

### Executing program

* There are a few important details to note regarding the contents of this workbook:
  * This workbook contains one macro and will require trust in order to run properly
  * There is sample data included that is for demonstration purposes only. This data can be removed before use.
  * The first tab is the "Employee Profile Template". This tab can be duplicated for each new employee between the "First" and "End" tabs. These tabs exist to mark the beginning and end of the employee profile section of the workbook.
  * The "Summary" tab provides a high-level overview of the entire year and is broken down into quarters. The "Update" macro button allows all values from each individual employee profile to be summarized for each corresponding cell in the "Earnings Summary" table.
  * The next tab is the "Projection Calculator" sheet and it contains several formulas and references that are combined to create the table trio.
    * Columns A and B are opened for viewing purposes but are not intended to be changed as they already reference the required cells and tables.
    * The "Reference Stats" section shows specific data for use when using the calculation tools beside it.
  * The "Employee List" tab contains a table with basic employee data. This table is referenced elsehwere in the program for calculations.
  * The "Department List" tab is a related table with only department information.
  * The remaining tabs were created by copying the "Employee Profile Template" and entering specific information such as "Name", "Department", "Employee ID", "Grant Type", and "Pay Rate". The only other required information to be entered for tracking are the hours worked per week. The earnings and all other parts of the workbook will automatically update based on those hours.

## Help
* To submit any issues please follow the steps from the guide below and submit them to my issues section.
  * [Submit Issues](https://github.com/AbsoluteAI/Employee-Master-List/issues)
  * [Github Isses Instructions](https://docs.github.com/en/issues/tracking-your-work-with-issues/learning-about-issues/about-issues)

## Authors

Alexander Dodd
[@linkedin.com](www.linkedin.com/in/alexander-j-dodd93)

## Version History

* 0.1
  * Initial Release
    * See [commit change](https://github.com/AbsoluteAI/Employee-Master-List/commits/main/) or See [release history](https://github.com/AbsoluteAI/Employee-Master-List/releases) 

## License

The MIT License (MIT)

Copyright (c) 2015 Chris Kibble

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
