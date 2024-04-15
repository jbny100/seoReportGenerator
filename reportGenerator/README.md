## SEO Report Generator 

SEO Report Generator is a Python-based automation tool designed to generate SEO performance reports from Google Search Console. It uses Selenium to navigate the web interface and OpenPyXL to handle Excel file operations, compiling data for indexed pages, 404 errors, top queries, top page searches and total clicks into an Excel spreadsheet with numerous tabs. 

## Features

- Automated login to Google Search Console
- Data extraction for indexed pages, 404 URLs, and performance metrics
- Report generation in Excel format
- Detailed logging of operations and error handling

## Prerequisites 

To run this script, you need:

- Python 3.8 or higher
- selenium: For automating web browser interaction
- openpyxl: For creating and updating Excel files 
- logging: For logging events and errors to a file 
- A Google Chrome browser
- ChromeDriver (make sure it's compatible with your Chrome version)

## Setup

1. Clone the repository to your local machine:
2. Install the required Python packages: pip install -r requirements.txt

## Configuration

Before running the script, you need to configure the following variables in the script:

- Username and password to enter in login() 
- Base Path: Set the base path directory where Excel reports will be saved.
- target_url in navigate_to_console() method in BrowserNavigator class.
- download_path: In __init__() method of BrowserNavigator. Download path for helper files.
- filename_prefix: The filenames of these downloaded spreadsheets before '.xlsx.'
- chrome_driver_path: The path to your ChromeDriver executable.

## Running the program

- The script can be run from terminal or directly from a text editor. 
- The program pauses for user to enter Google verification code and CAPTCHA
- Certain CSS selectors may differ for the Google Search Console profile for your website.

Output:

The script will create an Excel workbook called 'Monthly_SEO_Metrics.xlsx' in the base path directory. The workbook will contain basic performance SEO metrics categorized into worksheets that is scraped from the Google Search Console profile for your website.

Two additional workboks, called 'Indxed_Pages.xlsx' and 'Total_Clicks.xlsx', will also be created in the same base path directory. These files are updated each time the program runs and then copied to the sheet 'Indexed Pages' and 'Total Clicks Last 3 Months' in Monthly_SEO_Metrics.xlsx.

Please note that your website must have a Google Search Console profile in order to run this program. As scraping can violate Google's terms of service, please use this script responsibly and only scrape data that you are authorized to access.

## Logging

This script uses Python's built-in `logging` module to provide status updates and debugging information throughout the program. These logs can help you understand what the script is doing, and they can be very useful for troubleshooting when something goes wrong.

The level of detail in the logs is controlled by the log level, which is set to `INFO` by default. This means that informational messages, warnings, and errors will be logged, but more detailed debug messages will not. You can change the log level to `DEBUG` if you need more detailed logs for troubleshooting.

Here's what the different log levels mean:

- `DEBUG`: Detailed information, typically useful only when diagnosing problems. This level includes everything.
- `INFO`: Confirmation that things are working as expected. This is the default log level.
- `WARNING`: An indication that something unexpected happened, or there may be some issue in the near future (e.g., 'disk space low'). The software is still working as expected.
- `ERROR`: More serious problem that prevented the software from performing some function.
- `CRITICAL`: A very serious error that may prevent the program from continuing to run.

The log messages will be printed to a log created in the base path called 'seo_report_generator.log.' Look for these messages to understand what the script is doing and to identify any problems.

Example of a log message:

´´´bash
INFO:root:Successfully logged in with username: your_username

## Possible Improvements

- Optimizing Workbook Handling 
	
	Creating a method that ensures any workbook changes are temporarily held until the final save, minimizing read/write operations on the disk. 

- Centralized Saving and Cleanup

	Use of a centralized save method (in run()) that not only saves the workbook but also cleans it up (removes unwanted sheets) just before finalizing. This avoids having to save Monthly_SEO_Metrics.xlsx at the end of individual methods. 

## Contributing

Contributions to the SEO Report Generator are welcome. Please fork the repository and submit a pull request with your enhancements.





