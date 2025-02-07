# Decommission Checklist Automation (WITH SENSITIVE DATA REPLACED BY "---")

This project automates the process of filling out a decommission checklist using Selenium and Excel. The tool reads asset names and their statuses from an Excel file and interacts with a web application (in Power Apps) to tick the appropriate checkboxes. It also marks the processed asset names in green in the Excel file.

## Features

- Automatically fills out the decommission checklist based on asset names and statuses from an Excel file.
- Handles cases where the checklist is already ticked or unavailable.
- Marks processed asset names in green in the Excel file.
- Runs in headless mode, allowing the user to work on other tasks simultaneously.
- Logs the execution time of the automation process.

## Prerequisites

- Python 3.12
- Microsoft Edge browser
- Edge WebDriver (compatible with your Edge browser version)
