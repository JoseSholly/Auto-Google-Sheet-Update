# Google Spreadsheet Automation with Python

This Python project demonstrates how to automate updates to a Google Spreadsheet using the Google Sheets API. The project utilizes a virtual environment (venv) for managing dependencies and packages listed in `requirements.txt`.

## Installation

### 1. Clone the Repository

Clone this repository to your local machine using Git:

```bash
git clone https://github.com/yourusername/google-sheets-automation.git
cd google-sheets-automation
```
### Set Up Virtual Environment (venv)
It's recommended to use a virtual environment to manage project dependencies. Here's how to set it up:

#### For macOS/Linux

``` bash
python3 -m venv venv
source venv/bin/activate
```

#### For Windows

``` bash
python -m venv venv
venv\Scripts\activate
```
### Install Dependencies
With the virtual environment activated, install the required packages from requirements.txt:

``` bash 
pip install -r requirements.txt
```

### Set Up Google Sheets API
To use the Google Sheets API, follow these steps:

- Go to the Google API Console.
- Create a new project.
- Enable the Google Sheets API for your project.
- Create credentials (OAuth client ID) for a desktop application.
- Download the credentials JSON file and save it in the project directory.
- Also, set up your SMTP server by adding your user email address and  app password
##
### Configure Spreadsheet ID
Open sheet_update.py and replace 'GOOGLE_APPLICATION_CREDENTIALS' with the ID of the Google Spreadsheet you want to update.

### Run the Script
You're now ready to run the script to update your Google Spreadsheet. Ensure your virtual environment is activated, then execute:
``` bash
python sheet_update.py
```

This will run the script, which will authenticate using the credentials file and update the specified Google Spreadsheet.
