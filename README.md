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

### Setting Up Google Service Account for API Access

#### 1. Create a New Project in Google Cloud Console

- Go to the [Google Cloud Console](https://console.cloud.google.com/).
- Click on "Select a project" at the top of the page and then click on "New Project".
- Enter a name for your project and click on "Create".

#### 2. Enable the Required API
- In the Google Cloud Console, navigate to "APIs & Services" > "Library".
- Search Google Sheets API and click on it.
- Click the "Enable" button to enable the API for your project.

#### 3. Create a Service Account
- In the Google Cloud Console, navigate to "APIs & Services" > "Credentials".
- Click on "Create credentials" and select "Service account".
- Enter a name and description for your service account.
- Click on "Create" to create the service account.
- Assign a Editor role to your service account to grant it the necessary permissions.

## 4. Generate and Download the Service Account Key (JSON)
- Find the newly created service account in the list and click on the three-dot menu icon.
- Select "Manage keys" > "Add key" > "Create new key".
- Choose the key type as JSON and click on "Create". This will download a JSON file containing your service account credentials.
- Add credentials to *cred* `folder and save as *service_acct_cred.json*

### Input data
- Open *input.txt* and input data
- Input the BFMR-ID and Total Claimed for each deals into inner list
- **[[BFMR-ID, Total Claimed], [BFMR-ID , Total Claimed], [BFMR-ID , Total Claimed]]**


### Run the Script
You're now ready to run the script to update your Google Spreadsheet. Ensure your virtual environment is activated, then execute:
``` bash
python sheet_update.py
```

This will run the script, which will authenticate using the credentials file and update the specified Google Spreadsheet.
