# Outlook Follow-Up Automation

This project contains a Python script (`outlook_followup.py`) that automates sending follow-up emails via Microsoft Outlook based on a list of contacts in an Excel file.

## Features

- Reads contacts (Name, Email ID) from an Excel file (`contacts.xlsx`).
- Connects to the local Microsoft Outlook application via COM interface (`win32com`).
- Searches the "Sent Items" folder for the *first* (oldest) email sent to each contact.
- Uses "Reply All" to prepend a customized follow-up message addressing the person by their *first name*.
- Replaces the old email subject line with a custom new subject line.
- Enforces a 30-second delay between consecutive emails to prevent rate limiting or spam flags.
- Ensures existing email HTML formatting is preserved by injecting the follow-up text directly after the `<body>` tag.

## Prerequisites

- **OS**: Windows (requires `win32com` to interact with the local Outlook application).
- **Microsoft Outlook**: Must be installed, configured, and running on the machine.
- **Python**: Version 3.6 or higher.

## Setup Instructions

1.  **Clone the repository** (or download the files).
2.  **Install dependencies**:
    ```bash
    pip install -r requirements.txt
    ```
    This will install `pandas`, `openpyxl`, and `pywin32`.

## Usage

1.  **Prepare the Excel file**:
    The script expects an Excel file named `contacts.xlsx` in the same directory, containing at least two columns: `Name` and `Email ID`.

    *Optional*: You can run the included helper script to generate a sample file:
    ```bash
    python generate_sample_excel.py
    ```

2.  **Run the script**:
    Ensure Microsoft Outlook is open and authenticated. You can edit the `new_subject_line` variable at the bottom of `outlook_followup.py` to set your desired subject. Then run:
    ```bash
    python outlook_followup.py
    ```

3.  **Process**:
    The script will output its progress to the console. For each contact, it will:
    - Search for the first sent email to their address.
    - If found, create a "Reply All" response.
    - Insert a greeting like "Hi [First Name]," and a brief follow-up message.
    - Send the email.
    - Wait 30 seconds before processing the next contact in the list.

## How it Works (`outlook_followup.py`)

1.  **Excel Data Loading**: It uses `pandas.read_excel()` to load the list of contacts.
2.  **Outlook Connection**: It uses `win32com.client.Dispatch("Outlook.Application")` to hook into the running Outlook instance and accesses the MAPI namespace to get the "Sent Items" folder (DefaultFolder 5).
3.  **Search Optimization**: It uses the `.Restrict()` method with an `@SQL` query to efficiently filter the Sent Items for emails matching the target email address, rather than iterating through the entire folder.
4.  **Sorting**: It sorts the filtered results by `[SentOn]` in ascending order (`False`) to grab the oldest/first email sent.
5.  **HTML Injection**: To preserve formatting, it uses a regular expression to locate the `<body>` tag in the original email's HTML body (`reply.HTMLBody`) and safely inserts the new follow-up text immediately after it.
