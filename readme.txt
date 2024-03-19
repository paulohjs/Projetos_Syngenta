service_now_download_report.py

DESCRIPTION:
This script automates the downloading of reports from a ServiceNow instance. It includes functions for browser automation, Microsoft account login, iframe navigation, report information capture, and file management.

REQUIREMENTS:
- Python 3.x
- Selenium WebDriver
- ChromeDriver
- Tkinter (for warning and error dialogs)
- webdriver_manager package

INSTALLATION:
To set up the necessary environment, run the `download_libraries.py` script included in the same directory as the main script. This script will check for the required Python packages and install them if they are not already installed.

To run the installation script, open a terminal or command prompt, navigate to the script's directory, and execute the following command:
python download_libraries.py

USAGE:
Ensure you have a user.json file in the same directory as the script with the following structure:

{
  "Users": [
    {
      "email": "user@example.com",
      "password": "yourpassword"
    }
    // Add more users if necessary
  ]
}

Run the script by importing it into your Python code and calling the dowload_report_sn function with the appropriate arguments:

from sndownload import dowload_report_sn

# Example usage
link = 'https://service-now-instance-url.com/report'
destination_path = 'C:/path/to/destination/folder'
email = 'user@example.com'
password = 'yourpassword'

dowload_report_sn(link, destination_path, email, password)
