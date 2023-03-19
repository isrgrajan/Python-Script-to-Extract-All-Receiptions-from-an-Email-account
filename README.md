# Extracting Email Receptions using IMAP

This Python script allows you to extract all email receptions in a specified mailbox and write the sender's email address and name to an Excel file. By using the Extracting Email Receptions using IMAP, it is possible to extract all email recipients from various email providers such as Gmail, Outlook, and Rediff. This API is specifically designed for extracting contacts, which can then be used to send newsletters or greetings to either customers or friends and family.

IMAP stands for Internet Message Access Protocol, which is a widely used protocol for email retrieval. It enables users to access and manipulate their email messages on a remote server, without downloading them to their local device. With IMAP, users can easily extract email recipients from their email accounts without compromising the security of their accounts.

This API can be a valuable tool for businesses and individuals who need to send mass emails or newsletters to their customers, subscribers, or acquaintances. By extracting the contact list, users can send personalised messages to their recipients, making the communication more effective and engaging. Additionally, this API can save users time and effort, as it eliminates the need to manually extract email addresses from their email accounts.

## Installation
1. Clone the repository to your local machine: `git clone https://github.com/isrgrajan/imap-email-receptions.git`
2. Install the required libraries: `pip install imaplib openpyxl`

## Usage
1. Open the `email_receptions.py` file in a text editor of your choice.
2. Replace the following variables with the correct values for your email account:
```
imap_host = 'imap.example.com'
username = 'user@example.com'
password = 'your_password'
```
3. Run the script using the following command:
```
python email_receptions.py
```
4. The script will create an Excel file named `email_receptions.xlsx` in the same directory as the script.

## Contributing
If you find any bugs or have suggestions for improvement, please feel free to open an issue or submit a pull request.

## License
This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details.




