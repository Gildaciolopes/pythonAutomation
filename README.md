# Sales Report

This project generates a sales report from an Excel file and sends an email with the results.
Is a reporting automation in python.

## Technologies Used

- Python 3.13.2
- Pandas (library for data manipulation and analysis)
- win32com.client (for sending emails via Outlook)

## Features

- Loads sales data from a `Vendas.xlsx` file
- Calculates:
  - Revenue per store
  - Quantity of products sold per store
  - Average ticket per product in each store
- Generates a formatted email with the results
- Sends the email automatically

### Requirements

Make sure you have installed:

- Python (3.13.2 version)
- Pandas (`pip install pandas`)
- pywin32 (`pip install pywin32`)
- Microsoft Outlook configured on your computer

### Execution

1. Place the `Vendas.xlsx` file in the same directory as the script.
2. Run the Python script:
   ```bash
   python script.py
   ```
3. The email will be sent automatically to `contato.gildaciolopes@gmail.com`.

## Notes

- Ensure that Outlook is correctly configured for sending emails.
- The script can be adapted to send emails to different recipients by modifying the line:
  ```python
  mail.To = 'email@example.com'
  ```
- If necessary, adjust the currency formats and email layout as preferred.

## Author

Gildácio Lopes
