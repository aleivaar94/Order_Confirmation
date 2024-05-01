# Imports and configurations
# Data handling
import pandas as pd
import numpy as np
import re
from datetime import datetime
# Google Cloud
import gspread
from google.oauth2.service_account import Credentials
# Environment variables
import os
from dotenv import load_dotenv
# Decode google sheets API credentials
import json
import base64
# Email
import smtplib
from email.message import EmailMessage
from email.utils import formataddr

# Load environment variables
load_dotenv()

# Decode the base64-encoded credentials
encoded_creds = os.getenv('ENCODED_CREDENTIALS')
creds_json = json.loads(base64.b64decode(encoded_creds))

# Google Sheets Authentication
scopes = ['https://www.googleapis.com/auth/spreadsheets']
creds = Credentials.from_service_account_info(creds_json, scopes=scopes)
client = gspread.authorize(creds)

# Google Sheets Access
orders_sheet_id = os.getenv('ORDERS_SHEET_ID')
orders_sheet = client.open_by_key(orders_sheet_id)
orders_worksheet = orders_sheet.sheet1
discounts_sheet_id = os.getenv('DISCOUNTS_SHEET_ID')
discounts_sheet = client.open_by_key(discounts_sheet_id)
discounts_worksheet = discounts_sheet.sheet1


def update_order_numbers(orders_worksheet):
    submitted_on_list = orders_worksheet.col_values(1)  # Column A: 'submitted on'
    order_number_list = orders_worksheet.col_values(10)  # Column J: 'order number'

    missing_orders_count = len(submitted_on_list) - len(order_number_list)
    print(f"Missing orders: {missing_orders_count}")

    next_row = len(order_number_list) + 1 # Set value for make_order_table function

    if missing_orders_count > 0:
        next_order_number = '00001' if not order_number_list or order_number_list[-1][-1].isalpha() else str(int(order_number_list[-1]) + 1).zfill(5)
        
        for _ in range(missing_orders_count):
            next_row = len(order_number_list) + 1
            update_order_cell = f'J{next_row}'
            orders_worksheet.update_acell(update_order_cell, next_order_number)
            print(f"Updated order number {next_order_number} at {update_order_cell}")
            order_number_list.append(next_order_number)
            next_order_number = str(int(next_order_number) + 1).zfill(5)

            order_html = make_order_table(orders_worksheet, discounts_worksheet, next_row) # Call make_order_table function to create the order table for the new order
            send_order_emails(orders_worksheet, order_html, sender_email, password, server_address, port) # Call send_order_emails function to send the email for the new order
    else:
        print("No new orders. No update necessary.")
        return None # Return None to skip make_order_table and send_order_emails functions.
    return next_row # Return the next row for make_order_table function


def make_order_table(orders_worksheet, discounts_worksheet, next_row):
    anxiety_tea_column = 'D'
    address_column = 'G'
    next_row = next_row
    data = orders_worksheet.get(f'{anxiety_tea_column}{next_row}:{address_column}{next_row}') # Get data up to address column to get promo code column values, even if they are null
    flat_list = [item for sublist in data for item in sublist]
    flat_list = flat_list[:3]  # Get anxiety, immune, and promo code columns

    # Initialize lists to store the extracted quantities and prices
    quantities = []
    prices = []

    # Extract quantities and prices for Anxiety Reset and Immune Harmony
    for item in flat_list[:2]:
        if item:  # Check if the string is not empty
            quantity_search = re.search(r'^(\d+)', item)
            price_search = re.search(r'\$(\d+)', item)
            if quantity_search and price_search: # If there's a regex match for both quantity and price append to the lists
                quantity = quantity_search.group(1)
                price = price_search.group(1)
                quantities.append(int(quantity))
                prices.append(int(price))
            else:
                # Append NaN where regex search fails to match
                quantities.append(np.nan)
                prices.append(np.nan)
        else:
            # Append NaN for empty strings
            quantities.append(np.nan)  # or np.nan
            prices.append(np.nan)  # or np.nan

    # Create DataFrame with product names as index
    order_df = pd.DataFrame({
        'Quantity': quantities,
        'Price': prices
    }, index=["Anxiety Reset", "Immune Harmony"])

    # Check for promo code and apply discount if applicable
    promo_code = flat_list[2].strip().lower() # Get promo code submitted
    discounts_data = discounts_worksheet.get_all_records()
    for record in discounts_data:
        if record['Discount Code'].strip().lower() == promo_code and record['Status'].strip().lower() == 'active':
            discount_factor = 1 - float(record['% Discount']) / 100
            order_df['Price'] = order_df['Price'].apply(lambda x: x * discount_factor)
            break

    order_df.loc['Shipping'] = [np.nan, 15] # Add a row for Shipping

    # Calculate totals
    total_quantity = order_df['Quantity'].sum(min_count=1)  # min_count=1 ensures NaN is ignored without resulting in NaN
    total_price = order_df['Price'].sum(min_count=1)

    order_df.loc['Total'] = [total_quantity, f'${total_price:.2f}']  # Add a row for totals

    # Convert dataframe to HTML
    order_html = order_df.to_html(index=True, na_rep='') # na_rep='' replaces NaN with an empty string

    return order_html


def send_order_emails(orders_worksheet, order_html, sender_email, password, server_address, port):
    # Records start at row 2, as the first row contains the headers
    rows = orders_worksheet.get_all_records()
    for index, row in enumerate(rows, start=2):  # Adjust the start index if necessary
        if row["order number"] and not row['send email']:  # Check if the order has a number and email not sent
            # Obtain customer and order details
            order_date = datetime.strptime(row['submitted on'], '%m/%d/%Y %H:%M:%S').strftime("%m/%d/%Y")
            recipient = row["email"]
            customer_name = row["name"]
            order_number = "{:0>5}".format(row["order number"])  # Zero-pad the order number

            if row['promo code']:
                promo_code = f"Promo code: {row['promo code']}"
            else:
                promo_code = ''
            
            # Prepare email content
            msg = EmailMessage()
            msg['Subject'] = 'Order Confirmation'
            msg['From'] = formataddr(('Small Business', sender_email))

            msg.add_alternative(
                f"""\
                <html>
                    <body>
                        <p>Hi {customer_name},</p>
                        <p>Thank you for supporting a small business. We have confirmed your order number <b>{order_number}</b> on <b>{order_date}</b>.  Below are the details of your purchase:</p>
                        {order_html}
                        <p>{promo_code}</p>
                        <p>Once we confirm your payment, we will process your order and send you your tracking number in less than 24hrs.</p>
                        <p><b>Payment steps:</b></p>
                        <ol>
                            <li>Send the full amount via Interac e-transfer to <b><a href="mailto:confirmation@smallbusiness.ca">confirmation@smallbusiness.ca</a></b></li>
                            <li>Use the name <b>Small Business</b> as the payee name in your e-transfer setup.</li>
                            <li>In the e-transfer message include your <b>order number</b>.</li>
                            <li>This email is registered for auto-deposit. This means you shouldn’t need to add in a secret question and answer. If your bank requires you to have a secret question and answer please use the secret question: <b>what do you support?</b> and secret answer: <b>small business</b></li>
                        </ol>
                        <p>If you have any questions or concerns about your order, please don't hesitate to reply to this email and we will get back to you as soon as possible.</p>
                        <p>--</p>
                        <p><b>Small Business Team</b></p>
                        <p><a href='www.smallbusiness.ca'>www.smallbusiness.ca</a></p>
                        <p><a href='www.instagram.com'>@smallbusiness</a></p>
                    </body>
                </html>
                """,
                subtype='html'
            )
            
            # Send the email
            try:
                with smtplib.SMTP_SSL(server_address, port) as server:
                    server.login(sender_email, password)  # Log in to the server
                    recipients = [verification_email, recipient]
                    server.sendmail(sender_email, recipients, msg.as_string())  # Send the email
                    orders_worksheet.update_cell(index, 11, 'TRUE')  # Update column (column K), the "send email" column
                    print(f"Email sent successfully to {recipient}, for order number {order_number}. 'Send email' column updated to TRUE for row {index}")
            except Exception as e:
                print(f"Failed to send email to {recipient} for {order_number}: {e}")
        
        else:
            continue # Skip the row if the order number is missing or email already sent

def action_notification_email():
    # Send a notification email every time GitHub action is executed
    msg = EmailMessage()
    msg['Subject'] = 'Order Confirmation GitHub Action Notification'
    msg['From'] = formataddr(('Small Business', sender_email))
    msg['To'] = '' # Insert email here

    msg.set_content(
        """\
        Hi Alejandro,
        The script execution has finished.
        """
    )

    try:
        with smtplib.SMTP_SSL(server_address, port) as server:
            server.login(sender_email, password)  # Log in to the server
            server.send_message(msg)  # Send the email
        print("GitHub action notification email sent to email.")
    except Exception as e:
        print(f"Failed to send GitHub action notification email to email: {e}")
    return

if __name__ == "__main__":
    # Email credentials
    sender_email = os.getenv('EMAIL_ADDRESS')
    password = os.getenv('EMAIL_PASSWORD')
    server_address = os.getenv('EMAIL_SERVER')
    port = os.getenv('PORT')
    verification_email = os.getenv('VERIFICATION_EMAIL')

    next_row = update_order_numbers(orders_worksheet)
    action_notification_email() # Send email notification for GitHub action
    if next_row is not None:
        order_html = make_order_table(orders_worksheet, discounts_worksheet, next_row)
        send_order_emails(orders_worksheet, order_html, sender_email, password, server_address, port)
    else:
        print("No new orders. No email sent.")