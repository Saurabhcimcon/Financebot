import pandas as pd
import streamlit as st
import smtplib
from email.message import EmailMessage
import os
import tempfile

def purchase_highlighter():
    st.write("This is the Purchase Page")

    # File uploader for user to upload CSV file
    uploaded_file = st.file_uploader("Upload CSV file", type=['csv'])

    # Input field for user to enter email address
    email = st.text_input("Enter your email address:")

    # Button to send email
    if st.button("Send Email") and uploaded_file is not None:
        if email is None or email.strip() == "":
            st.error("Please provide an email address.")
        else:
            # Read the CSV file
            df = pd.read_csv(uploaded_file, encoding='utf-8')

            # Extract column names from the second row
            column_names = df.iloc[1]

            # Set the column names
            df.columns = column_names

            # Drop the first two rows and reset index
            df = df.drop([0, 1]).reset_index(drop=True)

            # Forward fill 'Party's Name' column to fill NaN values
            df["Party's Name"] = df["Party's Name"].fillna(method='ffill')

            # Drop any remaining rows with index 0
            df = df.drop(0)

            # Create a dictionary to store party names and pending values
            party_pending_dict = {}
            for index, row in df.iterrows():
                party_name = row["Party's Name"]
                pending_value = row["Pending"]
                if party_name in party_pending_dict:
                    party_pending_dict[party_name].append(pending_value)
                else:
                    party_pending_dict[party_name] = [pending_value]

            # Get keys with only 'Dr' values
            keys_dr_only = [key for key, values in party_pending_dict.items() if
                            all(value.endswith(' Dr') for value in values if pd.notna(value))]

            # Function to highlight rows with keys_dr_only in yellow
            def highlight_dr_only(row):
                return ['background-color: yellow' if row["Party's Name"] in keys_dr_only else '' for _ in row]

            # Apply highlight function to DataFrame
            styled_df = df.style.apply(highlight_dr_only, axis=1)

            # Sending email with highlighted dataframe attached
            send_email(email, styled_df)


def send_email(receiver_email, styled_df):
    # SMTP configuration for your organization's mail server
    smtp_server = 'smtp.office365.com'
    port = 587  # For starttls
    sender_email = 'billing@cimconautomation.com'  # Enter your email address
    password = 'Boq61126'  # Enter your email password

    # Create EmailMessage object
    msg = EmailMessage()
    msg['Subject'] = 'Highlighted Purchase Data'
    msg['From'] = sender_email
    msg['To'] = receiver_email
    
    msg.set_content('Please find the highlighted purchase data attached.')

    # Write the styled DataFrame to an Excel file in a temporary directory
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as temp_file:
        styled_df.to_excel(temp_file.name, index=False)

        # Attach the Excel file to the email
        with open(temp_file.name, 'rb') as f:
            file_data = f.read()

    # Remove the temporary Excel file
    os.unlink(temp_file.name)

    msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename='Purchase_Highlighted.xlsx')

    # Send the email
    with smtplib.SMTP(smtp_server, port) as server:
        server.starttls()
        server.login(sender_email, password)
        server.send_message(msg)
