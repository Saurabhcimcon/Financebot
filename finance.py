import pandas as pd
import smtplib
from email.message import EmailMessage
import os
import tempfile
import streamlit as st

def finance_highlighter():
    uploaded_file = st.file_uploader("Upload CSV file", type=['csv'])
    email = st.text_input("Enter your email address:")

    if st.button("Send Email") and uploaded_file is not None:
        if email is None or email.strip() == "":
            st.error("Please provide an email address.")
        else:
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

            # Get keys with both 'Cr' and 'Dr' values
            keys_cr_and_dr = [key for key, values in party_pending_dict.items() if
                              any(value.endswith(' Cr') for value in values if pd.notna(value)) and
                              any(value.endswith(' Dr') for value in values if pd.notna(value))]

            # Function to highlight rows with keys_cr_and_dr in yellow
            def highlight_cr_and_dr(row):
                return ['background-color: yellow' if row["Party's Name"] in keys_cr_and_dr else '' for _ in row]
            df['Pending'] = df['Pending'].astype(str).apply(lambda x: f"+{x.replace(' Cr', '')}" if "Cr" in x else x)
            df['Pending'] = df['Pending'].astype(str).apply(lambda x: f"-{x.replace(' Dr', '')}" if "Dr" in x else x)

            # Apply highlight function to DataFrame
            styled_df = df.style.apply(highlight_cr_and_dr, axis=1)

            # Function to highlight all rows associated with a party name where 'TDS' is present
            def highlight_tds_party(df):
                styled_df = pd.DataFrame('', index=df.index, columns=df.columns)
                for index, row in df.iterrows():
                    party_name = row["Party's Name"]
                    if 'TDS' in str(row).upper():
                        styled_df.loc[df["Party's Name"] == party_name] = 'background-color: red'
                return styled_df

            # Apply red highlighting to entire rows associated with the party name where 'TDS' is present
            styled_df = styled_df.apply(highlight_tds_party, axis=None)

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
    msg['Subject'] = 'Highlighted Finance Data'
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg.set_content('Please find the highlighted finance data attached.')

    # Write the styled DataFrame to an Excel file in a temporary directory
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as temp_file:
        styled_df.to_excel(temp_file.name, index=False)

        # Attach the Excel file to the email
        with open(temp_file.name, 'rb') as f:
            file_data = f.read()

    # Remove the temporary Excel file
    os.unlink(temp_file.name)

    msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename='Finance_Highlighted.xlsx')

    # Send the email
    with smtplib.SMTP(smtp_server, port) as server:
        server.starttls()
        server.login(sender_email, password)
        server.send_message(msg)
