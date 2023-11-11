# If issues with imported libraries you may need to install a few things
# !pip install pandas
# !pip install openpxyl
# !pip install fsspec

# Import Libraries
import pandas as pd
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText


# Load the CSV file from my device into pandas.DataFrame
df = pd.read_csv("C://Desktop/Soccer.csv")

# Rename header columns from csv file to match requirements
df.columns = ["ID", "Player", "Age", "Nationality", "Overall", "Money", "Position", "H", "W"]
# Was really tempted to rename the header files with clear information for how height and weight are formated
# df.columns = ["ID", "Player", "Age", "Nationality", "Overall", "Money", "Position", "H (ft.inches)", "W (lbs)"]

# Clean data/remove special characters from Money, H, and W
df["Money"] = df["Money"].str.replace("€", "").str.replace("K", "000")      # Replace the 'Euro' symbol with a space and the 'K' value for 'thousands' with '000'
df["H"] = df["H"].str.replace("'", ".")                                     # Remove the 'comma' from Height and replace with period to represent a float
df["W"] = df["W"].str.replace("lbs", "")                                    # Remove the 'lbs' from Weight and only leave the number behind

# Convert/Ensure ID, Player, Age, Nationality, Overall, Position, Height, and Weight are proper data types
df["ID"] = df["ID"].astype(int)                                         # Convert/Ensure ID column is 'int' type
df["Player"] = df["Player"].astype(str)                                 # Convert/Ensure Player column is 'string' type
df["Age"] = df["Age"].astype(int)                                       # Convert/Ensure Age column is 'int' type
df["Nationality"] = df["Nationality"].astype(str)                       # Convert/Ensure Nationality column is 'string' type
df["Overall"] = df["Overall"].astype(int)                               # Convert/Ensure Overall column is 'int' type
df["Money"] = df["Money"].astype(int)                                   # Convert/Ensure Money column is 'int' type
df["Position"] = df["Position"].astype(str)                             # Convert/Ensure Position column is 'string' type
df["H"] = df["H"].astype(float)                                         # Convert/Ensure Height column is 'float' type
df["W"] = df["W"].astype(float)                                         # Convert/Ensure Weight column is 'float' type
# If we were using my header files
# df["H (ft.inches)"] = df["H (ft.inches)"].str.replace("'", ".").astype(float)
# df["W (lbs)"] = df["W (lbs)"].str.replace("lbs", "").astype(float)

# Export dataframe by converting to excel file 'By Country' and store on my device
output_file_path = "C://Desktop/By Country.xlsx"
df.to_excel(output_file_path, index = False)


# Send Excel file via email
from_email = "Enter the senders email"                                    # Variable to hold the sender email address
from_password = "Enter password for sender email"                         # Variable to hold password for sender email address
to_email = "Enter the recievers email"                                    # Variable to hold the reciever email address
subject_email = "Soccer Data"                                             # Variable to hold the subject of the email
body = "The report you requested earlier is ready for you to review."     # Variable to hold the body of the email

# Create the email content using the variables above
msg = MIMEMultipart()                                            # Use MIMEMultipart function to build email message as msg
msg["From"] = from_email                                         # Message is being sent from the email entered in variable above
msg["To"] = to_email                                             # Message is being sent to the email specified in variable above
msg["Subject"] = subject_email                                   # Subject of email clarified in the variable above
body = MIMEText(body)                                            # Convert the body of the email entered above to a MIME compatible string
msg.attach(body)                                                 # Attach the body to the main message for email

# Attach the Excel file to the message for email
with open(output_file_path, "rb") as file:                                                   # Open as read only in binary format and then build attachment as xlsx
    attach = MIMEApplication(file.read(), _subtype = "xlsx")                                 # Convert the file to be MIME compatible and include the type 'xlsx'
    attach.add_header("Soccer Content", "attachment", filename = "By Country.xlsx")          # Build header for attachment in email
    msg.attach(attach)                                                                       # Add the built attachment to the main msg for the email

# Use SMTP to connect to email server and send the email
try:                                                                    # Use try/catch block for exception handling for success and failure
    server = smtplib.SMTP("smtp.gmail.com", 587)                        # Use function to make connection with gmail server using gmail port number, SSL Port number is 465
    server.starttls()                                                   # Create secure transport layer 'tunnel transport layer security'
    server.login(from_email, from_password)                             # Login in to my gmail account with email and password
    server.sendmail(from_email, to_email, msg.as_string())              # Send the info to server with the msg to send, which email to send it to and which email is doing the sending
    server.quit()                                                       # Close connection with gmail server
    print("Email sent successfully.")                                   # Print successful message if gets here with no error

    # Optional Objective: 'By Country' file is deleted ONLY after the email sends successfully
    os.remove(output_file_path)                                         # After email has sent successfully remove file from device
    print("File deleted successfully.")                                 # Print successful message file was deleted

# Capture exception thrown from try block if code errors
except Exception as e:                                                  # Catch the exception thrown if the try block errors at any point
    print("Email could not be sent. Error:", str(e))                    # Print prompt with error message from try block