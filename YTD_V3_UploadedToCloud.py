#V3 - Add date prefix for easier referencing + csv renaming

import csv
import pandas as pd
import schedule
import time
from datetime import datetime

# Read the CSV file into a DataFrame
sheet = pd.read_csv(
    "YTD and Spring semester total hours  - Total Hours Summa.csv")  # Replace with the path to your CSV file

# Create an empty dictionary to store messages for each manager
messages_dict = {}

# Group the DataFrame by "Manager Email"
grouped_sheet = sheet.groupby("Manager Email")

# Loop through each group
for email, group in grouped_sheet:
    # Get the manager's name and email address
    manager_name = group["Manager Name"].iloc[0]
    manager_email = email

    # Print a message to confirm the manager's name and email address
    print(f"Creating message for {manager_name} at {manager_email}...")

    # Loop through each intern in the group
    message = f"Dear {manager_name},\n\nWe hope you are doing well.\n\n\nBelow, you can find the Year-to-Date and Semester total hours of your WS interns.\n\n"

    for index, row in group.iterrows():
        student_first_name = row["Student First Name"]
        student_last_name = row["Student Last Name"]
        ytd = row["Year-To-Date Total Hours"]
        spring = row["Spring Total Hours"]
        left_hours = row["Total Spring hours left"]
        message += f"{student_first_name} {student_last_name}: YTD = {ytd}, Spring hours = {spring}, Hours left to work = {left_hours}\n\n"

    message += "Best,\nThe Student Finance Team"

    # Add the message to the messages_dict, with the manager's email as the key
    messages_dict[manager_email] = message

    # Generate the filename with the current date as a prefix
    current_date = datetime.now().strftime("%Y-%m-%d")
    filename = f"{current_date}_YTD_ManagersReport_SY2324.csv"

    # Write the messages to the CSV file
    with open(filename, "w", newline="") as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(["Manager Email", "Message"])
        for email, message in messages_dict.items():
            writer.writerow([email, message])

    print(f"CSV file '{filename}' created successfully.")

# Schedule the job to run every Saturday at 8:00 AM
schedule.every().saturday.at("08:00").do(generate_messages)

# Run the scheduler indefinitely
while True:
    schedule.run_pending()
    time.sleep(1)

