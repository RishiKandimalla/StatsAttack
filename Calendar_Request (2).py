import os
import csv
import win32com.client
import datetime as dt

staff_member_email = "tsx4wu@virginia.edu"
staff_member_name = "Zachary Denison"
rds_sne_group = "Research Librarianship"
department = "Your Department"  # Replace with your actual department
grant_related = "No"  # Set this depending on your context
school = "Your School"  # Replace with the actual school
source_software = "Outlook"

def parse_user_input_date(date_str):
    """Parse the user input date in YYYY-MM-DD format."""
    try:
        return dt.datetime.strptime(date_str, "%Y-%m-%d")
    except ValueError:
        print("Invalid date format. Please use YYYY-MM-DD format.")
        exit(1)

# Ask the user for the date range
start_date_input = input("Enter the start date (YYYY-MM-DD): ")
end_date_input = input("Enter the end date (YYYY-MM-DD): ")

# Convert user input into datetime objects
start_date = parse_user_input_date(start_date_input)
end_date = parse_user_input_date(end_date_input)

# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
calendar = outlook.GetDefaultFolder(9).Items
calendar.IncludeRecurrences = True
calendar.Sort("[Start]")

# Apply the user-specified date range with correct format for Outlook restriction
start_date_formatted_for_restriction = start_date.strftime("%m/%d/%Y 12:00 AM")
end_date_formatted_for_restriction = end_date.strftime("%m/%d/%Y 11:59 PM")
restriction = "[Start] >= '" + start_date_formatted_for_restriction + "' AND [End] <= '" + end_date_formatted_for_restriction + "'"
calendar = calendar.Restrict(restriction)

# Define the path to the Downloads folder
downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
csv_file_name = os.path.join(downloads_path, "outlook_export.csv")

# Open the CSV file for writing
with open(csv_file_name, mode='w', newline='', encoding='utf-8') as file:
    writer = csv.writer(file)
    # Write the header row
    writer.writerow([
        "Start Date", "Internal Notes", "Entered By", "Additional Notes",
        "Additional Staff", "Additional Users", "ARL Interaction Type", 
        "Attendee Type", "Date of the interaction", "Department", 
        "Description", "Grant Related?", "Medium",
        "Pre-post-time", "Primary User Name", "Primary User's Computing ID", 
        "RDS+SNE Group", "Referral", "School", "Session Duration", 
        "Source/Software", "Staff", "Topic"
    ])
    
    for item in calendar:
        start_date = item.Start.Format("%Y-%m-%d")
        internal_notes = item.Subject + ": " + item.Body
        entered_by = staff_member_name
        additional_notes = item.Subject + ": " + item.Body
        additional_staff = "N/A"
        additional_users = "; ".join([attendee.Name for attendee in item.Recipients if "@staff" not in attendee.Address])
        arl_interaction_type = "N/A"  # Replace with actual interaction type if known
        attendee_type = "N/A"  # Replace with actual attendee type if known
        date_of_the_interaction = item.Start.Format("%Y-%m-%d")
        description = item.Subject
        medium = "Zoom" if "zoom" in item.Body.lower() else "In-person"
        pre_post_time = "N/A"  # Set this depending on your context
        primary_user_name = staff_member_name  # Assuming the staff member is the primary user
        primary_users_computing_id = staff_member_email.split('@')[0]
        referral = "N/A"  # Set this depending on your context
        session_duration = str(item.Duration) + " minutes"
        # Write the row to the CSV file
        writer.writerow([
            start_date, internal_notes, entered_by, additional_notes,
            additional_staff, additional_users, arl_interaction_type, 
            attendee_type, date_of_the_interaction, department, 
            description, grant_related, medium,
            pre_post_time, primary_user_name, primary_users_computing_id, 
            rds_sne_group, referral, school, session_duration, 
            source_software, staff_member_name, description  # Assuming 'Topic' is the same as 'Description'
        ])

       


