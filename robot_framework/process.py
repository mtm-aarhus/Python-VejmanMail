from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement
from email.message import EmailMessage
import smtplib
import requests
import os
from datetime import datetime, timedelta 

def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement | None = None) -> None:
        
    orchestrator_connection.log_info("Getting token")

    # Get credentials
    token = orchestrator_connection.get_credential("VejmanToken").password
    orchestrator_connection.log_info("Defining variables")
    # Date Variables
    yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    today = datetime.now().strftime("%Y-%m-%d")

    # URLs and headers
    urls = [	f'https://vejman.vd.dk/permissions/getcases?pmCaseStates=1%2C2%2C3%2C6&pmCaseFields=state%2Ctype%2Ccase_number%2Cauthority_reference_number%2Cwebgtno%2Cstart_date%2Cend_date%2Capplicant_folder_number%2Cconnected_case%2Cstreet_name%2Capplicant%2Crovm_equipment_type%2Cinitials&pmCaseWorker=all&pmCaseTypes=%27rovm%27%2C%27gt%27&pmCaseVariant=all&pmCaseTags=ignorerTags&pmCaseShowAttachments=false&dontincludemap=1&endDateFrom={yesterday}&endDateTo={today}&_=1715179724504&token={token}',
        f'https://vejman.vd.dk/permissions/getcases?pmCaseStates=8&pmCaseFields=state%2Ctype%2Ccase_number%2Cauthority_reference_number%2Cwebgtno%2Cstart_date%2Cend_date%2Capplicant_folder_number%2Cconnected_case%2Cstreet_name%2Capplicant%2Crovm_equipment_type%2Cinitials&pmCaseWorker=all&pmCaseTypes=%27rovm%27%2C%27gt%27&pmCaseVariant=all&pmCaseTags=ignorerTags&pmCaseShowAttachments=false&dontincludemap=1&&endDateFrom={today}&endDateTo={today}&_=1715093019127&token={token}',
        f'https://vejman.vd.dk/permissions/getcases?pmCaseStates=1%2C3%2C6%2C8%2C12&pmCaseFields=state%2Ctype%2Ccase_number%2Cauthority_reference_number%2Cwebgtno%2Cstart_date%2Cend_date%2Capplicant_folder_number%2Cconnected_case%2Cstreet_name%2Capplicant%2Crovm_equipment_type%2Cinitials&pmCaseWorker=all&pmCaseTypes=%27rovm%27%2C%27gt%27&pmCaseVariant=all&pmCaseTags=ignorerTags&pmCaseShowAttachments=false&dontincludemap=1&startDateFrom={today}&startDateTo={today}&_=1715095095761&token={token}'
        ]
    

    headers = ["Udløbne tilladelser", "Færdigmeldte tilladelser", "Nye tilladelser"]

    # Custom Headers for Table (Order is enforced here)
    custom_headers = [
        ("case_id", "Sag"),
        ("initials", "Behandler"),
        ("state", "Status"),
        ("type", "Ansøgning"),
        ("connected_case", "Relateret"),
        ("end_date", "Slutdato"),
        ("start_date", "Startdato"),
        ("applicant", "Ansøger"),
        ("rovm_equipment_type", "Udstyr"),
        ("applicant_folder_number", "Sagsmappenr"),
        ("authority_reference_number", "Kommentar"),
        ("street_name", "Vejnavn"),
    ]
    orchestrator_connection.log_info("Generating HTML table")
    # Generate HTML Table
    html_table = ""
    for idx, url in enumerate(urls):
        orchestrator_connection.log_info(f'Looking at url no. {idx}')
        response = requests.get(url)
        response.raise_for_status()
        cases = response.json().get("cases", [])

        html_table += f"<h2>{headers[idx]}</h2>"

        if not cases:
            html_table += "<p>Ingen tilladelser</p>"
            continue

        # Filter cases by initials
        filtered_cases = [case for case in cases if case.get("initials") in ["MAMASA", "LERV"]]

        # Fetch detailed case information and update start_date and end_date
        updated_cases = []
        for case in filtered_cases:
            case_id = case.get("case_id")
            if not case_id:
                continue

            # Fetch detailed case information
            orchestrator_connection.log_info("Connecting to API")
            case_response = requests.get(f"https://vejman.vd.dk/permissions/getcase?caseid={case_id}&token={token}")
            case_response.raise_for_status()
            case_details = case_response.json().get("data", {})

            # Update start_date and end_date with detailed case information
            orchestrator_connection.log_info("Updating start and end date")
            case["start_date"] = case_details.get("start_date", case.get("start_date"))
            case["end_date"] = case_details.get("end_date", case.get("end_date"))
            updated_cases.append(case)

        # Apply time-based filtering for "Udløbne tilladelser"
        if headers[idx] == "Udløbne tilladelser":
            def should_include_case(case):
                end_date_str = case.get("end_date")
                if not end_date_str:
                    return False
                
                # Parse end_date into a datetime object
                try:
                    end_date = datetime.strptime(end_date_str, "%d-%m-%Y %H:%M:%S")
                except ValueError:
                    print('End time wrong format')

                # Check time constraints
                yesterday = datetime.now() - timedelta(days=1)
                today = datetime.now()
                return (
                    (end_date.date() == yesterday.date() and end_date.hour >= 8) or
                    (end_date.date() == today.date() and end_date.hour < 8 and not (end_date.hour == 0 and end_date.minute == 0))
                )

            # Filter cases based on the "shouldIncludeCase" logic
            orchestrator_connection.log_info("Updating cases")
            updated_cases = [case for case in updated_cases if should_include_case(case)]

        if not updated_cases:
            html_table += f"<p>Ingen tilladelser</p>"
            continue

        # Start building the table
        orchestrator_connection.log_info("Building table")
        temp_table = "<table border='1'><tr>"
        for _, header in custom_headers:
            temp_table += f"<th>{header}</th>"
        temp_table += "</tr>"

        # Populate rows
        for case in updated_cases:
            case_id = case.get("case_id")
            case_number = case.get("case_number", "")

            # Build the row
            temp_table += "<tr>"
            for key, _ in custom_headers:
                value = case.get(key, "")
                if value is None:  # If the value is None, leave it blank
                    value = ""
                if key == "case_id":
                    temp_table += f"<td><a href='https://vejman.vd.dk/permissions/update.jsp?caseid={case_id}'>{case_number}</a></td>"
                else:
                    temp_table += f"<td>{value}</td>"
            temp_table += "</tr>"
        temp_table += "</table>"
        html_table += temp_table
    orchestrator_connection.log_info("Sending email")
    SMTP_SERVER = "smtp.adm.aarhuskommune.dk"
    SMTP_PORT = 25
    SCREENSHOT_SENDER = "vejmanmail@aarhus.dk"

    jadt = orchestrator_connection.get_constant("jadt").value
    balas = orchestrator_connection.get_constant("balas").value

    # Send Email
    if html_table.strip():
        msg = EmailMessage()
        msg['To'] = balas
        msg['From'] = SCREENSHOT_SENDER
        msg['Subject'] = "Daglig liste over tilladelser i Vejman"
        msg.set_content("Please enable HTML to view this message.")
        msg.add_alternative(html_table, subtype='html')
        msg['Bcc'] = balas

        try:
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
                smtp.send_message(msg)
                orchestrator_connection.log_info("Email sent successfully.")
        except Exception as e:
            orchestrator_connection.log_error(f"Failed to send email: {e}")
