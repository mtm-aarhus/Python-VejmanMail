from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement
from email.message import EmailMessage
import smtplib
import requests
import os
from datetime import datetime, timedelta 



def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement | None = None) -> None:
    
    # Initialize Orchestrator Connection
    orchestrator_connection = OrchestratorConnection("VejmanMail", os.getenv('OpenOrchestratorSQL'), os.getenv('OpenOrchestratorKey'), None)

    # Get credentials
    token = orchestrator_connection.get_credential("VejmanToken").password

    # Date Variables
    yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    today = datetime.now().strftime("%Y-%m-%d")

    # URLs and headers
    urls = [
        f"https://vejman.vd.dk/permissions/getcases?pmCaseStates=1%2C2%2C3%2C6&pmCaseFields=state%2Ctype%2Ccase_number%2Cauthority_reference_number%2Cwebgtno%2Cstart_date%2Cend_date%2Capplicant_folder_number%2Cconnected_case%2Cstreet_name%2Capplicant%2Crovm_equipment_type%2Cinitials&pmCaseWorker=all&pmCaseTypes=%27rovm%27%2C%27gt%27&pmCaseVariant=all&pmCaseTags=ignorerTags&pmCaseShowAttachments=false&dontincludemap=1&endDateFrom={yesterday}&endDateTo={today}&_=1715179724504&token={token}",
        f"https://vejman.vd.dk/permissions/getcases?pmCaseStates=8&pmCaseFields=state%2Ctype%2Ccase_number%2Cauthority_reference_number%2Cwebgtno%2Cstart_date%2Cend_date%2Capplicant_folder_number%2Cconnected_case%2Cstreet_name%2Capplicant%2Crovm_equipment_type%2Cinitials&pmCaseWorker=all&pmCaseTypes=%27rovm%27%2C%27gt%27&pmCaseVariant=all&pmCaseTags=ignorerTags&pmCaseShowAttachments=false&dontincludemap=1&endDateFrom={today}&endDateTo={today}&_=1715093019127&token={token}",
        f"https://vejman.vd.dk/permissions/getcases?pmCaseStates=1%2C3%2C6%2C8%2C12&pmCaseFields=state%2Ctype%2Ccase_number%2Cauthority_reference_number%2Cwebgtno%2Cstart_date%2Cend_date%2Capplicant_folder_number%2Cconnected_case%2Cstreet_name%2Capplicant%2Crovm_equipment_type%2Cinitials&pmCaseWorker=all&pmCaseTypes=%27rovm%27%2C%27gt%27&pmCaseVariant=all&pmCaseTags=ignorerTags&pmCaseShowAttachments=false&dontincludemap=1&startDateFrom={today}&startDateTo={today}&_=1715095095761&token={token}"
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

    # Generate HTML Table
    html_table = ""

    for idx, url in enumerate(urls):
        response = requests.get(url)
        response.raise_for_status()
        cases = response.json().get("cases", [])

        html_table += f"<h2>{headers[idx]}</h2>"

        if not cases:
            html_table += "<p>Ingen tilladelser</p>"
            continue

        # Filter cases by initials
        filtered_cases = [case for case in cases if case.get("initials") in ["MAMASA", "LERV"]]

        if not filtered_cases:
            html_table += f"<p>Ingen tilladelser for initials: ['MAMASA', 'LERV']</p>"
            continue

        # Start building the table
        temp_table = "<table border='1'><tr>"
        for _, header in custom_headers:
            temp_table += f"<th>{header}</th>"
        temp_table += "</tr>"

        # Populate rows
        for case in filtered_cases:
            case_id = case.get("case_id")
            if not case_id:
                continue

            # Fetch case details
            case_details = requests.get(f"https://vejman.vd.dk/permissions/getcase?caseid={case_id}&token={token}").json().get("data", {})
            case["start_date"] = case_details.get("start_date", case.get("start_date"))
            case["end_date"] = case_details.get("end_date", case.get("end_date"))

            # Build the row
            temp_table += "<tr>"
            for key, _ in custom_headers:
                value = case.get(key, "")
                if value is None:  # If the value is None, leave it blank
                    value = ""
                if key == "case_id":
                    case_number = case.get("case_number", "")
                    temp_table += f"<td><a href='https://vejman.vd.dk/permissions/update.jsp?caseid={value}'>{case_number}</a></td>"
                else:
                    temp_table += f"<td>{value}</td>"
            temp_table += "</tr>"
        temp_table += "</table>"
        html_table += temp_table

    # Send Email
    if html_table.strip():
        msg = EmailMessage()
        msg['To'] = orchestrator_connection.get_constant("balas").value
        msg['From'] = SCREENSHOT_SENDER
        msg['Subject'] = "Daglig liste over tilladelser i Vejman"
        msg.set_content("Please enable HTML to view this message.")
        msg.add_alternative(html_table, subtype='html')
        msg['Bcc'] = ERROR_EMAIL

        try:
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
                smtp.send_message(msg)
                orchestrator_connection.log_info("Email sent successfully.")
        except Exception as e:
            orchestrator_connection.log_error(f"Failed to send email: {e}")
