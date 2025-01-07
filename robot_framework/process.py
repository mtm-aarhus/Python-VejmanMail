from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from email.message import EmailMessage
import smtplib
import requests
import os
from datetime import datetime, timedelta 


def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement | None = None) -> None:
    orchestrator_connection = OrchestratorConnection("VejmanMail", os.getenv('OpenOrchestratorSQL'),os.getenv('OpenOrchestratorKey'), None)

    send_to_fællesmail = True

    # Mail til sagsbehandler
    orchestrator_connection.log_info("Running VejmanMail")

    # Initialize variables
    yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    today = datetime.now().strftime("%Y-%m-%d")
    token = orchestrator_connection.get_credential("VejmanToken").password

    # URLs and headers
    urls = [
        f"https://vejman.vd.dk/permissions/getcases?pmCaseStates=1%2C2%2C3%2C6&pmCaseFields=state%2Ctype%2Ccase_number%2Cauthority_reference_number%2Cwebgtno%2Cstart_date%2Cend_date%2Capplicant_folder_number%2Cconnected_case%2Cstreet_name%2Capplicant%2Crovm_equipment_type%2Cinitials&pmCaseWorker=all&pmCaseTypes=%27rovm%27%2C%27gt%27&pmCaseVariant=all&pmCaseTags=ignorerTags&pmCaseShowAttachments=false&dontincludemap=1&endDateFrom={yesterday}&endDateTo={today}&_=1715179724504&token={token}",
        f"https://vejman.vd.dk/permissions/getcases?pmCaseStates=8&pmCaseFields=state%2Ctype%2Ccase_number%2Cauthority_reference_number%2Cwebgtno%2Cstart_date%2Cend_date%2Capplicant_folder_number%2Cconnected_case%2Cstreet_name%2Capplicant%2Crovm_equipment_type%2Cinitials&pmCaseWorker=all&pmCaseTypes=%27rovm%27%2C%27gt%27&pmCaseVariant=all&pmCaseTags=ignorerTags&pmCaseShowAttachments=false&dontincludemap=1&endDateFrom={today}&endDateTo={today}&_=1715093019127&token={token}",
        f"https://vejman.vd.dk/permissions/getcases?pmCaseStates=1%2C3%2C6%2C8%2C12&pmCaseFields=state%2Ctype%2Ccase_number%2Cauthority_reference_number%2Cwebgtno%2Cstart_date%2Cend_date%2Capplicant_folder_number%2Cconnected_case%2Cstreet_name%2Capplicant%2Crovm_equipment_type%2Cinitials&pmCaseWorker=all&pmCaseTypes=%27rovm%27%2C%27gt%27&pmCaseVariant=all&pmCaseTags=ignorerTags&pmCaseShowAttachments=false&dontincludemap=1&startDateFrom={today}&startDateTo={today}&_=1715095095761&token={token}"
    ]
    headers = ["Udløbne tilladelser", "Færdigmeldte tilladelser", "Nye tilladelser"]

    html_table = ""

    # Loop through each URL
    for idx, url in enumerate(urls):
        print(f"Henter {headers[idx]}")
        response = requests.get(url)
        response.raise_for_status()
        json_object = response.json()

        # Extract cases
        cases = json_object.get("cases", [])
        html_table += f"<h2>{headers[idx]}</h2>"

        if not cases:
            html_table += "<p>Ingen tilladelser</p>"
            continue

        # Filter cases based on initials
        filtered_cases = [
            case for case in cases
            if case.get("initials") in ["MAMASA", "LERV"]
        ]

        if not filtered_cases:
            html_table += "<p>Ingen tilladelser</p>"
            continue

        # Prepare table headers
        columns_to_exclude = {"geometrycount", "manualpinpoint", "ispinpointed", "message_state", "webgtno", "case_number"}
        custom_headers = {
            "case_id": "Sag",
            "initials": "Behandler",
            "state": "Status",
            "type": "Ansøgning",
            "end_date": "Slutdato",
            "applicant_folder_number": "Sagsmappenr",
            "start_date": "Startdato",
            "applicant": "Ansøger",
            "rovm_equipment_type": "Udstyr",
            "authority_reference_number": "Kommentar",
            "street_name": "Vejnavn",
            "connected_case": "Relateret",
        }

        # Temporary table to store rows
        temp_table = "<table border='1'><tr>"
        for key in filtered_cases[0]:
            if key not in columns_to_exclude:
                header = custom_headers.get(key, key)
                temp_table += f"<th>{header}</th>"
        temp_table += "</tr>"

        # Populate table rows
        table_has_rows = False
        for case in filtered_cases:
            case_id = case.get("case_id")
            if not case_id:
                continue

            # Fetch detailed case information
            print(f"Henter info om {case_id}")
            case_response = requests.get(f"https://vejman.vd.dk/permissions/getcase?caseid={case_id}&token={token}")
            case_response.raise_for_status()
            case_details = case_response.json().get("data", {})

            # Update start_date and end_date
            case["start_date"] = case_details.get("start_date")
            case["end_date"] = case_details.get("end_date")

            # Parse and format dates
            end_date = datetime.strptime(case["end_date"], "%d-%m-%Y %H:%M:%S")
            start_date = datetime.strptime(case["start_date"], "%d-%m-%Y %H:%M:%S")

            should_include_case = True
            if headers[idx] == "Udløbne tilladelser":
                should_include_case = (
                    (end_date.date() == datetime.now().date() and end_date.hour < 8) or
                    (end_date.date() == (datetime.now() - timedelta(days=1)).date() and end_date.hour >= 8)
                )

            if should_include_case:
                table_has_rows = True
                temp_table += "<tr>"
                for key, value in case.items():
                    if key not in columns_to_exclude:
                        if key == "case_id":
                            case_number = case.get("case_number", "")
                            temp_table += f"<td><a href='https://vejman.vd.dk/permissions/update.jsp?caseid={value}'>{case_number}</a></td>"
                        else:
                            temp_table += f"<td>{value}</td>"
                temp_table += "</tr>"
        temp_table += "</table>"

        # Append the table only if it has rows
        if table_has_rows:
            html_table += temp_table
        else:
            html_table += "<p>Ingen tilladelser</p>"

    # Output the HTML
    if send_to_fællesmail == True:
        subject = "Daglig liste over tilladelser i Vejman"
        body = html_table

        # Create the email message
        msg = EmailMessage()
        msg['To'] = orchestrator_connection.get_constant("balas")#orchestrator_connection.get_constant("VejArealMail")
        msg['From'] = SCREENSHOT_SENDER
        msg['Subject'] = subject
        msg.set_content("Please enable HTML to view this message.")
        msg.add_alternative(body, subtype='html')
        msg['Bcc'] = ERROR_EMAIL

        # Send the email using SMTP
        try:
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
                smtp.send_message(msg)
                orchestrator_connection.log_info("VejmanMail sent")
        except Exception as e:
            print(f"Failed to send email: {e}")
        