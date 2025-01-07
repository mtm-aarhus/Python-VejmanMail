#Connect to orchestrator
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement
import os
import json

orchestrator_connection = OrchestratorConnection("Dokumentliste i Python", os.getenv('OpenOrchestratorSQL'),os.getenv('OpenOrchestratorKey'), None)
queue_json = {
        "SagsID": "GEO-2024-199848",
        "MailModtager": "jadt@aarhus.dk",
        "PodioID": "2923285810",
        "DeskProID": "2070",
        "DeskProTitel": "Test",
    }

orchestrator_connection.create_queue_element("AktbobDokumentlisteQueue", "test", json.dumps(queue_json))


