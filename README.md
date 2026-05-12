# SP_RAG_gebiedsdossiers

Scripts om gebiedsdossiers te parsen. Documenten worden op basis van de TOC geindexeerd en omgezet naar numerieke vectoren. Met een AI model en cosinus gelijknis kunnen de paragraven doorzocht worden en kan data worden gestructureerd. Deze verkenning is uitgevoerd voor het project Signaleren en Prioriteren (2025)

# Scripts:
extract.py : bevat code om de documenten te parsen en een tweede deel waarmee de uitvraag met het LLM (AI) model kan worden uitgevoerd. 
parse_write.py : bevat alleen de code om documenten te parsen, deze kan worden gebruikt om een job te starten op het HPC.
Precisie_Recall.py: berekent score voor de resultaten, zie memo 2025

