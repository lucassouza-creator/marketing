#!/usr/bin/env python3
"""
LearnWorlds → Google Sheets exporter
--------------------------------------
Extrai dados do curso da LearnWorlds e cria uma planilha Google Sheets
com 7 abas de análise completa.

Uso:
    python run_sheets_export.py <course_id>
    python run_sheets_export.py formacao-analise-dados-corteva

Variáveis de ambiente necessárias (.env):
    LEARNWORLDS_URL             URL base da API (ex: https://academy.indicium.ai/admin/api)
    LEARNWORLDS_CLIENT_ID       Client ID da LearnWorlds
    LEARNWORLDS_CLIENT_SECRET   Client Secret da LearnWorlds
    GOOGLE_SERVICE_ACCOUNT_FILE Caminho para o JSON da service account (padrão: google-credentials.json)
    GOOGLE_SHARE_EMAIL          Email para compartilhar a planilha (opcional)
"""

import sys
import os
from dotenv import load_dotenv

load_dotenv()

from learnworlds.client import LearnWorldsClient
from learnworlds.extractor import collect_full_course_data
from learnworlds.sheets_exporter import export_to_sheets


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    course_id = sys.argv[1]
    share_email = os.getenv("GOOGLE_SHARE_EMAIL")

    print(f"Iniciando extração para o curso: {course_id}\n")
    client = LearnWorldsClient()
    course_data = collect_full_course_data(client, course_id)

    url = export_to_sheets(
        course_data,
        share_email=share_email,
    )

    print(f"Planilha disponível em:\n  {url}")
