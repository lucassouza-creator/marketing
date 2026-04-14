"""LearnWorlds API client with OAuth2 client credentials authentication."""

import os
import time
import requests
from dotenv import load_dotenv

load_dotenv()

BASE_URL = os.getenv("LEARNWORLDS_URL", "https://academy.indicium.ai/admin/api")
CLIENT_ID = os.getenv("LEARNWORLDS_CLIENT_ID")
CLIENT_SECRET = os.getenv("LEARNWORLDS_CLIENT_SECRET")


class LearnWorldsClient:
    def __init__(self):
        if not CLIENT_ID or not CLIENT_SECRET:
            raise ValueError(
                "LEARNWORLDS_CLIENT_ID and LEARNWORLDS_CLIENT_SECRET must be set in .env"
            )
        self.base_url = BASE_URL.rstrip("/")
        self.client_id = CLIENT_ID
        self.client_secret = CLIENT_SECRET
        self._access_token = None
        self._token_expires_at = 0

    def _authenticate(self):
        """Obtain OAuth2 access token via client credentials flow."""
        url = f"{self.base_url}/v2/oauth2/access_token"
        response = requests.post(
            url,
            headers={
                "Content-Type": "application/x-www-form-urlencoded",
                "Lw-Client": self.client_id,
            },
            data={
                "client_id": self.client_id,
                "client_secret": self.client_secret,
                "grant_type": "client_credentials",
            },
        )
        response.raise_for_status()
        data = response.json()
        self._access_token = data["tokenData"]["token"]
        expires_in = data["tokenData"].get("expires", 3600)
        self._token_expires_at = time.time() + expires_in - 60  # 60s buffer
        print(f"[auth] Token obtained, expires in {expires_in}s")

    def _ensure_token(self):
        if not self._access_token or time.time() >= self._token_expires_at:
            self._authenticate()

    def _headers(self):
        self._ensure_token()
        return {
            "Authorization": f"Bearer {self._access_token}",
            "Lw-Client": self.client_id,
            "Content-Type": "application/json",
        }

    def get(self, path, params=None):
        url = f"{self.base_url}{path}"
        response = requests.get(url, headers=self._headers(), params=params)
        response.raise_for_status()
        return response.json()

    def get_paginated(self, path, params=None, page_size=25):
        """Iterate through all pages of a paginated endpoint."""
        params = params or {}
        params["page"] = 1
        params["itemsPerPage"] = page_size

        while True:
            data = self.get(path, params=params)
            items = data.get("data", [])
            yield from items

            meta = data.get("meta", {})
            total_pages = meta.get("totalPages", 1)
            current_page = meta.get("page", 1)

            if current_page >= total_pages:
                break
            params["page"] += 1
