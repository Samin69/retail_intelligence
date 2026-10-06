import asyncio
import base64
import time
from dataclasses import dataclass
from typing import Optional

import httpx
import streamlit as st


class Settings:
    def __init__(self):
        self.databricks_host = str(st.secrets["DATABRICKS_HOST"]).rstrip("/")
        self.databricks_client_id = str(st.secrets["DATABRICKS_CLIENT_ID"])
        self.databricks_client_secret = str(st.secrets["DATABRICKS_CLIENT_SECRET"])
        self.genie_space_id = str(st.secrets["GENIE_SPACE_ID"])
        self.genie_poll_interval_seconds = float(st.secrets.get("GENIE_POLL_INTERVAL_SECONDS", 1.0))
        self.genie_timeout_seconds = int(st.secrets.get("GENIE_TIMEOUT_SECONDS", 1800))
        self.genie_stream_retries = int(st.secrets.get("GENIE_STREAM_RETRIES", 2))
        self.genie_tcp_keepalive = str(st.secrets.get("GENIE_TCP_KEEPALIVE", "true")).lower() in {"1", "true", "yes", "on"}
        self.postgres_host = str(st.secrets.get("POSTGRES_HOST", "127.0.0.1"))
        self.postgres_port = int(st.secrets.get("POSTGRES_PORT", 5432))
        self.postgres_database = str(st.secrets["POSTGRES_DATABASE"])
        self.postgres_user = str(st.secrets["POSTGRES_USER"])
        self.postgres_password = str(st.secrets["POSTGRES_PASSWORD"])
        self.postgres_schema = str(st.secrets.get("POSTGRES_SCHEMA", "genie_app"))
        self.app_username = str(st.secrets["APP_USERNAME"])
        self.app_password = str(st.secrets["APP_PASSWORD"])


@st.cache_resource
def get_settings():
    return Settings()


settings = get_settings()


class DatabricksAuthError(Exception):
    pass


class DatabricksAuth:
    """Databricks OAuth token cache that is safe across Streamlit reruns.

    Streamlit reruns can execute async work on different event loops. Therefore
    this object deliberately does NOT retain an AsyncClient. The access token is
    process-cached, while each token request owns and closes its HTTP client on
    the same event loop that created it.
    """

    def __init__(self):
        self._access_token: Optional[str] = None
        self._expires_at = 0.0

    async def get_access_token(self, force_refresh: bool = False) -> str:
        if not force_refresh and self._access_token and time.time() < self._expires_at - 60:
            return self._access_token

        credentials = f"{settings.databricks_client_id}:{settings.databricks_client_secret}"
        encoded = base64.b64encode(credentials.encode("ascii")).decode("ascii")

        async with httpx.AsyncClient(
            timeout=httpx.Timeout(connect=30.0, read=60.0, write=30.0, pool=10.0)
        ) as client:
            response = await client.post(
                f"{settings.databricks_host}/oidc/v1/token",
                headers={
                    "Authorization": f"Basic {encoded}",
                    "Content-Type": "application/x-www-form-urlencoded",
                },
                data={"grant_type": "client_credentials", "scope": "genie"},
            )

        if response.status_code >= 400:
            raise DatabricksAuthError(
                f"Databricks OAuth failed: HTTP {response.status_code}: {response.text[:1000]}"
            )

        payload = response.json()
        token = payload.get("access_token")
        if not token:
            raise DatabricksAuthError("Databricks OAuth response did not contain access_token.")

        self._access_token = token
        self._expires_at = time.time() + int(payload.get("expires_in", 3600))
        return token

    async def invalidate_token(self, token: str | None = None):
        self._access_token = None
        self._expires_at = 0.0

    async def close(self):
        # Kept for compatibility with the existing client lifecycle.
        return None


@st.cache_resource
def get_databricks_auth():
    return DatabricksAuth()


databricks_auth = get_databricks_auth()
