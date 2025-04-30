# pipeline.py
"""
Automated GMD sanity‑check pipeline
----------------------------------
Downloads a model extract from the GMD portal, fetches the most‑recent
validation report from SharePoint, injects key values from the extract
into the report, re‑uploads the updated report to SharePoint, and sends
a Teams notification.

External deps (add to requirements.txt):
  playwright>=1.44
  office365-rest-python-client>=2.5
  pandas>=2.2
  openpyxl>=3.1
  python-docx>=1.1
  msal>=1.27 (optional if you prefer MSAL for auth)
  requests>=2.31

Environment variables expected (use a vault/secret store in prod):
  # GMD portal
  GMD_URL           – e.g. "https://gmd.mybank.com"
  GMD_USERNAME      – SSO/LDAP username
  GMD_PASSWORD      – password or app‑specific token

  # SharePoint / Graph
  SP_SITE_URL       – e.g. "https://mybank.sharepoint.com/sites/MRM"
  SP_DOC_LIB        – document library name, e.g. "ValidationReports"
  SP_TENANT_ID      – Azure AD tenant id
  SP_CLIENT_ID      – app registration client id
  SP_CLIENT_SECRET  – app registration secret

  # Notification
  TEAMS_WEBHOOK_URL – Incoming webhook for Teams channel
"""

from __future__ import annotations

import asyncio
import json
import logging
import os
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Dict, Optional

import pandas as pd
import requests
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from office365.sharepoint.files.file_creation_information import FileCreationInformation
from playwright.async_api import async_playwright
import docx

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s ‑ %(levelname)s ‑ %(name)s ‑ %(message)s",
)
LOGGER = logging.getLogger("gmd_pipeline")


# ---------------------------------------------------------------------------
# 1. GMD download helpers
# ---------------------------------------------------------------------------

async def download_gmd_extract(model_id: str, download_dir: Path) -> Path:
    """Login to GMD and download the Excel extract for *model_id*.

    Returns the path to the downloaded file.
    """
    gmd_url = os.environ["GMD_URL"]
    username = os.environ["GMD_USERNAME"]
    password = os.environ["GMD_PASSWORD"]

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()

        # 1. Navigate to login page and authenticate (update selectors!)
        await page.goto(f"{gmd_url}/login")
        await page.fill("input[name=username]", username)
        await page.fill("input[name=password]", password)
        await page.click("button[type=submit]")
        await page.wait_for_load_state("networkidle")

        # 2. Navigate to model page
        await page.goto(f"{gmd_url}/model/{model_id}")
        await page.wait_for_load_state("networkidle")

        # 3. Click export‑to‑Excel button
        with page.expect_download() as download_info:
            await page.click("text=Export to Excel")  # selector example
        download = await download_info.value
        dest = download_dir / f"gmd_extract_{model_id}.xlsx"
        await download.save_as(dest)

        await context.close()
        await browser.close()

    LOGGER.info("GMD extract saved at %s", dest)
    return dest


# ---------------------------------------------------------------------------
# 2. SharePoint helpers
# ---------------------------------------------------------------------------

class SharePointClient:  # minimal wrapper
    def __init__(self):
        self._site_url = os.environ["SP_SITE_URL"]
        self._doc_lib = os.environ["SP_DOC_LIB"]
        self._ctx = ClientContext(self._site_url).with_credentials(
            UserCredential(os.environ["GMD_USERNAME"], os.environ["GMD_PASSWORD"])
        )

    def latest_report(self, model_id: str, dest_dir: Path) -> Path:
        """Download the most‑recent report file for *model_id* into *dest_dir*.
        Assumes files are named like **<model_id>_YYYYMMDD.docx**.
        Returns local path.
        """
        folder = self._ctx.web.lists.get_by_title(self._doc_lib).root_folder
        files = folder.files.expand(["ListItemAllFields"]).get().execute_query()
        target_files = [
            f for f in files if f.name.lower().startswith(model_id.lower())
        ]
        if not target_files:
            raise RuntimeError("No report found for model %s" % model_id)
        latest: File = max(
            target_files, key=lambda f: f.time_last_modified  # noqa: E501
        )
        local = dest_dir / latest.name
        with open(local, "wb") as fp:
            latest.download(fp).execute_query()
        LOGGER.info("Downloaded latest report to %s", local)
        return local

    def upload_report(self, local_path: Path, remote_name: Optional[str] = None):
        if remote_name is None:
            remote_name = local_path.name
        folder = self._ctx.web.lists.get_by_title(self._doc_lib).root_folder
        info = FileCreationInformation()
        info.overwrite = True
        info.url = remote_name
        with open(local_path, "rb") as content:
            info.content = content.read()
        folder.files.add(info).execute_query()
        LOGGER.info("Uploaded updated report %s", remote_name)


# ---------------------------------------------------------------------------
# 3. Data merge helpers
# ---------------------------------------------------------------------------

def inject_values_into_docx(extract_path: Path, report_path: Path) -> Path:
    """Replace placeholders in *report_path* using values in *extract_path*.

    Placeholders are assumed to be like {{FieldName}} inside the Word doc.
    Column names in the extract must match FieldName.
    Returns path to the new document.
    """
    df = pd.read_excel(extract_path, sheet_name=0)  # adjust sheet if needed
    record: Dict[str, str] = df.iloc[0].to_dict()  # simple example – first row

    doc = docx.Document(report_path)
    for p in doc.paragraphs:
        for key, val in record.items():
            placeholder = f"{{{{{key}}}}}"
            if placeholder in p.text:
                inline = p.runs
                for i in range(len(inline)):
                    if placeholder in inline[i].text:
                        inline[i].text = inline[i].text.replace(placeholder, str(val))
    # Save as new file next to original
    new_path = report_path.with_name(report_path.stem + "_updated.docx")
    doc.save(new_path)
    LOGGER.info("Report updated ➜ %s", new_path)
    return new_path


# ---------------------------------------------------------------------------
# 4. Notification helper
# ---------------------------------------------------------------------------

def send_teams_message(title: str, text: str):
    url = os.environ.get("TEAMS_WEBHOOK_URL")
    if not url:
        LOGGER.warning("No TEAMS_WEBHOOK_URL set – skipping Teams notification")
        return
    payload = {
        "@type": "MessageCard",
        "@context": "http://schema.org/extensions",
        "summary": title,
        "themeColor": "0076D7",
        "title": title,
        "text": text,
    }
    resp = requests.post(url, json=payload, timeout=10)
    resp.raise_for_status()
    LOGGER.info("Teams notification sent")


# ---------------------------------------------------------------------------
# 5. Orchestration
# ---------------------------------------------------------------------------

async def run_pipeline(model_id: str):
    LOGGER.info("Starting pipeline for model %s", model_id)
    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        extract = await download_gmd_extract(model_id, tmp_dir)

        sp = SharePointClient()
        report = sp.latest_report(model_id, tmp_dir)

        updated = inject_values_into_docx(extract, report)
        sp.upload_report(updated)

    send_teams_message(
        title=f"GMD sanity‑check complete – {model_id}",
        text=f"Updated validation report has been uploaded for **{model_id}** on {datetime.utcnow():%Y‑%m‑%d %H:%M} UTC.",
    )


# ---------------------------------------------------------------------------
# 6. CLI entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Run GMD pipeline for a model id")
    parser.add_argument("model_id", help="Model identifier as used in GMD/SharePoint")
    args = parser.parse_args()

    try:
        asyncio.run(run_pipeline(args.model_id))
    except Exception as exc:
        LOGGER.exception("Pipeline failed: %s", exc)
        send_teams_message(
            title=f"GMD sanity‑check FAILED – {args.model_id}",
            text=f"Error: {exc}",
        )
        raise
