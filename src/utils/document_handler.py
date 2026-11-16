import requests
import json
import hashlib
import base64

from urllib.parse import urlparse


def download_contracts(payload):
    leasing_data = payload.get("leasingaftaler", {})
    documents = leasing_data.get("aftalekontrakter", [])

    if isinstance(documents, str):
        documents = json.loads(documents)

    attachments = []

    for document in documents:
        file_url = document.get("url")
        file_name = document.get("title", "document")
        mime_type = document.get(
            "mimeType",
            "application/octet-stream"
        )

        if not file_url:
            continue

        parsed_url = urlparse(file_url)

        # TODO: Verify allowed XFlow domains
        if parsed_url.scheme != "https":
            raise ValueError("Invalid document URL")

        # TODO: Verify authentication requirements with XFlow
        response = requests.get(
            file_url,
            timeout=30,
            allow_redirects=False
        )

        response.raise_for_status()

        file_content = response.content

        expected_size = document.get("size")

        if (
            expected_size is not None
            and len(file_content) != expected_size
        ):
            raise ValueError(
                f"File size mismatch: {file_name}"
            )

        # TODO: Verify XFlow checksum format
        expected_checksum = document.get("checksum")

        if expected_checksum:
            actual_hash = hashlib.sha256(
                file_content
            ).digest()

            actual_checksum = base64.b64encode(
                actual_hash
            ).decode("ascii")

            if actual_checksum != expected_checksum:
                raise ValueError(
                    f"Checksum mismatch: {file_name}"
                )

        attachments.append({
            "filename": file_name,
            "mime_type": mime_type,
            "content": file_content
        })

    return attachments
