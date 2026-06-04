"""One-off: upload brand logo assets to the blob container under brand/."""
import os
from pathlib import Path
from dotenv import load_dotenv
from azure.storage.blob import BlobServiceClient, ContentSettings

load_dotenv()

conn = (os.getenv("BLOB_CONNECTION_STRING") or "").strip()
container = (os.getenv("CONTAINER") or "marlu-online-quote").strip()
if not conn:
    raise SystemExit("BLOB_CONNECTION_STRING not set")

svc = BlobServiceClient.from_connection_string(conn)
cc = svc.get_container_client(container)
try:
    cc.create_container()
except Exception:
    pass

uploads = [
    (Path(__file__).parent / "brand_assets" / "westlink-logo.png", "brand/westlink-logo.png", "image/png"),
    (Path(__file__).parent / "brand_assets" / "marlu-logo.jpg",     "brand/marlu-logo.jpg",     "image/jpeg"),
]

for src, blob_name, ctype in uploads:
    if not src.exists():
        print(f"SKIP (missing): {src}")
        continue
    data = src.read_bytes()
    cc.get_blob_client(blob_name).upload_blob(
        data, overwrite=True, content_settings=ContentSettings(content_type=ctype))
    print(f"UPLOADED: {blob_name}  <-  {src.name}  ({len(data):,} bytes)  container={container}")

print("done")
