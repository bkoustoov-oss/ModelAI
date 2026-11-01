"""Create or inspect the local OTA signing key without adding it to the repository."""

import base64
import os
from pathlib import Path

from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey
from platformdirs import user_config_dir

def main():
    key_dir = Path(user_config_dir(
        "LAYA Agentic Studio", "Bkoustoov")) / "security"
    key_dir.mkdir(parents=True, exist_ok=True)
    key_path = key_dir / "update-signing-key.pem"
    if key_path.exists():
        private_key = serialization.load_pem_private_key(
            key_path.read_bytes(), password=None)
    else:
        private_key = Ed25519PrivateKey.generate()
        pem = private_key.private_bytes(
            serialization.Encoding.PEM,
            serialization.PrivateFormat.PKCS8,
            serialization.NoEncryption())
        with open(key_path, "xb") as output:
            output.write(pem)
        try:
            os.chmod(key_path, 0o600)
        except OSError:
            pass

    raw_private = private_key.private_bytes(
        serialization.Encoding.Raw,
        serialization.PrivateFormat.Raw,
        serialization.NoEncryption())
    raw_public = private_key.public_key().public_bytes(
        serialization.Encoding.Raw,
        serialization.PublicFormat.Raw)
    print(f"Private key file: {key_path}")
    print("Keep this file private and backed up securely.")
    print("GitHub Actions secret LAYA_UPDATE_SIGNING_KEY_B64 value:")
    print(base64.b64encode(raw_private).decode("ascii"))
    print("Public key base64 (embed in update_service.py):")
    print(base64.b64encode(raw_public).decode("ascii"))


if __name__ == "__main__":
    main()
