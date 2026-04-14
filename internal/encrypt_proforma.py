"""Encrypt the pro forma xlsx with AES-256-GCM for client-side WebCrypto decryption.

Output format: [16 bytes salt][12 bytes IV][ciphertext + GCM auth tag]
Client decrypts with PBKDF2-SHA256 (100k iterations) + AES-GCM.
"""
import os
import sys
from pathlib import Path
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

HERE = Path(__file__).parent
SRC = HERE.parent / "deliverables" / "1601-central-proforma.xlsx"
OUT = HERE.parent / "deliverables" / "1601-central-proforma.xlsx.enc"

PASSWORD = sys.argv[1].encode() if len(sys.argv) > 1 else b"co-llab"
ITERATIONS = 100_000

salt = os.urandom(16)
iv = os.urandom(12)
plaintext = SRC.read_bytes()

kdf = PBKDF2HMAC(
    algorithm=hashes.SHA256(),
    length=32,
    salt=salt,
    iterations=ITERATIONS,
)
key = kdf.derive(PASSWORD)

aesgcm = AESGCM(key)
ciphertext = aesgcm.encrypt(iv, plaintext, None)

OUT.write_bytes(salt + iv + ciphertext)
print(f"encrypted {len(plaintext):,} bytes → {len(salt + iv + ciphertext):,} bytes")
print(f"salt: {salt.hex()}")
print(f"iv:   {iv.hex()}")
print(f"out:  {OUT}")
