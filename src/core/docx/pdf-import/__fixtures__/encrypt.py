# Encrypts copies of lettre.pdf (a small letter the Word PDF export of this repository wrote) with pypdf (an independent implementation of the PDF standard
# security handler), so the reader's decryption is checked against another writer.
# Run from this folder: python encrypt.py
from pypdf import PdfReader, PdfWriter
from pypdf.constants import UserAccessPermissions

VARIANTS = [
    # name, user password, owner password, algorithm
    ("lettre-rc4-owner-only", "", "proprietaire", "RC4-128"),
    ("lettre-aes128", "secret", "proprietaire", "AES-128"),
    ("lettre-aes256", "secret", "proprietaire", "AES-256"),
]

for name, user, owner, algorithm in VARIANTS:
    reader = PdfReader("lettre.pdf")
    writer = PdfWriter(clone_from=reader)
    # Print and copy allowed, editing not (bits 3 and 5).
    writer.encrypt(user_password=user, owner_password=owner, algorithm=algorithm, permissions_flag=UserAccessPermissions.PRINT | UserAccessPermissions.EXTRACT)
    with open(f"{name}.pdf", "wb") as out:
        writer.write(out)
    print(name)
