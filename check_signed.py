import os

signed_folder = "to_sign"
metadata_folder = "notifications"

print("=== ΕΛΕΓΧΟΣ ΑΝΤΙΣΤΟΙΧΙΣΗΣ ===")
print()

for filename in os.listdir(signed_folder):

    if not filename.lower().endswith("_signed.pdf"):
        continue

    print("SIGNED:")
    print(repr(filename))

    # Αυτό ακριβώς κάνει τώρα το sort.py
    name_without_signed = filename[:-len("_signed.pdf")]
    base_name = name_without_signed.rsplit("_", 1)[0]

    expected_docx = base_name + ".docx"

    print()
    print("ΨΑΧΝΕΙ DOCX:")
    print(repr(expected_docx))

    docx_path = os.path.join(
        metadata_folder,
        expected_docx
    )

    print()
    print("ΥΠΑΡΧΕΙ;", os.path.exists(docx_path))

    # Βρες DOCX που περιέχει τον ίδιο ΑΜ
    parts = base_name.split("_")

    numbers = [
        x for x in parts
        if x.isdigit()
    ]

    if numbers:

        am = numbers[-1]

        print()
        print("DOCX ΜΕ ΙΔΙΟ ΑΜ:")

        for docx in os.listdir(metadata_folder):

            if (
                docx.lower().endswith(".docx")
                and am in docx
            ):
                print(repr(docx))

    print()
    print("=" * 70)
    print()
