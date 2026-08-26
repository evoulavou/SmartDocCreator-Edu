import os
import unicodedata

signed_folder = "to_sign"
metadata_folder = "notifications"

for filename in os.listdir(signed_folder):

    if not filename.lower().endswith("_signed.pdf"):
        continue

    name_without_signed = filename[:-len("_signed.pdf")]
    base_name = name_without_signed.rsplit("_", 1)[0]
    expected_docx = base_name + ".docx"

    # βρίσκουμε ΑΜ
    parts = base_name.split("_")
    numbers = [x for x in parts if x.isdigit()]

    if not numbers:
        continue

    am = numbers[-1]

    for actual_docx in os.listdir(metadata_folder):

        if not actual_docx.lower().endswith(".docx"):
            continue

        if am not in actual_docx:
            continue

        print("EXPECTED:")
        print(repr(expected_docx))

        print("\nACTUAL:")
        print(repr(actual_docx))

        print("\nΙΣΑ ΚΑΝΟΝΙΚΑ;")
        print(expected_docx == actual_docx)

        print("\nΙΣΑ ΜΕ NFC NORMALIZATION;")
        print(
            unicodedata.normalize("NFC", expected_docx)
            ==
            unicodedata.normalize("NFC", actual_docx)
        )

        print("\nΔΙΑΦΟΡΕΣ ΧΑΡΑΚΤΗΡΩΝ:")

        e = expected_docx
        a = actual_docx

        for i in range(max(len(e), len(a))):

            c1 = e[i] if i < len(e) else "<END>"
            c2 = a[i] if i < len(a) else "<END>"

            if c1 != c2:

                if c1 != "<END>":
                    info1 = (
                        repr(c1),
                        f"U+{ord(c1):04X}",
                        unicodedata.name(c1, "UNKNOWN")
                    )
                else:
                    info1 = ("<END>", "", "")

                if c2 != "<END>":
                    info2 = (
                        repr(c2),
                        f"U+{ord(c2):04X}",
                        unicodedata.name(c2, "UNKNOWN")
                    )
                else:
                    info2 = ("<END>", "", "")

                print(i, info1, "!=", info2)

        print("\n" + "=" * 70)
