import os

folder = "to_sign"

print("=== ΠΡΑΓΜΑΤΙΚΑ FILENAMES ΑΠΟ ΤΑ WINDOWS ===")
print()

for filename in os.listdir(folder):

    if "_signed" not in filename.lower():
        continue

    print("ΟΝΟΜΑ:")
    print(filename)

    print("\nREPR:")
    print(repr(filename))

    print("\nΜΗΚΟΣ:")
    print(len(filename))

    print("\nΤΕΛΕΥΤΑΙΟΙ 30 ΧΑΡΑΚΤΗΡΕΣ:")
    print(repr(filename[-30:]))

    print("\nUNICODE:")
    for i, char in enumerate(filename):
        print(
            i,
            repr(char),
            f"U+{ord(char):04X}"
        )

    print("\n" + "=" * 70 + "\n")
