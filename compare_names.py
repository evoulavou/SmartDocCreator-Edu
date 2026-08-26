name1 = "notification_s1_Διάθεση οργ.υπεράριθμων_for_ΓΑΚΟΠΟΥΛΟΣ_200301_Εισαγωγή_GNGJD_signed.pdf"
name2 = "ΒΑΛΕ_ΕΔΩ_ΤΟ_ΠΡΑΓΜΑΤΙΚΟ_ΟΝΟΜΑ_ΤΟΥ_ΥΠΟΓΕΓΡΑΜΜΕΝΟΥ.pdf"

print("Name1:", repr(name1))
print("Name2:", repr(name2))
print()

for i, (a, b) in enumerate(zip(name1, name2)):
    if a != b:
        print(
            f"Διαφορά στη θέση {i}: "
            f"{repr(a)} U+{ord(a):04X} "
            f"!= {repr(b)} U+{ord(b):04X}"
        )

if len(name1) != len(name2):
    print()
    print("Διαφορετικό μήκος:")
    print("name1:", len(name1))
    print("name2:", len(name2))
