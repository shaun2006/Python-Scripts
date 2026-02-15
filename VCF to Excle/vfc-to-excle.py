import pandas as pd
import argparse


def parse_vcf(vcf_file):
    contacts = []

    with open(vcf_file, 'r', encoding='utf-8') as file:
        contact = {}

        for line in file:
            line = line.strip()

            if line.startswith("BEGIN:VCARD"):
                contact = {}

            elif line.startswith("FN:"):
                contact["Full Name"] = line.replace("FN:", "")

            elif line.startswith("TEL"):
                phone = line.split(":")[-1]
                contact.setdefault("Phone Numbers", []).append(phone)

            elif line.startswith("EMAIL"):
                email = line.split(":")[-1]
                contact.setdefault("Emails", []).append(email)

            elif line.startswith("ORG:"):
                contact["Organization"] = line.replace("ORG:", "")

            elif line.startswith("ADR"):
                address = line.split(":")[-1].replace(";", " ")
                contact["Address"] = address

            elif line.startswith("END:VCARD"):
                contact["Phone Numbers"] = ", ".join(contact.get("Phone Numbers", []))
                contact["Emails"] = ", ".join(contact.get("Emails", []))
                contacts.append(contact)

    return contacts


def main():
    parser = argparse.ArgumentParser(description="Convert VCF to Excel")
    parser.add_argument("input", help="Input VCF file")
    parser.add_argument("-o", "--output", help="Output Excel file")

    args = parser.parse_args()

    input_file = args.input
    output_file = args.output if args.output else input_file.replace(".vcf", ".xlsx")

    contacts = parse_vcf(input_file)
    df = pd.DataFrame(contacts)
    df.to_excel(output_file, index=False)

    print(f"✅ Conversion complete! Saved as {output_file}")


if __name__ == "__main__":
    main()

