#!/usr/bin/env python3
import argparse
import pandas as pd

SUPERSCRIPT_MAP = str.maketrans("0123456789", "⁰¹²³⁴⁵⁶⁷⁸⁹")

def superscript(num: int) -> str:
    return str(num).translate(SUPERSCRIPT_MAP)

def clean_name(name: str) -> str:
    return " ".join(str(name).split())  # trim whitespace

def split_name(fullname: str):
    parts = fullname.strip().split()
    if len(parts) == 1:
        return fullname, ""  # unlikely
    first = " ".join(parts[:-1])
    last = parts[-1]
    return first, last

def format_name(name, aff, is_pi, mode, offset=0):
    aff_number = aff + offset
    pi_mark = "*" if is_pi else ""
    name = clean_name(name)

    if mode == "github":
        return f"{name}{pi_mark}<sup>{aff_number}</sup>"
    elif mode == "plain":
        return f"{name}{pi_mark}{superscript(aff_number)}"
    else:
        return ""  # LaTeX handled separately

def main():
    parser = argparse.ArgumentParser(description="Format CF GMS Consortium author list.")
    parser.add_argument("excel_file", help="Path to canadian_cf_gms_collaborators.xlsx")
    parser.add_argument("--pi", action="store_true", help="Show only site PIs")
    parser.add_argument("--latex", action="store_true", help="Output in LaTeX format")
    parser.add_argument("--github", action="store_true", help="Output in GitHub/Markdown format")
    parser.add_argument("--start", type=int, default=0, help="Offset for affiliation numbers")
    args = parser.parse_args()

    # Decide mode
    mode = "plain"
    if args.latex:
        mode = "latex"
    elif args.github:
        mode = "github"

    # Load sheets
    current = pd.read_excel(args.excel_file, sheet_name="current")
    sites = pd.read_excel(args.excel_file, sheet_name="sites")
    former = pd.read_excel(args.excel_file, sheet_name="former")

    # Filter based on --pi flag
    if args.pi:
        current = current[current["PI"].astype(str).str.strip() == "*"]
        former = former[former["PI"].astype(str).str.strip() == "*"]

    if mode == "latex":
        # Consortium
        print("Canadian Cystic Fibrosis Gene Modifier Consortium\n")
        # Authors
        for _, row in current.iterrows():
            name = clean_name(row["Collaborator Name"])
            first, last = split_name(name)
            aff = int(row["Affiliation #"]) + args.start
            is_pi = str(row.get("PI", "")).strip() == "*"
            star = "" if args.pi else ("*" if is_pi else "")

            line = (
                f"\\author{star}[{aff}]{{\\fnm{{{first}}} \\sur{{{last}}}}}\\email{{}}"
            )
            print(line)

        if not args.pi:
                print ("\n* indicates site Principal Investigator\n")

        # Affiliations
        for _, row in sites.iterrows():
            aff = int(row["Affiliation #"]) + args.start
            affil_line = (
                f"\\affil[{aff}]{{\\orgname{{{row['Name']}}}, "
                f"\\orgaddress{{\\city{{{row['City']}}}, "
                f"\\state{{{row['Province']}}}, "
                f"\\country{{{row['Country']}}}}}" + "}"
            )
            print(affil_line)

    elif mode == "github":
        formatted_current = []
        for _, row in current.iterrows():
            aff = int(row["Affiliation #"])
            is_pi = str(row.get("PI", "")).strip() == "*"
            formatted = format_name(
                row["Collaborator Name"], aff,
                (not args.pi and is_pi),  # mark with * only if not filtering
                mode, args.start
            )
            formatted_current.append(formatted)

        print("Canadian Cystic Fibrosis Gene Modifier Consortium\n")
        print(", ".join(formatted_current))

        if not former.empty:
            formatted_former = []
            for _, row in former.iterrows():
                aff = int(row["Affiliation #"])
                is_pi = str(row.get("PI", "")).strip() == "*"
                formatted = format_name(
                    row["Collaborator Name"], aff,
                    (not args.pi and is_pi),
                    mode, args.start
                )
                formatted_former.append(formatted)

            print("\n## Past contributors\n")
            print(", ".join(formatted_former))

        print("\n---\n")

        # Affiliations
        shield_colors = {
            "ON": "d7212e",
            "AB": "39b577",
            "BC": "2d91cc",
            "SK": "f0e64a",
            "MB": "ee5934",
            "QC": "dc0b89",
            "NB": "965ca0",
            "NS": "ac2188",
            "NL": "752d8d",
        }
        for _, row in sites.iterrows():
            aff = int(row["Affiliation #"]) + args.start
            province = row["Province"]
            city = row["City"]
            country = row["Country"]
            name = row["Name"]
            color = shield_colors.get(province, "555555")
            badge = f"![{province}](https://img.shields.io/badge/{aff}-{province}-%23{color})"
            print(f"{badge} {name}, {city}, {province}, {country}.  ")

    else:  # plain text mode
        formatted_current = []
        for _, row in current.iterrows():
            aff = int(row["Affiliation #"])
            is_pi = str(row.get("PI", "")).strip() == "*"
            formatted = format_name(
                row["Collaborator Name"], aff,
                (not args.pi and is_pi),
                mode, args.start
            )
            formatted_current.append(formatted)

        print("Canadian Cystic Fibrosis Gene Modifier Consortium\n")
        print(", ".join(formatted_current))

        if not args.pi:
                print ("\n* indicates site Principal Investigator")
        print("\nAffiliations\n")
        for _, row in sites.iterrows():
            aff = int(row["Affiliation #"]) + args.start
            aff_text = f"{aff}. {row['Name']}, {row['City']}, {row['Province']}, {row['Country']}."
            print(aff_text)


if __name__ == "__main__":
    main()

