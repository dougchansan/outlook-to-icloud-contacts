#!/usr/bin/env python3
"""
Step 2: Import Outlook vCard contacts into macOS Contacts.app under a specified group.
Deduplicates against existing contacts by name, email, and phone number.

Usage:
    python3 import_to_contacts_app.py [--group GROUP] [--dry-run] [--vcf FILE]

Options:
    --group GROUP   Contacts.app group name to import into (default: "GM")
    --dry-run       Show what would be imported without making changes
    --vcf FILE      Path to vCard file (default: outlook_contacts.vcf)
"""

import argparse
import re
import subprocess
import sys


def load_existing_contacts():
    """Load existing contacts from Contacts.app for deduplication."""
    script = '''
    tell application "Contacts"
        set allPeople to every person
        set output to ""
        repeat with p in allPeople
            set pName to name of p
            set pEmails to ""
            try
                repeat with e in emails of p
                    set pEmails to pEmails & (value of e) & ";"
                end repeat
            end try
            set pPhones to ""
            try
                repeat with ph in phones of p
                    set pPhones to pPhones & (value of ph) & ";"
                end repeat
            end try
            set output to output & pName & "\t" & pEmails & "\t" & pPhones & linefeed
        end repeat
        return output
    end tell
    '''

    result = subprocess.run(
        ["osascript", "-e", script],
        capture_output=True, text=True, timeout=120
    )

    names = set()
    emails = set()
    phones = set()

    for line in result.stdout.strip().split("\n"):
        parts = line.split("\t")
        if not parts or not parts[0]:
            continue
        names.add(parts[0].strip().lower())
        if len(parts) > 1:
            for email in parts[1].split(";"):
                email = email.strip().lower()
                if email:
                    emails.add(email)
        if len(parts) > 2:
            for phone in parts[2].split(";"):
                digits = re.sub(r"\D", "", phone)
                if len(digits) >= 7:
                    phones.add(digits[-10:])

    return names, emails, phones


def parse_vcards(filepath):
    """Parse a multi-vCard file into individual vCard strings."""
    vcards = []
    current = []
    with open(filepath, "r") as f:
        for line in f:
            line = line.rstrip()
            if line.lower() == "begin:vcard":
                current = [line]
            elif line.lower() == "end:vcard":
                current.append(line)
                vcards.append("\n".join(current))
                current = []
            elif current:
                current.append(line)
    return vcards


def extract_vcard_fields(vc):
    """Extract structured fields from a vCard string."""
    info = {}

    fn_match = re.search(r"fn;[^:]*:(.+)", vc, re.IGNORECASE)
    info["fn"] = fn_match.group(1).strip() if fn_match else ""

    n_match = re.search(r"n;[^:]*:([^;\n]*);([^;\n]*)", vc, re.IGNORECASE)
    if n_match:
        info["last_name"] = n_match.group(1).strip()
        info["first_name"] = n_match.group(2).strip()
    else:
        parts = info["fn"].split(None, 1)
        info["first_name"] = parts[0] if parts else ""
        info["last_name"] = parts[1] if len(parts) > 1 else ""

    org_match = re.search(r"org;[^:]*:([^;\n]+)", vc, re.IGNORECASE)
    info["company"] = org_match.group(1).strip() if org_match else ""

    title_match = re.search(r"title;[^:]*:(.+)", vc, re.IGNORECASE)
    info["job_title"] = title_match.group(1).strip() if title_match else ""

    info["emails"] = []
    for m in re.finditer(r"email;([^:]*):(.+)", vc, re.IGNORECASE):
        params = m.group(1).lower()
        addr = m.group(2).strip()
        label = "work" if "work" in params else ("home" if "home" in params else "other")
        info["emails"].append((label, addr))

    info["phones"] = []
    for m in re.finditer(r"tel;([^:]*):(.+)", vc, re.IGNORECASE):
        params = m.group(1).lower()
        number = m.group(2).strip()
        if "cell" in params or "mobile" in params:
            label = "mobile"
        elif "work" in params:
            label = "work"
        elif "home" in params:
            label = "home"
        elif "fax" in params:
            label = "work fax"
        elif "pager" in params:
            label = "pager"
        else:
            label = "other"
        info["phones"].append((label, number))

    note_match = re.search(r"note;[^:]*:(.+)", vc, re.IGNORECASE)
    info["note"] = note_match.group(1).strip().replace("\\n", "\n") if note_match else ""
    if info["note"].strip() == "":
        info["note"] = ""

    return info


def is_duplicate(info, existing_names, existing_emails, existing_phones):
    """Check if a contact already exists by email, phone, or exact name."""
    for _, email in info["emails"]:
        if email.lower() in existing_emails:
            return True, f"email: {email}"
    for _, phone in info["phones"]:
        digits = re.sub(r"\D", "", phone)
        if len(digits) >= 7 and digits[-10:] in existing_phones:
            return True, f"phone: {phone}"
    if info["fn"] and info["fn"].lower() in existing_names:
        return True, f"name: {info['fn']}"
    return False, ""


def escape_applescript(s):
    """Escape a string for AppleScript."""
    return s.replace("\\", "\\\\").replace('"', '\\"')


def build_contact_applescript(info):
    """Build AppleScript code to create a single contact."""
    first = escape_applescript(info["first_name"])
    last = escape_applescript(info["last_name"])
    company = escape_applescript(info["company"])
    title = escape_applescript(info["job_title"])
    note = escape_applescript(info["note"])

    props = []
    if first:
        props.append(f'first name:"{first}"')
    if last:
        props.append(f'last name:"{last}"')
    if company:
        props.append(f'organization:"{company}"')
    if title:
        props.append(f'job title:"{title}"')
    if note:
        props.append(f'note:"{note}"')

    if not props:
        return None

    lines = [f"set newPerson to make new person with properties {{{', '.join(props)}}}"]

    for label, addr in info["emails"]:
        lines.append(
            f'make new email at end of emails of newPerson '
            f'with properties {{label:"{escape_applescript(label)}", '
            f'value:"{escape_applescript(addr)}"}}'
        )

    for label, number in info["phones"]:
        lines.append(
            f'make new phone at end of phones of newPerson '
            f'with properties {{label:"{escape_applescript(label)}", '
            f'value:"{escape_applescript(number)}"}}'
        )

    lines.append("add newPerson to gmGroup")
    return "\n        ".join(lines)


def ensure_group_exists(group_name):
    """Create the Contacts.app group if it doesn't exist."""
    script = f'''
    tell application "Contacts"
        try
            set gmGroup to group "{group_name}"
        on error
            set gmGroup to make new group with properties {{name:"{group_name}"}}
            save
        end try
        return name of gmGroup
    end tell
    '''
    result = subprocess.run(
        ["osascript", "-e", script],
        capture_output=True, text=True, timeout=15
    )
    if result.returncode != 0:
        print(f"Error creating group: {result.stderr.strip()}", file=sys.stderr)
        sys.exit(1)


def reassign_group(group_name):
    """
    Re-assign recently created contacts to the group.
    Works around a Contacts.app bug where group assignment in the same
    AppleScript call as contact creation doesn't persist.
    """
    script = f'''
    tell application "Contacts"
        set gmGroup to group "{group_name}"
        set recentPeople to every person whose creation date > (current date) - 1 * hours
        set addCount to 0
        repeat with p in recentPeople
            add p to gmGroup
            set addCount to addCount + 1
        end repeat
        save
        return addCount as text
    end tell
    '''
    result = subprocess.run(
        ["osascript", "-e", script],
        capture_output=True, text=True, timeout=120
    )
    if result.returncode == 0:
        return int(result.stdout.strip())
    return 0


def main():
    parser = argparse.ArgumentParser(
        description="Import Outlook contacts into macOS Contacts.app"
    )
    parser.add_argument("--group", default="GM", help="Contacts group name (default: GM)")
    parser.add_argument("--dry-run", action="store_true", help="Preview without importing")
    parser.add_argument("--vcf", default="outlook_contacts.vcf", help="Path to vCard file")
    args = parser.parse_args()

    print("Loading existing contacts for deduplication...")
    existing_names, existing_emails, existing_phones = load_existing_contacts()
    print(f"  {len(existing_names)} names, {len(existing_emails)} emails, {len(existing_phones)} phones")

    print(f"Parsing vCards from {args.vcf}...")
    vcards = parse_vcards(args.vcf)
    print(f"  {len(vcards)} vCards found")

    to_import = []
    dupes = []
    blanks = 0

    for vc in vcards:
        info = extract_vcard_fields(vc)
        if not info["fn"] and not info["emails"] and not info["phones"]:
            blanks += 1
            continue
        is_dupe, reason = is_duplicate(info, existing_names, existing_emails, existing_phones)
        if is_dupe:
            dupes.append((info["fn"], reason))
            continue
        to_import.append(info)

    print(f"\n  Duplicates: {len(dupes)}")
    print(f"  Blank/empty: {blanks}")
    print(f"  New to import: {len(to_import)}")

    if dupes:
        print("\nSkipping duplicates:")
        for name, reason in sorted(dupes):
            print(f"  SKIP: {name} ({reason})")

    if not to_import:
        print("\nNothing to import.")
        return

    if args.dry_run:
        print(f"\n[DRY RUN] Would import {len(to_import)} contacts into '{args.group}':")
        for info in sorted(to_import, key=lambda x: x["fn"]):
            print(f"  + {info['fn']}")
        return

    ensure_group_exists(args.group)

    print(f"\nImporting {len(to_import)} contacts into '{args.group}'...")
    batch_size = 10
    imported = 0
    failed = 0

    for i in range(0, len(to_import), batch_size):
        batch = to_import[i : i + batch_size]
        contact_blocks = []
        for info in batch:
            block = build_contact_applescript(info)
            if block:
                contact_blocks.append(block)

        if not contact_blocks:
            continue

        contacts_code = "\n        ".join(contact_blocks)
        script = f'''
tell application "Contacts"
    set gmGroup to group "{args.group}"
    try
        {contacts_code}
        save
        return "ok"
    on error errMsg
        return "ERROR: " & errMsg
    end try
end tell
'''
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True, text=True, timeout=60,
        )
        if result.returncode == 0 and result.stdout.strip() == "ok":
            imported += len(contact_blocks)
            batch_num = i // batch_size + 1
            total_batches = (len(to_import) + batch_size - 1) // batch_size
            print(f"  Batch {batch_num}/{total_batches}: {len(contact_blocks)} imported ({imported} total)")
        else:
            batch_num = i // batch_size + 1
            print(f"  Batch {batch_num} failed, trying individually...")
            for info in batch:
                block = build_contact_applescript(info)
                if not block:
                    continue
                single = f'''
tell application "Contacts"
    set gmGroup to group "{args.group}"
    try
        {block}
        save
        return "ok"
    on error errMsg
        return "ERROR: " & errMsg
    end try
end tell
'''
                r = subprocess.run(
                    ["osascript", "-e", single],
                    capture_output=True, text=True, timeout=30,
                )
                if r.returncode == 0 and r.stdout.strip() == "ok":
                    imported += 1
                else:
                    failed += 1
                    print(f"    FAILED: {info['fn']}")

    # Re-assign group membership (workaround for Contacts.app persistence bug)
    print("\nAssigning group membership...")
    assigned = reassign_group(args.group)
    print(f"  {assigned} contacts assigned to '{args.group}'")

    print(f"\nDone! Imported {imported}, failed {failed}.")


if __name__ == "__main__":
    main()
