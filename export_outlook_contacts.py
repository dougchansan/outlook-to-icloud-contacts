#!/usr/bin/env python3
"""
Step 1: Export contacts from Microsoft Outlook for Mac as vCards via AppleScript.
Saves to outlook_contacts.vcf in the current directory.
"""

import subprocess
import sys


def main():
    print("Exporting contacts from Microsoft Outlook...")

    script = '''
    tell application "Microsoft Outlook"
        set allContacts to every contact
        set vcardOutput to ""
        repeat with c in allContacts
            try
                set vcardOutput to vcardOutput & (vcard data of c) & linefeed
            end try
        end repeat
        return vcardOutput
    end tell
    '''

    result = subprocess.run(
        ["osascript", "-e", script],
        capture_output=True, text=True, timeout=300
    )

    if result.returncode != 0:
        print(f"Error: {result.stderr.strip()}", file=sys.stderr)
        sys.exit(1)

    vcf_path = "outlook_contacts.vcf"
    with open(vcf_path, "w") as f:
        f.write(result.stdout)

    count = result.stdout.lower().count("begin:vcard")
    print(f"Exported {count} contacts to {vcf_path}")


if __name__ == "__main__":
    main()
