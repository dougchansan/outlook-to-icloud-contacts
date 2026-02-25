# Outlook to iCloud Contacts

Export contacts from Microsoft Outlook for Mac and import them into macOS Contacts.app (syncs to iCloud) with deduplication.

## Requirements

- macOS
- Microsoft Outlook for Mac (local app)
- Python 3
- Contacts.app

## Usage

### Step 1: Export from Outlook

```bash
python3 export_outlook_contacts.py
```

Exports all Outlook contacts to `outlook_contacts.vcf` via AppleScript.

### Step 2: Import to Contacts.app

```bash
# Preview what will be imported (no changes made)
python3 import_to_contacts_app.py --dry-run

# Import into default "GM" group
python3 import_to_contacts_app.py

# Import into a custom group
python3 import_to_contacts_app.py --group "Work"

# Use a different vCard file
python3 import_to_contacts_app.py --vcf my_contacts.vcf
```

## Deduplication

Contacts are matched against existing Contacts.app entries by:

1. **Email address** (exact match, case-insensitive)
2. **Phone number** (last 10 digits, ignoring formatting)
3. **Full name** (exact match, case-insensitive)

Duplicates are skipped. Existing contacts are never modified.

## Syncing to Other Devices

Once imported into Contacts.app, contacts sync via iCloud to any device signed into the same Apple ID. To sync to a work phone with a different Apple ID, add your iCloud account as a secondary account on that device (Settings > Contacts > Accounts).
