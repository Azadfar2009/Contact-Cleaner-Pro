#!/usr/bin/env python3
"""
Contact Cleaner Pro
Organize, deduplicate & normalize your contacts.
Made with ❤️ by Mohammad Azadfar | +98 912 419 8445
Intelligence powered by DeepSeek AI
"""

import csv
import re
import sys
import os
from collections import defaultdict
from pathlib import Path

# ---------- colorama with Windows console fix ----------
try:
    import colorama
    from colorama import Fore, Style, init, just_fix_windows_console
    just_fix_windows_console()
    init(autoreset=True)
    COLORS_OK = True
except ImportError:
    class Fore:
        RED = GREEN = YELLOW = BLUE = CYAN = MAGENTA = LIGHTCYAN_EX = LIGHTMAGENTA_EX = RESET = ''
    class Style:
        BRIGHT = NORMAL = RESET_ALL = ''
    COLORS_OK = False

# ------------------------- settings -------------------------
OUTPUT_CSV = "cleaned_contacts.csv"
OUTPUT_XLSX = "cleaned_contacts.xlsx"
OUTPUT_VCF = "cleaned_contacts.vcf"
LOG_FILE = "changes_log.txt"
# ------------------------------------------------------------

# ---------- Iranian prefix tables ----------
MOBILE_PREFIXES = {
    '910','911','912','913','914','915','916','917','918','919',
    '990','991','992','993','994',
    '900','901','902','903','904','905','930','933','935','936','937','938','939','941',
    '920','921','922','923',
    '999'
}

LANDLINE_CODES = {
    '21','31','41','44','11','51','71','61','34','54','45',
    '84','26','77','38','24','23','28','25','87','83','74',
    '17','13','66','86','76','81','35','58','56'
}

# ---------- helpers ----------

def color_text(text, color=Fore.RESET, style=Style.NORMAL):
    if COLORS_OK:
        return f"{style}{color}{text}{Style.RESET_ALL}"
    return text


def clean_raw_phone(raw):
    if not raw:
        return ""
    s = str(raw).strip()
    if s.startswith('+'):
        cleaned = '+' + re.sub(r'[^\d]', '', s[1:])
    else:
        cleaned = re.sub(r'[^\d]', '', s)
    if re.match(r'^9\d{9}$', cleaned):
        cleaned = '0' + cleaned
    return cleaned


def to_normalized_for_compare(phone):
    cleaned = clean_raw_phone(phone)
    if not cleaned:
        return ""
    if cleaned.startswith('+'):
        return cleaned
    if cleaned.startswith('0'):
        return '+98' + cleaned[1:]
    if cleaned.startswith('98') and len(cleaned) > 2:
        return '+' + cleaned
    return cleaned


def is_phone_like(phone):
    if not phone:
        return False
    cleaned = clean_raw_phone(phone)
    return bool(re.match(r'^\+?\d{7,15}$', cleaned))


def is_iran_mobile(normalized):
    if not normalized.startswith('+98'):
        return False
    digits = normalized[1:]
    if len(digits) == 12:
        return digits[2:5] in MOBILE_PREFIXES
    return False


def is_iran_landline(normalized):
    if not normalized.startswith('+98'):
        return False
    digits = normalized[1:]
    if len(digits) == 12:
        return digits[2:4] in LANDLINE_CODES
    if len(digits) == 13:
        return digits[2:5] in LANDLINE_CODES
    return False


def classify_phone(normalized):
    if is_iran_mobile(normalized):
        return 'Mobile'
    if is_iran_landline(normalized):
        return 'Landline'
    return 'Other'


def looks_like_mobile(phone):
    c = clean_raw_phone(phone)
    if not c:
        return False
    if re.match(r'^09\d{9}$', c):
        return True
    if re.match(r'^\+989\d{9}$', c):
        return True
    return False


def normalize_with_country_code(phone, country_code):
    if not phone:
        return phone
    if phone.startswith('+') and not phone.startswith('+98'):
        return phone
    if phone.startswith('0'):
        return '+' + country_code + phone[1:]
    if phone.startswith('+98'):
        return '+' + country_code + phone[4:]
    return phone


def deduplicate_phones(raw_list, norm_list):
    seen = set()
    new_raw, new_norm = [], []
    for r, n in zip(raw_list, norm_list):
        if n not in seen:
            seen.add(n)
            new_raw.append(r)
            new_norm.append(n)
    return new_raw, new_norm


# ---------- file detection & parsing ----------

def find_input_files():
    files = os.listdir('.')
    csv_files = [f for f in files if f.lower().endswith('.csv')]
    vcf_files = [f for f in files if f.lower().endswith('.vcf')]
    return csv_files, vcf_files


def choose_file(csv_files, vcf_files):
    all_files = csv_files + vcf_files
    if not all_files:
        print(color_text("No CSV or VCF file found.", Fore.RED))
        sys.exit(1)
    if len(all_files) == 1:
        return all_files[0]
    print(color_text("Multiple contact files found:", Fore.CYAN))
    for i, f in enumerate(all_files, 1):
        print(f"  {i}. {f}")
    while True:
        try:
            choice = int(input(color_text("Choose a file number: ", Fore.GREEN)))
            if 1 <= choice <= len(all_files):
                return all_files[choice - 1]
        except ValueError:
            pass
        print(color_text("Invalid choice.", Fore.YELLOW))


def parse_csv(filename):
    with open(filename, encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        fieldnames = reader.fieldnames
        if not fieldnames:
            print(color_text("CSV file is empty.", Fore.RED))
            sys.exit(1)
        rows = list(reader)

    name_col = phone_col = None
    for col in fieldnames:
        low = col.lower()
        if not name_col and ('name' in low or 'first' in low or 'نام' in low):
            name_col = col
        if not phone_col and ('phone' in low or 'mobile' in low or 'tel' in low or 'تلفن' in low):
            phone_col = col

    if not name_col or not phone_col:
        print(color_text("Detected columns:", Fore.CYAN), fieldnames)
        if not name_col:
            name_col = input(color_text("Enter name column: ", Fore.GREEN)).strip()
        if not phone_col:
            phone_col = input(color_text("Enter phone column: ", Fore.GREEN)).strip()

    if name_col not in fieldnames or phone_col not in fieldnames:
        print(color_text("Column not found.", Fore.RED))
        sys.exit(1)

    other_cols = [c for c in fieldnames if c not in (name_col, phone_col)]
    contacts = []
    for row in rows:
        name = row.get(name_col, '').strip()
        if not name:
            continue
        raw_phone = row.get(phone_col, '').strip()
        cleaned = clean_raw_phone(raw_phone)
        if cleaned:
            phones_raw = [cleaned]
            phones_norm = [to_normalized_for_compare(cleaned)]
        else:
            phones_raw = []
            phones_norm = []
        phones_raw, phones_norm = deduplicate_phones(phones_raw, phones_norm)
        extra = {col: row[col].strip() for col in other_cols if row.get(col, '').strip()}
        contacts.append({
            'name': name,
            'phones_raw': phones_raw,
            'phones_norm': phones_norm,
            'extra': extra
        })
    return contacts


def parse_vcf(filename):
    with open(filename, encoding='utf-8') as f:
        content = f.read()

    vcards = re.split(r'END:VCARD\s*\n?', content, flags=re.IGNORECASE)
    contacts = []
    for block in vcards:
        if not block.strip() or 'BEGIN:VCARD' not in block.upper():
            continue
        name = ''
        fn_match = re.search(r'^FN:(.*)', block, re.MULTILINE | re.IGNORECASE)
        if fn_match:
            name = fn_match.group(1).strip()
        else:
            n_match = re.search(r'^N:(.*)', block, re.MULTILINE | re.IGNORECASE)
            if n_match:
                parts = n_match.group(1).split(';')
                name = ' '.join(filter(None, parts[:3]))
        if not name:
            name = 'Unknown'

        raw_phones = []
        norm_phones = []
        for tel_match in re.finditer(r'^TEL(?:;[^:]*)*:([^\n\r]+)', block, re.MULTILINE | re.IGNORECASE):
            raw = tel_match.group(1).strip()
            cleaned = clean_raw_phone(raw)
            if cleaned:
                raw_phones.append(cleaned)
                norm_phones.append(to_normalized_for_compare(cleaned))
        raw_phones, norm_phones = deduplicate_phones(raw_phones, norm_phones)

        extra = {}
        for field in ['EMAIL', 'ADR', 'ORG', 'TITLE', 'NOTE', 'BDAY']:
            vals = re.findall(rf'^{field}(?:;[^:]*)*:(.*)', block, re.MULTILINE | re.IGNORECASE)
            if vals:
                extra[field] = '; '.join(v.strip() for v in vals)

        contacts.append({
            'name': name,
            'phones_raw': raw_phones,
            'phones_norm': norm_phones,
            'extra': extra
        })
    return contacts


# ---------- analysis ----------

def run_analysis(contacts):
    name_groups = defaultdict(list)
    phone_map = defaultdict(list)
    exact_dup_map = defaultdict(list)
    non_phone_indices = []
    extra_contacts = []

    for idx, c in enumerate(contacts):
        name = c['name']
        norms = c['phones_norm']
        name_groups[name].append(idx)

        if not norms or not any(is_phone_like(p) for p in c['phones_raw']):
            non_phone_indices.append(idx)

        unique_norms = list(dict.fromkeys(norms))
        for p in unique_norms:
            if p:
                phone_map[p].append(idx)
                exact_dup_map[(name, p)].append(idx)

        if c['extra']:
            extra_contacts.append((idx, name, c['extra']))

    duplicate_names = {n: idxs for n, idxs in name_groups.items() if len(idxs) > 1}
    duplicate_phones = {p: idxs for p, idxs in phone_map.items() if len(idxs) > 1}
    exact_duplicates = {k: idxs for k, idxs in exact_dup_map.items() if len(idxs) > 1}

    return {
        'total': len(contacts),
        'duplicate_names': duplicate_names,
        'duplicate_phones': duplicate_phones,
        'exact_duplicates': exact_duplicates,
        'non_phone_indices': non_phone_indices,
        'extra_contacts': extra_contacts
    }


def show_summary(analysis):
    print(color_text("=" * 50, Fore.CYAN))
    print(color_text("CURRENT STATUS", Fore.CYAN + Style.BRIGHT))
    print(f"Total contacts: {analysis['total']}")
    print(f"Duplicate name groups: {len(analysis['duplicate_names'])}")
    print(f"Duplicate phone numbers: {len(analysis['duplicate_phones'])}")
    print(f"Exact duplicate entries: {len(analysis['exact_duplicates'])}")
    print(f"Contacts with extra fields: {len(analysis['extra_contacts'])}")
    print(f"Entries with no valid phone: {len(analysis['non_phone_indices'])}")
    print(color_text("=" * 50, Fore.CYAN))
    print()


# ---------- interactive menus ----------

def main_menu():
    print()
    print()
    print(color_text("=" * 50, Fore.CYAN))
    print()
    print()
    print(color_text("MAIN MENU", Fore.CYAN + Style.BRIGHT))
    print("1. Fix Duplicate Phones (shared numbers)")
    print("2. Fix Non-Phone Entries")
    print("3. Fix Contacts with Extra Fields")
    print("4. Fix Duplicate Name Groups (merge/delete)")
    print("5. Show current status")
    print("6. Proceed to final output")
    return input(color_text("Your choice (1-6): ", Fore.GREEN)).strip()


# ---- Phone duplicates ----

def build_phone_list(contacts, handled_phones):
    phone_map = defaultdict(list)
    for idx, c in enumerate(contacts):
        for p in c['phones_norm']:
            phone_map[p].append(idx)
    phone_list = []
    for p, idxs in phone_map.items():
        if p in handled_phones or len(idxs) < 2:
            continue
        names = {contacts[i]['name'] for i in idxs}
        if len(names) < 2:
            continue
        phone_list.append((p, idxs))
    return phone_list


def fix_duplicate_phones(contacts, log_lines):
    handled_phones = set()
    phone_list = build_phone_list(contacts, handled_phones)
    if not phone_list:
        print(color_text("No cross-name duplicate phones found.", Fore.GREEN))
        return

    print(color_text("Duplicate phone numbers (cross-name):", Fore.YELLOW))
    for p, idxs in phone_list:
        names = {contacts[i]['name'] for i in idxs}
        print(f"  {p} -> {', '.join(names)}")
    print()

    while True:
        if not phone_list:
            break
        phone, indices = phone_list[0]
        unique_indices = list(dict.fromkeys(indices))

        print(color_text(f"\nPhone: {phone}", Fore.YELLOW + Style.BRIGHT))
        for idx in unique_indices:
            c = contacts[idx]
            raw_phones = ', '.join(c['phones_raw'])
            norm_phones = ', '.join(c['phones_norm'])
            print(f"  [{idx}] {c['name']}")
            print(f"       raw : {raw_phones}")
            print(f"       norm: {norm_phones}")
        choice = input(color_text("(f)ix, (q)uick merge, (s)kip, (b)ack to main menu: ", Fore.GREEN)).strip().lower()
        if choice == 'b':
            return
        elif choice == 's':
            handled_phones.add(phone)
            log_lines.append(f"Skipped duplicate phone {phone}")
            phone_list = build_phone_list(contacts, handled_phones)
            continue
        elif choice == 'q':
            base_idx = unique_indices[0]
            base = contacts[base_idx]
            for idx in unique_indices:
                if idx == base_idx:
                    continue
                c = contacts[idx]
                base['phones_raw'].extend(c['phones_raw'])
                base['phones_norm'].extend(c['phones_norm'])
                if not base['extra'] and c['extra']:
                    base['extra'] = c['extra']
                c['_deleted'] = True
            base['phones_raw'], base['phones_norm'] = deduplicate_phones(base['phones_raw'], base['phones_norm'])
            contacts[:] = [c for c in contacts if not c.get('_deleted', False)]
            print(color_text(f"Quick merge into {base['name']} completed.", Fore.GREEN))
            log_lines.append(f"Quick merge on {phone} into {base['name']}")
            # rebuild list, don't add to handled because phone is gone after merge
            phone_list = build_phone_list(contacts, handled_phones)
            continue
        elif choice == 'f':
            while True:
                sub = input(color_text("Options: (r)emove from one, (m)erge all, (i)gnore, (b)ack to phone list: ", Fore.GREEN)).strip().lower()
                if sub == 'b':
                    # return to phone list (outer loop will restart)
                    phone_list = build_phone_list(contacts, handled_phones)
                    break
                elif sub == 'i':
                    handled_phones.add(phone)
                    log_lines.append(f"Ignored duplicate phone {phone}")
                    phone_list = build_phone_list(contacts, handled_phones)
                    break
                elif sub == 'r':
                    try:
                        idx_str = input("Remove from which contact (index): ").strip()
                        idx = int(idx_str)
                        if idx in unique_indices:
                            c = contacts[idx]
                            if phone in c['phones_norm']:
                                norm_idx = c['phones_norm'].index(phone)
                                del c['phones_raw'][norm_idx]
                                del c['phones_norm'][norm_idx]
                                print(color_text(f"Phone {phone} removed from {c['name']}.", Fore.GREEN))
                                log_lines.append(f"Removed {phone} from {c['name']}")
                        else:
                            print("Index not among those sharing this phone.")
                    except ValueError:
                        print("Invalid index.")
                    # after removal rebuild and restart from top
                    phone_list = build_phone_list(contacts, handled_phones)
                    break
                elif sub == 'm':
                    print("Choose the base contact (keep its name & extra fields):")
                    for idx in unique_indices:
                        print(f"  [{idx}] {contacts[idx]['name']}")
                    base_str = input("Base index: ").strip()
                    try:
                        base_idx = int(base_str)
                        if base_idx not in unique_indices:
                            print("Invalid index.")
                            continue
                    except ValueError:
                        print("Invalid number.")
                        continue
                    base = contacts[base_idx]
                    for idx in unique_indices:
                        if idx == base_idx:
                            continue
                        c = contacts[idx]
                        base['phones_raw'].extend(c['phones_raw'])
                        base['phones_norm'].extend(c['phones_norm'])
                        if not base['extra'] and c['extra']:
                            base['extra'] = c['extra']
                        c['_deleted'] = True
                    base['phones_raw'], base['phones_norm'] = deduplicate_phones(base['phones_raw'], base['phones_norm'])
                    contacts[:] = [c for c in contacts if not c.get('_deleted', False)]
                    print(color_text(f"Merged contacts into {base['name']}.", Fore.GREEN))
                    log_lines.append(f"Merged contacts sharing {phone} into {base['name']}")
                    phone_list = build_phone_list(contacts, handled_phones)
                    break
                else:
                    print(color_text("Invalid option.", Fore.YELLOW))
            # after inner break, outer loop will restart with updated phone_list
        else:
            print(color_text("Invalid choice.", Fore.YELLOW))
            phone_list = build_phone_list(contacts, handled_phones)


# ---- Non-phone entries ----

def fix_non_phone_entries(contacts, analysis, log_lines):
    non_idx = analysis['non_phone_indices']
    if not non_idx:
        print(color_text("No non-phone entries.", Fore.GREEN))
        return
    non_idx = [i for i in non_idx if not contacts[i].get('_deleted', False)]
    if not non_idx:
        return

    print(color_text(f"Found {len(non_idx)} entries without valid phone numbers.", Fore.YELLOW))
    while True:
        choice = input(color_text("Options: (d)elete all, (r)eview one by one, (b)ack: ", Fore.GREEN)).strip().lower()
        if choice == 'b':
            return
        elif choice == 'd':
            confirm = input(color_text("Are you sure? (y/n): ", Fore.RED)).strip().lower()
            if confirm == 'y':
                for i in non_idx:
                    contacts[i]['_deleted'] = True
                print(color_text(f"Deleted {len(non_idx)} entries.", Fore.GREEN))
                log_lines.append(f"Deleted all {len(non_idx)} non-phone entries")
                contacts[:] = [c for c in contacts if not c.get('_deleted', False)]
                return
            else:
                continue
        elif choice == 'r':
            i = 0
            while i < len(non_idx):
                idx = non_idx[i]
                c = contacts[idx]
                raw_phones = ', '.join(c['phones_raw']) if c['phones_raw'] else '(empty)'
                print(f"\n[{idx}] {c['name']}  phones: {raw_phones}")
                sub = input(color_text("(d)elete, (k)eep, (b)ack to sub-menu: ", Fore.GREEN)).strip().lower()
                if sub == 'd':
                    c['_deleted'] = True
                    print(color_text(f"Deleted {c['name']}.", Fore.GREEN))
                    log_lines.append(f"Deleted non-phone entry: {c['name']}")
                    contacts[:] = [c for c in contacts if not c.get('_deleted', False)]
                    non_idx = [j for j in non_idx if not contacts[j].get('_deleted', False)]
                    i -= 1
                elif sub == 'k':
                    log_lines.append(f"Kept non-phone entry: {c['name']}")
                elif sub == 'b':
                    break
                else:
                    print(color_text("Invalid option.", Fore.YELLOW))
                    continue
                i += 1
            break
        else:
            print(color_text("Invalid option.", Fore.YELLOW))


# ---- Extra fields ----

def fix_extra_fields(contacts, analysis, log_lines):
    field_counts = defaultdict(int)
    for idx, name, extra in analysis['extra_contacts']:
        if contacts[idx].get('_deleted', False):
            continue
        for field in extra:
            field_counts[field] += 1

    if not field_counts:
        print(color_text("No extra fields.", Fore.GREEN))
        return

    while True:
        print(color_text("Extra field types:", Fore.CYAN))
        fields_list = list(field_counts.items())
        for i, (field, cnt) in enumerate(fields_list, 1):
            print(f"  {i}. {field} ({cnt} contacts)")
        print(f"  B. Back to main menu")
        choice = input(color_text("Select field to manage (1-{}) or B: ".format(len(fields_list)), Fore.GREEN)).strip()
        if choice.lower() == 'b':
            return
        try:
            idx_choice = int(choice) - 1
            if idx_choice < 0 or idx_choice >= len(fields_list):
                print(color_text("Invalid choice.", Fore.YELLOW))
                continue
        except ValueError:
            print(color_text("Invalid input.", Fore.YELLOW))
            continue
        field = fields_list[idx_choice][0]
        while True:
            print(color_text(f"\nManaging field '{field}'.", Fore.YELLOW))
            sub = input(color_text("Options: (d)elete from all, (r)eview one by one, (b)ack to field list: ", Fore.GREEN)).strip().lower()
            if sub == 'b':
                break
            elif sub == 'd':
                confirm = input(color_text(f"Remove '{field}' from ALL contacts? (y/n): ", Fore.RED)).strip().lower()
                if confirm == 'y':
                    count_removed = 0
                    for idx, _, _ in analysis['extra_contacts']:
                        if contacts[idx].get('_deleted', False):
                            continue
                        if field in contacts[idx]['extra']:
                            del contacts[idx]['extra'][field]
                            count_removed += 1
                    print(color_text(f"Removed '{field}' from {count_removed} contacts.", Fore.GREEN))
                    log_lines.append(f"Removed field '{field}' from {count_removed} contacts.")
                    break
            elif sub == 'r':
                contacts_with_field = [(idx, name, extra) for idx, name, extra in analysis['extra_contacts']
                                       if field in extra and not contacts[idx].get('_deleted', False)]
                if not contacts_with_field:
                    print(color_text("No contacts with this field remaining.", Fore.GREEN))
                    break
                i = 0
                while i < len(contacts_with_field):
                    idx, name, extra = contacts_with_field[i]
                    print(f"\n[{idx}] {name}")
                    print(f"  {field}: {contacts[idx]['extra'][field]}")
                    sub2 = input(color_text("(e)rase, (k)eep, (b)ack to sub-menu: ", Fore.GREEN)).strip().lower()
                    if sub2 == 'e':
                        del contacts[idx]['extra'][field]
                        print(color_text(f"Erased '{field}' from {name}.", Fore.GREEN))
                        log_lines.append(f"Erased '{field}' from {name}")
                        contacts_with_field.pop(i)
                        continue
                    elif sub2 == 'k':
                        pass
                    elif sub2 == 'b':
                        break
                    else:
                        print(color_text("Invalid option.", Fore.YELLOW))
                        continue
                    i += 1
                break
            else:
                print(color_text("Invalid option.", Fore.YELLOW))
        field_counts = defaultdict(int)
        for idx, name, extra in analysis['extra_contacts']:
            if contacts[idx].get('_deleted', False):
                continue
            for f in extra:
                field_counts[f] += 1
        if not field_counts:
            print(color_text("All extra fields removed.", Fore.GREEN))
            return


# ---- Duplicate name groups ----

def build_name_groups(contacts, handled_names):
    groups = defaultdict(list)
    for idx, c in enumerate(contacts):
        groups[c['name']].append(idx)
    # Return only those with >1 and not skipped
    return {n: idxs for n, idxs in groups.items() if len(idxs) > 1 and n not in handled_names}


def fix_duplicate_names(contacts, log_lines):
    handled_names = set()
    name_groups = build_name_groups(contacts, handled_names)
    if not name_groups:
        print(color_text("No duplicate name groups.", Fore.GREEN))
        return

    print(color_text("Duplicate name groups:", Fore.YELLOW))
    for name, idxs in name_groups.items():
        print(f"  {name} ({len(idxs)} entries)")
    print()

    name_list = list(name_groups.keys())

    while True:
        if not name_list:
            break
        name = name_list[0]
        indices = name_groups[name]
        # safety: ensure still valid
        valid_indices = [i for i in indices if not contacts[i].get('_deleted', False)]
        if len(valid_indices) < 2:
            # group no longer duplicate
            name_list.pop(0)
            continue

        print(color_text(f"\n--- {name} ---", Fore.YELLOW + Style.BRIGHT))
        for idx in valid_indices:
            c = contacts[idx]
            print(f"  [{idx}] phones: {', '.join(c['phones_raw'])}")
        choice = input(color_text("(f)ix, (q)uick merge, (s)kip, (b)ack to main menu: ", Fore.GREEN)).strip().lower()
        if choice == 'b':
            return
        elif choice == 's':
            handled_names.add(name)
            log_lines.append(f"Skipped name group: {name}")
            name_groups = build_name_groups(contacts, handled_names)
            name_list = list(name_groups.keys())
            continue
        elif choice == 'q':
            base_idx = valid_indices[0]
            base = contacts[base_idx]
            for idx in valid_indices:
                if idx == base_idx:
                    continue
                c = contacts[idx]
                base['phones_raw'].extend(c['phones_raw'])
                base['phones_norm'].extend(c['phones_norm'])
                if not base['extra'] and c['extra']:
                    base['extra'] = c['extra']
                c['_deleted'] = True
            base['phones_raw'], base['phones_norm'] = deduplicate_phones(base['phones_raw'], base['phones_norm'])
            contacts[:] = [c for c in contacts if not c.get('_deleted', False)]
            print(color_text(f"Quick merge into {base['name']} completed.", Fore.GREEN))
            log_lines.append(f"Quick merge group '{name}' into {base['name']}")
            name_groups = build_name_groups(contacts, handled_names)
            name_list = list(name_groups.keys())
            continue
        elif choice == 'f':
            while True:
                sub = input(color_text("Options: (m)erge, (d)elete all, (i)gnore, (b)ack to group list: ", Fore.GREEN)).strip().lower()
                if sub == 'b':
                    name_groups = build_name_groups(contacts, handled_names)
                    name_list = list(name_groups.keys())
                    break
                elif sub == 'i':
                    handled_names.add(name)
                    log_lines.append(f"Ignored group: {name}")
                    name_groups = build_name_groups(contacts, handled_names)
                    name_list = list(name_groups.keys())
                    break
                elif sub == 'd':
                    for idx in valid_indices:
                        contacts[idx]['_deleted'] = True
                    print(color_text(f"Deleted all entries for '{name}'.", Fore.GREEN))
                    log_lines.append(f"Deleted group: {name}")
                    contacts[:] = [c for c in contacts if not c.get('_deleted', False)]
                    name_groups = build_name_groups(contacts, handled_names)
                    name_list = list(name_groups.keys())
                    break
                elif sub == 'm':
                    print("Choose the base contact (keep its name & extra fields):")
                    for idx in valid_indices:
                        print(f"  [{idx}] {contacts[idx]['name']}")
                    base_str = input("Base index: ").strip()
                    try:
                        base_idx = int(base_str)
                        if base_idx not in valid_indices:
                            print("Invalid index.")
                            continue
                    except ValueError:
                        print("Invalid number.")
                        continue
                    base = contacts[base_idx]
                    for idx in valid_indices:
                        if idx == base_idx:
                            continue
                        c = contacts[idx]
                        base['phones_raw'].extend(c['phones_raw'])
                        base['phones_norm'].extend(c['phones_norm'])
                        if not base['extra'] and c['extra']:
                            base['extra'] = c['extra']
                        c['_deleted'] = True
                    base['phones_raw'], base['phones_norm'] = deduplicate_phones(base['phones_raw'], base['phones_norm'])
                    contacts[:] = [c for c in contacts if not c.get('_deleted', False)]
                    print(color_text(f"Merged '{name}' group into {base['name']}.", Fore.GREEN))
                    log_lines.append(f"Merged group '{name}' into base {base['name']}")
                    name_groups = build_name_groups(contacts, handled_names)
                    name_list = list(name_groups.keys())
                    break
                else:
                    print(color_text("Invalid option.", Fore.YELLOW))
            # after inner break, outer loop continues with updated name_list
        else:
            print(color_text("Invalid choice.", Fore.YELLOW))
            name_groups = build_name_groups(contacts, handled_names)
            name_list = list(name_groups.keys())


# ---------- main ----------

def main():
    # prerequisites
    missing = []
    for lib in ['colorama', 'pandas', 'openpyxl']:
        try:
            __import__(lib)
        except ImportError:
            missing.append(lib)
    if missing:
        print(color_text("Missing libraries: " + ', '.join(missing), Fore.YELLOW))
        print("Install with: pip install " + ' '.join(missing))
        if 'pandas' in missing or 'openpyxl' in missing:
            print(color_text("(Excel output will be disabled.)", Fore.YELLOW))
    else:
        print(color_text("All prerequisites satisfied.", Fore.GREEN))
    print()

    # banner
    print(color_text("=" * 55, Fore.CYAN))
    print(color_text("  CONTACT CLEANER PRO", Fore.CYAN + Style.BRIGHT))
    print(color_text("  Tidy up, merge duplicates, fix Iranian numbers", Fore.CYAN))
    print(color_text("  ----------------------------------------------", Fore.CYAN))
    print(color_text("  Crafted with care by Mohammad Azadfar", Fore.GREEN))
    print(color_text("  +98 912 419 8445", Fore.YELLOW))
    print(color_text("  Intelligence powered by DeepSeek AI", Fore.BLUE))
    print(color_text("=" * 55, Fore.CYAN))
    print()

    # file selection
    csv_files, vcf_files = find_input_files()
    filename = choose_file(csv_files, vcf_files)
    print(color_text(f"\nLoading: {filename}\n", Fore.CYAN))

    if filename.lower().endswith('.vcf'):
        contacts = parse_vcf(filename)
    else:
        contacts = parse_csv(filename)

    if not contacts:
        print(color_text("No valid contacts found.", Fore.YELLOW))
        return

    log_lines = []
    analysis = run_analysis(contacts)
    show_summary(analysis)

    while True:
        choice = main_menu()
        if choice == '1':
            fix_duplicate_phones(contacts, log_lines)
        elif choice == '2':
            fix_non_phone_entries(contacts, analysis, log_lines)
        elif choice == '3':
            fix_extra_fields(contacts, analysis, log_lines)
        elif choice == '4':
            fix_duplicate_names(contacts, log_lines)
        elif choice == '5':
            analysis = run_analysis(contacts)
            show_summary(analysis)
        elif choice == '6':
            break
        else:
            print(color_text("Invalid choice.", Fore.YELLOW))
        analysis = run_analysis(contacts)

    # Final normalization
    print(color_text("\nFINAL NORMALIZATION", Fore.CYAN + Style.BRIGHT))
    print("1. Add Iran prefix (+98)")
    print("2. Keep numbers as they are")
    print("3. Enter a custom country code (e.g., +1)")
    norm_choice = input(color_text("Your choice: ", Fore.GREEN)).strip()

    country_code = None
    if norm_choice == '1':
        country_code = '98'
    elif norm_choice == '3':
        code = input("Enter country code (without +): ").strip()
        if code:
            country_code = code
        else:
            print("Invalid code. Keeping numbers unchanged.")

    # Prepare final output
    final_contacts = []
    for c in contacts:
        mobiles, landlines, others = [], [], []
        for p in c['phones_raw']:
            norm = to_normalized_for_compare(p)
            cat = classify_phone(norm) if norm.startswith('+98') else 'Other'
            if country_code and cat in ('Mobile', 'Landline'):
                p = normalize_with_country_code(p, country_code)
            norm2 = to_normalized_for_compare(p)
            final_cat = classify_phone(norm2)
            if final_cat == 'Mobile':
                mobiles.append(p)
            elif final_cat == 'Landline':
                landlines.append(p)
            else:
                others.append(p)
        final_contacts.append({
            'name': c['name'],
            'mobiles': mobiles,
            'landlines': landlines,
            'others': others,
            'extra': c['extra']
        })

    # Final review of "Other" numbers that look like mobile
    other_like_mobile = []
    for c in final_contacts:
        for p in c['others']:
            if looks_like_mobile(p):
                other_like_mobile.append((c['name'], p, c))

    if other_like_mobile:
        print(color_text("\nFINAL REVIEW: numbers in 'Other' that look like mobile", Fore.CYAN + Style.BRIGHT))
        print(f"Found {len(other_like_mobile)} such numbers.")
        print("(1) Move all to Mobile")
        print("(2) Skip all (keep in Other)")
        print("(3) Review one by one")
        rev_choice = input(color_text("Your choice: ", Fore.GREEN)).strip()
        if rev_choice == '1':
            for name, p, c in other_like_mobile:
                c['mobiles'].append(p)
                c['others'].remove(p)
            print(color_text("All moved to Mobile.", Fore.GREEN))
        elif rev_choice == '2':
            print(color_text("All kept in Other.", Fore.GREEN))
        elif rev_choice == '3':
            for name, p, c in other_like_mobile:
                print(f"\n{color_text(name, Fore.YELLOW)}: {p}")
                mv = input(color_text("Move to (m)obile, (s)kip: ", Fore.GREEN)).strip().lower()
                if mv == 'm':
                    c['mobiles'].append(p)
                    c['others'].remove(p)
                    print(color_text("Moved.", Fore.GREEN))
                else:
                    print("Skipped.")
        else:
            print(color_text("Invalid choice. Keeping unchanged.", Fore.YELLOW))
    else:
        print(color_text("\nNo 'Other' numbers that look like mobile found.", Fore.GREEN))

    # Recompute max column counts based on actual data to avoid empty columns
    max_mob = max((len(c['mobiles']) for c in final_contacts), default=0)
    max_land = max((len(c['landlines']) for c in final_contacts), default=0)
    max_oth = max((len(c['others']) for c in final_contacts), default=0)

    fieldnames = ['Name']
    fieldnames += [f'Mobile {i+1}' for i in range(max_mob)]
    fieldnames += [f'Landline {i+1}' for i in range(max_land)]
    fieldnames += [f'Other Phone {i+1}' for i in range(max_oth)]

    rows = []
    for c in final_contacts:
        row = {'Name': c['name']}
        for i in range(max_mob):
            row[f'Mobile {i+1}'] = c['mobiles'][i] if i < len(c['mobiles']) else ''
        for i in range(max_land):
            row[f'Landline {i+1}'] = c['landlines'][i] if i < len(c['landlines']) else ''
        for i in range(max_oth):
            row[f'Other Phone {i+1}'] = c['others'][i] if i < len(c['others']) else ''
        rows.append(row)

    # Save CSV
    with open(OUTPUT_CSV, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    print(color_text(f"\nCSV saved: {OUTPUT_CSV}", Fore.GREEN))

    # Save VCF
    with open(OUTPUT_VCF, 'w', encoding='utf-8') as f:
        for c in final_contacts:
            f.write("BEGIN:VCARD\nVERSION:3.0\n")
            f.write(f"FN:{c['name']}\n")
            f.write(f"N:;{c['name']};;;\n")
            for phone in c['mobiles']:
                f.write(f"TEL;TYPE=CELL:{phone}\n")
            for phone in c['landlines']:
                f.write(f"TEL;TYPE=HOME:{phone}\n")
            for phone in c['others']:
                f.write(f"TEL;TYPE=OTHER:{phone}\n")
            if c['extra']:
                note_parts = [f"{k}: {v}" for k, v in c['extra'].items()]
                f.write(f"NOTE:{' | '.join(note_parts)}\n")
            f.write("END:VCARD\n")
    print(color_text(f"VCF saved: {OUTPUT_VCF}", Fore.GREEN))

    # Save Excel
    try:
        import pandas as pd
        df = pd.DataFrame(rows)
        df.to_excel(OUTPUT_XLSX, index=False, engine='openpyxl')
        print(color_text(f"Excel saved: {OUTPUT_XLSX}", Fore.GREEN))
    except ImportError:
        print(color_text("Excel output skipped (pandas/openpyxl not installed).", Fore.YELLOW))

    # Write log
    with open(LOG_FILE, 'w', encoding='utf-8') as log:
        log.write("Contact Cleaner Pro - Change Log\n")
        log.write("=" * 30 + "\n")
        for line in log_lines:
            log.write(line + "\n")
    print(color_text(f"Log saved: {LOG_FILE}", Fore.CYAN))

    # Goodbye
    print(color_text("\n" + "=" * 50, Fore.CYAN))
    print(color_text("Thank you for using Contact Cleaner Pro!", Fore.GREEN + Style.BRIGHT))
    print(color_text("We hope your contacts are now clean and organized.", Fore.CYAN))
    print(color_text("=" * 50, Fore.CYAN))
    input(color_text("Press any key to exit...", Fore.GREEN))


if __name__ == "__main__":
    main()
