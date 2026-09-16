"""
region_program_contacts.py
----------------------------
Resolves the DMP interface's "FWS Region" and "FWS Program" selections into
mdEditor contact UUIDs, using an mdEditor contacts export (the
"Manage Contacts > Export" JSON file for the Region/Program organization
records).

This replaces the notebook's old hardcoded block:

    akfwsrow = pd.DataFrame(columns=['uuid','Role'])
    akfwsrow.loc[0] = ['821858df-5d0e-445a-b027-014f5ef68782','administrator']
    akfwsrow.loc[1] = ['821858df-5d0e-445a-b027-014f5ef68782','distributor']
    akfwsrow.loc[2] = ['821858df-5d0e-445a-b027-014f5ef68782','publisher']

...which always pointed at the Alaska Region org contact no matter what the
project actually was. Now that the interface asks which FWS Region and
FWS Program(s) a project belongs to, we look up the matching org contacts
instead.

Usage inside the notebook (replaces the old akfwsrow block in the
"final_contacts" cell):

    import region_program_contacts as rpc
    admrows = rpc.build_admin_rows(df, region_program_contacts_path)
    final_contacts = pd.concat([final_contacts, admrows]).reset_index(drop=True)

CHANGELOG
---------
- FWS Region and FWS Program contact IDs are now HARDCODED (REGION_CONTACT_ID
  / PROGRAM_CONTACT_ID below), sourced from the actual mdEditor contacts
  export cross-referenced against each organization's real contactId. This
  is the primary lookup path now -- no CSV file is required for it.
  REGION_CONTACT_NAME / PROGRAM_CONTACT_NAME (name-matching) remain as a
  fallback for cases where a hardcoded ID isn't present in whatever
  contacts export is loaded (e.g. an older or different mdEditor instance).
- build_admin_rows() still accepts optional fws_regions_csv_path /
  fws_programs_csv_path parameters for advanced use -- if supplied, a CSV-
  based ID lookup is tried BEFORE the hardcoded tables (so a local CSV can
  override the hardcoded defaults without editing this file). Both remain
  optional and default to None; omitting them is the normal case now.
- load_contact_records() requires record["type"] == "contacts" before
  attempting to parse a record, to support fuller mdEditor exports that
  also contain "records", "settings", "schemas", "custom-profiles", and
  "profiles" entries alongside contacts.
- Added load_contact_records_by_id(), the by-contactId counterpart to
  load_contact_records() (which is keyed by name), used by the ID-based
  lookup path.
"""

import csv
import json
import re

import pandas as pd

# ---------------------------------------------------------------------------
# PRIMARY lookup: hardcoded mdEditor contact IDs for each FWS Region /
# Program, sourced from the actual FWSRegion_Program_Contacts_mdeditor-*.json
# export (cross-referenced each organization's real contactId against its
# name). If your organization's mdEditor contact IDs for these orgs ever
# change (e.g. a record gets recreated), update the values below to match.
# ---------------------------------------------------------------------------
REGION_CONTACT_ID = {
    "1 – Pacific Region": "95aa955c-140b-4744-ae07-138ff8d5a220",
    "2 – Southwest Region": "ae95f439-0914-4528-8bd6-f948c3c01a38",
    "3 – Midwest Region": "fe83831d-b35d-4023-95fb-ca91aecb01cf",
    "4 – Southeast Region": "42f6e946-635d-4a6a-b317-d011f6cf04a2",
    "5 – Northeast Region": "e3efce92-4061-4ac5-b45b-33d98eabc69a",
    "6 – Mountain-Prairie Region": "cabb6cef-48ef-42bc-94db-815ae0d8392c",
    "7 – Alaska Region": "821858df-5d0e-445a-b027-014f5ef68782",
    "8 – Pacific Southwest Region": "7d8f22bd-a9c0-4333-9b7a-dc0b310a85c1",
    "HQ – Headquarters": "a92632e3-caf2-4a13-8840-2dd35e4d92cc",
}
PROGRAM_CONTACT_ID = {
    "Ecological Services": "b74f0135-d04b-4989-aac1-3f4f38661436",
    "Fish and Aquatic Conservation": "f899557c-d2f3-4257-8f56-4bcd7b0d12d2",
    "Migratory Birds": "5a8ffdde-358e-4276-9289-596b7f23c22d",
    "National Wildlife Refuge System": "9aa1b724-b0a6-4f65-9290-fa680127bd72",
    "Office of Subsistence Management": "a0d6a50b-b5e1-42dd-8bc5-fd4a29f84f21",
    "Science Applications": "51a0a513-4930-48ab-9f8b-ec4ab8d47a7f",
}

# ---------------------------------------------------------------------------
# FALLBACK lookup (name-matching): used only if a region/program's
# hardcoded/CSV ID above isn't found in the loaded contacts export.
#
# These don't match character-for-character (e.g. the interface says
# "Migratory Birds", the contact record says "Migratory Bird Management"),
# so this table is what bridges the two. If your organization's contact
# names change, update the right-hand side here to match.
# ---------------------------------------------------------------------------
REGION_CONTACT_NAME = {
    "1 – Pacific Region": "U.S. Fish and Wildlife Service, Pacific Region",
    "2 – Southwest Region": "U.S. Fish and Wildlife Service, Southwest Region",
    "3 – Midwest Region": "U.S. Fish and Wildlife Service, Midwest Region",
    "4 – Southeast Region": "U.S. Fish and Wildlife Service, Southeast Region",
    "5 – Northeast Region": "U.S Fish and Wildlife Service, Northeast Region",  # sic - no period after "U.S" in source data
    "6 – Mountain-Prairie Region": "U.S. Fish and Wildlife Service, Mountain Prairie Region",
    "7 – Alaska Region": "U.S. Fish and Wildlife Service, Alaska Region",
    "8 – Pacific Southwest Region": "U.S. Fish and Wildlife Service, Pacific Southwest Region",
    "HQ – Headquarters": "U.S. Fish and Wildlife Service, Headquarters",
}
PROGRAM_CONTACT_NAME = {
    "Ecological Services": "Ecological Services",
    "Fish and Aquatic Conservation": "Fisheries and Aquatic Conservation",
    "Migratory Birds": "Migratory Bird Management",
    "National Wildlife Refuge System": "National Wildlife Refuge System",
    "Office of Subsistence Management": "Office of Subsistence Management",
    "Science Applications": "Science Applications",
}


def load_contact_records(contacts_json_path):
    """
    Parse an mdEditor contacts export and return {contact name: {"uuid":..., "json":...}}.

    "json" here is the record re-serialized in the same shape final_contacts
    expects in its own 'json' column (matching what the notebook's
    existing_contacts['contact_json'] looks like) -- i.e. the *outer*
    mdEditor record {"id":..., "attributes":{"json": "<escaped inner json>"}, "type":...},
    JSON-dumped to a string. Carrying this directly means we don't depend on
    these same org contacts also happening to exist in whatever file
    `contact_folder` points at -- this file is self-sufficient.

    Only entries with type == "contacts" are considered.
    """
    with open(contacts_json_path, "r", encoding="utf-8") as fh:
        raw = json.load(fh)

    records = {}
    for record in raw.get("data", []):
        if record.get("type") != "contacts":
            continue
        try:
            inner = json.loads(record["attributes"]["json"])
        except (KeyError, ValueError, TypeError):
            continue
        name = inner.get("name")
        contact_id = inner.get("contactId")
        if name and contact_id:
            records[name] = {"uuid": contact_id, "json": json.dumps(record)}
    return records


def load_contact_records_by_id(contacts_json_path):
    """
    Same as load_contact_records(), but keyed by contactId (uuid) instead of
    by name. Used for the ID-based lookup path (REGION_CONTACT_ID /
    PROGRAM_CONTACT_ID, or an optional CSV override).
    """
    with open(contacts_json_path, "r", encoding="utf-8") as fh:
        raw = json.load(fh)

    records = {}
    for record in raw.get("data", []):
        if record.get("type") != "contacts":
            continue
        try:
            inner = json.loads(record["attributes"]["json"])
        except (KeyError, ValueError, TypeError):
            continue
        contact_id = inner.get("contactId")
        if contact_id:
            records[contact_id] = {"uuid": contact_id, "json": json.dumps(record)}
    return records


def load_contact_name_to_uuid(contacts_json_path):
    """Convenience wrapper: {contact name: uuid} only (no json payload)."""
    return {name: rec["uuid"] for name, rec in load_contact_records(contacts_json_path).items()}


def _read_csv_rows(csv_path):
    """Minimal CSV -> list-of-dicts reader using the stdlib csv module.
    Uses utf-8-sig to tolerate a BOM, which Excel-exported CSVs commonly
    include. Only used for the optional CSV-override path."""
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as fh:
        return [dict(row) for row in csv.DictReader(fh)]


def load_region_id_map(fws_regions_csv_path):
    """
    Optional override: {region_number (int): mdEditor_Contact_ID (str)} from
    a CSV with columns Region_ID and mdEditor_Contact_ID. Not needed for
    normal use -- see REGION_CONTACT_ID above, which is the default.
    """
    mapping = {}
    for row in _read_csv_rows(fws_regions_csv_path):
        rid = (row.get("Region_ID") or "").strip()
        cid = (row.get("mdEditor_Contact_ID") or "").strip()
        if not rid or not cid:
            continue
        try:
            mapping[int(rid)] = cid
        except ValueError:
            continue
    return mapping


def load_program_id_map(fws_programs_csv_path):
    """
    Optional override: {fws_program (str, exact dropdown text): mdEditorID
    (str)} from a CSV with columns fws_program and mdEditorID. Not needed
    for normal use -- see PROGRAM_CONTACT_ID above, which is the default.
    """
    mapping = {}
    for row in _read_csv_rows(fws_programs_csv_path):
        name = (row.get("fws_program") or "").strip()
        cid = (row.get("mdEditorID") or "").strip()
        if not name or not cid:
            continue
        mapping[name] = cid
    return mapping


def _region_number_from_label(region_label):
    """'7 – Alaska Region' -> 7. Matches the leading integer regardless of
    which dash character separates it from the region name. Headquarters
    ("HQ – Headquarters") does not start with a digit but is Region 9 in
    CMT.csv / fws_regions.csv, so it needs an explicit mapping rather than
    relying on the regex fallback."""
    if not region_label:
        return None
    explicit = {
        "1 – Pacific Region": 1, "2 – Southwest Region": 2, "3 – Midwest Region": 3,
        "4 – Southeast Region": 4, "5 – Northeast Region": 5, "6 – Mountain-Prairie Region": 6,
        "7 – Alaska Region": 7, "8 – Pacific Southwest Region": 8, "HQ – Headquarters": 9,
    }
    if region_label in explicit:
        return explicit[region_label]
    m = re.match(r"^\s*(\d+)", str(region_label))
    return int(m.group(1)) if m else None


def build_admin_rows(df, contacts_json_path, include_region_distributor_publisher=True,
                      fws_regions_csv_path=None, fws_programs_csv_path=None):
    """
    Build the dataframe of rows to append to `final_contacts`, assigning:
      - the project's FWS Region as 'administrator'
        (and, by default, also 'distributor' and 'publisher' -- the same
        three roles the old hardcoded Alaska Region contact used to cover)
      - each of the project's FWS Program(s) as 'administrator'

    `df` is the dataframe produced by dmp_json_to_dataframe.build_dataframe().
    Returns a DataFrame with columns ['uuid','Role','json'], ready to concat
    onto final_contacts. The 'json' column is pre-filled (from
    contacts_json_path itself) so the notebook doesn't try to generate a new
    stub contact record for these -- they already exist in mdEditor under a
    permanent uuid, so we want to reference them, not recreate them.

    Lookup order for each Region/Program value:
      1. fws_regions_csv_path / fws_programs_csv_path, if given (optional
         override -- lets a local CSV take precedence without editing this
         file).
      2. REGION_CONTACT_ID / PROGRAM_CONTACT_ID (hardcoded, the default).
      3. REGION_CONTACT_NAME / PROGRAM_CONTACT_NAME (name-matching fallback).

    Prints a warning (and skips that entry) if a selected region/program
    still can't be resolved by any method, rather than failing silently.
    """
    records_by_name = load_contact_records(contacts_json_path)
    records_by_id = load_contact_records_by_id(contacts_json_path)

    region_id_map = load_region_id_map(fws_regions_csv_path) if fws_regions_csv_path else {}
    program_id_map = load_program_id_map(fws_programs_csv_path) if fws_programs_csv_path else {}

    rows = []

    region_value = df.loc[df.field == "fwsRegion", "value"]
    region_value = region_value.values[0] if len(region_value) else ""
    if region_value:
        record = None
        region_num = _region_number_from_label(region_value)
        if region_num is not None and region_num in region_id_map:
            record = records_by_id.get(region_id_map[region_num])
        if record is None and region_value in REGION_CONTACT_ID:
            record = records_by_id.get(REGION_CONTACT_ID[region_value])
        if record is None:
            contact_name = REGION_CONTACT_NAME.get(region_value)
            record = records_by_name.get(contact_name) if contact_name else None
        if record:
            rows.append({"uuid": record["uuid"], "Role": "administrator", "json": record["json"]})
            if include_region_distributor_publisher:
                rows.append({"uuid": record["uuid"], "Role": "distributor", "json": record["json"]})
                rows.append({"uuid": record["uuid"], "Role": "publisher", "json": record["json"]})
        else:
            print(f"[region_program_contacts] WARNING: no contact match found for FWS Region {region_value!r} "
                  f"-- it will not be added as a contact on any metadata record.")

    program_values = df.loc[df.field.str.startswith("FWSProgram"), "value"].dropna().tolist()
    for program_value in program_values:
        if not program_value:
            continue
        record = None
        if program_value in program_id_map:
            record = records_by_id.get(program_id_map[program_value])
        if record is None and program_value in PROGRAM_CONTACT_ID:
            record = records_by_id.get(PROGRAM_CONTACT_ID[program_value])
        if record is None:
            contact_name = PROGRAM_CONTACT_NAME.get(program_value)
            record = records_by_name.get(contact_name) if contact_name else None
        if record:
            rows.append({"uuid": record["uuid"], "Role": "administrator", "json": record["json"]})
        else:
            print(f"[region_program_contacts] WARNING: no contact match found for FWS Program {program_value!r} "
                  f"-- it will not be added as a contact on any metadata record.")

    return pd.DataFrame(rows, columns=["uuid", "Role", "json"])


if __name__ == "__main__":
    import sys
    import dmp_json_to_dataframe as bridge

    dmp_json = sys.argv[1] if len(sys.argv) > 1 else "test_data.json"
    contacts_json = sys.argv[2] if len(sys.argv) > 2 else "FWSRegion_Program_Contacts_mdeditor-20260811-235138.json"

    out_df, _, _ = bridge.build_dataframe(dmp_json)
    admin_rows = build_admin_rows(out_df, contacts_json)
    print(admin_rows)
