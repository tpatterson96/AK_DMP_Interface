r"""
dmp_desktop_app.py
-------------------
Desktop wrapper for the AK DMP Interface, using pywebview.

Loads DMP_Interface.html in a native window (no browser required) and
exposes a small Python API to the page's JavaScript via
`window.pywebview.api.*`. This lets the interface do things a plain browser
tab structurally cannot:

  - Native Open/Save file dialogs (real filesystem paths, not blob downloads)
  - Actually import and run dmp_json_to_dataframe.py and
    region_program_contacts.py -- the real bridge scripts, unmodified --
    instead of the JavaScript port used when this interface runs as a plain
    web page.
  - Persist shared file-location settings to a small local JSON file next
    to this script, instead of browser localStorage.

Install
-------
    pip install pywebview pandas

    (pandas is already a requirement of dmp_json_to_dataframe.py itself.)

Run
---
    python dmp_desktop_app.py

Packaging as a single .exe (optional -- lets end users run this without
Python installed themselves)
---------------------------------------------------------------------
    pip install pyinstaller
    pyinstaller --onefile --add-data "DMP_Interface.html;." dmp_desktop_app.py

Expected folder layout (same folder as this script, or point bridgeDir at
wherever these actually live via the Setup tab):
    DMP_Interface.html
    dmp_json_to_dataframe.py
    region_program_contacts.py
    dmp_desktop_app.py            <- this file
    dmp_app_config.json           <- created automatically on first run
"""

import importlib
import json
import os
import re
import sys
import tempfile
import traceback
import urllib.request
import uuid
from datetime import datetime, timezone

import webview

APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(APP_DIR, "dmp_app_config.json")
HTML_PATH = os.path.join(APP_DIR, "DMP_Interface.html")

DEFAULT_CONFIG = {
    "bridgeDir": APP_DIR,
    "contactFolder": "",
    "tmpDir": "",
    "dmpSpreadsheet": "",
    "regionProgramContactsPath": "",
    "contactsExportPath": "",
}


# =============================================================================
# CONFIG PERSISTENCE
# =============================================================================
def load_config():
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as fh:
                saved = json.load(fh)
            cfg = dict(DEFAULT_CONFIG)
            cfg.update(saved)
            return cfg
        except Exception:
            pass
    return dict(DEFAULT_CONFIG)


def save_config(cfg):
    with open(CONFIG_PATH, "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, indent=2)


def _import_bridge_modules(bridge_dir):
    """
    Imports the REAL bridge scripts. Re-imported (and reloaded) on every
    call rather than once at startup, so that editing
    dmp_json_to_dataframe.py / region_program_contacts.py, or changing
    bridgeDir on the Setup tab, takes effect without restarting the app.
    """
    if bridge_dir and bridge_dir not in sys.path:
        sys.path.insert(0, bridge_dir)
    import dmp_json_to_dataframe as bridge
    import region_program_contacts as rpc
    importlib.reload(bridge)
    importlib.reload(rpc)
    return bridge, rpc


# =============================================================================
# CONTACT RESOLUTION (Python port of the notebook's "Pulling contacts" cells,
# operating on the REAL dataframe produced by the real bridge script)
# =============================================================================
def _field_val(df, field, default=""):
    row = df.loc[df.field == field, "value"]
    return row.values[0] if len(row) else default


def _now_iso():
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3] + "Z"


def _match_contact_by_email(email, contacts_records):
    if not email:
        return None
    email_l = email.strip().lower()
    if not email_l:
        return None
    for rec in contacts_records:
        try:
            inner = json.loads(rec["attributes"]["json"])
        except Exception:
            continue
        emails = [str(e).lower() for e in inner.get("electronicMailAddress", [])]
        if email_l in emails:
            return {"uuid": inner.get("contactId"), "name": inner.get("name"), "record": rec}
    return None


def _stub_contact(first, last, email, phone="", position="", org_uuid=None):
    cid = str(uuid.uuid4())
    name = (str(first or "") + " " + str(last or "")).strip() or "Unnamed contact"
    inner = {
        "contactId": cid,
        "isOrganization": False,
        "name": name,
        "memberOfOrganization": [org_uuid] if org_uuid else [],
        "logoGraphic": [],
        "phone": [{"phoneNumber": re.sub(r"[^0-9]", "", phone)}] if phone else [],
        "address": [],
        "electronicMailAddress": [email] if email else [],
        "onlineResource": [],
        "hoursOfService": [],
        "contactType": "federal",
        "externalIdentifier": [],
    }
    if position:
        inner["positionName"] = position
    record = {
        "id": cid[:8],
        "attributes": {"json": json.dumps(inner), "date-updated": _now_iso(), "rev": None},
        "type": "contacts",
        "meta": {"title": name, "icon": "users", "export": True},
    }
    return {"uuid": cid, "name": name, "record": record}


def _resolve_person_fields(df, prefix, contacts_records, new_contacts):
    """prefix e.g. 'contactSteward', 'dmpCreator' -- matches the
    <prefix>FirstName/<prefix>LastName/<prefix>Email/<prefix>Phone/<prefix>Position
    field naming used throughout dmp_json_to_dataframe.py's output."""
    first = _field_val(df, prefix + "FirstName")
    last = _field_val(df, prefix + "LastName")
    email = _field_val(df, prefix + "Email")
    phone = _field_val(df, prefix + "Phone")
    position = _field_val(df, prefix + "Position")
    if not (first or last or email):
        return None
    match = _match_contact_by_email(email, contacts_records)
    if match:
        return match
    stub = _stub_contact(first, last, email, phone, position)
    new_contacts.append(stub["record"])
    return stub


def _role_entry(contact, role):
    return {"party": [{"contactId": contact["uuid"]}], "role": role}


def _load_contacts_records(path, warnings, label):
    if not path:
        warnings.append("No " + label + " file configured (see the Setup tab).")
        return []
    if not os.path.exists(path):
        warnings.append(label + " file not found at: " + path)
        return []
    try:
        with open(path, "r", encoding="utf-8") as fh:
            raw = json.load(fh)
        return raw.get("data", [])
    except Exception as e:
        warnings.append("Could not read " + label + ": " + str(e))
        return []


def _resolve_region_program(admin_rows_df):
    """admin_rows_df is the REAL return value of
    region_program_contacts.build_admin_rows() -- columns ['uuid','Role','json'].
    Region gets distributor/publisher (in addition to administrator);
    Programs only ever get administrator -- that distinction is how we tell
    them apart here without re-deriving the name-matching logic ourselves."""
    if admin_rows_df is None or len(admin_rows_df) == 0:
        return None, []
    uuid_roles = {}
    for _, row in admin_rows_df.iterrows():
        uuid_roles.setdefault(row["uuid"], set()).add(row["Role"])
    region_contact = None
    program_contacts = []
    for u, roles in uuid_roles.items():
        c = {"uuid": u}
        if "distributor" in roles or "publisher" in roles:
            region_contact = c
        else:
            program_contacts.append(c)
    return region_contact, program_contacts


def _fetch_profiles_schemas(urls, warnings):
    out = []
    for label, url in (urls or {}).items():
        if not url:
            continue
        try:
            with urllib.request.urlopen(url, timeout=10) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            out.append({"label": label, "url": url, "json": data})
        except Exception as e:
            warnings.append("Could not fetch " + label + " from " + url + ": " + str(e))
    return out


# =============================================================================
# DRAFT ASSEMBLY -- built from the REAL dataframe returned by
# dmp_json_to_dataframe.build_dataframe(), not a re-derivation of its logic.
# =============================================================================
def _assemble_draft(df, admin_rows_df, contacts_records, warnings):
    new_contacts = []

    steward = _resolve_person_fields(df, "contactSteward", contacts_records, new_contacts)
    custodian = _resolve_person_fields(df, "contactCustodian", contacts_records, new_contacts)
    trustee = _resolve_person_fields(df, "contactTrustee", contacts_records, new_contacts)
    creator = _resolve_person_fields(df, "dmpCreator", contacts_records, new_contacts)
    updater = None
    if _field_val(df, "dmpUpdaterEmail") or _field_val(df, "dmpUpdaterFirstName"):
        updater = _resolve_person_fields(df, "dmpUpdater", contacts_records, new_contacts)

    region_contact, program_contacts = _resolve_region_program(admin_rows_df)

    project_uuid = str(uuid.uuid4())

    keywords_raw = _field_val(df, "projectKeywords")
    keyword_list = [k.strip() for k in keywords_raw.split(",") if k.strip()] if keywords_raw else []

    point_of_contact = []
    if steward:
        point_of_contact.append(_role_entry(steward, "pointOfContact"))
    if region_contact:
        point_of_contact.append(_role_entry(region_contact, "administrator"))
    if custodian:
        point_of_contact.append(_role_entry(custodian, "custodian"))
    for pc in program_contacts:
        point_of_contact.append(_role_entry(pc, "administrator"))
    if trustee:
        point_of_contact.append(_role_entry(trustee, "owner"))

    metadata_contact = []
    if creator:
        metadata_contact.append(_role_entry(creator, "author"))
        metadata_contact.append(_role_entry(creator, "pointOfContact"))
    if updater:
        metadata_contact.append(_role_entry(updater, "author"))
    if region_contact:
        metadata_contact.append(_role_entry(region_contact, "publisher"))

    citation_responsible_party = []
    if steward:
        citation_responsible_party.append(_role_entry(steward, "pointOfContact"))
    if region_contact:
        citation_responsible_party.append(_role_entry(region_contact, "publisher"))

    start_date = _field_val(df, "projectStartDate")
    ongoing = _field_val(df, "projectOngoing") == "True"
    end_date = _field_val(df, "projectEndDate")

    date_arr = []
    if start_date:
        date_arr.append({"date": start_date + "T00:00:00.000Z", "dateType": "start"})
    if not ongoing and end_date:
        date_arr.append({"date": end_date + "T00:00:00.000Z", "dateType": "end"})

    time_period = None
    if start_date or (not ongoing and end_date):
        time_period = {}
        if start_date:
            time_period["startDateTime"] = start_date + "T00:00:00.000Z"
        if not ongoing and end_date:
            time_period["endDateTime"] = end_date + "T00:00:00.000Z"

    status = "proposed"
    try:
        today = datetime.now().date()
        if not ongoing and end_date and today >= datetime.strptime(end_date, "%Y-%m-%d").date():
            status = "completed"
        elif start_date and today >= datetime.strptime(start_date, "%Y-%m-%d").date():
            status = "onGoing"
    except ValueError:
        pass

    project_record = {
        "schema": {"name": "mdJson", "version": "2.7.0"},
        "metadata": {
            "metadataInfo": {
                "defaultMetadataLocale": {"characterSet": "UTF-8", "country": "USA", "language": "eng"},
                "metadataContact": metadata_contact or None,
                "metadataDate": [{"date": _now_iso(), "dateType": "creation"},
                                  {"date": _now_iso(), "dateType": "lastUpdate"}],
                "metadataIdentifier": {"identifier": project_uuid, "namespace": "urn:uuid"},
                "metadataStatus": "initiated",
            },
            "resourceInfo": {
                "abstract": _field_val(df, "projectAbstract") or None,
                "purpose": _field_val(df, "projectPurpose") or None,  # NEW -- mdJSON resourceInfo.purpose
                "citation": {
                    "date": date_arr or None,
                    "title": _field_val(df, "projectTitle") or "Untitled project",
                    "responsibleParty": citation_responsible_party or None,
                },
                "defaultResourceLocale": {"characterSet": "UTF-8", "country": "USA", "language": "eng"},
                "keyword": (
                    [{
                        "keyword": [{"keyword": k} for k in keyword_list],
                        "keywordType": "theme",
                        "thesaurus": {"identifier": [{"identifier": "custom"}],
                                      "title": "Keywords from Data Management Plan"},
                        "fullPath": True,
                    }] if keyword_list else None
                ),
                "pointOfContact": point_of_contact or None,
                "resourceType": [{"name": _field_val(df, "projectShortTitle") or _field_val(df, "projectTitle") or "Untitled project",
                                   "type": "project"}],
                "status": [status],
                "timePeriod": time_period,
            },
        },
    }

    # ---- Products ----
    prod_nums = sorted(
        {v for v in df.loc[df.prod_num.notna(), "prod_num"].tolist() if v},
        key=lambda x: (len(x), x),
    )
    products = []
    for pn in prod_nums:
        sub = df[df.prod_num == pn]

        def pv(field_base):
            row = sub.loc[sub.field == field_base + pn, "value"]
            return row.values[0] if len(row) else ""

        if pv("productSkipMeta") == "True":
            continue

        prod_uuid = str(uuid.uuid4())

        originators = []
        j = 1
        while (sub.field == f"productOriginatorFirstName{pn}-{j}").any() or \
              (sub.field == f"productOriginatorEmail{pn}-{j}").any():
            first_row = sub.loc[sub.field == f"productOriginatorFirstName{pn}-{j}", "value"]
            last_row = sub.loc[sub.field == f"productOriginatorLastName{pn}-{j}", "value"]
            email_row = sub.loc[sub.field == f"productOriginatorEmail{pn}-{j}", "value"]
            first = first_row.values[0] if len(first_row) else ""
            last = last_row.values[0] if len(last_row) else ""
            email = email_row.values[0] if len(email_row) else ""
            if first or last or email:
                match = _match_contact_by_email(email, contacts_records)
                if not match:
                    match = _stub_contact(first, last, email)
                    new_contacts.append(match["record"])
                originators.append(match)
            j += 1

        meta_authors = []
        j = 1
        while (sub.field == f"productMetaAuthorFirstName{pn}-{j}").any():
            first_row = sub.loc[sub.field == f"productMetaAuthorFirstName{pn}-{j}", "value"]
            last_row = sub.loc[sub.field == f"productMetaAuthorLastName{pn}-{j}", "value"]
            email_row = sub.loc[sub.field == f"productMetaAuthorEmail{pn}-{j}", "value"]
            first = first_row.values[0] if len(first_row) else ""
            last = last_row.values[0] if len(last_row) else ""
            email = email_row.values[0] if len(email_row) else ""
            if first or last or email:
                match = _match_contact_by_email(email, contacts_records)
                if not match:
                    match = _stub_contact(first, last, email)
                    new_contacts.append(match["record"])
                meta_authors.append(match)
            j += 1

        point_of_contact_prod = []
        if steward:
            point_of_contact_prod.append(_role_entry(steward, "pointOfContact"))
        if region_contact:
            point_of_contact_prod.append(_role_entry(region_contact, "administrator"))
        for o in originators:
            point_of_contact_prod.append(_role_entry(o, "originator"))
        if custodian:
            point_of_contact_prod.append(_role_entry(custodian, "custodian"))

        metadata_contact_prod = []
        for a in meta_authors:
            metadata_contact_prod.append(_role_entry(a, "author"))
        if steward:
            metadata_contact_prod.append(_role_entry(steward, "pointOfContact"))
        if region_contact:
            metadata_contact_prod.append(_role_entry(region_contact, "publisher"))

        skip_vals = ("", "na", "n/a", "tbd")
        qa, qc = pv("productQualityAssurance"), pv("productQualityControl")
        lineage_parts = []
        if qa.strip().lower() not in skip_vals:
            lineage_parts.append("Quality assurance: " + qa)
        if qc.strip().lower() not in skip_vals:
            lineage_parts.append("Quality control: " + qc)

        restriction = pv("productRestriction") or "open access"
        constraint = {
            "type": "use",
            "legal": {"accessConstraint": [restriction], "useConstraint": ["unrestricted"] if restriction == "open access" else []},
            "security": {"classification": "unclassified"},
        }
        justification = pv("productRestrictionJustification")
        if restriction != "open access" and justification:
            constraint["useLimitation"] = [justification]

        products.append({
            "schema": {"name": "mdJson", "version": "2.7.0"},
            "metadata": {
                "metadataInfo": {
                    "metadataIdentifier": {"identifier": prod_uuid, "namespace": "urn:uuid"},
                    "metadataStatus": "initiated",
                    "defaultMetadataLocale": {"language": "eng", "characterSet": "UTF-8", "country": "USA"},
                    "metadataContact": metadata_contact_prod or None,
                    "metadataDate": [{"date": _now_iso(), "dateType": "creation"},
                                      {"date": _now_iso(), "dateType": "lastUpdate"}],
                },
                "resourceInfo": {
                    "resourceType": [{"type": pv("productResourceType") or None,
                                       "name": pv("productName") or pv("productTitle") or "Untitled product"}],
                    "status": ["proposed"],
                    "citation": {"title": pv("productTitle") or "Untitled product"},
                    "abstract": pv("productAbstract") or None,
                    "defaultResourceLocale": {"language": "eng", "characterSet": "UTF-8", "country": "USA"},
                    "pointOfContact": point_of_contact_prod or None,
                    "constraint": [constraint],
                    "format": (pv("productFormatOther") if pv("productFormat") == "other" else pv("productFormat")) or None,
                },
                "resourceLineage": [{"statement": "\n\n".join(lineage_parts)}] if lineage_parts else None,
                "associatedResource": [{"associationType": "parentProject", "mdRecordId": project_uuid}],
            },
        })

    referenced = set()
    for c in [steward, custodian, trustee, creator, updater, region_contact] + program_contacts:
        if c:
            referenced.add(c["uuid"])
    contacts_out = list(new_contacts)
    for rec in contacts_records:
        try:
            inner = json.loads(rec["attributes"]["json"])
        except Exception:
            continue
        if inner.get("contactId") in referenced:
            contacts_out.append(rec)

    return {
        "generatedAt": _now_iso(),
        "bridgeFieldCount": int(len(df)),
        "newContactsCreated": len(new_contacts),
        "project": project_record,
        "products": products,
        "contacts": contacts_out,
    }


# =============================================================================
# JS-EXPOSED API (window.pywebview.api.<method> from DMP_Interface.html)
# Only public (non-underscore) methods on this class are exposed to JS.
# =============================================================================
class Api:
    def __init__(self):
        self.config = load_config()
        self.window = None  # assigned after window creation, see main()

    # ---------------- native file dialogs ----------------

    def pick_json_file(self):
        """Open dialog restricted to JSON. Returns {"path","content"} or None."""
        result = self.window.create_file_dialog(
            webview.OPEN_DIALOG, file_types=("JSON Files (*.json)", "All files (*.*)")
        )
        if not result:
            return None
        path = result[0] if isinstance(result, (list, tuple)) else result
        with open(path, "r", encoding="utf-8") as fh:
            content = fh.read()
        return {"path": path, "content": content}

    def pick_folder(self):
        result = self.window.create_file_dialog(webview.FOLDER_DIALOG)
        if not result:
            return None
        return result[0] if isinstance(result, (list, tuple)) else result

    def pick_file(self):
        result = self.window.create_file_dialog(webview.OPEN_DIALOG, file_types=("All files (*.*)",))
        if not result:
            return None
        return result[0] if isinstance(result, (list, tuple)) else result

    def save_text_file(self, default_name, content):
        """Save dialog; writes `content` to the chosen path. Returns the path, or None if cancelled."""
        result = self.window.create_file_dialog(webview.SAVE_DIALOG, save_filename=default_name)
        if not result:
            return None
        path = result[0] if isinstance(result, (list, tuple)) else result
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(content)
        return path

    # ---------------- config persistence ----------------

    def get_config(self):
        return self.config

    def set_config(self, updates):
        self.config.update(updates or {})
        save_config(self.config)
        return self.config

    # ---------------- the real bridge-script pipeline ----------------

    def run_metadata_draft(self, dmp_json_str, options=None):
        """
        Runs the ACTUAL dmp_json_to_dataframe.py and region_program_contacts.py
        against the current DMP data (passed as a JSON string from the page),
        then assembles a draft mdJSON bundle from the real dataframe output.
        """
        options = options or {}
        warnings = []

        try:
            bridge, rpc = _import_bridge_modules(self.config.get("bridgeDir") or APP_DIR)
        except Exception as e:
            return {"error": "Could not import the bridge scripts: " + str(e),
                    "trace": traceback.format_exc()}

        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False, encoding="utf-8") as tmp:
                tmp.write(dmp_json_str)
                tmp_path = tmp.name
            df, dmp_path, dmpvers = bridge.build_dataframe(tmp_path)
        except Exception as e:
            return {"error": "dmp_json_to_dataframe.build_dataframe() failed: " + str(e),
                    "trace": traceback.format_exc()}
        finally:
            if tmp_path:
                try:
                    os.remove(tmp_path)
                except OSError:
                    pass

        region_program_path = self.config.get("regionProgramContactsPath")
        admin_rows_df = None
        if region_program_path and os.path.exists(region_program_path):
            try:
                admin_rows_df = rpc.build_admin_rows(df, region_program_path)
            except Exception as e:
                warnings.append("region_program_contacts.build_admin_rows() failed: " + str(e))
        else:
            warnings.append("No Region/Program contacts file configured or found (see the Setup tab / Python Pipeline Paths). FWS Region/Program contact IDs are hardcoded in region_program_contacts.py, but the full contact records still need this file to be embedded in the output.")

        contacts_records = _load_contacts_records(
            self.config.get("contactsExportPath"), warnings, "contacts export"
        )

        try:
            draft = _assemble_draft(df, admin_rows_df, contacts_records, warnings)
        except Exception as e:
            return {"error": "Draft assembly failed: " + str(e), "trace": traceback.format_exc()}

        if options.get("attachProfilesSchemas"):
            draft["profilesSchemas"] = _fetch_profiles_schemas(options.get("profileSchemaUrls") or {}, warnings)

        draft["warnings"] = warnings
        draft["_draftNotice"] = (
            "Generated by the REAL Python bridge scripts (dmp_json_to_dataframe.py, "
            "region_program_contacts.py) running inside the desktop app -- not the "
            "JavaScript approximation used when this interface runs as a plain web page. "
            "Still does not reproduce mdEditor's own internal record/profile/schema ID "
            "envelope, PRIMR ID extraction, geoJSON geometry embedding, or the SharePoint "
            "Excel export -- run the full DMP_to_metadata_from_interface.ipynb notebook for those."
        )
        return draft


# =============================================================================
# ENTRY POINT
# =============================================================================
def main():
    if not os.path.exists(HTML_PATH):
        print("ERROR: DMP_Interface.html not found next to this script:", HTML_PATH)
        sys.exit(1)

    api = Api()
    window = webview.create_window(
        "AK Region DMP Interface",
        HTML_PATH,
        js_api=api,
        width=1320,
        height=880,
        min_size=(980, 640),
    )
    api.window = window
    webview.start()


if __name__ == "__main__":
    main()
