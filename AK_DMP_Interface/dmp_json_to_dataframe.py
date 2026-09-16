r"""
dmp_json_to_dataframe.py
-------------------------
Bridge module for the AK DMP Interface.

The original DMP_to_metadata.ipynb notebook builds a dataframe called `df`
by scraping content-control tags out of a filled-in Word DMP (.docx/.docm).
This module builds THE SAME dataframe (identical columns, identical field
naming/numbering scheme, identical value formatting) directly from the JSON
file exported by DMP_Interface.html.

Because the shape of `df` is identical to what the notebook produced before,
every cell after the old docx-scraping cells (contacts matching, mdEditor
JSON generation, SharePoint spreadsheet export, etc.) keeps working with NO
changes required.

Usage inside the notebook (replaces the old "read blank/completed docx"
cells):

    import dmp_json_to_dataframe as bridge
    df, dmp, dmpvers = bridge.build_dataframe(r"C:\path\to\dmp_data.json")

`dmp`     -> a file path string, used later in the notebook when the DMP
             itself is registered as a product (kept for compatibility).
`dmpvers` -> the DMP template version string (kept for compatibility with
             the productSkipMeta / productExists branch later in the
             notebook).

CHANGELOG
---------
- Added `projectPurpose` field (maps to DMP_Interface.html's Project
  Details -> Purpose box, and to mdJSON metadata.resourceInfo.purpose).
  New, optional field -- older dmp_data.json exports without a "purpose"
  key simply yield an empty string here, same as any other optional field.
"""

import json
import re

import pandas as pd

# Fields that should be rendered as True/False strings (matches the
# original docx checkbox translation, just skipping the intermediate
# unicode-checkbox step since we already have real booleans).
_CHECKBOX_KEYS = {
    "projectOngoing",
    "contactCustodianNonDOI",
    "metaStandardMdJSON",
    "metaStandardOther",
    "storageOneDrive",
    "storageTeams",
    "storageExternal",
    "storageNetwork",
    "storageOther",
    "rdrCheckbox",
    "productSkipMeta",
    "productOriginatorNonDOI",
    "productSpatialMatchesProject",
}


def _b(val):
    """Format a python bool (or bool-ish value) the same way the notebook expects."""
    return "True" if bool(val) else "False"


def _s(val):
    """Coerce to a stripped string, treating None as empty string."""
    if val is None:
        return ""
    return str(val).strip()


def _cost_center_code(display_value):
    """Extract the parenthetical cost-center code, e.g.
    'Kenai NWR (FF07RKNA00)' -> 'FF07RKNA00' -- matches the regex the
    original notebook uses on the scraped docx text."""
    if not display_value:
        return ""
    match = re.findall(r"(?<=[(])[A-Z0-9]+?(?=[)])", display_value)
    return match[0] if match else display_value


def _phone_digits(val):
    """Strip everything but digits, matching the original phone formatting."""
    return re.sub(r"[^0-9]", "", _s(val))


def _numbered_group(rows_out, base_tag, values, checkbox=False):
    """
    Append (field, value) pairs for a repeatable single-tag group
    (FWSProgram, costCenter, dataStandard, etc.), replicating the original
    behavior where a tag ONLY gets a numeric suffix if it occurs more than
    once in the document.
    """
    values = [v for v in values if v not in (None, "")]
    if len(values) == 0:
        return
    if len(values) == 1:
        val = values[0]
        rows_out.append((base_tag, _b(val) if checkbox else _s(val)))
    else:
        for i, val in enumerate(values, start=1):
            rows_out.append((base_tag + str(i), _b(val) if checkbox else _s(val)))


def _paired_group(rows_out, tag_pairs, list_of_dicts):
    """
    Append interleaved (field, value) pairs for a repeatable group made of
    MULTIPLE tags per item (e.g. repoName+repoURL, or
    recordsSchedule+recordsType+recordsDisposition), replicating the
    "only number if count>1" rule per-tag.
    """
    n = len(list_of_dicts)
    if n == 0:
        return
    for tag, key in tag_pairs:
        for i, item in enumerate(list_of_dicts, start=1):
            field = tag if n == 1 else tag + str(i)
            rows_out.append((field, _s(item.get(key, ""))))


def build_dataframe(json_path, dmp_version_default="1.3"):
    """
    Read the JSON exported by DMP_Interface.html and return
    (df, dmp, dmpvers) ready to drop into the notebook in place of the old
    docx-scraping cells.
    """
    with open(json_path, "r", encoding="utf-8") as fh:
        data = json.load(fh)

    rows = []  # list of (field, value) tuples, built in document order

    # ---- DMP header info -------------------------------------------------
    d = data.get("dmp", {})
    rows.append(("dmpVersion", _s(d.get("version") or dmp_version_default)))
    rows.append(("dmpCreateDate", _s(d.get("createDate"))))
    rows.append(("dmpCreatorFirstName", _s(d.get("creatorFirstName"))))
    rows.append(("dmpCreatorLastName", _s(d.get("creatorLastName"))))
    rows.append(("dmpCreatorEmail", _s(d.get("creatorEmail"))))
    rows.append(("dmpUpdateDate", _s(d.get("updateDate"))))
    rows.append(("dmpUpdaterFirstName", _s(d.get("updaterFirstName"))))
    rows.append(("dmpUpdaterLastName", _s(d.get("updaterLastName"))))
    rows.append(("dmpUpdaterEmail", _s(d.get("updaterEmail"))))
    rows.append(("dmpSubmitDate", _s(d.get("submitDate"))))
    rows.append(("dmpDMFirstName", _s(d.get("dmFirstName"))))
    rows.append(("dmpDMLastName", _s(d.get("dmLastName"))))
    rows.append(("dmpDMEmail", _s(d.get("dmEmail"))))

    # ---- Project details ---------------------------------------------------
    p = data.get("project", {})
    rows.append(("fwsRegion", _s(p.get("region"))))
    _numbered_group(rows, "FWSProgram", p.get("programs", []))
    _numbered_group(
        rows,
        "costCenter",
        [_cost_center_code(c) for c in p.get("costCenters", [])],
    )
    rows.append(("projectTitle", _s(p.get("title"))))
    rows.append(("projectStartDate", _s(p.get("startDate"))))
    rows.append(("projectEndDate", _s(p.get("endDate"))))
    rows.append(("projectOngoing", _b(p.get("ongoing"))))
    rows.append(("projectAbstract", _s(p.get("abstract"))))
    rows.append(("projectPurpose", _s(p.get("purpose"))))  # NEW -- maps to mdJSON resourceInfo.purpose
    #rows.append(("projectKeywords", _s(p.get("keywords"))))  
    # The "keywords" field in the DMP_Interface.html JSON is now split into two separate arrays: "themeKeywords" and "placeKeywords". The original notebook code combined these into a single "projectKeywords" field, but we now store them separately for clarity and backward compatibility.
    rows.append(("projectThemeKeywords", _s(p.get("themeKeywords"))))
    rows.append(("projectPlaceKeywords", _s(p.get("placeKeywords"))))
    # Backward-compat combined field, still consumed by older downstream
    # code (e.g. the SharePoint list export cell in the notebook) that
    # expects a single comma-separated projectKeywords value.
    _combined_kw = ", ".join(
        [v for v in [_s(p.get("themeKeywords")), _s(p.get("placeKeywords"))] if v]
    )
    rows.append(("projectKeywords", _combined_kw))


    rows.append(("projectUIDList", _s(p.get("uidList"))))
    rows.append(("projectSpatialDesc", _s(p.get("spatialDesc"))))
    rows.append(("projectSpatialURL", _s(p.get("spatialURL"))))

    # ---- Project personnel / contacts --------------------------------------
    c = data.get("contacts", {})

    primary = c.get("primary", {})
    rows.append(("contactPrimaryFirstName", _s(primary.get("firstName"))))
    rows.append(("contactPrimaryLastName", _s(primary.get("lastName"))))
    rows.append(("contactPrimaryPosition", _s(primary.get("position"))))
    rows.append(("contactPrimaryEmail", _s(primary.get("email"))))
    rows.append(("contactPrimaryPhone", _phone_digits(primary.get("phone"))))
    rows.append(("contactPrimaryOrg", _s(primary.get("org"))))

    steward = c.get("steward", {})
    rows.append(("contactStewardFirstName", _s(steward.get("firstName"))))
    rows.append(("contactStewardLastName", _s(steward.get("lastName"))))
    rows.append(("contactStewardPosition", _s(steward.get("position"))))
    rows.append(("contactStewardEmail", _s(steward.get("email"))))
    rows.append(("contactStewardPhone", _phone_digits(steward.get("phone"))))
    rows.append(("contactStewardOrg", _s(steward.get("org"))))

    custodian = c.get("custodian", {})
    rows.append(("contactCustodianFirstName", _s(custodian.get("firstName"))))
    rows.append(("contactCustodianLastName", _s(custodian.get("lastName"))))
    rows.append(("contactCustodianPosition", _s(custodian.get("position"))))
    rows.append(("contactCustodianEmail", _s(custodian.get("email"))))
    rows.append(("contactCustodianPhone", _phone_digits(custodian.get("phone"))))
    rows.append(("contactCustodianNonDOI", _b(custodian.get("nonDOI"))))
    rows.append(("contactCustodianOrg", _s(custodian.get("org"))))

    trustee = c.get("trustee", {})
    rows.append(("contactTrusteeFirstName", _s(trustee.get("firstName"))))
    rows.append(("contactTrusteeLastName", _s(trustee.get("lastName"))))
    rows.append(("contactTrusteePosition", _s(trustee.get("position"))))
    rows.append(("contactTrusteeEmail", _s(trustee.get("email"))))
    rows.append(("contactTrusteePhone", _phone_digits(trustee.get("phone"))))

    others = c.get("other", [])
    n_other = len(others)
    for i, oc in enumerate(others, start=1):
        suffix = "" if n_other == 1 else str(i)
        rows.append(("contactOtherFirstName" + suffix, _s(oc.get("firstName"))))
        rows.append(("contactOtherLastName" + suffix, _s(oc.get("lastName"))))
        rows.append(("contactOtherPosition" + suffix, _s(oc.get("position"))))
        rows.append(("contactOtherEmail" + suffix, _s(oc.get("email"))))
        rows.append(("contactOtherPhone" + suffix, _phone_digits(oc.get("phone"))))
        rows.append(("contactOtherRole" + suffix, _s(oc.get("role"))))
        rows.append(("contactOtherOrg" + suffix, _s(oc.get("org"))))

    # ---- Metadata & data standards -----------------------------------------
    ms = data.get("metadataStandards", {})
    rows.append(("metaStandardMdJSON", _b(ms.get("mdJSON"))))
    rows.append(("metaStandardOther", _b(ms.get("other"))))
    rows.append(("metaStandardOtherType", _s(ms.get("otherType"))))

    _numbered_group(rows, "dataStandard", data.get("dataStandards", []))

    # ---- Storage / backup / review -----------------------------------------
    st = data.get("storage", {})
    for key, tag_prefix in [
        ("oneDrive", "storageOneDrive"),
        ("teams", "storageTeams"),
        ("external", "storageExternal"),
        ("network", "storageNetwork"),
        ("other", "storageOther"),
    ]:
        block = st.get(key, {})
        rows.append((tag_prefix, _b(block.get("checked"))))
        loc_key = "url" if "url" in block else "loc"
        rows.append((tag_prefix + ("URL" if key in ("oneDrive", "teams", "network") else "Loc"),
                      _s(block.get(loc_key))))

    rows.append(("storageLocationURL", _s(st.get("locationURL"))))
    rows.append(("backupFreq", _s(st.get("backupFreq"))))
    rows.append(("backupFreqOther", _s(st.get("backupFreqOther"))))
    rows.append(("filePermissions", _s(st.get("filePermissions"))))
    rows.append(("dataReviewSchedule", _s(st.get("reviewSchedule"))))

    # ---- Distribution / RDR / repositories / records -----------------------
    dist = data.get("distribution", {})
    rows.append(("rdrCheckbox", _b(dist.get("rdrChecked"))))
    rows.append(("rdrURL", _s(dist.get("rdrURL"))))
    rows.append(("projectShortTitle", _s(dist.get("shortTitle"))))

    _paired_group(
        rows,
        [("repoName", "name"), ("repoURL", "url")],
        dist.get("repositories", []),
    )
    _paired_group(
        rows,
        [
            ("recordsSchedule", "schedule"),
            ("recordsType", "type"),
            ("recordsDisposition", "disposition"),
        ],
        dist.get("recordsSchedules", []),
    )

    # ---- Products -----------------------------------------------------------
    for i, prod in enumerate(data.get("products", []), start=1):
        rows.append(("productNo" + str(i), str(i)))
        rows.append(("productTitle" + str(i), _s(prod.get("title"))))
        rows.append(("productName" + str(i), _s(prod.get("name"))))
        rows.append(("productURL" + str(i), _s(prod.get("url"))))
        rows.append(("productFileVersioning" + str(i), _s(prod.get("fileVersioning"))))
        rows.append(("productMetaURL" + str(i), _s(prod.get("metaURL"))))
        rows.append(("productSkipMeta" + str(i), _b(prod.get("skipMeta"))))
        rows.append(("productResourceType" + str(i), _s(prod.get("resourceType"))))
        rows.append(("productFormat" + str(i), _s(prod.get("format"))))
        rows.append(("productFormatOther" + str(i), _s(prod.get("formatOther"))))
        rows.append(("productAbstract" + str(i), _s(prod.get("abstract"))))

        for j, orig in enumerate(prod.get("originators", []), start=1):
            base = str(i) + "-" + str(j)
            rows.append(("productOriginatorFirstName" + base, _s(orig.get("firstName"))))
            rows.append(("productOriginatorLastName" + base, _s(orig.get("lastName"))))
            rows.append(("productOriginatorEmail" + base, _s(orig.get("email"))))
            rows.append(("productOriginatorNonDOI" + base, _b(orig.get("nonDOI"))))
            rows.append(("productOriginatorOrg" + base, _s(orig.get("org"))))

        rows.append(("productQualityAssurance" + str(i), _s(prod.get("qualityAssurance"))))
        rows.append(("productQualityControl" + str(i), _s(prod.get("qualityControl"))))
        rows.append(("productResources" + str(i), _s(prod.get("resources"))))

        for j, ma in enumerate(prod.get("metaAuthors", []), start=1):
            base = str(i) + "-" + str(j)
            rows.append(("productMetaAuthorFirstName" + base, _s(ma.get("firstName"))))
            rows.append(("productMetaAuthorLastName" + base, _s(ma.get("lastName"))))
            rows.append(("productMetaAuthorEmail" + base, _s(ma.get("email"))))
            rows.append(("productMetaAuthorOrg" + base, _s(ma.get("org"))))

        rows.append(("productRestriction" + str(i), _s(prod.get("restriction"))))
        rows.append(("productRestrictionJustification" + str(i), _s(prod.get("restrictionJustification"))))
        rows.append(("productAdditional" + str(i), _s(prod.get("additional"))))
        rows.append(("productSpatialMatchesProject" + str(i), _b(prod.get("spatialMatchesProject"))))
        rows.append(("productSpatialDesc" + str(i), _s(prod.get("spatialDesc"))))
        rows.append(("productSpatialURL" + str(i), _s(prod.get("spatialURL"))))

    # ---- Source data ----------------------------------------------------------
    src_list = data.get("sourceData", [])
    n_src = len(src_list)
    for i, src in enumerate(src_list, start=1):
        suffix = "" if n_src == 1 else str(i)
        rows.append(("sourceDataDescription" + suffix, _s(src.get("description"))))
        rows.append(("sourceDataLoc" + suffix, _s(src.get("location"))))

    # ---- Physical samples -------------------------------------------------
    for i, samp in enumerate(data.get("samples", []), start=1):
        rows.append(("sampleNo" + str(i), str(i)))
        rows.append(("sampleName" + str(i), _s(samp.get("name"))))
        rows.append(("sampleDescription" + str(i), _s(samp.get("description"))))
        rows.append(("sampleLabel" + str(i), _s(samp.get("label"))))
        rows.append(("sampleQAQC" + str(i), _s(samp.get("qaqc"))))
        rows.append(("sampleLoc" + str(i), _s(samp.get("loc"))))
        rows.append(("sampleChain" + str(i), _s(samp.get("chain"))))
        rows.append(("sampleFate" + str(i), _s(samp.get("fate"))))
        rows.append(("sampleProducts" + str(i), _s(samp.get("products"))))

    # ---- Assemble dataframe -------------------------------------------------
    df = pd.DataFrame(rows, columns=["field", "value"])

    # prod_num / sample_num -- identical logic to the original notebook cell
    df["prod_num"] = df["field"].str.extract(r"([0-9]{0,}-?[0-9]{1,})")
    df["prod_num"] = df["prod_num"].str.replace(r"-[0-9]{1,}", "", regex=True)
    df.loc[(df.field.str.startswith("product")) & (df.prod_num.isna()), "prod_num"] = "1"
    df.loc[df.field.str.startswith("product") == False, "prod_num"] = None  # noqa: E712

    df["sample_num"] = df["field"].str.extract(r"([0-9]{0,}-?[0-9]{1,})")
    df.loc[(df.field.str.startswith("sample")) & (df.sample_num.isna()), "sample_num"] = "1"
    df.loc[df.field.str.startswith("sample") == False, "sample_num"] = None  # noqa: E712

    dmp_path = _s(d.get("documentPath")) or json_path
    dmpvers = _s(d.get("version")) or dmp_version_default

    return df, dmp_path, dmpvers


if __name__ == "__main__":
    import sys

    path = sys.argv[1] if len(sys.argv) > 1 else "dmp_data.json"
    out_df, out_dmp, out_ver = build_dataframe(path)
    pd.set_option("display.max_columns", None)
    pd.set_option("display.max_colwidth", 60)
    print(f"dmp = {out_dmp!r}")
    print(f"dmpvers = {out_ver!r}")
    print(out_df)
