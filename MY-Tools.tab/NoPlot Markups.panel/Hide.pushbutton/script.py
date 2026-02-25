# -*- coding: utf-8 -*-

__doc__ = "Hide annotations across the project by prefix and remember IDs + last prefix."

from pyrevit import revit, DB, forms, script

from System.Collections.Generic import List

import os

import json

doc = revit.doc

logger = script.get_logger()


# -----------------------------

# Storage helpers (per document)

# -----------------------------

def safe_file_part(s):
    bad = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']

    for ch in bad:
        s = s.replace(ch, "_")

    return s


def get_store_path():
    base = os.environ.get("APPDATA", "")

    folder = os.path.join(base, "pyRevit_PrefixAnno")

    if not os.path.exists(folder):
        os.makedirs(folder)

    doc_key = safe_file_part(doc.PathName if doc.PathName else doc.Title)

    return os.path.join(folder, "{}.json".format(doc_key))


def load_store():
    path = get_store_path()

    if not os.path.exists(path):
        return {"last_prefix": "NOPLOT", "prefix_ids": {}}

    try:

        with open(path, "r") as f:

            data = json.load(f)

        if "last_prefix" not in data:
            data["last_prefix"] = "NOPLOT"

        if "prefix_ids" not in data:
            data["prefix_ids"] = {}

        return data

    except:

        return {"last_prefix": "NOPLOT", "prefix_ids": {}}


def save_store(data):
    with open(get_store_path(), "w") as f:
        json.dump(data, f)


store = load_store()

default_prefix = store.get("last_prefix", "NOPLOT")

# -----------------------------

# Ask prefix (remember last)

# -----------------------------

prefix = forms.ask_for_string(

    default=default_prefix,

    prompt="Enter prefix to hide across the project:",

    title="Hide Annotations by Prefix (Project-wide)"

)

if prefix is None:
    script.exit()

prefix = prefix.strip()

if not prefix:
    forms.alert("Prefix is empty. Cancelled.", exitscript=True)

prefix_upper = prefix.upper()

# save last prefix immediately

store["last_prefix"] = prefix

save_store(store)


# -----------------------------

# Helpers

# -----------------------------

def safe_str(val):
    try:

        if val is None:
            return ""

        return str(val)

    except:

        return ""


def starts_with_prefix(text):
    return safe_str(text).strip().upper().startswith(prefix_upper)


def get_type_name_and_family(elem):
    type_name = ""

    family_name = ""

    try:

        tid = elem.GetTypeId()

        if tid and tid != DB.ElementId.InvalidElementId:

            et = doc.GetElement(tid)

            if et:

                try:

                    p = et.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)

                    if p:
                        type_name = safe_str(p.AsString())

                except:

                    pass

                if not type_name:

                    try:

                        type_name = safe_str(et.Name)

                    except:

                        pass

                try:

                    if hasattr(et, "FamilyName"):

                        family_name = safe_str(et.FamilyName)

                    elif hasattr(et, "Family") and et.Family:

                        family_name = safe_str(et.Family.Name)

                except:

                    pass

    except:

        pass

    if not type_name:

        try:

            type_name = safe_str(elem.Name)

        except:

            pass

    return type_name, family_name


def is_annotation_like(elem):
    try:

        cat = elem.Category

        if cat is None:
            return False

        # Most annotations

        try:

            if cat.CategoryType == DB.CategoryType.Annotation:
                return True

        except:

            pass

        # Detail lines / detail curves

        if isinstance(elem, DB.DetailCurve):
            return True

    except:

        pass

    return False


def matches_prefix(elem):
    # 1) Text note text

    try:

        if isinstance(elem, DB.TextNote):

            if starts_with_prefix(elem.Text):
                return True

    except:

        pass

    # 2) Dimension text fields

    try:

        if isinstance(elem, DB.Dimension):

            for a in ["Prefix", "Suffix", "Above", "Below", "ValueOverride"]:

                try:

                    if starts_with_prefix(getattr(elem, a)):
                        return True

                except:

                    pass

    except:

        pass

    # 3) Detail line style name

    try:

        if isinstance(elem, DB.DetailCurve):

            try:

                gs = elem.LineStyle

                if gs and starts_with_prefix(gs.Name):
                    return True

            except:

                pass

    except:

        pass

    # 4) Tag text if available

    try:

        if isinstance(elem, DB.IndependentTag):

            try:

                if hasattr(elem, "TagText") and starts_with_prefix(elem.TagText):
                    return True

            except:

                pass

    except:

        pass

    # 5) Type / family fallback

    tname, fname = get_type_name_and_family(elem)

    if starts_with_prefix(tname) or starts_with_prefix(fname):
        return True

    # 6) String parameter scan fallback (very important)

    try:

        for p in elem.Parameters:

            try:

                if p.StorageType == DB.StorageType.String:

                    s = p.AsString()

                    if starts_with_prefix(s):
                        return True

            except:

                pass

    except:

        pass

    return False


def get_owner_view(elem):
    try:

        ovid = elem.OwnerViewId

        if ovid and ovid != DB.ElementId.InvalidElementId:
            return doc.GetElement(ovid)

    except:

        pass

    return None


# -----------------------------

# Scan whole project + group IDs by owner view

# -----------------------------

all_elems = DB.FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()

by_view = {}  # {viewIdInt: [ElementId, ...]}

matched_ids = []

checked = 0

for elem in all_elems:

    try:

        if elem is None:
            continue

        if not is_annotation_like(elem):
            continue

        owner_view = get_owner_view(elem)

        if owner_view is None:
            continue

        try:

            if not elem.CanBeHidden(owner_view):
                continue

        except:

            continue

        checked += 1

        if not matches_prefix(elem):
            continue

        vid = owner_view.Id.IntegerValue

        if vid not in by_view:
            by_view[vid] = []

        by_view[vid].append(elem.Id)

        matched_ids.append(elem.Id.IntegerValue)

    except Exception as ex:

        try:

            logger.debug("Skip element: {}".format(ex))

        except:

            pass

if not matched_ids:
    forms.alert("No matching annotations found for prefix '{}'.".format(prefix), exitscript=True)

# -----------------------------

# Hide by owner view

# -----------------------------

views_touched = 0

hidden_count = 0

with revit.Transaction("Hide annotations by prefix (project-wide)"):
    for vid_int, eids in by_view.items():

        try:

            view = doc.GetElement(DB.ElementId(vid_int))

            if view is None:
                continue

            id_list = List[DB.ElementId]()

            for eid in eids:
                id_list.Add(eid)

            if id_list.Count > 0:
                view.HideElements(id_list)

                views_touched += 1

                hidden_count += id_list.Count

        except Exception as ex:

            try:

                logger.debug("Hide failed in view {}: {}".format(vid_int, ex))

            except:

                pass

# -----------------------------

# Save hidden IDs by prefix (project-wide)

# -----------------------------

prefix_ids = store.get("prefix_ids", {})

existing = prefix_ids.get(prefix_upper, [])

existing_set = set([int(x) for x in existing])

for i in matched_ids:
    existing_set.add(i)

prefix_ids[prefix_upper] = sorted(list(existing_set))

store["prefix_ids"] = prefix_ids

save_store(store)

forms.alert(

    "Done.\nPrefix: {}\nChecked annotations: {}\nHidden: {}\nViews touched: {}\nSaved IDs for unhide.".format(

        prefix, checked, hidden_count, views_touched

    ),

    title="Hide Annotations by Prefix"

)
