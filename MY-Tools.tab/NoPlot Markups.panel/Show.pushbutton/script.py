# -*- coding: utf-8 -*-

__doc__ = "Unhide annotations across the project by prefix using saved IDs; remembers last prefix."

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

    prompt="Enter prefix to unhide across the project:",

    title="Unhide Annotations by Prefix (Project-wide)"

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

prefix_ids = store.get("prefix_ids", {})

saved_ids = prefix_ids.get(prefix_upper, [])

if not saved_ids:
    forms.alert(

        "No saved IDs found for prefix '{}'.\nUse the Hide script first (project-wide version).".format(prefix),

        exitscript=True

    )

# -----------------------------

# Group saved IDs by owner view

# -----------------------------

by_view = {}  # {viewIdInt: [ElementId, ...]}

valid_count = 0

missing_count = 0

for raw_id in saved_ids:

    try:

        eid = DB.ElementId(int(raw_id))

        elem = doc.GetElement(eid)

        if elem is None:
            missing_count += 1

            continue

        try:

            ovid = elem.OwnerViewId

            if not ovid or ovid == DB.ElementId.InvalidElementId:
                continue

        except:

            continue

        view_id_int = ovid.IntegerValue

        if view_id_int not in by_view:
            by_view[view_id_int] = []

        by_view[view_id_int].append(eid)

        valid_count += 1

    except Exception as ex:

        try:

            logger.debug("Bad saved id {}: {}".format(raw_id, ex))

        except:

            pass

if valid_count == 0:
    forms.alert(

        "Saved IDs exist, but none were found in this model anymore.",

        exitscript=True

    )

# -----------------------------

# Unhide by owner view

# -----------------------------

views_touched = 0

unhidden_count = 0

still_saved = []

with revit.Transaction("Unhide annotations by prefix (project-wide)"):
    for vid_int, eids in by_view.items():

        try:

            view = doc.GetElement(DB.ElementId(vid_int))

            if view is None:
                continue

            id_list = List[DB.ElementId]()

            for eid in eids:

                # keep IDs in memory in case unhide fails for some

                try:

                    id_list.Add(eid)

                except:

                    pass

            if id_list.Count == 0:
                continue

            view.UnhideElements(id_list)

            views_touched += 1

            unhidden_count += id_list.Count

        except Exception as ex:

            # If a whole view batch fails, keep those IDs saved

            try:

                logger.debug("Unhide failed in view {}: {}".format(vid_int, ex))

            except:

                pass

            for eid in eids:
                still_saved.append(eid.IntegerValue)

# -----------------------------

# Clean saved IDs for this prefix

# (remove successfully unhidden + missing ones)

# -----------------------------

# If any batch failed, keep only failed IDs. Otherwise clear all.

if len(still_saved) > 0:

    prefix_ids[prefix_upper] = sorted(list(set(still_saved)))

else:

    prefix_ids[prefix_upper] = []

store["prefix_ids"] = prefix_ids

save_store(store)

forms.alert(

    "Done.\nPrefix: {}\nUnhidden: {}\nViews touched: {}\nMissing/deleted IDs skipped: {}".format(

        prefix, unhidden_count, views_touched, missing_count

    ),

    title="Unhide Annotations by Prefix"

)
