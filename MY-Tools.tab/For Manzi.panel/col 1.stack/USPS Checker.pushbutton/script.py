# -*- coding: utf-8 -*-

from pyrevit import revit, DB, script
from collections import defaultdict
import re
import math
from datetime import datetime

doc = revit.doc
output = script.get_output()

# ------------------------------------------------------------
# Settings
# ------------------------------------------------------------

Z_TOL = 0.05
SINK_CABINET_CHECK_DEPTH = 10.0

GREEN_CHECK = u"<span style='color:#008000; font-weight:900;'>&#10003;</span>"


# ------------------------------------------------------------
# Standard output labels
# ------------------------------------------------------------

LABEL_WATER_CLOSET = "WATER CLOSET"
LABEL_URINAL = "URINAL"
LABEL_LAVATORY = "LAVATORY"

LABEL_RA1_TOILET_TISSUE = "RA-1 TOILET TISSUE DISPENSER"
LABEL_RA3_PAPER_TOWEL = "RA-3 PAPER TOWEL DISPENSER"
LABEL_RA5_NAPKIN_DISPENSER = "RA-5 NAPKIN DISPENSER"
LABEL_RA6_NAPKIN_DISPOSAL = "RA-6 NAPKIN DISPOSAL"
LABEL_RA7_SOAP = "RA-7 SOAP DISPENSER"
LABEL_RA9_MIRROR_36 = 'RA-9 MIRROR 36" HEIGHT'
LABEL_RA10_MIRROR_60 = "RA-10 MIRROR 60 HEIGHT"
LABEL_RA11_SEAT_COVER = "RA-11 TOILET SEAT COVER DISPENSER"

LABEL_A1_REFRIGERATOR = "A-1 REFRIGERATOR"
LABEL_A2_MICROWAVE = "A-2 MICROWAVE"
LABEL_A3_ICE_MACHINE = "A-3 ICE MACHINE"


# ------------------------------------------------------------
# Revision Check keywords
# ------------------------------------------------------------

SURFACE_MOUNTED_WASTE_FAMILY = "USPS_Surface-Mounted_Waste_Disposal"

AD100_SHEET_NUMBER = "AD100"
LOG_GENERAL_NOTES_TEXT = "LOG GENERAL NOTES"

BMEU_DRAFTING_VIEW_NAME = "TYPICAL BMEU 732 MOUNTING HEIGHT"

WIRE_SCREEN_VIEW_NAME = "WIRE SCREEN ENCLOSURE ELEVATION (TYP.)"
WIRE_SCREEN_GAP_TEXT = '1/2" MAX. ALLOWABLE GAP'

SINK_BASE_DRAFTING_VIEW_NAME = "SECTION THROUGH TYP. SINK BASE"
SINK_BASE_SWING_WITH_DOORS_TEXT = "swing with door"

MATERIALS_FINISH_LEGEND_SCHEDULE_NAME = "MATERIALS FINISH LEGEND"
MATERIALS_WALL_BASE_TEXT = 'WALL BASE 6"'
MATERIALS_RUNNING_BOND_TEXT = "1/3 offset running bond"
MATERIALS_GRAYSON_TEXT = "GRAYSON"


# ------------------------------------------------------------
# Basic helpers
# ------------------------------------------------------------

def get_param_text(elem, bip):
    p = elem.get_Parameter(bip)
    if p:
        return p.AsString() or ""
    return ""


def get_room_number(room):
    return get_param_text(room, DB.BuiltInParameter.ROOM_NUMBER)


def get_room_name(room):
    return get_param_text(room, DB.BuiltInParameter.ROOM_NAME)


def get_room_display_name(room):
    num = get_room_number(room)
    name = get_room_name(room)

    if num and name:
        return "{} - {}".format(num, name)
    elif num:
        return num
    elif name:
        return name
    else:
        return "Room Element ID {}".format(room.Id)


def get_room_id_key(room):
    try:
        return room.Id.IntegerValue
    except:
        return str(room.Id)


def get_family_name(inst):
    try:
        return inst.Symbol.Family.Name or ""
    except:
        return ""


def get_type_name(inst):
    try:
        p = inst.Symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if p:
            return p.AsString() or inst.Name
    except:
        pass

    return inst.Name


def get_family_type_name(inst):
    fam = get_family_name(inst)
    typ = get_type_name(inst)

    if fam and typ:
        return "{} : {}".format(fam, typ)
    elif fam:
        return fam
    else:
        return typ


def get_name_blob(inst):
    return get_family_type_name(inst).lower()


def get_element_id_key(elem):
    try:
        return elem.Id.IntegerValue
    except:
        return str(elem.Id)


def get_bbox(elem):
    try:
        return elem.get_BoundingBox(None)
    except:
        return None


def normalize_text(text):
    if not text:
        return ""
    return re.sub(r'[^a-z0-9]+', ' ', text.lower()).strip()


def compact_text(text):
    if not text:
        return ""
    return re.sub(r'[^a-z0-9]+', '', text.lower())


def contains_phrase_in_name_blob(inst, phrase):
    name_blob = get_name_blob(inst)
    normalized = normalize_text(name_blob)
    compact = compact_text(name_blob)

    phrase_normalized = normalize_text(phrase)
    phrase_compact = compact_text(phrase)

    return (
        phrase_normalized in normalized or
        phrase_compact in compact
    )


def text_blob_contains(source_text, target_text):
    if not source_text or not target_text:
        return False

    source_upper = source_text.upper()
    target_upper = target_text.upper()

    if target_upper in source_upper:
        return True

    source_norm = normalize_text(source_text)
    target_norm = normalize_text(target_text)

    if target_norm and target_norm in source_norm:
        return True

    return False


def type_has_number(inst, number_text):
    typ = get_type_name(inst).lower()
    pattern = r'(?<!\d)' + re.escape(str(number_text)) + r'(?!\d)'
    return re.search(pattern, typ) is not None


def get_message_label(display_label):
    return display_label.upper()


# ------------------------------------------------------------
# Category helpers
# ------------------------------------------------------------

def is_category(inst, bic):
    try:
        return (
            inst.Category and
            inst.Category.Id.IntegerValue == int(bic)
        )
    except:
        return False


def is_plumbing_fixture(inst):
    return is_category(inst, DB.BuiltInCategory.OST_PlumbingFixtures)


def is_generic_annotation(inst):
    return is_category(inst, DB.BuiltInCategory.OST_GenericAnnotation)


# ------------------------------------------------------------
# Phase helpers
# ------------------------------------------------------------

def is_existing_phase_element(elem):
    try:
        phase_id = elem.CreatedPhaseId

        if phase_id == DB.ElementId.InvalidElementId:
            return False

        phase = doc.GetElement(phase_id)

        if not phase:
            return False

        return "existing" in phase.Name.lower()

    except:
        return False


def is_demolished_element(elem):
    try:
        demo_phase_id = elem.DemolishedPhaseId

        if demo_phase_id == DB.ElementId.InvalidElementId:
            return False

        return True

    except:
        return False


def room_has_any_non_existing_detected_family(room_id, room_detected_ids, room_non_existing_detected_ids):
    detected_ids = room_detected_ids.get(room_id, set())

    if not detected_ids:
        return False

    non_existing_ids = room_non_existing_detected_ids.get(room_id, set())

    return len(non_existing_ids) > 0


# ------------------------------------------------------------
# Output helpers
# ------------------------------------------------------------

def html_escape(value):
    if value is None:
        return ""

    text = str(value)
    text = text.replace("&", "&amp;")
    text = text.replace("<", "&lt;")
    text = text.replace(">", "&gt;")
    text = text.replace('"', "&quot;")
    text = text.replace("'", "&#39;")
    return text


def get_revit_file_name():
    try:
        path = doc.PathName

        if path:
            file_name = path.split("\\")[-1].split("/")[-1]

            if file_name:
                return file_name
    except:
        pass

    try:
        if doc.Title:
            return doc.Title
    except:
        pass

    return "Current Revit Model"


def print_report_header():
    output.print_html(
        "<div style='font-size:24px; font-weight:bold; margin-bottom:2px;'>"
        "USPS REVISION AND QUANTITY CHECK"
        "</div>"
    )

    output.print_html(
        "<div style='font-size:13px; color:#666666; margin-bottom:8px;'>"
        "{}"
        "</div>".format(html_escape(get_revit_file_name()))
    )

    output.print_md("---")


def print_green_check_line(label, count):
    output.print_md(u"- **{}** Count: {} {}".format(label, count, GREEN_CHECK))


def print_plain_count_line(label, count):
    output.print_md("- **{}** Count: {}".format(label, count))


def print_revision_ok_line(message):
    output.print_html(
        "<div style='line-height:1.35; font-size:14px;'>"
        "- <b>{}</b> {}"
        "</div>".format(
            html_escape(message.upper()),
            GREEN_CHECK
        )
    )


def print_revision_red_line(message):
    output.print_html(
        "<div style='color:red; font-weight:bold; font-size:14px; line-height:1.35;'>"
        "- {}"
        "</div>".format(html_escape(message.upper()))
    )


def print_red_message(message):
    output.print_html(
        "<div style='color:red; font-weight:bold; font-size:14px; margin-left:52px;'>"
        "{}"
        "</div>".format(html_escape(message.upper()))
    )


def print_orange_message(message):
    output.print_html(
        "<div style='color:#d97706; font-weight:bold; font-size:14px; margin-left:52px;'>"
        "{}"
        "</div>".format(html_escape(message))
    )


def print_room_divider():
    output.print_md("---")


def validate_equal_to_required(label, actual_count, required_count):
    msg_label = get_message_label(label)

    if actual_count == required_count:
        print_green_check_line(label, actual_count)
    else:
        print_plain_count_line(label, actual_count)

        diff = actual_count - required_count

        if diff > 0:
            print_red_message("{} HAS {} EXTRA.".format(msg_label, diff))
        else:
            print_red_message("{} IS SHORT BY {}.".format(msg_label, abs(diff)))


def validate_minimum_required(label, actual_count, required_count):
    msg_label = get_message_label(label)

    if actual_count >= required_count:
        print_green_check_line(label, actual_count)
    else:
        print_plain_count_line(label, actual_count)
        print_red_message("{} IS SHORT BY {}.".format(msg_label, required_count - actual_count))

def print_room_area_check_line(room):
    area_sf = room.Area

    if area_sf >= 750:
        output.print_md("- **ROOM AREA**: {:.0f} SF".format(area_sf))
        print_red_message("Room area is 750 SF or greater. This room may need two out-swinging doors, please check.")
    else:
        output.print_md(u"- **ROOM AREA**: {:.0f} SF {}".format(area_sf, GREEN_CHECK))


# ------------------------------------------------------------
# Filter helpers
# ------------------------------------------------------------ro

def is_model_family_instance(inst):
    try:
        return (
            inst.Category and
            inst.Category.CategoryType == DB.CategoryType.Model
        )
    except:
        return False


def is_seat_cover(inst):
    return (
        contains_phrase_in_name_blob(inst, "seat cover") or
        contains_phrase_in_name_blob(inst, "toilet seat cover")
    )


def is_soap_dispenser(inst):
    return "soap" in get_name_blob(inst)


def is_toilet_tissue_dispenser(inst):
    return (
        contains_phrase_in_name_blob(inst, "toilet paper") or
        contains_phrase_in_name_blob(inst, "toilet tissue")
    )


def is_napkin_disposal(inst):
    return contains_phrase_in_name_blob(inst, "napkin disposal")


def is_napkin_dispenser(inst):
    return (
        contains_phrase_in_name_blob(inst, "napkin dispenser") or
        "tampon" in get_name_blob(inst)
    )


def is_paper_towel_dispenser(inst):
    return (
        contains_phrase_in_name_blob(inst, "towel dispenser") or
        contains_phrase_in_name_blob(inst, "paper towel")
    )


def is_mirror_family(inst):
    return "mirror" in get_family_name(inst).lower()


def is_mirror_36(inst):
    return is_mirror_family(inst) and type_has_number(inst, "36")


def is_mirror_60(inst):
    return is_mirror_family(inst) and type_has_number(inst, "60")


def is_lavatory(inst):
    if not is_plumbing_fixture(inst):
        return False

    name_blob = get_name_blob(inst)

    return (
        "lavatory" in name_blob or
        "sink" in name_blob
    )


def is_restroom_accessory(inst):
    return (
        is_seat_cover(inst) or
        is_soap_dispenser(inst) or
        is_mirror_36(inst) or
        is_mirror_60(inst) or
        is_toilet_tissue_dispenser(inst) or
        is_napkin_disposal(inst) or
        is_napkin_dispenser(inst) or
        is_paper_towel_dispenser(inst)
    )


def is_toilet_or_urinal_plumbing_fixture(inst):
    if not is_plumbing_fixture(inst):
        return False

    if is_restroom_accessory(inst):
        return False

    fam_name = get_family_name(inst).lower()

    return (
        "toilet" in fam_name or
        "urinal" in fam_name
    )


def get_restroom_fixture_type(inst):
    if not is_plumbing_fixture(inst):
        return None

    if is_restroom_accessory(inst):
        return None

    fam_name = get_family_name(inst).lower()

    if "urinal" in fam_name:
        return "urinal"

    if "toilet" in fam_name:
        return "water_closet"

    return None


def is_locker(inst):
    return "locker" in get_name_blob(inst)


def is_locker_room(room):
    return "locker" in get_room_name(room).lower()


def is_breakroom(room):
    return "breakroom" in get_room_name(room).lower()


def is_bmeu_room(room):
    return "bmeu" in get_room_name(room).lower()


def is_refrigerator(inst):
    return "refrigerator" in get_name_blob(inst)


def is_microwave(inst):
    return "microwave" in get_name_blob(inst)


def is_ice_machine(inst):
    family_name = get_family_name(inst).lower()
    name_blob = get_name_blob(inst)
    normalized = normalize_text(name_blob)

    return (
        "idt" in family_name or
        "ice machine" in normalized or
        "ice maker" in normalized or
        re.search(r'(^|[^a-z0-9])ice([^a-z0-9]|$)', name_blob) is not None
    )


def is_breakroom_appliance(inst):
    return (
        is_refrigerator(inst) or
        is_microwave(inst) or
        is_ice_machine(inst)
    )


def is_breakroom_sink(inst):
    if not is_plumbing_fixture(inst):
        return False

    return "sink" in get_name_blob(inst)


def is_base_cabinet(inst):
    return contains_phrase_in_name_blob(inst, "base cabinet")


def is_surface_mounted_waste_disposal(inst):
    family_name = get_family_name(inst)
    name_blob = get_name_blob(inst)

    return (
        compact_text(family_name) == compact_text(SURFACE_MOUNTED_WASTE_FAMILY) or
        compact_text(SURFACE_MOUNTED_WASTE_FAMILY) in compact_text(name_blob)
    )


def is_swing_gate(inst):
    return contains_phrase_in_name_blob(inst, "swing gate")


def is_revision_family(inst):
    return (
        is_surface_mounted_waste_disposal(inst) or
        is_swing_gate(inst) or
        is_breakroom_sink(inst) or
        is_base_cabinet(inst)
    )


def get_locker_tier(inst):
    name_blob = get_name_blob(inst)

    double_keywords = [
        "double tier",
        "double-tier",
        "double_tier",
        "doubletier",
        "double locker",
        "dbl tier",
        "dbl-tier",
        "dbl_tier",
        "2 tier",
        "2-tier",
        "2_tier",
        "2tier",
        "two tier",
        "two-tier",
        "two_tier",
        "twotier",
    ]

    single_keywords = [
        "single tier",
        "single-tier",
        "single_tier",
        "singletier",
        "single locker",
        "1 tier",
        "1-tier",
        "1_tier",
        "1tier",
        "one tier",
        "one-tier",
        "one_tier",
        "onetier",
    ]

    for kw in double_keywords:
        if kw in name_blob:
            return "double"

    if re.search(r'(^|[^a-z0-9])2\s*t([^a-z0-9]|$)', name_blob):
        return "double"

    for kw in single_keywords:
        if kw in name_blob:
            return "single"

    if re.search(r'(^|[^a-z0-9])1\s*t([^a-z0-9]|$)', name_blob):
        return "single"

    return None


# ------------------------------------------------------------
# View / Sheet / Schedule helpers
# ------------------------------------------------------------

def get_textnote_text(text_note):
    try:
        return text_note.Text or ""
    except:
        return ""


def find_sheet_by_number(sheet_number):
    sheets = list(
        DB.FilteredElementCollector(doc)
        .OfClass(DB.ViewSheet)
        .ToElements()
    )

    for sheet in sheets:
        try:
            if sheet.SheetNumber.upper() == sheet_number.upper():
                return sheet
        except:
            pass

    return None


def find_drafting_view_by_name(view_name):
    views = list(
        DB.FilteredElementCollector(doc)
        .OfClass(DB.View)
        .ToElements()
    )

    for view in views:
        try:
            if view.IsTemplate:
                continue

            if view.ViewType != DB.ViewType.DraftingView:
                continue

            if view.Name.upper() == view_name.upper():
                return view
        except:
            pass

    return None

def get_dimension_text_blob(dim_obj):
    texts = []

    # These correspond to dimension text fields:
    # Above / Below / Prefix / Suffix / Replace With Text
    for attr_name in ["Above", "Below", "Prefix", "Suffix", "ValueOverride"]:
        try:
            value = getattr(dim_obj, attr_name)
            if value:
                texts.append(value)
        except:
            pass

    return "\n".join(texts)


def dimension_contains_text(dimension, target_text):
    # Check the dimension itself, including Replace With Text / ValueOverride
    if text_blob_contains(get_dimension_text_blob(dimension), target_text):
        return True

    # Check each segment, for multi-segment dimensions
    try:
        segments = dimension.Segments

        if segments:
            for seg in segments:
                if text_blob_contains(get_dimension_text_blob(seg), target_text):
                    return True
    except:
        pass

    return False


def dimension_text_exists_in_view(view, target_text):
    try:
        dimensions = list(
            DB.FilteredElementCollector(doc, view.Id)
            .OfClass(DB.Dimension)
            .ToElements()
        )
    except:
        dimensions = []

    for dim in dimensions:
        if dimension_contains_text(dim, target_text):
            return True

    return False


def text_exists_in_view(view, target_text):
    # Check normal text notes
    try:
        text_notes = list(
            DB.FilteredElementCollector(doc, view.Id)
            .OfClass(DB.TextNote)
            .ToElements()
        )
    except:
        text_notes = []

    for text_note in text_notes:
        if text_blob_contains(get_textnote_text(text_note), target_text):
            return True

    # Check dimension text fields, including Replace With Text
    if dimension_text_exists_in_view(view, target_text):
        return True

    return False


def get_schedule_body_text(view_schedule):
    all_text = []

    try:
        table_data = view_schedule.GetTableData()
        body = table_data.GetSectionData(DB.SectionType.Body)

        first_row = body.FirstRowNumber
        last_row = body.LastRowNumber
        first_col = body.FirstColumnNumber
        last_col = body.LastColumnNumber

        for row in range(first_row, last_row + 1):
            for col in range(first_col, last_col + 1):
                cell_text = ""

                try:
                    cell_text = view_schedule.GetCellText(DB.SectionType.Body, row, col)
                except:
                    try:
                        cell_text = body.GetCellText(row, col)
                    except:
                        cell_text = ""

                if cell_text:
                    all_text.append(cell_text)

    except:
        pass

    return "\n".join(all_text)


def find_schedule_by_name(schedule_name):
    schedules = list(
        DB.FilteredElementCollector(doc)
        .OfClass(DB.ViewSchedule)
        .ToElements()
    )

    for schedule in schedules:
        try:
            if schedule.Name.upper() == schedule_name.upper():
                return schedule
        except:
            pass

    return None


def schedule_contains_text(schedule, target_text):
    if not schedule:
        return False

    try:
        if text_blob_contains(schedule.Name, target_text):
            return True
    except:
        pass

    schedule_text = get_schedule_body_text(schedule)

    return text_blob_contains(schedule_text, target_text)


def generic_annotation_name_matches(inst, target_text):
    if not is_generic_annotation(inst):
        return False

    search_text = ""

    try:
        search_text += get_family_name(inst) + "\n"
    except:
        pass

    try:
        search_text += get_type_name(inst) + "\n"
    except:
        pass

    try:
        search_text += inst.Name + "\n"
    except:
        pass

    return text_blob_contains(search_text, target_text)


def generic_annotation_family_exists_in_view(view, target_text):
    try:
        annotations = list(
            DB.FilteredElementCollector(doc, view.Id)
            .OfCategory(DB.BuiltInCategory.OST_GenericAnnotation)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except:
        annotations = []

    for ann in annotations:
        try:
            if generic_annotation_name_matches(ann, target_text):
                return True
        except:
            pass

    return False


def generic_annotation_family_exists_on_sheet(sheet, target_text):
    if not sheet:
        return False

    # Generic Annotation directly placed on the sheet.
    if generic_annotation_family_exists_in_view(sheet, target_text):
        return True

    # Generic Annotation inside views placed on the sheet.
    try:
        viewports = list(
            DB.FilteredElementCollector(doc, sheet.Id)
            .OfClass(DB.Viewport)
            .ToElements()
        )
    except:
        viewports = []

    for viewport in viewports:
        try:
            placed_view = doc.GetElement(viewport.ViewId)

            if placed_view and generic_annotation_family_exists_in_view(placed_view, target_text):
                return True
        except:
            pass

    return False


def text_exists_on_sheet(sheet, target_text):
    if not sheet:
        return False

    try:
        text_notes = list(
            DB.FilteredElementCollector(doc, sheet.Id)
            .OfClass(DB.TextNote)
            .ToElements()
        )
    except:
        text_notes = []

    for text_note in text_notes:
        if text_blob_contains(get_textnote_text(text_note), target_text):
            return True

    try:
        viewports = list(
            DB.FilteredElementCollector(doc, sheet.Id)
            .OfClass(DB.Viewport)
            .ToElements()
        )
    except:
        viewports = []

    for viewport in viewports:
        try:
            placed_view = doc.GetElement(viewport.ViewId)

            if placed_view and text_exists_in_view(placed_view, target_text):
                return True
        except:
            pass

    try:
        schedule_instances = list(
            DB.FilteredElementCollector(doc, sheet.Id)
            .OfClass(DB.ScheduleSheetInstance)
            .ToElements()
        )
    except:
        schedule_instances = []

    for schedule_instance in schedule_instances:
        try:
            schedule = doc.GetElement(schedule_instance.ScheduleId)

            if schedule:
                if text_blob_contains(schedule.Name, target_text):
                    return True

                if schedule_contains_text(schedule, target_text):
                    return True
        except:
            pass

    return False

def text_exists_in_any_sheet_drafting_view(target_text):
    seen_view_ids = set()

    try:
        viewports = list(
            DB.FilteredElementCollector(doc)
            .OfClass(DB.Viewport)
            .ToElements()
        )
    except:
        viewports = []

    for viewport in viewports:
        try:
            placed_view = doc.GetElement(viewport.ViewId)

            if not placed_view:
                continue

            if placed_view.IsTemplate:
                continue

            if placed_view.ViewType != DB.ViewType.DraftingView:
                continue

            view_id = get_element_id_key(placed_view)

            if view_id in seen_view_ids:
                continue

            seen_view_ids.add(view_id)

            if text_exists_in_view(placed_view, target_text):
                return True

        except:
            pass

    return False
# ------------------------------------------------------------
# Room / equipment overlap logic
# ------------------------------------------------------------

def z_ranges_overlap(bb1, bb2, tol=Z_TOL):
    if not bb1 or not bb2:
        return False

    return not (
        bb1.Max.Z < bb2.Min.Z - tol or
        bb1.Min.Z > bb2.Max.Z + tol
    )


def bbox_xy_sample_points(bb):
    minx = bb.Min.X
    maxx = bb.Max.X
    miny = bb.Min.Y
    maxy = bb.Max.Y

    midx = (minx + maxx) / 2.0
    midy = (miny + maxy) / 2.0

    pts = [
        DB.XYZ(midx, midy, 0),

        DB.XYZ(minx, miny, 0),
        DB.XYZ(minx, maxy, 0),
        DB.XYZ(maxx, miny, 0),
        DB.XYZ(maxx, maxy, 0),

        DB.XYZ(midx, miny, 0),
        DB.XYZ(midx, maxy, 0),
        DB.XYZ(minx, midy, 0),
        DB.XYZ(maxx, midy, 0),
    ]

    return pts


def equipment_overlaps_room_by_points(equip, room):
    equip_bb = get_bbox(equip)
    room_bb = get_bbox(room)

    if not equip_bb or not room_bb:
        return False

    if not z_ranges_overlap(equip_bb, room_bb):
        return False

    room_mid_z = (room_bb.Min.Z + room_bb.Max.Z) / 2.0

    for xy_pt in bbox_xy_sample_points(equip_bb):
        test_pt = DB.XYZ(xy_pt.X, xy_pt.Y, room_mid_z)

        try:
            if room.IsPointInRoom(test_pt):
                return True
        except:
            pass

    return False


def get_location_point_xy(elem):
    try:
        loc = elem.Location

        if hasattr(loc, "Point"):
            pt = loc.Point
            if pt:
                return pt

        if hasattr(loc, "Curve"):
            curve = loc.Curve
            if curve:
                return curve.Evaluate(0.5, True)
    except:
        pass

    return None


def equipment_inside_room_strict(equip, room):
    equip_bb = get_bbox(equip)
    room_bb = get_bbox(room)

    if not equip_bb or not room_bb:
        return False

    if not z_ranges_overlap(equip_bb, room_bb):
        return False

    room_mid_z = (room_bb.Min.Z + room_bb.Max.Z) / 2.0

    loc_pt = get_location_point_xy(equip)

    if loc_pt:
        try:
            test_pt = DB.XYZ(loc_pt.X, loc_pt.Y, room_mid_z)
            if room.IsPointInRoom(test_pt):
                return True
        except:
            pass

    try:
        center_x = (equip_bb.Min.X + equip_bb.Max.X) / 2.0
        center_y = (equip_bb.Min.Y + equip_bb.Max.Y) / 2.0
        center_pt = DB.XYZ(center_x, center_y, room_mid_z)

        if room.IsPointInRoom(center_pt):
            return True
    except:
        pass

    return False


def bbox_xy_intersects(bb1, bb2):
    if not bb1 or not bb2:
        return False

    if bb1.Max.X < bb2.Min.X or bb1.Min.X > bb2.Max.X:
        return False

    if bb1.Max.Y < bb2.Min.Y or bb1.Min.Y > bb2.Max.Y:
        return False

    return True


def base_cabinet_is_below_sink_within_10ft(sink, cabinet):
    sink_bb = get_bbox(sink)
    cab_bb = get_bbox(cabinet)

    if not sink_bb or not cab_bb:
        return False

    if not bbox_xy_intersects(sink_bb, cab_bb):
        return False

    if cab_bb.Min.Z > sink_bb.Max.Z + Z_TOL:
        return False

    if cab_bb.Max.Z < sink_bb.Min.Z - SINK_CABINET_CHECK_DEPTH - Z_TOL:
        return False

    return True


# ------------------------------------------------------------
# Collect rooms
# ------------------------------------------------------------

rooms = list(
    DB.FilteredElementCollector(doc)
    .OfCategory(DB.BuiltInCategory.OST_Rooms)
    .WhereElementIsNotElementType()
    .ToElements()
)

rooms = [r for r in rooms if r.Area > 0]

room_by_id = {}

for room in rooms:
    room_by_id[get_room_id_key(room)] = room


# ------------------------------------------------------------
# Collect Specialty Equipment
# ------------------------------------------------------------

specialty_equipment = list(
    DB.FilteredElementCollector(doc)
    .OfCategory(DB.BuiltInCategory.OST_SpecialityEquipment)
    .WhereElementIsNotElementType()
    .ToElements()
)


# ------------------------------------------------------------
# Collect family instances
# ------------------------------------------------------------

water_closet_urinal_families = []
lavatory_families = []
accessory_families = []
locker_families = []
breakroom_appliance_families = []
revision_families = []

all_family_instances = list(
    DB.FilteredElementCollector(doc)
    .OfClass(DB.FamilyInstance)
    .WhereElementIsNotElementType()
    .ToElements()
)

for inst in all_family_instances:
    if not is_model_family_instance(inst):
        continue

    if is_demolished_element(inst):
        continue

    if is_toilet_or_urinal_plumbing_fixture(inst):
        water_closet_urinal_families.append(inst)

    if is_lavatory(inst):
        lavatory_families.append(inst)

    if is_restroom_accessory(inst):
        accessory_families.append(inst)

    if is_locker(inst):
        locker_families.append(inst)

    if is_breakroom_appliance(inst):
        breakroom_appliance_families.append(inst)

    if is_revision_family(inst):
        revision_families.append(inst)


# ------------------------------------------------------------
# Combine lists and remove duplicates
# ------------------------------------------------------------

equipment = []
seen_ids = set()

for inst in specialty_equipment + water_closet_urinal_families + lavatory_families + accessory_families + locker_families + breakroom_appliance_families + revision_families:
    if is_demolished_element(inst):
        continue

    eid = get_element_id_key(inst)

    if eid not in seen_ids:
        equipment.append(inst)
        seen_ids.add(eid)


# ------------------------------------------------------------
# Count by room
# ------------------------------------------------------------

room_detected_family_ids = defaultdict(set)
room_non_existing_detected_family_ids = defaultdict(set)

room_water_closet_count = defaultdict(int)
room_urinal_count = defaultdict(int)
room_lavatory_count = defaultdict(int)
room_wc_urinal_total = defaultdict(int)

room_seat_cover_count = defaultdict(int)
room_soap_dispenser_count = defaultdict(int)
room_mirror_36_count = defaultdict(int)
room_mirror_60_count = defaultdict(int)
room_toilet_tissue_dispenser_count = defaultdict(int)
room_napkin_disposal_count = defaultdict(int)
room_napkin_dispenser_count = defaultdict(int)
room_paper_towel_dispenser_count = defaultdict(int)

room_refrigerator_count = defaultdict(int)
room_microwave_count = defaultdict(int)
room_ice_machine_count = defaultdict(int)

room_single_tier_locker_count = defaultdict(int)
room_double_tier_instance_count = defaultdict(int)
room_double_tier_actual_locker_count = defaultdict(int)
room_total_actual_locker_count = defaultdict(int)

for equip in equipment:
    if is_demolished_element(equip):
        continue

    for room in rooms:
        if equipment_overlaps_room_by_points(equip, room):
            rid = get_room_id_key(room)
            eid = get_element_id_key(equip)

            room_detected_family_ids[rid].add(eid)

            if not is_existing_phase_element(equip):
                room_non_existing_detected_family_ids[rid].add(eid)

            restroom_fixture_type = get_restroom_fixture_type(equip)

            if restroom_fixture_type:
                room_wc_urinal_total[rid] += 1

                if restroom_fixture_type == "water_closet":
                    room_water_closet_count[rid] += 1
                elif restroom_fixture_type == "urinal":
                    room_urinal_count[rid] += 1

            if is_lavatory(equip):
                room_lavatory_count[rid] += 1

            if is_toilet_tissue_dispenser(equip):
                room_toilet_tissue_dispenser_count[rid] += 1

            if is_paper_towel_dispenser(equip):
                room_paper_towel_dispenser_count[rid] += 1

            if is_napkin_dispenser(equip):
                room_napkin_dispenser_count[rid] += 1

            if is_napkin_disposal(equip):
                room_napkin_disposal_count[rid] += 1

            if is_soap_dispenser(equip):
                room_soap_dispenser_count[rid] += 1

            if is_mirror_36(equip):
                room_mirror_36_count[rid] += 1

            if is_mirror_60(equip):
                room_mirror_60_count[rid] += 1

            if is_seat_cover(equip):
                room_seat_cover_count[rid] += 1

            if is_breakroom(room):
                if is_refrigerator(equip):
                    room_refrigerator_count[rid] += 1

                if is_microwave(equip):
                    room_microwave_count[rid] += 1

                if is_ice_machine(equip):
                    room_ice_machine_count[rid] += 1

            if is_locker_room(room) and is_locker(equip):
                locker_tier = get_locker_tier(equip)

                if locker_tier == "double":
                    room_double_tier_instance_count[rid] += 1
                    room_double_tier_actual_locker_count[rid] += 2
                    room_total_actual_locker_count[rid] += 2

                elif locker_tier == "single":
                    room_single_tier_locker_count[rid] += 1
                    room_total_actual_locker_count[rid] += 1


# ------------------------------------------------------------
# Revision-specific strict findings
# ------------------------------------------------------------

surface_waste_findings = []
bmeu_swing_gate_findings = []
sink_cabinet_included = False

bmeu_rooms = []
breakroom_rooms_for_revision = []

for room in rooms:
    if is_bmeu_room(room):
        bmeu_rooms.append(room)

    if is_breakroom(room):
        breakroom_rooms_for_revision.append(room)

for inst in all_family_instances:
    if not is_model_family_instance(inst):
        continue

    if is_demolished_element(inst):
        continue

    if is_surface_mounted_waste_disposal(inst):
        for room in rooms:
            rid = get_room_id_key(room)

            if room_wc_urinal_total.get(rid, 0) <= 0:
                continue

            if equipment_inside_room_strict(inst, room):
                surface_waste_findings.append(
                    [
                        get_family_type_name(inst),
                        get_room_display_name(room)
                    ]
                )

    if is_swing_gate(inst):
        for room in bmeu_rooms:
            if equipment_inside_room_strict(inst, room):
                bmeu_swing_gate_findings.append(
                    [
                        get_family_type_name(inst),
                        get_room_display_name(room)
                    ]
                )

breakroom_sinks = []
breakroom_base_cabinets = []

for inst in all_family_instances:
    if not is_model_family_instance(inst):
        continue

    if is_demolished_element(inst):
        continue

    if is_breakroom_sink(inst):
        for room in breakroom_rooms_for_revision:
            if equipment_inside_room_strict(inst, room):
                breakroom_sinks.append([inst, room])

    if is_base_cabinet(inst):
        for room in breakroom_rooms_for_revision:
            if equipment_inside_room_strict(inst, room):
                breakroom_base_cabinets.append([inst, room])

for sink_pair in breakroom_sinks:
    sink_inst = sink_pair[0]
    sink_room = sink_pair[1]
    sink_room_id = get_room_id_key(sink_room)

    for cabinet_pair in breakroom_base_cabinets:
        cabinet_inst = cabinet_pair[0]
        cabinet_room = cabinet_pair[1]
        cabinet_room_id = get_room_id_key(cabinet_room)

        if cabinet_room_id != sink_room_id:
            continue

        if base_cabinet_is_below_sink_within_10ft(sink_inst, cabinet_inst):
            sink_cabinet_included = True
            break

    if sink_cabinet_included:
        break


# ------------------------------------------------------------
# Sort helper
# ------------------------------------------------------------

def room_sort_key(room):
    return "{} {}".format(get_room_number(room), get_room_name(room))


# ------------------------------------------------------------
# Print Report Header
# ------------------------------------------------------------

print_report_header()
def print_report_footer():
    generated_time = datetime.now().strftime("%Y-%m-%d %I:%M %p")

    output.print_md("---")
    output.print_html(
        "<div style='font-size:11px; color:#888888; margin-top:8px;'>"
        "Report generated by USPS QC Checker V9 | {}"
        "</div>".format(html_escape(generated_time))
    )


# ------------------------------------------------------------
# Print Revision Check
# ------------------------------------------------------------

output.print_md("## Revision Check")

ad100_sheet = find_sheet_by_number(AD100_SHEET_NUMBER)

if ad100_sheet:
    if generic_annotation_family_exists_on_sheet(ad100_sheet, LOG_GENERAL_NOTES_TEXT):
        print_revision_ok_line("LOG GENERAL NOTES detected on AD100")
    else:
        print_revision_red_line("AD100 sheet does not appear to include LOG GENERAL NOTES, please check.")
else:
    print_revision_red_line("AD100 sheet was not found, please check.")

if surface_waste_findings:
    for finding in surface_waste_findings:
        family_name = finding[0]
        room_display = finding[1]

        print_revision_red_line(
            "Surface mounted trash bin detected: {} in {}. Please note: for projects after May 8, 2026, trash bins need to be freestanding.".format(
                family_name,
                room_display
            )
        )
else:
    print_revision_ok_line("No surface mounted trash bin detected")


if bmeu_rooms:
    bmeu_drafting_view = find_drafting_view_by_name(BMEU_DRAFTING_VIEW_NAME)

    if bmeu_drafting_view:
        print_revision_ok_line("TYPICAL BMEU 732 MOUNTING HEIGHT drafting view detected")
    else:
        print_revision_red_line("Drafting View TYPICAL BMEU 732 MOUNTING HEIGHT was not found. If BMEU 732 is being used in the project, please add this detail.")

if text_exists_in_any_sheet_drafting_view(WIRE_SCREEN_GAP_TEXT):
    print_revision_ok_line('1/2" MAX. ALLOWABLE GAP detected in drafting view on sheet')
else:
    print_revision_red_line('1/2" MAX. ALLOWABLE GAP was not detected in any drafting view placed on sheets, please check.')

materials_schedule = find_schedule_by_name(MATERIALS_FINISH_LEGEND_SCHEDULE_NAME)

if materials_schedule:
    if schedule_contains_text(materials_schedule, MATERIALS_WALL_BASE_TEXT):
        print_revision_ok_line('MATERIALS FINISH LEGEND includes WALL BASE 6"')
    else:
        print_revision_red_line('MATERIALS FINISH LEGEND does not appear to include WALL BASE 6", please check.')

    if schedule_contains_text(materials_schedule, MATERIALS_RUNNING_BOND_TEXT):
        print_revision_ok_line("MATERIALS FINISH LEGEND includes 1/3 OFFSET RUNNING BOND")
    else:
        print_revision_red_line("MATERIALS FINISH LEGEND does not appear to include 1/3 offset running bond, please check.")

    if schedule_contains_text(materials_schedule, MATERIALS_GRAYSON_TEXT):
        print_revision_ok_line("RFT-1 IN MATERIALS FINISH LEGEND COLOR UPDATED TO GRAYSON")
    else:
        print_revision_red_line("RFT-1 IN MATERIALS FINISH LEGEND does not appear to be updated to GRAYSON, please check.")
else:
    print_revision_red_line("MATERIALS FINISH LEGEND schedule was not found, please check.")

if bmeu_rooms:
    if bmeu_swing_gate_findings:
        for finding in bmeu_swing_gate_findings:
            family_name = finding[0]
            room_display = finding[1]

            print_revision_red_line(
                "Swing gate detected: {} in {}. Please note: beginning June 8, 2026, swing gate should no longer be ordered for BMEU.".format(
                    family_name,
                    room_display
                )
            )
    else:
        print_revision_ok_line("No swing gate detected inside BMEU room")

if breakroom_rooms_for_revision:
    if sink_cabinet_included:
        print_revision_ok_line("Breakroom sink cabinet doors included (model check)")
    else:
        print_revision_red_line("NO CABINET FAMILY UNDER BREAKROOM SINK WAS DETECTED, please note: beginning January 14, 2026, sink cabinet needs to have doors.")

sink_base_view = find_drafting_view_by_name(SINK_BASE_DRAFTING_VIEW_NAME)

if sink_base_view:
    if text_exists_in_view(sink_base_view, SINK_BASE_SWING_WITH_DOORS_TEXT):
        print_revision_ok_line("SINK BASE CABINET DETAIL UPDATED TO INCLUDE DOORS AND OPERABLE TOE PANELS FOR FRONT APPROACH")
    else:
        print_revision_red_line("SINK BASE CABINET DETAIL has not been updated to the latest version. Doors and toe panel swing with doors are required, please check.")
else:
    print_revision_red_line("SECTION THROUGH TYP. SINK BASE drafting view was not found. Sink base cabinet detail may not be updated to the latest version, please check.")


# ------------------------------------------------------------
# Print Restroom Checklist
# ------------------------------------------------------------

output.print_md("---")
output.print_md("## Restroom Checklist")

restroom_room_ids = sorted(
    [
        rid for rid in room_wc_urinal_total.keys()
        if room_has_any_non_existing_detected_family(
            rid,
            room_detected_family_ids,
            room_non_existing_detected_family_ids
        )
    ],
    key=lambda rid: room_sort_key(room_by_id[rid])
)

if not restroom_room_ids:
    output.print_md("No in-scope rooms with Plumbing Fixture toilet / urinal families found.")
else:
    for rid in restroom_room_ids:
        room = room_by_id[rid]

        water_closet_count = room_water_closet_count.get(rid, 0)
        urinal_count = room_urinal_count.get(rid, 0)
        lavatory_count = room_lavatory_count.get(rid, 0)
        wc_urinal_total = room_wc_urinal_total.get(rid, 0)

        toilet_tissue_count = room_toilet_tissue_dispenser_count.get(rid, 0)
        paper_towel_count = room_paper_towel_dispenser_count.get(rid, 0)
        napkin_dispenser_count = room_napkin_dispenser_count.get(rid, 0)
        napkin_disposal_count = room_napkin_disposal_count.get(rid, 0)
        soap_dispenser_count = room_soap_dispenser_count.get(rid, 0)
        mirror_36_count = room_mirror_36_count.get(rid, 0)
        mirror_60_count = room_mirror_60_count.get(rid, 0)
        seat_cover_count = room_seat_cover_count.get(rid, 0)

        print_room_divider()
        output.print_md("### {}".format(get_room_display_name(room)))

        output.print_md("#### Plumbing Fixture")
        print_plain_count_line(LABEL_WATER_CLOSET, water_closet_count)
        print_plain_count_line(LABEL_URINAL, urinal_count)
        print_plain_count_line(LABEL_LAVATORY, lavatory_count)
        print_plain_count_line("WATER CLOSET + URINAL TOTAL", wc_urinal_total)

        if wc_urinal_total >= 6:
            print_orange_message("Ambulatory stall is required, please check.")

        output.print_md("#### Accessories")

        validate_equal_to_required(
            LABEL_RA1_TOILET_TISSUE,
            toilet_tissue_count,
            water_closet_count
        )

        paper_towel_required = int(math.ceil(lavatory_count / 2.0))

        validate_minimum_required(
            LABEL_RA3_PAPER_TOWEL,
            paper_towel_count,
            paper_towel_required
        )

        if water_closet_count > 1 and urinal_count == 0:
            validate_minimum_required(
                LABEL_RA5_NAPKIN_DISPENSER,
                napkin_dispenser_count,
                1
            )
        else:
            print_plain_count_line(LABEL_RA5_NAPKIN_DISPENSER, napkin_dispenser_count)

        if urinal_count > 0:
            if napkin_disposal_count == 0:
                print_green_check_line(LABEL_RA6_NAPKIN_DISPOSAL, napkin_disposal_count)
            else:
                print_plain_count_line(LABEL_RA6_NAPKIN_DISPOSAL, napkin_disposal_count)
                print_red_message("{} IS NOT NEEDED.".format(get_message_label(LABEL_RA6_NAPKIN_DISPOSAL)))
        else:
            validate_equal_to_required(
                LABEL_RA6_NAPKIN_DISPOSAL,
                napkin_disposal_count,
                water_closet_count
            )

        validate_equal_to_required(
            LABEL_RA7_SOAP,
            soap_dispenser_count,
            lavatory_count
        )

        validate_equal_to_required(
            LABEL_RA9_MIRROR_36,
            mirror_36_count,
            lavatory_count
        )

        if wc_urinal_total > 2:
            validate_minimum_required(
                LABEL_RA10_MIRROR_60,
                mirror_60_count,
                1
            )
        else:
            print_plain_count_line(LABEL_RA10_MIRROR_60, mirror_60_count)

        validate_equal_to_required(
            LABEL_RA11_SEAT_COVER,
            seat_cover_count,
            water_closet_count
        )


# ------------------------------------------------------------
# Print Breakroom Checklist
# ------------------------------------------------------------

output.print_md("---")
output.print_md("## Breakroom Checklist")

breakroom_rooms = []

for room in rooms:
    if is_breakroom(room):
        breakroom_rooms.append(room)

breakroom_rooms = sorted(breakroom_rooms, key=room_sort_key)

if not breakroom_rooms:
    output.print_md("No breakrooms found.")
else:
    for room in breakroom_rooms:
        rid = get_room_id_key(room)

        refrigerator_count = room_refrigerator_count.get(rid, 0)
        microwave_count = room_microwave_count.get(rid, 0)
        ice_machine_count = room_ice_machine_count.get(rid, 0)

        print_room_divider()
        output.print_md("### {}".format(get_room_display_name(room)))

        print_room_area_check_line(room)

        validate_minimum_required(
            LABEL_A1_REFRIGERATOR,
            refrigerator_count,
            2
        )

        validate_minimum_required(
            LABEL_A2_MICROWAVE,
            microwave_count,
            2
        )

        validate_minimum_required(
            LABEL_A3_ICE_MACHINE,
            ice_machine_count,
            2
        )


# ------------------------------------------------------------
# Print Locker Room Checklist
# ------------------------------------------------------------

output.print_md("---")
output.print_md("## Locker Room Checklist")

locker_rooms = []

for room in rooms:
    rid = get_room_id_key(room)

    if (
        is_locker_room(room) and
        room_has_any_non_existing_detected_family(
            rid,
            room_detected_family_ids,
            room_non_existing_detected_family_ids
        )
    ):
        locker_rooms.append(room)

locker_rooms = sorted(locker_rooms, key=room_sort_key)

if not locker_rooms:
    output.print_md("No in-scope locker rooms found.")
else:
    for room in locker_rooms:
        rid = get_room_id_key(room)

        single_tier_count = room_single_tier_locker_count.get(rid, 0)
        double_tier_instance_count = room_double_tier_instance_count.get(rid, 0)
        double_tier_actual_count = room_double_tier_actual_locker_count.get(rid, 0)
        total_actual_count = room_total_actual_locker_count.get(rid, 0)

        print_room_divider()
        output.print_md("### {}".format(get_room_display_name(room)))

        output.print_md("#### Locker Count")
        print_plain_count_line("Single tier locker", single_tier_count)
        print_plain_count_line("Double tier family instance", double_tier_instance_count)
        print_plain_count_line("Double tier actual locker", double_tier_actual_count)
        print_plain_count_line("Total actual locker", total_actual_count)

        accessible_required = int(math.ceil(total_actual_count * 0.05))

        if accessible_required > 0:
            if accessible_required == 1:
                print_orange_message("1 accessible locker is required, please check.")
            else:
                print_orange_message("{} accessible lockers are required, please check.".format(accessible_required))

print_report_footer()