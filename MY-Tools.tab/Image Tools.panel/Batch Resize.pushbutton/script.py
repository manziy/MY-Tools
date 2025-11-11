# -*- coding: utf-8 -*-
# Combined: auto-resize images/PDFs in active view, then auto-pair & move to frames
# - Shift+Click to edit resize targets (same as original resize script)
# - Step 1: resize
# - Step 2: re-collect and move into frames

from Autodesk.Revit.DB import *
import re

# ---------- UI & config helpers ----------
try:
    from pyrevit import forms, script
    def _alert(msg, title="Auto-resize images by orientation"):
        forms.alert(msg, title=title, warn_icon=False)
except Exception:
    script = None
    def _alert(msg, title="Auto-resize images by orientation"):
        print("[{}] {}".format(title, msg))

# Detect Shift modifier
try:
    from System.Windows.Input import Keyboard, ModifierKeys
    _shift_down = (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift
except Exception:
    _shift_down = False

uidoc = __revit__.ActiveUIDocument
doc   = uidoc.Document
view  = doc.ActiveView

# ------------------- RESIZE PART (your first script) -------------------

# Defaults
DEF_LANDSCAPE_WIDTH_FT  = 28.0   # non-sheets
DEF_PORTRAIT_HEIGHT_FT  = 25.0
DEF_SHEET_LANDSCAPE_IN  = 28.0   # sheets
DEF_SHEET_PORTRAIT_IN   = 25.0

def _get_cfg():
    if script:
        return script.get_config()
    class _Dummy(object): pass
    return _Dummy()

_cfg = _get_cfg()
if not hasattr(_cfg, 'landscape_ft'):       _cfg.landscape_ft       = DEF_LANDSCAPE_WIDTH_FT
if not hasattr(_cfg, 'portrait_ft'):        _cfg.portrait_ft        = DEF_PORTRAIT_HEIGHT_FT
if not hasattr(_cfg, 'sheet_landscape_in'): _cfg.sheet_landscape_in = DEF_SHEET_LANDSCAPE_IN
if not hasattr(_cfg, 'sheet_portrait_in'):  _cfg.sheet_portrait_in  = DEF_SHEET_PORTRAIT_IN

def _save_cfg():
    try:
        if script:
            script.save_config()
    except Exception as e:
        _alert("Could not save settings:\n{}".format(e))

def _parse_length(s, default_unit='ft'):
    if s is None:
        return None
    s = s.strip().lower()
    if not s:
        return None
    m = re.match(r"^\s*(?P<ft>-?\d+(?:\.\d+)?)\s*'\s*(?P<in>\d*(?:\.\d+)?)\s*\"?\s*$", s)
    if m:
        ft  = float(m.group('ft') or 0)
        ins = float(m.group('in') or 0)
        return ft + ins/12.0 if default_unit == 'ft' else ft*12.0 + ins
    m = re.match(r"^\s*(?P<in>-?\d+(?:\.\d+)?)\s*(?:\"|in|inch|inches)\s*$", s)
    if m:
        val = float(m.group('in'))
        return val if default_unit == 'in' else val/12.0
    m = re.match(r"^\s*(?P<ft>-?\d+(?:\.\d+)?)\s*(?:'|ft|feet|foot)\s*$", s)
    if m:
        val = float(m.group('ft'))
        return val if default_unit == 'ft' else val*12.0
    try:
        return float(s)
    except:
        return None

def _maybe_prompt_settings():
    global _cfg
    is_sheet = isinstance(view, ViewSheet)
    unit = 'in' if is_sheet else 'ft'
    cur_w = _cfg.sheet_landscape_in if is_sheet else _cfg.landscape_ft
    cur_h = _cfg.sheet_portrait_in  if is_sheet else _cfg.portrait_ft
    try:
        w_str = forms.ask_for_string(
            prompt="Landscape width ({}). Current: {}".format(unit, cur_w),
            default=str(cur_w)
        )
        h_str = forms.ask_for_string(
            prompt="Portrait height ({}). Current: {}".format(unit, cur_h),
            default=str(cur_h)
        )
    except Exception:
        return

    new_w = _parse_length(w_str, unit)
    new_h = _parse_length(h_str, unit)

    if new_w is None or new_w <= 0:
        _alert("Invalid landscape width. Keeping previous value: {}".format(cur_w))
        new_w = cur_w
    if new_h is None or new_h <= 0:
        _alert("Invalid portrait height. Keeping previous value: {}".format(cur_h))
        new_h = cur_h

    if is_sheet:
        _cfg.sheet_landscape_in = new_w
        _cfg.sheet_portrait_in  = new_h
    else:
        _cfg.landscape_ft = new_w
        _cfg.portrait_ft  = new_h
    _save_cfg()

# Let user edit if Shift
if _shift_down:
    _maybe_prompt_settings()

# Compute targets
if isinstance(view, ViewSheet):
    target_w_ft = (_cfg.sheet_landscape_in or DEF_SHEET_LANDSCAPE_IN) / 12.0
    target_h_ft = (_cfg.sheet_portrait_in  or DEF_SHEET_PORTRAIT_IN ) / 12.0
else:
    target_w_ft = (_cfg.landscape_ft or DEF_LANDSCAPE_WIDTH_FT)
    target_h_ft = (_cfg.portrait_ft  or DEF_PORTRAIT_HEIGHT_FT)

# Collect images/PDFs for resize
imgs_for_resize = list(FilteredElementCollector(doc, view.Id).OfClass(ImageInstance))

pdfs_for_resize = []
for imp in FilteredElementCollector(doc, view.Id).OfClass(ImportInstance):
    try:
        typ = doc.GetElement(imp.GetTypeId())
        nm = None
        p = typ.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME) if typ else None
        if p and p.HasValue:
            nm = p.AsString()
        if not nm and typ:
            nm = getattr(typ, "Name", None)
        if not nm:
            nm = getattr(imp, "Name", "")
        if nm and nm.lower().endswith(".pdf"):
            pdfs_for_resize.append(imp)
    except:
        pass

EPS = 1e-9

def _bbox_wh(el, v):
    bb = el.get_BoundingBox(v)
    if not bb:
        return None, None, None
    w = max(0.0, bb.Max.X - bb.Min.X)
    h = max(0.0, bb.Max.Y - bb.Min.Y)
    c = XYZ((bb.Min.X + bb.Max.X) * 0.5,
            (bb.Min.Y + bb.Max.Y) * 0.5,
            (bb.Min.Z + bb.Max.Z) * 0.5)
    return w, h, c

changed_img = 0
changed_pdf = 0

# --- Transaction 1: RESIZE ---
if imgs_for_resize or pdfs_for_resize:
    t1 = Transaction(doc, "Auto-resize images by orientation")
    t1.Start()

    # ImageInstance resize
    for el in imgs_for_resize:
        try:
            w_bb, h_bb, _ = _bbox_wh(el, view)
            if not w_bb or not h_bb or (w_bb <= EPS and h_bb <= EPS):
                continue
            portrait = h_bb > w_bb
            s_rel = (target_h_ft / h_bb) if portrait else (target_w_ft / w_bb)
            if s_rel > 0 and abs(s_rel - 1.0) > 1e-6:
                was_pinned = getattr(el, "Pinned", False)
                if was_pinned: el.Pinned = False
                if hasattr(el, "LockProportions"):
                    el.LockProportions = True
                cur_ws = el.WidthScale or 1.0
                new_ws = cur_ws * s_rel
                if new_ws > EPS:
                    el.WidthScale = new_ws
                    changed_img += 1
                if was_pinned: el.Pinned = True
        except Exception as e:
            print("ImageInstance {} skipped: {}".format(el.Id, e))

    # PDF resize
    for imp in pdfs_for_resize:
        try:
            w_bb, h_bb, center = _bbox_wh(imp, view)
            if not w_bb or not h_bb or (w_bb <= EPS and h_bb <= EPS):
                continue
            portrait = h_bb > w_bb
            s_rel = (target_h_ft / h_bb) if portrait else (target_w_ft / w_bb)
            if s_rel > 0 and abs(s_rel - 1.0) > 1e-6:
                was_pinned = getattr(imp, "Pinned", False)
                if was_pinned: imp.Pinned = False
                ElementTransformUtils.ScaleElement(doc, imp.Id, center, s_rel)
                if was_pinned: imp.Pinned = True
                changed_pdf += 1
        except Exception as e:
            print("PDF {} skipped: {}".format(imp.Id, e))

    t1.Commit()
else:
    # no images at all; we can still try to move, but likely nothing to move
    pass

# ------------------- MOVE PART (your second script) -------------------

# After resize, re-collect frames and photos so bbox is up-to-date
OFFSET_RIGHT_FT = 14.0
OFFSET_UP_FT    = 15.0 + 6.0/12.0   # 15'-6"
EPS = 1e-9

def _is_frame(fi):
    try:
        return isinstance(fi, FamilyInstance) and fi.Category and \
               fi.Category.Id.IntegerValue == int(BuiltInCategory.OST_GenericAnnotation)
    except:
        return False

def _is_pdf_import(el):
    if not isinstance(el, ImportInstance):
        return False
    try:
        typ = doc.GetElement(el.GetTypeId())
        nm = None
        if typ:
            p = typ.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
            if p and p.HasValue:
                nm = p.AsString()
            if not nm:
                nm = getattr(typ, "Name", None)
        if not nm:
            nm = getattr(el, "Name", "")
        return bool(nm) and nm.lower().endswith(".pdf")
    except:
        return False

def _bbox_center(el, v):
    bb = el.get_BoundingBox(v)
    if not bb:
        return None
    return XYZ((bb.Min.X + bb.Max.X) * 0.5,
               (bb.Min.Y + bb.Max.Y) * 0.5,
               (bb.Min.Z + bb.Max.Z) * 0.5)

def _frame_origin(fi, v):
    loc = fi.Location
    if isinstance(loc, LocationPoint):
        return loc.Point
    return _bbox_center(fi, v)

# collect
frames = [e for e in FilteredElementCollector(doc, view.Id).OfClass(FamilyInstance) if _is_frame(e)]
photos_img = list(FilteredElementCollector(doc, view.Id).OfClass(ImageInstance))
photos_pdf = [e for e in FilteredElementCollector(doc, view.Id).OfClass(ImportInstance) if _is_pdf_import(e)]
photos = photos_img + photos_pdf

if frames and photos:
    frame_targets = []
    for f in frames:
        o = _frame_origin(f, view)
        if o:
            frame_targets.append((f, o + XYZ(OFFSET_RIGHT_FT, OFFSET_UP_FT, 0.0)))
    if frame_targets:
        frames_only, targets = zip(*frame_targets)
        frames_only = list(frames_only)
        targets = list(targets)

        photo_centers = []
        for p in photos:
            c = _bbox_center(p, view)
            if c:
                photo_centers.append((p, c))
        if photo_centers:
            photos_only, centers = zip(*photo_centers)
            photos_only = list(photos_only)
            centers = list(centers)

            # build pair distances
            pairs = []
            for i, tp in enumerate(targets):
                for j, cp in enumerate(centers):
                    dx = tp.X - cp.X
                    dy = tp.Y - cp.Y
                    d2 = dx*dx + dy*dy
                    pairs.append((d2, i, j))
            pairs.sort(key=lambda t: t[0])

            assigned_f = set()
            assigned_p = set()
            assignments = []
            max_pairs = min(len(frames_only), len(photos_only))
            for d2, i, j in pairs:
                if i in assigned_f or j in assigned_p:
                    continue
                assignments.append((i, j))
                assigned_f.add(i)
                assigned_p.add(j)
                if len(assignments) >= max_pairs:
                    break

            if assignments:
                t2 = Transaction(doc, "Auto-pair frames/photos and move")
                t2.Start()
                for i, j in assignments:
                    target_pt = targets[i]
                    photo     = photos_only[j]
                    center    = centers[j]
                    delta = target_pt - center
                    if delta.GetLength() <= EPS:
                        continue
                    was_pinned = getattr(photo, "Pinned", False)
                    try:
                        if was_pinned:
                            photo.Pinned = False
                        ElementTransformUtils.MoveElement(doc, photo.Id, delta)
                    finally:
                        if was_pinned:
                            photo.Pinned = True
                t2.Commit()

# Optional small heads-up if literally nothing resized:
if (not imgs_for_resize and not pdfs_for_resize):
    # but we won't pop up if move happened
    pass
