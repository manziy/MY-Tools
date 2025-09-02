# Auto-resize images by orientation in ACTIVE VIEW
# - Landscape: width = 28'-0"  (non-sheets) / 28" (sheets)   [user-configurable via Shift+Click]
# - Portrait : height = 25'-0" (non-sheets) / 25" (sheets)   [user-configurable via Shift+Click]
# Keeps aspect ratio; does not rotate.
#
# Supports:
#   - ImageInstance (PNG/JPG/etc): bbox-based relative scaling with LockProportions=True
#   - ImportInstance (PDF): uniform ScaleElement about center
#
# Shift+Click the button to edit and save the targets for the current view type.
#
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

# -------- Defaults (used if no saved config) --------
DEF_LANDSCAPE_WIDTH_FT  = 28.0   # 28'-0" on non-sheets
DEF_PORTRAIT_HEIGHT_FT  = 25.0   # 25'-0" on non-sheets
DEF_SHEET_LANDSCAPE_IN  = 28.0   # 28" on sheets
DEF_SHEET_PORTRAIT_IN   = 25.0   # 25" on sheets
# ----------------------------------------------------

# Load/save per-script config
def _get_cfg():
    if script:
        return script.get_config()
    class _Dummy(object): pass
    return _Dummy()
_cfg = _get_cfg()

# Ensure keys exist
if not hasattr(_cfg, 'landscape_ft'):      _cfg.landscape_ft      = DEF_LANDSCAPE_WIDTH_FT
if not hasattr(_cfg, 'portrait_ft'):       _cfg.portrait_ft       = DEF_PORTRAIT_HEIGHT_FT
if not hasattr(_cfg, 'sheet_landscape_in'): _cfg.sheet_landscape_in = DEF_SHEET_LANDSCAPE_IN
if not hasattr(_cfg, 'sheet_portrait_in'):  _cfg.sheet_portrait_in  = DEF_SHEET_PORTRAIT_IN

def _save_cfg():
    try:
        if script:
            script.save_config()
    except Exception as e:
        _alert("Could not save settings:\n{}".format(e))

# Length parser: supports 28, 28.5, 28'-6", 28' 6", 30", 30 in, 2.5'
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

# Prompt settings when Shift is held (per current view context)
def _maybe_prompt_settings():
    global _cfg
    is_sheet = isinstance(view, ViewSheet)
    unit = 'in' if is_sheet else 'ft'
    cur_w = _cfg.sheet_landscape_in if is_sheet else _cfg.landscape_ft
    cur_h = _cfg.sheet_portrait_in  if is_sheet else _cfg.portrait_ft

    try:
        # Two simple prompts; user can enter ft-in formats (e.g., 28'-6") or plain numbers
        w_str = forms.ask_for_string(
            prompt="Landscape width ({}). Current: {}".format(unit, cur_w),
            default=str(cur_w)
        )
        h_str = forms.ask_for_string(
            prompt="Portrait height ({}). Current: {}".format(unit, cur_h),
            default=str(cur_h)
        )
    except Exception:
        # Fallback to console input not really ideal inside Revit; if forms not available, skip
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

# If Shift+Click, let user edit + save targets first
if _shift_down:
    _maybe_prompt_settings()

# Compute targets for this view
if isinstance(view, ViewSheet):
    target_w_ft = (_cfg.sheet_landscape_in or DEF_SHEET_LANDSCAPE_IN) / 12.0
    target_h_ft = (_cfg.sheet_portrait_in  or DEF_SHEET_PORTRAIT_IN ) / 12.0
else:
    target_w_ft = (_cfg.landscape_ft or DEF_LANDSCAPE_WIDTH_FT)
    target_h_ft = (_cfg.portrait_ft  or DEF_PORTRAIT_HEIGHT_FT)

# Collect images (PNG/JPG/etc)
imgs = list(FilteredElementCollector(doc, view.Id).OfClass(ImageInstance))

# Collect PDFs (ImportInstance with type/name ending in .pdf)
pdf_imps = []
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
            pdf_imps.append(imp)
    except:
        pass

if not imgs and not pdf_imps:
    _alert("No ImageInstance or PDF ImportInstance found in the active view.")
    raise Exception("Nothing to resize.")

EPS = 1e-9
changed_img = 0
changed_pdf = 0

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

t = Transaction(doc, "Auto-resize images by orientation")
t.Start()

# --- ImageInstance: bbox-based RELATIVE scaling (rotation-aware) ---
for el in imgs:
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

# --- PDF ImportInstance: uniform relative scale by bbox width/height ---
for imp in pdf_imps:
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

t.Commit()

# Silent on success as requested earlier. If nothing changed, give a light heads-up.
if changed_img == 0 and changed_pdf == 0:
    _alert("No images/PDFs needed resizing.\n(Targets: W={:.3f} ft, H={:.3f} ft)".format(target_w_ft, target_h_ft))
