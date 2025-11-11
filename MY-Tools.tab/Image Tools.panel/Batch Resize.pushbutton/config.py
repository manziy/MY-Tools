# Config UI for "Auto-resize images by orientation"
# Shift+Click runs this script automatically and shows the black dot on the button.
from Autodesk.Revit.DB import *
import re

try:
    from pyrevit import forms, script
    def _alert(msg, title="Auto-resize (Config)"):
        forms.alert(msg, title=title, warn_icon=False)
except Exception:
    script = None
    def _alert(msg, title="Auto-resize (Config)"):
        print("[{}] {}".format(title, msg))

uidoc = __revit__.ActiveUIDocument
doc   = uidoc.Document
view  = doc.ActiveView

# Defaults (only used if no saved config yet)
DEF_LANDSCAPE_WIDTH_FT  = 28.0
DEF_PORTRAIT_HEIGHT_FT  = 25.0
DEF_SHEET_LANDSCAPE_IN  = 28.0
DEF_SHEET_PORTRAIT_IN   = 25.0

# Load config (per-script)
def _cfg():
    if script: return script.get_config()
    class _Dummy: pass
    return _Dummy()
cfg = _cfg()

# ensure keys
if not hasattr(cfg, 'landscape_ft'):       cfg.landscape_ft       = DEF_LANDSCAPE_WIDTH_FT
if not hasattr(cfg, 'portrait_ft'):        cfg.portrait_ft        = DEF_PORTRAIT_HEIGHT_FT
if not hasattr(cfg, 'sheet_landscape_in'): cfg.sheet_landscape_in = DEF_SHEET_LANDSCAPE_IN
if not hasattr(cfg, 'sheet_portrait_in'):  cfg.sheet_portrait_in  = DEF_SHEET_PORTRAIT_IN

def _save():
    try:
        if script: script.save_config()
    except Exception as e:
        _alert("Could not save settings:\n{}".format(e))

# Parse "28'-6\"", 28, 28.5, 30", 30 in, 2.5', etc.
def _parse_len(s, unit='ft'):
    if s is None: return None
    t = s.strip().lower()
    if not t: return None
    m = re.match(r"^\s*(?P<ft>-?\d+(?:\.\d+)?)\s*'\s*(?P<in>\d*(?:\.\d+)?)\s*\"?\s*$", t)
    if m:
        ft  = float(m.group('ft') or 0)
        ins = float(m.group('in') or 0)
        return ft + ins/12.0 if unit == 'ft' else ft*12.0 + ins
    m = re.match(r"^\s*(?P<in>-?\d+(?:\.\d+)?)\s*(?:\"|in|inch|inches)\s*$", t)
    if m:
        v = float(m.group('in'))
        return v if unit == 'in' else v/12.0
    m = re.match(r"^\s*(?P<ft>-?\d+(?:\.\d+)?)\s*(?:'|ft|feet|foot)\s*$", t)
    if m:
        v = float(m.group('ft'))
        return v if unit == 'ft' else v*12.0
    try:
        return float(t)
    except:
        return None

# Configure values for the CURRENT view context (sheet vs non-sheet)
is_sheet = isinstance(view, ViewSheet)
unit = 'in' if is_sheet else 'ft'
cur_w = cfg.sheet_landscape_in if is_sheet else cfg.landscape_ft
cur_h = cfg.sheet_portrait_in  if is_sheet else cfg.portrait_ft

try:
    w_str = forms.ask_for_string(
        prompt="Landscape width ({}). Current: {}".format(unit, cur_w),
        default=str(cur_w),
        title="Auto-resize ({} settings)".format("Sheet" if is_sheet else "Non-sheet")
    )
    h_str = forms.ask_for_string(
        prompt="Portrait height ({}). Current: {}".format(unit, cur_h),
        default=str(cur_h),
        title="Auto-resize ({} settings)".format("Sheet" if is_sheet else "Non-sheet")
    )
except Exception:
    _alert("Config UI unavailable.")
    raise

new_w = _parse_len(w_str, unit)
new_h = _parse_len(h_str, unit)

if new_w is None or new_w <= 0:
    _alert("Invalid landscape width. Keeping previous value: {}".format(cur_w))
    new_w = cur_w
if new_h is None or new_h <= 0:
    _alert("Invalid portrait height. Keeping previous value: {}".format(cur_h))
    new_h = cur_h

if is_sheet:
    cfg.sheet_landscape_in = new_w
    cfg.sheet_portrait_in  = new_h
else:
    cfg.landscape_ft = new_w
    cfg.portrait_ft  = new_h

_save()
