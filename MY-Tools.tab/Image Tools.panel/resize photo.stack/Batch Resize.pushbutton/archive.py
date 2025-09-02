# Auto-resize images by orientation in ACTIVE VIEW
# - Landscape: width = 28'-0"  (non-sheets) / 28" (sheets)
# - Portrait : height = 25'-0" (non-sheets) / 25" (sheets)
# Keeps aspect ratio; does not rotate.
#
# Supports:
#   - ImageInstance (PNG/JPG/etc): uses bbox-based relative scaling with LockProportions=True
#   - ImportInstance (PDF): uniform ScaleElement about center
#
# Run this in the view that contains your images/PDFs.

from Autodesk.Revit.DB import *
try:
    from pyrevit import forms
    def _alert(msg, title="Auto-resize images by orientation"):
        forms.alert(msg, title=title, warn_icon=False)
except Exception:
    def _alert(msg, title="Auto-resize images by orientation"):
        print("[{}] {}".format(title, msg))

uidoc = __revit__.ActiveUIDocument
doc   = uidoc.Document
view  = doc.ActiveView

# -------- Defaults --------
# Non-sheet (Legend/Drafting/Model) => model FEET
LANDSCAPE_WIDTH_FT  = 28.0   # 28'-0"
PORTRAIT_HEIGHT_FT  = 25.0   # 25'-0"

# Sheet views => paper INCHES
SHEET_LANDSCAPE_WIDTH_IN  = 28.0
SHEET_PORTRAIT_HEIGHT_IN  = 25.0
# --------------------------

if isinstance(view, ViewSheet):
    target_w_ft = SHEET_LANDSCAPE_WIDTH_IN / 12.0
    target_h_ft = SHEET_PORTRAIT_HEIGHT_IN / 12.0
else:
    target_w_ft = LANDSCAPE_WIDTH_FT
    target_h_ft = PORTRAIT_HEIGHT_FT

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

        # Orientation comes from on-page bbox (respects current rotation)
        portrait = h_bb > w_bb

        # Desired relative scale so that the bbox matches target dimension
        if portrait:
            s_rel = target_h_ft / h_bb
        else:
            s_rel = target_w_ft / w_bb

        # Avoid tiny nudges
        if s_rel > 0 and abs(s_rel - 1.0) > 1e-6:
            was_pinned = getattr(el, "Pinned", False)
            if was_pinned:
                el.Pinned = False

            # Uniform scaling via WidthScale with LockProportions
            if hasattr(el, "LockProportions"):
                el.LockProportions = True

            cur_ws = el.WidthScale or 1.0
            new_ws = cur_ws * s_rel
            # Guard against invalid or zero
            if new_ws > EPS:
                el.WidthScale = new_ws
                changed_img += 1

            if was_pinned:
                el.Pinned = True
    except Exception as e:
        print("ImageInstance {} skipped: {}".format(el.Id, e))

# --- PDF ImportInstance: uniform relative scale by bbox width/height ---
for imp in pdf_imps:
    try:
        w_bb, h_bb, center = _bbox_wh(imp, view)
        if not w_bb or not h_bb or (w_bb <= EPS and h_bb <= EPS):
            continue

        portrait = h_bb > w_bb
        if portrait:
            s_rel = target_h_ft / h_bb
        else:
            s_rel = target_w_ft / w_bb

        if s_rel > 0 and abs(s_rel - 1.0) > 1e-6:
            was_pinned = getattr(imp, "Pinned", False)
            if was_pinned:
                imp.Pinned = False
            ElementTransformUtils.ScaleElement(doc, imp.Id, center, s_rel)
            if was_pinned:
                imp.Pinned = True
            changed_pdf += 1
    except Exception as e:
        print("PDF {} skipped: {}".format(imp.Id, e))

t.Commit()

_alert("Done.\n"
       "Targets (this view):\n"
       "  - Landscape width = {:.3f} ft\n"
       "  - Portrait height = {:.3f} ft\n"
       "- Resized images: {}\n"
       "- Resized PDFs:   {}\n"
       "(Active view: {})".format(
           target_w_ft, target_h_ft, changed_img, changed_pdf, view.ViewType))
