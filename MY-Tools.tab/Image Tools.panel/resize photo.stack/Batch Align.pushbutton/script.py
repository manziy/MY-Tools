# -*- coding: utf-8 -*-
# Quiet auto-pair & move:
# - No popups or prints on success.
# - Active view only (e.g., Legend).
# - Frames: Generic Annotation (OST_GenericAnnotation).
# - Photos: ImageInstance + ImportInstance ending with ".pdf".
# - Target = frame origin + ( +14'-0" X, +15'-6" Y ).

from Autodesk.Revit.DB import *

uidoc = __revit__.ActiveUIDocument
doc   = uidoc.Document
view  = doc.ActiveView

# ---- Offsets (feet) ----
OFFSET_RIGHT_FT = 14.0
OFFSET_UP_FT    = 15.0 + 6.0/12.0   # 15'-6"
EPS = 1e-9
# ------------------------

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

# Collect frames & photos in active view
frames = [e for e in FilteredElementCollector(doc, view.Id).OfClass(FamilyInstance) if _is_frame(e)]
if not frames:
    # Quiet exit
    import sys; sys.exit()

photos_img = list(FilteredElementCollector(doc, view.Id).OfClass(ImageInstance))
photos_pdf = [e for e in FilteredElementCollector(doc, view.Id).OfClass(ImportInstance) if _is_pdf_import(e)]
photos = photos_img + photos_pdf
if not photos:
    # Quiet exit
    import sys; sys.exit()

# Compute frame targets (origin + offset); drop frames without a usable origin
frame_targets = []
for f in frames:
    o = _frame_origin(f, view)
    if o:
        frame_targets.append((f, o + XYZ(OFFSET_RIGHT_FT, OFFSET_UP_FT, 0.0)))
if not frame_targets:
    import sys; sys.exit()

frames, targets = zip(*frame_targets)
frames  = list(frames)
targets = list(targets)

# Compute photo centers; drop those without bbox
photo_centers = []
for p in photos:
    c = _bbox_center(p, view)
    if c:
        photo_centers.append((p, c))
if not photo_centers:
    import sys; sys.exit()

photos, centers = zip(*photo_centers)
photos  = list(photos)
centers = list(centers)

# Build all pair distances (XY only), sort ascending
pairs = []
for i, tp in enumerate(targets):
    for j, cp in enumerate(centers):
        dx = tp.X - cp.X
        dy = tp.Y - cp.Y
        d2 = dx*dx + dy*dy
        pairs.append((d2, i, j))
pairs.sort(key=lambda t: t[0])

# Greedy one-to-one assignment
assigned_f = set()
assigned_p = set()
assignments = []
max_pairs = min(len(frames), len(photos))
for d2, i, j in pairs:
    if i in assigned_f or j in assigned_p:
        continue
    assignments.append((i, j))
    assigned_f.add(i)
    assigned_p.add(j)
    if len(assignments) >= max_pairs:
        break

if not assignments:
    import sys; sys.exit()

# Move photos (quiet)
t = Transaction(doc, "Auto-pair frames/photos and move")
t.Start()
for i, j in assignments:
    target_pt = targets[i]
    photo     = photos[j]
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
t.Commit()
