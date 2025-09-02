# -*- coding: utf-8 -*-
# Alternating workflow: Frame -> Photo -> Frame -> Photo ... (Esc to finish)
# For each pair:
#   - Compute frame origin (Generic Annotation/Detail Component/Group)
#   - Move the photo so its center = frame origin + offsets

import clr
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

uidoc = __revit__.ActiveUIDocument
doc   = uidoc.Document
view  = doc.ActiveView

# ----- Offsets (feet) -----
OFFSET_RIGHT_FT = 14.0              # +X (right)
OFFSET_UP_FT    = 15.0 + 6.0/12.0   # +Y (up) = 15'-6"
# --------------------------

class FrameSelFilter(ISelectionFilter):
    def AllowElement(self, e):
        try:
            cat = e.Category
            if cat and cat.Id.IntegerValue in (
                int(BuiltInCategory.OST_GenericAnnotation),
                int(BuiltInCategory.OST_DetailComponents)
            ):
                return True
            if isinstance(e, Group):
                return True
        except:
            pass
        return False
    def AllowReference(self, ref, pt):
        return False

class PhotoSelFilter(ISelectionFilter):
    def AllowElement(self, e):
        return isinstance(e, ImageInstance) or isinstance(e, ImportInstance)
    def AllowReference(self, ref, pt):
        return False

def center_of_bbox(el, v):
    bb = el.get_BoundingBox(v)
    if not bb:
        return None
    return XYZ((bb.Min.X + bb.Max.X) * 0.5,
               (bb.Min.Y + bb.Max.Y) * 0.5,
               (bb.Min.Z + bb.Max.Z) * 0.5)

def frame_origin(frame_el, v):
    # Prefer true family origin (LocationPoint)
    if isinstance(frame_el, FamilyInstance):
        loc = frame_el.Location
        if isinstance(loc, LocationPoint):
            return loc.Point
        c = center_of_bbox(frame_el, v)
        if c:
            return c

    # If Group, try to find a GA inside; else group center
    if isinstance(frame_el, Group):
        try:
            for mid in frame_el.GetMemberIds():
                mem = doc.GetElement(mid)
                if isinstance(mem, FamilyInstance):
                    cat = mem.Category
                    if cat and cat.Id.IntegerValue == int(BuiltInCategory.OST_GenericAnnotation):
                        loc = mem.Location
                        if isinstance(loc, LocationPoint):
                            return loc.Point
                        c = center_of_bbox(mem, v)
                        if c:
                            return c
        except:
            pass
        c = center_of_bbox(frame_el, v)
        if c:
            return c

    # Fallback
    return center_of_bbox(frame_el, v)

# --------- Main alternating loop ---------
while True:
    # 1) Pick a frame
    try:
        ref_frame = uidoc.Selection.PickObject(ObjectType.Element, FrameSelFilter(), "Pick a frame (ESC to finish)")
    except:
        break  # Esc ends the whole tool

    frame = doc.GetElement(ref_frame.ElementId)
    target_view = doc.GetElement(frame.OwnerViewId) or view

    base_origin = frame_origin(frame, target_view)
    if base_origin is None:
        # Could not resolve; skip this cycle and ask for another frame
        continue

    target_point = base_origin + XYZ(OFFSET_RIGHT_FT, OFFSET_UP_FT, 0.0)

    # 2) Pick a photo for this frame
    try:
        ref_photo = uidoc.Selection.PickObject(ObjectType.Element, PhotoSelFilter(), "Pick a photo for this frame (ESC to finish)")
    except:
        break  # Esc ends the whole tool

    photo = doc.GetElement(ref_photo.ElementId)

    # Require same host view (legend elements are view-specific)
    if photo.OwnerViewId != frame.OwnerViewId:
        # Skip if not same view; loop back to pick another frame
        continue

    c = center_of_bbox(photo, target_view)
    if not c:
        continue

    delta = target_point - c
    if delta.IsZeroLength():
        continue

    was_pinned = getattr(photo, "Pinned", False)

    t = Transaction(doc, "Move photo to offset")
    t.Start()
    try:
        if was_pinned:
            photo.Pinned = False
        ElementTransformUtils.MoveElement(doc, photo.Id, delta)
    finally:
        if was_pinned:
            photo.Pinned = True
        t.Commit()
