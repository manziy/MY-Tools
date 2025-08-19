# ===== SAVE THIS AS: script.py =====
# -*- coding: utf-8 -*-
"""
pyRevit pushbutton (READ-ONLY)
Displays revisions from a selected linked model in a WPF DataGrid.
Columns: Sequence | Revision Number | Date | Description
Requires a sibling file named 'ui.xaml' in the same folder.
"""
from __future__ import print_function
import os, clr

from Autodesk.Revit import DB
from pyrevit import forms

# Clipboard
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import Clipboard

# Data for binding
clr.AddReference("System.Data")
from System import String
from System.Data import DataTable

uiapp = __revit__
uidoc = uiapp.ActiveUIDocument
if not uidoc:
    forms.alert("No active document.")
    raise SystemExit

doc = uidoc.Document

# --- pick the main model link ---
links = list(DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance))
if not links:
    forms.alert("No Revit links found in this model.")
    raise SystemExit

if len(links) == 1:
    link_inst = links[0]
else:
    items = []
    for li in links:
        ldoc = li.GetLinkDocument()
        name = (ldoc.Title + " (loaded)") if ldoc else (li.Name + " (unloaded)")
        items.append({"name": name, "li": li})
    items.sort(key=lambda x: x["name"].lower())
    picked = forms.SelectFromList.show([it["name"] for it in items],
                                       title="Pick the main model link",
                                       button_name="Use This Link",
                                       multiselect=False)
    if not picked:
        raise SystemExit
    link_inst = [it["li"] for it in items if it["name"] == picked][0]

lnkdoc = link_inst.GetLinkDocument()
if lnkdoc is None:
    forms.alert("Selected link is not loaded. Please load it and try again.")
    raise SystemExit

# --- collect revisions from linked doc ---
revs = list(DB.FilteredElementCollector(lnkdoc).OfClass(DB.Revision))
revs.sort(key=lambda r: r.SequenceNumber)

if not revs:
    forms.alert("No revisions found in linked model: " + lnkdoc.Title)
    raise SystemExit

# robustly read revision number across versions

def get_rev_number(rev):
    try:
        num = rev.RevisionNumber
        if num is not None:
            return str(num)
    except Exception:
        pass
    for bip_name in ("PROJECT_REVISION_REVISION_NUM", "REVISION_NUMBER"):
        try:
            bip = getattr(DB.BuiltInParameter, bip_name)
            p = rev.get_Parameter(bip)
            if p:
                s = p.AsString()
                if s:
                    return s
        except Exception:
            pass
    return ""

# build DataTable -> bind to DataGrid

dt = DataTable("Revisions")
dt.Columns.Add("Sequence", String)
dt.Columns.Add("Number", String)
dt.Columns.Add("Date", String)
dt.Columns.Add("Description", String)

for r in revs:
    row = dt.NewRow()
    row["Sequence"] = str(int(r.SequenceNumber))
    row["Number"] = get_rev_number(r)
    # IMPORTANT: use escaped sequences, not literal tab/CR/LF
    row["Date"] = (r.RevisionDate or "").replace("\t", " ").replace("\r", " ").replace("\n", " ")
    row["Description"] = (r.Description or "").replace("\t", " ").replace("\r", " ").replace("\n", " ")
    dt.Rows.Add(row)

# WPF window using pyRevit's WPFWindow wrapper
class RevWindow(forms.WPFWindow):
    def __init__(self, xaml_path, link_doc, table):
        self._link_doc = link_doc
        self._table = table
        forms.WPFWindow.__init__(self, xaml_path)
        # header
        self.TitleText.Text = "Linked model: {}".format(link_doc.Title)
        # grid
        self.RevGrid.ItemsSource = table.DefaultView
        # buttons
        self.CloseBtn.Click += self._on_close
        self.CopyBtn.Click += self._on_copy

    def _on_close(self, sender, args):
        self.Close()

    def _on_copy(self, sender, args):
        lines = ["Sequence\tNumber\tDate\tDescription"]
        for row in self._table.Rows:
            lines.append("{}	{}	{}	{}".format(row["Sequence"], row["Number"], row["Date"], row["Description"]))
        try:
            Clipboard.SetText("\n".join(lines))
            try:
                forms.toast("Copied to clipboard")
            except Exception:
                pass
        except Exception:
            forms.alert("Couldn't access clipboard. Select rows and press Ctrl+C.")

# locate XAML next to script
xaml_path = os.path.join(os.path.dirname(__file__), 'ui.xaml')
if not os.path.exists(xaml_path):
    forms.alert("ui.xaml not found next to script.py.\nPlace ui.xaml in the same folder and try again.")
    raise SystemExit

# show
RevWindow(xaml_path, lnkdoc, dt).ShowDialog()