# coding: utf-8
# Sheet Index Revision Dots
# - Stable per-revision columns named "SEQ n" (sequence-based)
# - UI (ui.xaml) lets you: pick revisions, set width & orientation, header content,
#   and optionally group headers with a custom title (Revit 2023+ only).
# - Remembers last settings; width only resets when ORIENTATION changes
# - Preview applies without closing; Update applies and closes
# - Deletes orphan SEQ params when revisions are removed
# - Orders columns by sequence (SEQ 1, SEQ 2, …)
# - Quick wait box with rotating messages

import clr, os, io, json

# -------------------- INSTANT WAIT (WinForms) --------------------
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import (Form, Label, ProgressBar, ProgressBarStyle,
                                  Application, FormBorderStyle, FormStartPosition, Timer)
from System.Drawing import (Size, Point, Font, FontStyle, Color)

class QuickSplash(object):
    def __init__(self, title=u"Sheet Index Revision Dots",
                 messages=None,
                 cycle_ms=1600,
                 width=420,
                 height=120):
        self.messages = messages or [u"Doing the heavy lifting for you…",
                                     u"Hang tight, almost there…"]
        self._idx = 0

        self.form = Form()
        self.form.Text = title
        self.form.StartPosition = FormStartPosition.CenterScreen
        self.form.FormBorderStyle = FormBorderStyle.FixedDialog
        self.form.MaximizeBox = False
        self.form.MinimizeBox = False
        self.form.TopMost = True
        self.form.ClientSize = Size(width, height)
        self.form.BackColor = Color.FromArgb(250, 250, 250)

        self.title = Label()
        self.title.Text = title
        self.title.AutoSize = True
        self.title.Font = Font("Segoe UI", 11, FontStyle.Bold)
        self.title.ForeColor = Color.FromArgb(17, 24, 39)
        self.title.Location = Point(12, 10)

        self.msg = Label()
        self.msg.Text = self.messages[0]
        self.msg.AutoSize = True
        self.msg.Font = Font("Segoe UI", 10, FontStyle.Regular)
        self.msg.ForeColor = Color.FromArgb(55, 65, 81)
        self.msg.Location = Point(12, 40)

        self.pb = ProgressBar()
        self.pb.Location = Point(12, 70)
        self.pb.Width = width - 24
        self.pb.Style = ProgressBarStyle.Marquee
        self.pb.MarqueeAnimationSpeed = 30

        self.form.Controls.Add(self.title)
        self.form.Controls.Add(self.msg)
        self.form.Controls.Add(self.pb)

        self._timer = Timer()
        self._timer.Interval = int(cycle_ms)
        def _tick(sender, e):
            if not self.messages:
                return
            self._idx = (self._idx + 1) % len(self.messages)
            try:
                self.msg.Text = self.messages[self._idx]
                Application.DoEvents()
            except:
                pass
        self._timer.Tick += _tick

    def show(self):
        self.form.Show()
        try:
            Application.DoEvents()
        except:
            pass
        self._timer.Start()

    def set_status(self, text):
        try:
            self.msg.Text = text
            Application.DoEvents()
        except:
            pass

    def close(self):
        try:
            self._timer.Stop()
        except:
            pass
        try:
            self.form.Close()
        except:
            pass

_splash = QuickSplash(messages=[u"Doing the heavy lifting for you…", u"Hang tight, almost there…"])
_splash.show()

_splash_is_closed = False
def close_splash_safe():
    global _splash_is_closed
    if not _splash_is_closed:
        try:
            _splash.close()
        except:
            pass
        _splash_is_closed = True

_user_cancelled = False

# -------------------- DOMAIN IMPORTS --------------------
from pyrevit import forms
from Autodesk.Revit.DB import (
    FilteredElementCollector, ViewSheet, ViewSchedule, Revision,
    BuiltInCategory, Category, Transaction, TransactionStatus,
    BuiltInParameterGroup, CategorySet, InstanceBinding,
    ExternalDefinitionCreationOptions, ScheduleHorizontalAlignment
)

# Optional orientation (works on 2023+)
try:
    from Autodesk.Revit.DB import ScheduleHeadingOrientation
except Exception:
    ScheduleHeadingOrientation = None

# WPF (ui.xaml)
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
from System import Uri
from System.Windows import Thickness, Visibility
from System.Windows.Media import Brushes
from System.Windows.Controls import CheckBox
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption
from System.Windows.Input import Keyboard, ModifierKeys  # <-- for Shift-click

uiapp = __revit__             # type: ignore
app   = uiapp.Application
doc   = uiapp.ActiveUIDocument.Document

# -------------------- CONFIG --------------------
SCHEDULE_NAME_HINT   = "Sheet Index"
INCLUDE_UNISSUED     = True
PARAM_PREFIX         = "SEQ "
PARAM_GROUP_NAME     = "pyRevit_RevDots"
PARAM_GROUP          = BuiltInParameterGroup.PG_TEXT
DOT                  = u"\u25CF"          # ●
AUTO_CREATE_SCHEDULE = True

VERT_DEFAULT_IN  = 0.5000   # vertical default width
HORIZ_DEFAULT_IN = 2.0000   # horizontal default width

GROUP_HEADER_DEFAULT = u"Revisions"       # Revit 2023+ only
# ------------------------------------------------

# ----------------------- helpers ----------------
def safe_get_name(obj):
    try:
        return obj.GetName()
    except TypeError:
        return obj.GetName(doc)
    except Exception:
        return getattr(obj, "Name", "")

def get_all_revisions():
    revs = list(FilteredElementCollector(doc).OfClass(Revision))
    if not INCLUDE_UNISSUED:
        revs = [r for r in revs if getattr(r, "Issued", False)]
    try:
        revs.sort(key=lambda r: r.SequenceNumber)
    except Exception:
        pass
    return revs

def ensure_shared_params_file():
    spf = app.SharedParametersFilename
    if spf and os.path.exists(spf):
        return app.OpenSharedParameterFile()
    script_dir = os.path.dirname(__file__)
    temp_spf = os.path.join(script_dir, "_rev_dots_sharedparams.txt")
    if not os.path.exists(temp_spf):
        with io.open(temp_spf, "w", encoding="utf-8") as f:
            f.write("# pyRevit Rev Dots shared parameters\n")
    app.SharedParametersFilename = temp_spf
    return app.OpenSharedParameterFile()

def get_or_create_sp_group(def_file, group_name):
    for g in def_file.Groups:
        if g.Name == group_name:
            return g
    return def_file.Groups.Create(group_name)

def ensure_sheet_param(def_file, param_name):
    group = get_or_create_sp_group(def_file, PARAM_GROUP_NAME)
    ext_def = None
    for d in group.Definitions:
        if d.Name == param_name:
            ext_def = d; break
    if not ext_def:
        try:
            from Autodesk.Revit.DB import ParameterType
            opts = ExternalDefinitionCreationOptions(param_name, ParameterType.Text)
        except Exception:
            from Autodesk.Revit.DB import SpecTypeId
            opts = ExternalDefinitionCreationOptions(param_name, SpecTypeId.String.Text)
        ext_def = group.Definitions.Create(opts)

    catset = CategorySet()
    catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_Sheets))
    binding = app.Create.NewInstanceBinding(catset)
    pb = doc.ParameterBindings
    if not pb.Insert(ext_def, binding, PARAM_GROUP):
        pb.ReInsert(ext_def, binding, PARAM_GROUP)

def find_or_create_sheetlist_schedule():
    cands = []
    for vs in FilteredElementCollector(doc).OfClass(ViewSchedule):
        try:
            if vs.Definition and vs.Definition.CategoryId.IntegerValue == int(BuiltInCategory.OST_Sheets):
                cands.append(vs)
        except Exception:
            pass
    if cands:
        for s in cands:
            if SCHEDULE_NAME_HINT and SCHEDULE_NAME_HINT.lower() in s.Name.lower():
                return s
        if len(cands) == 1:
            return cands[0]
        close_splash_safe()
        pick = forms.SelectFromList.show([s.Name for s in cands], title="Select Sheet Index schedule", multiselect=False)
        if pick:
            for s in cands:
                if s.Name == pick:
                    return s
        global _user_cancelled
        _user_cancelled = True
        return None
    if not AUTO_CREATE_SCHEDULE:
        return None
    try:
        cat_id = Category.GetCategory(doc, BuiltInCategory.OST_Sheets).Id
        sched  = ViewSchedule.CreateSchedule(doc, cat_id)
        sched.Name = "Sheet Index"
        return sched
    except Exception:
        return None

def schedule_has_field(schedule, field_name):
    sd = schedule.Definition
    for i in range(sd.GetFieldCount()):
        f = sd.GetField(i)
        if safe_get_name(f) == field_name:
            return True
    return False

def remove_field_from_schedule(schedule, field_name):
    sd = schedule.Definition
    for i in reversed(range(sd.GetFieldCount())):
        f = sd.GetField(i)
        if safe_get_name(f) == field_name:
            try:
                sd.RemoveField(i)
            except Exception:
                try:
                    sd.RemoveField(f.FieldId)
                except Exception:
                    return False
    return True

def add_field_to_schedule(schedule, field_name, column_heading=None, column_width_ft=None, heading_orientation="vertical"):
    sd = schedule.Definition
    target_field = None

    for sf in sd.GetSchedulableFields():
        if safe_get_name(sf) != field_name:
            continue
        if not schedule_has_field(schedule, field_name):
            target_field = sd.AddField(sf)
        else:
            for i in range(sd.GetFieldCount()):
                f = sd.GetField(i)
                if safe_get_name(f) == field_name:
                    target_field = f
                    break
        break

    if not target_field:
        return False

    try:
        if column_heading is not None:
            target_field.ColumnHeading = column_heading
    except Exception:
        pass
    try:
        target_field.HorizontalAlignment = ScheduleHorizontalAlignment.Center
    except Exception:
        pass

    if ScheduleHeadingOrientation:
        try:
            if heading_orientation == "horizontal":
                target_field.HeadingOrientation = ScheduleHeadingOrientation.Horizontal
            else:
                target_field.HeadingOrientation = ScheduleHeadingOrientation.Vertical
        except Exception:
            pass

    try:
        if column_width_ft is None:
            column_width_ft = (HORIZ_DEFAULT_IN/12.0 if heading_orientation == "horizontal" else VERT_DEFAULT_IN/12.0)
        target_field.GridColumnWidth = column_width_ft
    except Exception:
        pass

    return True

def set_sheet_text(sheet, name, value):
    p = sheet.LookupParameter(name)
    if p and p.AsString() != value:
        p.Set(value)
        return True
    return False

def parse_width_inches(text_in):
    s = (text_in or "").strip().lower()
    if not s:
        return None
    for token in ['inches', 'inch', 'in', '"']:
        s = s.replace(token, '')
    s = s.strip()
    try:
        val = float(s)
        return val if val > 0 else None
    except Exception:
        return None

def try_load_logo(img_control):
    folder = os.path.dirname(__file__)
    for name in ["logo.png", "logo.jpg", "logo.jpeg", "logo.bmp", "logo.ico"]:
        path = os.path.join(folder, name)
        if os.path.exists(path):
            try:
                bi = BitmapImage()
                bi.BeginInit()
                bi.UriSource = Uri(path)
                bi.CacheOption = BitmapCacheOption.OnLoad
                bi.EndInit()
                img_control.Source = bi
                img_control.Visibility = Visibility.Visible
                return True
            except Exception:
                pass
    return False

def seq_from_stable_name(stable_name):
    try:
        s = stable_name[len(PARAM_PREFIX):].strip()
        num = []
        for ch in s:
            if ch.isdigit():
                num.append(ch)
            else:
                break
        return int(''.join(num)) if num else 10**9
    except Exception:
        return 10**9

def reorder_rev_fields(schedule, names_in_desired_order, headings_by_name, width_feet, heading_orientation):
    for name in names_in_desired_order:
        remove_field_from_schedule(schedule, name)
    for name in names_in_desired_order:
        heading = headings_by_name.get(name, name)
        add_field_to_schedule(schedule, name, column_heading=heading, column_width_ft=width_feet, heading_orientation=heading_orientation)

def all_rev_project_param_names():
    names = set()
    try:
        pb = doc.ParameterBindings
        it = pb.ForwardIterator(); it.Reset()
        while it.MoveNext():
            defn = it.Key
            nm = getattr(defn, "Name", "")
            try:
                is_str = isinstance(nm, basestring)
            except:
                is_str = isinstance(nm, str)
            if is_str and nm.startswith(PARAM_PREFIX):
                names.add(nm)
    except Exception:
        pass
    return names

def unbind_project_parameter_by_name(name):
    try:
        pb = doc.ParameterBindings
        it = pb.ForwardIterator(); it.Reset()
        while it.MoveNext():
            defn = it.Key
            if getattr(defn, "Name", "") == name:
                return pb.Remove(defn)
    except Exception:
        pass
    return False

def make_heading_text(date_text, desc_text, mode_key, orient_key):
    d = (date_text or "").strip()
    c = (desc_text or "").strip()
    if mode_key == "date":
        return d
    if mode_key == "desc":
        return c
    if orient_key == "horizontal":
        if d and c:
            return u"{}\n{}".format(d, c)     # two lines, no hyphen
        return d or c
    else:
        if d and c:
            return u"{} - {}".format(d, c)    # one line with hyphen
        return d or c

# -------- Group/UnGroup for Revit 2023+ only (modern API) --------
def _seq_column_bounds(schedule, prefix):
    sd = schedule.Definition
    seq_idxs = []
    for i in range(sd.GetFieldCount()):
        nm = safe_get_name(sd.GetField(i))
        if isinstance(nm, str) and nm.startswith(prefix):
            seq_idxs.append(i)
    if not seq_idxs:
        return None
    return (min(seq_idxs), max(seq_idxs))

def _group_headers_modern(schedule, left, right, title):
    try:
        if hasattr(schedule, "CanUngroupHeaders") and schedule.CanUngroupHeaders(0, left, 0, right):
            try:
                schedule.UngroupHeaders(0, left, 0, right)
            except Exception:
                pass
        if hasattr(schedule, "CanGroupHeaders") and not schedule.CanGroupHeaders(0, left, 0, right):
            try:
                doc.Regenerate()
            except Exception:
                pass
        schedule.GroupHeaders(0, left, 0, right, title)
        return True
    except Exception:
        return False

def _ungroup_headers_modern(schedule, left, right):
    try:
        if hasattr(schedule, "CanUngroupHeaders") and schedule.CanUngroupHeaders(0, left, 0, right):
            schedule.UngroupHeaders(0, left, 0, right)
            return True
    except Exception:
        pass
    return False

def group_rev_headers(schedule, prefix, title):
    # 2023+: use modern API; 2022 will simply return False
    try:
        if not hasattr(schedule, "GroupHeaders"):
            return False
        bounds = _seq_column_bounds(schedule, prefix)
        if not bounds:
            return False
        left, right = bounds
        try:
            doc.Regenerate()
        except Exception:
            pass
        return _group_headers_modern(schedule, left, right, title)
    except Exception:
        return False

def ungroup_rev_headers(schedule, prefix):
    try:
        if not hasattr(schedule, "UngroupHeaders"):
            return False
        bounds = _seq_column_bounds(schedule, prefix)
        if not bounds:
            return False
        left, right = bounds
        try:
            doc.Regenerate()
        except Exception:
            pass
        return _ungroup_headers_modern(schedule, left, right)
    except Exception:
        return False

# ---------------- Settings (persist across runs) ----------------
def _settings_path():
    return os.path.join(os.path.dirname(__file__), "_rev_dots_settings.json")

def load_settings():
    try:
        with io.open(_settings_path(), "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_settings(mode_key, orient_key, width_in, group_enabled, group_title):
    try:
        data = {
            "mode": mode_key,
            "orient": orient_key,
            "width_in": float(width_in),
            "group_enabled": bool(group_enabled),
            "group_title": group_title or GROUP_HEADER_DEFAULT
        }
        with io.open(_settings_path(), "w", encoding="utf-8") as f:
            json.dump(data, f)
    except Exception:
        pass

# -------------------- PREP (splash visible) --------------------
try:
    _splash.set_status("Checking document…")
    if doc.IsFamilyDocument:
        close_splash_safe()
        forms.alert("Open a project document (not a family) and run again.")
        raise SystemExit

    _splash.set_status("Loading settings…")
    _settings = load_settings()

    _splash.set_status("Scanning revisions…")
    revs = get_all_revisions()
    if not revs:
        close_splash_safe()
        forms.alert("No revisions found (or all are unissued and filtering is set to Issued-only).")
        raise SystemExit

    _splash.set_status("Locating Sheet Index…")
    schedule = find_or_create_sheetlist_schedule()
    if not schedule:
        if _user_cancelled:
            raise SystemExit
        close_splash_safe()
        forms.alert("No Sheet List schedule found/created. Create one, then run again.")
        raise SystemExit

    _splash.set_status("Preparing shared parameters…")
    def_file = ensure_shared_params_file()
    if not def_file:
        close_splash_safe()
        forms.alert("Could not open or create a Shared Parameters file.")
        raise SystemExit

    _splash.set_status("Restoring previous choices…")
    default_orient = _settings.get("orient", "vertical")
    default_mode   = _settings.get("mode", "both")
    default_group_enabled = bool(_settings.get("group_enabled", True))
    default_group_title   = _settings.get("group_title", GROUP_HEADER_DEFAULT)
    if "width_in" in _settings:
        default_width_in = float(_settings.get("width_in", VERT_DEFAULT_IN))
    else:
        default_width_in = HORIZ_DEFAULT_IN if default_orient == "horizontal" else VERT_DEFAULT_IN

    _splash.set_status("Building revision list…")
    entries = []  # (stable_name, rev_id, date_text, desc_text)
    seen = set()
    for r in revs:
        stable_name = u"{}{}".format(PARAM_PREFIX, r.SequenceNumber)
        i = 2
        while stable_name in seen:
            stable_name = u"{}{} ({})".format(PARAM_PREFIX, r.SequenceNumber, i); i += 1
        seen.add(stable_name)

        date_text = (getattr(r, "RevisionDate", "") or "").strip()
        desc_text = (getattr(r, "Description", "") or "").strip()
        entries.append((stable_name, r.Id, date_text, desc_text))

finally:
    pass

# ---------- helpers that depend on schedule/entries ----------
def current_existing_names(schedule, entries):
    names = set()
    sd = schedule.Definition
    existing = set()
    for i in range(sd.GetFieldCount()):
        existing.add(safe_get_name(sd.GetField(i)))
    for (stable_name, _rid, _d, _c) in entries:
        if stable_name in existing:
            names.add(stable_name)
    return names

def perform_update(checked_names, mode_key, orient_key, width_inches, group_enabled, group_title, show_alert=True):
    width_feet = float(width_inches) / 12.0
    existing_now = current_existing_names(schedule, entries)
    to_show = set(checked_names)
    to_hide = existing_now - to_show

    def make_heading(date_text, desc_text):
        return make_heading_text(date_text, desc_text, mode_key, orient_key)
    parts = {st: (rid, d, c) for (st, rid, d, c) in entries}
    headings_by_name = {st: make_heading(d, c) for st, (rid, d, c) in parts.items()}

    current_rev_stable_names = set(st for (st, _rid, _d, _c) in entries)
    project_rev_param_names = all_rev_project_param_names()
    orphan_param_names = sorted(list(project_rev_param_names - current_rev_stable_names), key=seq_from_stable_name)

    t = Transaction(doc, "Sheet Index Revision Dots — Apply")
    deleted_params = hidden = updated = 0
    try:
        t.Start()

        for pname in orphan_param_names:
            remove_field_from_schedule(schedule, pname)
            if unbind_project_parameter_by_name(pname):
                deleted_params += 1

        for stable_name in to_hide:
            if remove_field_from_schedule(schedule, stable_name):
                hidden += 1

        for stable_name in to_show:
            ensure_sheet_param(def_file, stable_name)
            add_field_to_schedule(
                schedule,
                stable_name,
                column_heading=headings_by_name.get(stable_name, stable_name),
                column_width_ft=width_feet,
                heading_orientation=orient_key
            )

        ordered_names = sorted(list(to_show), key=seq_from_stable_name)
        if ordered_names:
            reorder_rev_fields(schedule, ordered_names, headings_by_name, width_feet, orient_key)

        if to_show:
            all_sheets = list(FilteredElementCollector(doc).OfClass(ViewSheet))
            for sh in all_sheets:
                try:
                    rev_ids = set(rid.IntegerValue for rid in sh.GetAllRevisionIds())
                except Exception:
                    rev_ids = set()
                for stable_name in to_show:
                    rid, d, c = parts[stable_name]
                    val = DOT if rid.IntegerValue in rev_ids else ""
                    if set_sheet_text(sh, stable_name, val):
                        updated += 1

        # Group or ungroup (only works on Revit 2023+)
        try:
            doc.Regenerate()
        except Exception:
            pass
        try:
            if group_enabled and to_show:
                group_rev_headers(schedule, prefix=PARAM_PREFIX, title=(group_title or GROUP_HEADER_DEFAULT))
            else:
                # user turned grouping off -> ungroup if currently grouped (2023+ only)
                ungroup_rev_headers(schedule, prefix=PARAM_PREFIX)
        except Exception:
            pass

        t.Commit()
        if show_alert:
            forms.toast(
                "Applied — Shown: {} | Hidden: {} | Updated cells: {} | Deleted params: {}\nWidth: {:.4f}\"  Orientation: {}  Grouped: {}".format(
                    len(to_show), hidden, updated, deleted_params, float(width_inches), orient_key.capitalize(),
                    "Yes" if group_enabled else "No"
                ),
                title="Sheet Index Revision Dots",
                appid="Sheet Index Revision Dots"
            )
        return True
    except Exception as e:
        try:
            if t.GetStatus() == TransactionStatus.Started:
                t.RollBack()
        except Exception:
            pass
        forms.alert("Failed to apply changes.\nAll changes were rolled back.\n\nDetails:\n{}".format(e))
        return False

# -------------------- Main UI wrapper (uses your ui.xaml) --------------------
class RevDotsUI(forms.WPFWindow):
    """Width resets to DEFAULT only when the user changes orientation."""
    def __init__(self, xaml_path, ui_items, start_width_in, default_mode, default_orient, default_group_enabled, default_group_title):
        forms.WPFWindow.__init__(self, xaml_path)
        self.result = None
        self._last_clicked_index = None  # <-- anchor for Shift-click ranges
        self._vert_default_str  = "{:.4f}".format(VERT_DEFAULT_IN)
        self._horiz_default_str = "{:.4f}".format(HORIZ_DEFAULT_IN)
        self._suppress_orient_event = True

        # Wire buttons
        self.UpdateBtn.Click    += self.on_update
        self.CancelBtn.Click    += self.on_cancel
        self.CheckAllBtn.Click  += self.on_check_all
        self.CheckNoneBtn.Click += self.on_check_none
        if hasattr(self, "PreviewBtn"):
            self.PreviewBtn.Click += self.on_preview

        # Initial values
        try:
            self.WidthBox.Text = "{:.4f}".format(start_width_in)
        except:
            pass
        try:
            idx = {"both":0, "date":1, "desc":2}.get(default_mode, 0)
            self.HeadingModeBox.SelectedIndex = idx
        except:
            pass
        try:
            self.OrientationBox.SelectedIndex = (1 if default_orient == "horizontal" else 0)
        except:
            pass

        # Group UI defaults
        try:
            self.GroupHeadersCheck.IsChecked = bool(default_group_enabled)
        except:
            pass
        try:
            self.GroupTitleBox.Text = default_group_title or GROUP_HEADER_DEFAULT
        except:
            pass
        try:
            self.GroupTitleBox.IsEnabled = bool(self.GroupHeadersCheck.IsChecked)
        except:
            pass

        # Handlers AFTER initial load
        try:
            self.OrientationBox.SelectionChanged += self.on_orientation_changed
        except:
            pass
        try:
            self.GroupHeadersCheck.Checked   += self.on_group_toggle
            self.GroupHeadersCheck.Unchecked += self.on_group_toggle
        except:
            pass
        self._suppress_orient_event = False

        # Optional logo
        try:
            if hasattr(self, "LogoImg"):
                try_load_logo(self.LogoImg)
        except:
            pass

        # Populate list (now with Shift-click range selection)
        for idx, it in enumerate(ui_items):
            cb = CheckBox()
            cb.Content = it["label"]
            cb.Tag = it["stable"]
            cb.Margin = Thickness(0, 3, 0, 3)
            cb.IsChecked = it["checked"]
            if it["checked"]:
                cb.Foreground = Brushes.DimGray
            self.RevList.Items.Add(cb)

            # --- Shift-click range selection ---
            def _make_click(idx_local, cb_local):
                def _on_click(sender, args):
                    try:
                        shift_down = (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift
                    except:
                        shift_down = False

                    if shift_down and self._last_clicked_index is not None:
                        start = min(self._last_clicked_index, idx_local)
                        end   = max(self._last_clicked_index, idx_local)
                        new_state = bool(cb_local.IsChecked)  # apply clicked state to the whole range
                        for j in range(start, end + 1):
                            item_cb = self.RevList.Items[j]
                            item_cb.IsChecked = new_state

                    # anchor for the next shift-range
                    self._last_clicked_index = idx_local
                return _on_click

            cb.Click += _make_click(idx, cb)

    def on_group_toggle(self, sender, args):
        try:
            self.GroupTitleBox.IsEnabled = bool(self.GroupHeadersCheck.IsChecked)
        except:
            pass

    def on_orientation_changed(self, sender, args):
        if getattr(self, "_suppress_orient_event", False):
            return
        try:
            if self.OrientationBox.SelectedIndex == 1:
                self.WidthBox.Text = self._horiz_default_str   # Horizontal -> 2.0000"
            else:
                self.WidthBox.Text = self._vert_default_str    # Vertical   -> 0.5000"
        except Exception:
            pass

    def on_check_all(self, sender, args):
        for i in range(self.RevList.Items.Count):
            cb = self.RevList.Items[i]
            cb.IsChecked = True

    def on_check_none(self, sender, args):
        for i in range(self.RevList.Items.Count):
            cb = self.RevList.Items[i]
            cb.IsChecked = False

    def _read_heading_mode(self):
        try:
            txt = self.HeadingModeBox.SelectedItem.Content.ToString().lower()
        except Exception:
            return "both"
        if "date only" in txt:         return "date"
        if "description only" in txt:  return "desc"
        return "both"

    def _read_orientation(self):
        try:
            txt = self.OrientationBox.SelectedItem.Content.ToString().lower()
        except Exception:
            return "vertical"
        if "horizontal" in txt:
            return "horizontal"
        return "vertical"

    def _read_grouping(self):
        try:
            enabled = bool(self.GroupHeadersCheck.IsChecked)
        except:
            enabled = True
        try:
            title = self.GroupTitleBox.Text or GROUP_HEADER_DEFAULT
        except:
            title = GROUP_HEADER_DEFAULT
        return enabled, title

    def _current_checked(self):
        checked = []
        for i in range(self.RevList.Items.Count):
            cb = self.RevList.Items[i]
            if bool(cb.IsChecked):
                checked.append(cb.Tag)
        return checked

    def _current_width_inches(self):
        width_in = parse_width_inches(self.WidthBox.Text)
        orient_key = self._read_orientation()
        if width_in is None:
            width_in = HORIZ_DEFAULT_IN if orient_key == "horizontal" else VERT_DEFAULT_IN
        return float(width_in)

    def on_preview(self, sender, args):
        checked    = self._current_checked()
        orient_key = self._read_orientation()
        mode_key   = self._read_heading_mode()
        width_in   = self._current_width_inches()
        group_enabled, group_title = self._read_grouping()
        save_settings(mode_key, orient_key, width_in, group_enabled, group_title)
        perform_update(checked, mode_key, orient_key, width_in, group_enabled, group_title, show_alert=True)

    def on_update(self, sender, args):
        checked    = self._current_checked()
        orient_key = self._read_orientation()
        mode_key   = self._read_heading_mode()
        width_in   = self._current_width_inches()
        group_enabled, group_title = self._read_grouping()
        save_settings(mode_key, orient_key, width_in, group_enabled, group_title)
        perform_update(checked, mode_key, orient_key, width_in, group_enabled, group_title, show_alert=False)
        self.result = {
            "checked": checked, "width_in": width_in, "mode": mode_key, "orient": orient_key,
            "group_enabled": group_enabled, "group_title": group_title
        }
        self.Close()

    def on_cancel(self, sender, args):
        global _user_cancelled
        _user_cancelled = True
        self.result = None
        self.Close()

# ------------------------ SHOW MAIN UI ----------------------------
existing_names = set()
try:
    existing_names = current_existing_names(schedule, entries)
except:
    existing_names = set()

ui_items = []
for (stable_name, rid, d, c) in entries:
    preview = (d + (" | " + c if (d and c) else (c or ""))).strip()
    label = u"[{}]  {}".format(stable_name, preview)
    ui_items.append({"label": label, "stable": stable_name, "checked": (stable_name in existing_names)})

close_splash_safe()

xaml_path = os.path.join(os.path.dirname(__file__), "ui.xaml")
dlg = RevDotsUI(
    xaml_path, ui_items,
    default_width_in, default_mode, default_orient,
    default_group_enabled, default_group_title
)
dlg.ShowDialog()

if dlg.result is None or _user_cancelled:
    raise SystemExit

checked_names = set(dlg.result.get("checked", []))
width_inches  = float(dlg.result.get("width_in", VERT_DEFAULT_IN))
mode_key      = dlg.result.get("mode", "both")
orient_key    = dlg.result.get("orient", "vertical")
group_enabled = bool(dlg.result.get("group_enabled", True))
group_title   = dlg.result.get("group_title", GROUP_HEADER_DEFAULT)

perform_update(checked_names, mode_key, orient_key, width_inches, group_enabled, group_title, show_alert=False)
from pyrevit import forms as _forms_toast
_forms_toast.toast("Done. Columns updated.", title="Sheet Index Revision Dots", appid="Sheet Index Revision Dots")
