# coding: utf-8
# Shortcut Manager (pyRevit) — Add / Remove under RevitYYYY.pulldown in THIS PANEL
# - Add: creates a .pushbutton under "<This Panel>.panel/Revit{YEAR}.pulldown"
# - Remove: lists existing ACC shortcut buttons inside all RevitYYYY.pulldown under THIS panel
# - Uses version-specific icons (RevitYYYY.png or "Revit YYYY".png) beside this script if present.

import os, io, re, shutil, subprocess
import clr, System
from pyrevit import forms
from Autodesk.Revit.DB import ModelPathUtils, OpenOptions, DetachFromCentralOption

# WPF refs (for older hosts)
try:
    clr.AddReference('PresentationCore')
    clr.AddReference('PresentationFramework')
    clr.AddReference('WindowsBase')
except Exception:
    pass

uiapp = __revit__
DEFAULT_REGION = 'US'
SCRIPT_DIR = os.path.dirname(__file__)
XAML_PATH = os.path.join(SCRIPT_DIR, 'ui.xaml')

# ---------- helpers ----------
def _sanitize_name(name):
    name = re.sub(r'[<>:"/\\|?*\x00-\x1F]+', '-', name).strip().rstrip('.')
    return name or 'Launcher'

def _find_icon(src_dir):
    for n in ['icon.png', 'button.png', 'Icon.png', 'icon-32.png', 'large.png']:
        p = os.path.join(src_dir, n)
        if os.path.exists(p):
            return p
    for fn in os.listdir(src_dir):
        if fn.lower().endswith('.png'):
            return os.path.join(src_dir, fn)
    return None

def _pick_version_icon(src_dir, version_str):
    # Revit2024.png  or  Revit 2024.png
    for name in [
        'Revit{}.png'.format(version_str),
        'Revit {}.png'.format(version_str),
        'revit{}.png'.format(version_str),
        'revit {}.png'.format(version_str),
    ]:
        p = os.path.join(src_dir, name)
        if os.path.exists(p):
            return p
    return None

def _unique_dir(base_dir, desired_name):
    out = os.path.join(base_dir, desired_name)
    if not os.path.exists(out):
        return out
    i = 2
    while True:
        cand = os.path.join(base_dir, '{} ({})'.format(desired_name, i))
        if not os.path.exists(cand):
            return cand
        i += 1

def _get_revit_major_version():
    try:
        return uiapp.Application.VersionNumber
    except Exception:
        ad = uiapp.ActiveUIDocument
        return ad.Document.Application.VersionNumber if ad else 'Unknown'

def _get_current_panel_dir():
    """.../<Panel>.panel that contains THIS button."""
    return os.path.dirname(SCRIPT_DIR)

def _get_or_create_version_pulldown_dir():
    """Return pulldown dir '<This Panel>.panel/Revit{YEAR}.pulldown' (create if missing), plus version string."""
    version_str = _get_revit_major_version()
    panel_dir = _get_current_panel_dir()
    target_pulldown_dir = os.path.join(panel_dir, 'Revit{}.pulldown'.format(version_str))
    if not os.path.exists(target_pulldown_dir):
        os.makedirs(target_pulldown_dir)
    return target_pulldown_dir, version_str

def _write_openonly_script(dst_dir):
    code = '''# coding: utf-8
import os, io, System
from Autodesk.Revit.DB import ModelPathUtils, OpenOptions, DetachFromCentralOption

uiapp = __revit__
LINK_FILE = os.path.join(os.path.dirname(__file__), 'link.txt')

def parse_kv(text):
    kv = {}
    for line in text.splitlines():
        if '=' in line:
            k, v = line.split('=', 1)
            kv[k.strip().upper()] = v.strip()
    return kv

def build_cloud_modelpath(region_str, proj_guid_str, model_guid_str):
    pg = System.Guid(proj_guid_str)
    mg = System.Guid(model_guid_str)
    try:
        return ModelPathUtils.ConvertCloudGUIDsToCloudPath(region_str, pg, mg)
    except Exception:
        return ModelPathUtils.ConvertCloudGUIDsToCloudPath(pg, mg)

def main():
    raw = io.open(LINK_FILE, 'r', encoding='utf-8').read()
    kv = parse_kv(raw)
    region = kv.get('REGION', 'US').upper()
    project_guid = kv.get('PROJECT_GUID')
    model_guid = kv.get('MODEL_GUID')
    cloud_path = build_cloud_modelpath(region, project_guid, model_guid)
    opts = OpenOptions()
    opts.DetachFromCentralOption = DetachFromCentralOption.DoNotDetach
    uiapp.OpenAndActivateDocument(cloud_path, opts, False)

if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        from pyrevit import forms
        forms.alert('Open failed.\\n\\n{}'.format(e), exitscript=True)
'''
    with io.open(os.path.join(dst_dir, 'script.py'), 'w', encoding='utf-8') as f:
        f.write(code)

def _write_bundle_yaml(dst_dir, title):
    """Create bundle.yaml to force a line break in ribbon label (e.g., 'Revit\\nPin')."""
    multiline = title.replace(' ', '\n', 1)
    yaml_text = u'title: |\n  {}\n'.format(multiline.replace('\n', '\n  '))
    with io.open(os.path.join(dst_dir, 'bundle.yaml'), 'w', encoding='utf-8') as f:
        f.write(yaml_text)

def _write_launcher_files(dst_dir, label, region, proj_guid, model_guid, version_str):
    contents = (
        'MODE=ACC\nREGION={}\nPROJECT_GUID={}\nMODEL_GUID={}\n'
    ).format(region, proj_guid, model_guid)
    with io.open(os.path.join(dst_dir, 'link.txt'), 'w', encoding='utf-8') as f:
        f.write(contents)

    _write_openonly_script(dst_dir)
    _write_bundle_yaml(dst_dir, label)  # ensures two-word labels wrap to two lines on the ribbon

    src_icon = _pick_version_icon(SCRIPT_DIR, version_str) or _find_icon(SCRIPT_DIR)
    if src_icon:
        try:
            shutil.copy2(src_icon, os.path.join(dst_dir, 'icon.png'))
        except Exception:
            pass

    tip = 'Opens cloud model: {}\nProject GUID: {}\nModel GUID: {}'.format(label, proj_guid, model_guid)
    try:
        with io.open(os.path.join(dst_dir, 'tooltip.txt'), 'w', encoding='utf-8') as f:
            f.write(tip)
    except Exception:
        pass

# ---- listing in pulldowns (THIS PANEL) ----
_PULDOWN_RE = re.compile(r'^revit(\d{4})$', re.IGNORECASE)

def _list_existing_shortcuts_in_pulldowns():
    """Scan RevitYYYY.pulldown under THIS panel; return (display, fullpath) for ACC shortcuts."""
    items = []
    panel_dir = _get_current_panel_dir()
    try:
        for entry in os.listdir(panel_dir):
            if not entry.lower().endswith('.pulldown'):
                continue
            base = os.path.splitext(entry)[0]  # e.g., 'Revit2024'
            if not _PULDOWN_RE.match(base):
                continue
            pulldown_dir = os.path.join(panel_dir, entry)
            if not os.path.isdir(pulldown_dir):
                continue

            for sub in os.listdir(pulldown_dir):
                if not sub.lower().endswith('.pushbutton'):
                    continue
                pbd = os.path.join(pulldown_dir, sub)
                link = os.path.join(pbd, 'link.txt')
                if os.path.exists(link):
                    try:
                        raw = io.open(link, 'r', encoding='utf-8').read()
                        u = raw.upper()
                        if ('MODE=ACC' in u) and ('PROJECT_GUID' in u) and ('MODEL_GUID' in u):
                            btn_name = os.path.splitext(sub)[0]
                            display = u'{}  /  {}'.format(base, btn_name)
                            items.append((display, pbd))
                    except Exception:
                        continue
    except Exception:
        pass
    items.sort(key=lambda x: x[0].lower())
    return items

# ---------- auto-reload ----------
def _autoreload_pyrevit():
    try:
        from pyrevit.loader import sessionmgr
        sessionmgr.reload_pyrevit()
        return True
    except Exception:
        pass
    si = None
    try:
        si = subprocess.STARTUPINFO(); si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    except Exception:
        pass
    try:
        subprocess.Popen(['pyrevit', 'reload'], startupinfo=si, shell=False); return True
    except Exception:
        pass
    appdata = os.environ.get('APPDATA') or ''
    localapp = os.environ.get('LOCALAPPDATA') or ''
    for exe in [
        os.path.join(appdata, 'pyRevit-Master', 'bin', 'pyrevit.exe'),
        os.path.join(appdata, 'pyRevit', 'bin', 'pyrevit.exe'),
        os.path.join(localapp, 'pyRevit', 'bin', 'pyrevit.exe'),
        r'C:\Program Files\pyRevit\pyrevit.exe',
        r'C:\Program Files (x86)\pyRevit\pyrevit.exe',
    ]:
        try:
            if os.path.exists(exe):
                subprocess.Popen([exe, 'reload'], startupinfo=si, shell=False); return True
        except Exception:
            continue
    return False

# ---------- WPF UI ----------
class ShortcutManager(forms.WPFWindow):
    def __init__(self, xaml_path, model_name, panel_label, default_label):
        forms.WPFWindow.__init__(self, xaml_path)
        self._result = None
        self._remove_map = {}  # display -> path

        # Fill header/body info (defensively; some header controls may be absent)
        try:
            self.DetectedModel.Text = model_name
        except Exception:
            pass
        try:
            self.PanelName.Text = panel_label  # no-op if not present in XAML
        except Exception:
            pass
        try:
            self.LabelInput.Text = default_label
        except Exception:
            pass

        # Default mode: Add
        self.mode_add(None, None)
        # Initial remove list
        self.refresh_remove_list(None, None)

    # --- Mode toggles ---
    def mode_add(self, sender, args):
        try:
            self.AddPanel.Visibility = System.Windows.Visibility.Visible
            self.RemovePanel.Visibility = System.Windows.Visibility.Collapsed
            self.CreateBtn.Visibility = System.Windows.Visibility.Visible
            self.RemoveSelectedBtn.Visibility = System.Windows.Visibility.Collapsed
            self.Subheader.Text = ""
            self.AddRadio.IsChecked = True
        except Exception:
            pass

    def mode_remove(self, sender, args):
        try:
            self.AddPanel.Visibility = System.Windows.Visibility.Collapsed
            self.RemovePanel.Visibility = System.Windows.Visibility.Visible
            self.CreateBtn.Visibility = System.Windows.Visibility.Collapsed
            self.RemoveSelectedBtn.Visibility = System.Windows.Visibility.Visible
            self.Subheader.Text = ""
            self.RemoveRadio.IsChecked = True
            self.refresh_remove_list(None, None)
        except Exception:
            pass

    # --- Buttons ---
    def create(self, sender, args):
        label = (self.LabelInput.Text or '').strip()
        if not label:
            forms.alert('Please enter a shortcut label.', title='Missing Label'); return
        reload_flag = bool(self.ReloadCheck.IsChecked)
        self._result = {'mode': 'add', 'label': label, 'reload': reload_flag}
        self.Close()

    def remove(self, sender, args):
        sels = list(self.RemoveList.SelectedItems) if self.RemoveList.SelectedItems else []
        if not sels:
            forms.alert('Select at least one shortcut to remove.'); return
        paths = [self._remove_map.get(s) for s in sels if s in self._remove_map]
        paths = [p for p in paths if p]
        self._result = {'mode': 'remove', 'paths': paths, 'reload': bool(self.ReloadCheckRemove.IsChecked)}
        self.Close()

    def cancel(self, sender, args):
        self._result = None
        self.Close()

    def refresh_remove_list(self, sender, args):
        try:
            self.RemoveList.Items.Clear()
            self._remove_map.clear()
            for display, path in _list_existing_shortcuts_in_pulldowns():
                self.RemoveList.Items.Add(display)
                self._remove_map[display] = path
        except Exception:
            pass

# ---------- flows ----------
def _add_shortcut_from_active_doc(label):
    uidoc = uiapp.ActiveUIDocument
    if not uidoc:
        forms.alert('Open the cloud model you want to create a shortcut for, then click again.', exitscript=True)
    doc = uidoc.Document
    if not getattr(doc, 'IsModelInCloud', False):
        forms.alert('The active document is not a cloud model.', exitscript=True)

    mp = doc.GetCloudModelPath()
    try:
        proj_guid = mp.GetProjectGUID(); model_guid = mp.GetModelGUID()
    except AttributeError:
        try:
            proj_guid = ModelPathUtils.GetProjectGUID(mp); model_guid = ModelPathUtils.GetModelGUID(mp)
        except Exception:
            forms.alert('Could not read GUIDs from this Revit version.', exitscript=True)

    pulldown_dir, version_str = _get_or_create_version_pulldown_dir()
    safe_label = _sanitize_name(label)
    new_bundle = safe_label + '.pushbutton'
    dst_dir = _unique_dir(pulldown_dir, new_bundle)
    os.makedirs(dst_dir)

    _write_launcher_files(dst_dir, label, DEFAULT_REGION, str(proj_guid), str(model_guid), version_str)
    return dst_dir, version_str

def _remove_shortcuts(paths):
    ok, failed = [], []
    for p in paths:
        try:
            shutil.rmtree(p)
            ok.append(p)
        except Exception as e:
            failed.append((p, str(e)))
    return ok, failed

# ---------- main ----------
def main():
    # Defaults for UI
    uidoc = uiapp.ActiveUIDocument
    docname = '(no document)'
    if uidoc and uidoc.Document:
        dn = uidoc.Document.Title
        docname = dn[:-4] if dn.lower().endswith('.rvt') else dn
    default_label = 'Open {}'.format(docname)
    panel_label = 'Revit {}'.format(_get_revit_major_version())

    # Show UI (fallback to simple add if XAML missing)
    result = None
    if os.path.exists(XAML_PATH):
        dlg = ShortcutManager(XAML_PATH, docname, panel_label, default_label)
        dlg.ShowDialog()
        result = dlg._result
        if not result:
            return
    else:
        label = forms.ask_for_string(prompt='Name for the new launcher button:', default=default_label, title='Create Model Launcher')
        if not label:
            return
        result = {'mode': 'add', 'label': label, 'reload': True}

    if result['mode'] == 'add':
        dst_dir, version_str = _add_shortcut_from_active_doc(result['label'])
        if result.get('reload'):
            reloaded = _autoreload_pyrevit()
        else:
            reloaded = False
        msg = ['Created shortcut under "Revit{}" pulldown:'.format(version_str), dst_dir, '']
        if result.get('reload'):
            msg.append('✅ pyRevit UI {}.'.format('reloaded' if reloaded else 'could not auto-reload — please click Reload'))
        forms.alert('\n'.join(msg), exitscript=False)

    elif result['mode'] == 'remove':
        ok, failed = _remove_shortcuts(result.get('paths', []))
        if result.get('reload'):
            _autoreload_pyrevit()
        lines = []
        if ok:
            lines += ['Removed:', '\n'.join(ok), '']
        if failed:
            lines += ['Failed:', '\n'.join(['{}  —  {}'.format(p, err) for p, err in failed])]
        if not lines:
            lines = ['No changes made.']
        forms.alert('\n'.join(lines), exitscript=False)

if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        forms.alert('Unexpected error.\n\n{}'.format(e), exitscript=True)
