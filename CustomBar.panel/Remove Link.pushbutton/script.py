# coding: utf-8
# Remove Pushbuttons — list ALL pushbuttons inside pulldown folders under THIS PANEL

import os, io, shutil, subprocess

# py2/py3 text shim
try:
    unicode
except NameError:
    unicode = str

def _alert(msg, title='Remove Buttons'):
    try:
        from pyrevit import forms
        forms.alert(msg, title=title, warn_icon=False)
    except Exception:
        print(msg)

def _get_panel_dir_from_here(script_path):
    """.../<This>.pushbutton/script.py -> .../<Panel>.panel"""
    btn_dir   = os.path.dirname(script_path)        # .../<This>.pushbutton
    panel_dir = os.path.dirname(btn_dir)            # .../<Panel>.panel
    return panel_dir, btn_dir

def _list_pulldown_pushbuttons(panel_dir, exclude_btn_dir=None):
    """
    Return [(display, pushbutton_dir)] for ALL .pushbutton bundles found
    inside any *.pulldown under the given panel.
    display format:  '<PulldownName>  /  <ButtonName>'
    """
    items = []
    try:
        for entry in os.listdir(panel_dir):
            if not entry.lower().endswith('.pulldown'):
                continue
            pd_name = os.path.splitext(entry)[0]              # e.g. 'CustomBarURL'
            pd_dir  = os.path.join(panel_dir, entry)
            if not os.path.isdir(pd_dir):
                continue

            for sub in os.listdir(pd_dir):
                if not sub.lower().endswith('.pushbutton'):
                    continue
                pbd = os.path.join(pd_dir, sub)
                if exclude_btn_dir and os.path.normpath(pbd) == os.path.normpath(exclude_btn_dir):
                    continue  # don't list this remover itself (in case it lives in a pulldown)
                btn_name = os.path.splitext(sub)[0]
                display  = u'{}  /  {}'.format(pd_name, btn_name)
                items.append((display, pbd))
    except Exception:
        pass

    items.sort(key=lambda x: x[0].lower())
    return items

# ---------- your proven pyRevit auto-reload ----------
def _autoreload_pyrevit():
    try:
        from pyrevit.loader import sessionmgr
        sessionmgr.reload_pyrevit(); return True
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
    appdata  = os.environ.get('APPDATA') or ''
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

# ---------- WPF UI (same ui.xaml as before) ----------
def _show_ui(xaml_path, items):
    try:
        import clr
        clr.AddReference('PresentationFramework'); clr.AddReference('WindowsBase')
        from pyrevit import forms
        import System

        class Dialog(forms.WPFWindow):
            def __init__(self, xaml, items_in):
                forms.WPFWindow.__init__(self, xaml)
                self.RemoveList = self.FindName('RemoveList')
                self.RemoveBtn  = self.FindName('RemoveBtn')
                self.CancelBtn  = self.FindName('CancelBtn')

                if self.RemoveList is not None:
                    self.RemoveList.SelectionMode = System.Windows.Controls.SelectionMode.Extended
                    for disp, _ in items_in:
                        self.RemoveList.Items.Add(disp)

                if self.RemoveBtn is not None: self.RemoveBtn.Click += self._ok
                if self.CancelBtn is not None: self.CancelBtn.Click += self._cancel

                try:
                    self.RemoveList.SelectionChanged += self._update_state
                except Exception: pass
                self._update_state(None, None)

                self.result = None

            def _update_state(self, s, a):
                try:
                    has_sel = self.RemoveList and self.RemoveList.SelectedItems and len(self.RemoveList.SelectedItems) > 0
                    if self.RemoveBtn is not None: self.RemoveBtn.IsEnabled = bool(has_sel)
                except Exception: pass

            def _ok(self, s, a):
                sels = list(self.RemoveList.SelectedItems) if self.RemoveList is not None else []
                self.result = [unicode(x) for x in sels] if sels else []
                self.Close()

            def _cancel(self, s, a):
                self.result = None
                self.Close()

        if not os.path.isfile(xaml_path):
            _alert('ui.xaml not found next to script.py', 'UI Error'); return None

        dlg = Dialog(xaml_path, items); dlg.ShowDialog()
        return dlg.result
    except Exception as e:
        _alert('Could not open UI.\n{}'.format(e), 'UI Error'); return None

def _remove_dirs(paths):
    ok, failed = [], []
    for p in paths:
        try:
            shutil.rmtree(p); ok.append(p)
        except Exception as e:
            failed.append((p, str(e)))
    return ok, failed

def main():
    here = os.path.abspath(__file__)
    panel_dir, my_btn_dir = _get_panel_dir_from_here(here)

    items = _list_pulldown_pushbuttons(panel_dir, exclude_btn_dir=my_btn_dir)
    if not items:
        _alert('No pushbuttons found under pulldown folders in this panel.', title='Nothing to Remove')
        return

    xaml_path = os.path.join(os.path.dirname(here), 'ui.xaml')
    selection = _show_ui(xaml_path, items)
    if selection is None:
        return

    disp_to_path = dict(items)
    targets = [disp_to_path.get(d) for d in selection if disp_to_path.get(d)]
    if not targets:
        return

    ok, failed = _remove_dirs(targets)
    _autoreload_pyrevit()

    msg = []
    if ok:     msg += [u'Removed:', u'\n'.join(ok), u'']
    if failed: msg += [u'Failed:',  u'\n'.join([u'{}  —  {}'.format(p, e) for p, e in failed])]
    if not msg: msg = [u'No changes.']
    _alert(u'\n'.join(msg), title='Removal Summary')

if __name__ == '__main__':
    main()
