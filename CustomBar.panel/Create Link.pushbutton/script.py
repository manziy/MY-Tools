# coding: utf-8
# Make Link Button — route into CustomBar* pulldowns + reliable auto-reload (no dupes)
# - URL   ->  <TAB>/CustomBar.panel/CustomBarURL.pulldown/<Button>.pushbutton
# - File  ->  <TAB>/CustomBar.panel/CustomBarFile.pulldown/<Button>.pushbutton
# - Folder->  <TAB>/CustomBar.panel/CustomBarFolder.pulldown/<Button>.pushbutton
# - Creates the panel/pulldown folders if missing
# - Auto icons: url.png / file.png / folder.png (fallback: icon.png)
# - Reload first; inject only if reload fails (injects into CustomBar panel as a visible fallback)

import os, io, re, shutil, subprocess

PARENT_PANEL_TITLE = 'CustomBar'   # where the pulldowns live

def _alert(msg, title='Create Link Button'):
    try:
        from pyrevit import forms
        forms.alert(msg, title=title, warn_icon=True)
    except Exception:
        pass

def _safe_name(name):
    s = re.sub(r'[<>:"/\\|?*\r\n]+', '_', (name or '').strip())
    return s or 'Link Button'

def _is_url(s):
    s = (s or '').strip().lower()
    return s.startswith('http://') or s.startswith('https://') or s.startswith('mailto:')

def _classify_target(target_text):
    """Return 'url' | 'file' | 'folder' without needing the future button dir."""
    t = os.path.expandvars(os.path.expanduser((target_text or '').strip()))
    if _is_url(t):
        return 'url'
    if os.path.isabs(t):
        try:
            if os.path.isdir(t): return 'folder'
            if os.path.isfile(t): return 'file'
        except Exception:
            pass
    # Heuristics for relative/non-existing paths:
    if t.endswith('/') or t.endswith('\\'):
        return 'folder'
    last = os.path.basename(t.rstrip('/\\'))
    if '.' in last:
        return 'file'
    return 'folder'

def _pulldown_title_for(kind):
    # Names: CustomBarURL / CustomBarFile / CustomBarFolder
    return {'url': 'CustomBarURL', 'file': 'CustomBarFile', 'folder': 'CustomBarFolder'}.get(kind, 'CustomBarFolder')

def _pick_icon_path(generator_dir, kind):
    for name in ([kind + '.png'] if kind in ('url', 'file', 'folder') else []) + ['icon.png']:
        p = os.path.join(generator_dir, name)
        if os.path.isfile(p):
            return p
    return None

# -------- opener script that goes into the new button --------
OPEN_LINK_SCRIPT = r'''# DO NOT MODIFY
import os

def _alert(msg, title='Open Link'):
    try:
        from pyrevit import forms
        forms.alert(msg, title=title, warn_icon=True)
    except Exception:
        pass

def _read_target(txt_path):
    try:
        with open(txt_path, 'r') as f:
            for raw in f:
                s = raw.strip()
                if not s or s.startswith('#'):
                    continue
                if s.lower().startswith('url='):
                    s = s[4:].strip()
                return os.path.expandvars(os.path.expanduser(s))
    except Exception as e:
        _alert('Could not read:\n{}\n\n{}'.format(txt_path, e))
    return None

def _is_url(s):
    s = (s or '').lower()
    return s.startswith('http://') or s.startswith('https://') or s.startswith('mailto:')

def main():
    bundle_dir = os.path.dirname(os.path.abspath(__file__))
    link_txt = os.path.join(bundle_dir, 'link.txt')

    target = _read_target(link_txt)
    if not target:
        _alert('Missing/empty link.txt next to script.py.\nPut a URL or file path on the first non-empty line.')
        return

    # Resolve relative file paths relative to the button folder
    if not _is_url(target) and not os.path.isabs(target):
        target = os.path.abspath(os.path.join(bundle_dir, target))

    try:
        os.startfile(target)  # fastest path on Windows
    except Exception as e:
        _alert('Failed to open:\n{}\n\n{}'.format(target, e))

if __name__ == '__main__':
    main()
'''

# -------- single-backend pickers (Cancel = cancel) --------
def _browse_for_file():
    try:
        import clr
        clr.AddReference('System.Windows.Forms')
        from System.Windows.Forms import OpenFileDialog, DialogResult
        dlg = OpenFileDialog()
        dlg.Filter = 'All files (*.*)|*.*'
        dlg.CheckFileExists = False
        return dlg.FileName if dlg.ShowDialog() == DialogResult.OK else None
    except Exception:
        return None

def _browse_for_folder():
    try:
        import clr
        clr.AddReference('System.Windows.Forms')
        from System.Windows.Forms import FolderBrowserDialog, DialogResult
        dlg = FolderBrowserDialog()
        dlg.Description = 'Select a folder'
        dlg.ShowNewFolderButton = True
        return dlg.SelectedPath if dlg.ShowDialog() == DialogResult.OK else None
    except Exception:
        return None

# -------- UI loader (uses your existing ui.xaml) --------
def _get_inputs_with_ui(xaml_path):
    try:
        import clr
        clr.AddReference('PresentationFramework')
        clr.AddReference('WindowsBase')
        from pyrevit import forms

        class Dialog(forms.WPFWindow):
            def __init__(self, xaml):
                forms.WPFWindow.__init__(self, xaml)
                self.LabelBox      = self.FindName('LabelBox')
                self.PathBox       = self.FindName('PathBox')
                self.PickFileBtn   = self.FindName('PickFileBtn')
                self.PickFolderBtn = self.FindName('PickFolderBtn')
                self.CreateBtn     = self.FindName('CreateBtn')
                self.CancelBtn     = self.FindName('CancelBtn')

                if self.LabelBox is not None:
                    try: self.LabelBox.Text = 'Button Name'
                    except: pass

                if self.PickFileBtn is not None:   self.PickFileBtn.Click   += self._pick_file
                if self.PickFolderBtn is not None: self.PickFolderBtn.Click += self._pick_folder
                if self.CreateBtn is not None:     self.CreateBtn.Click     += self._ok
                if self.CancelBtn is not None:     self.CancelBtn.Click     += self._cancel

                try:
                    self.LabelBox.TextChanged += self._update_state
                    self.PathBox.TextChanged  += self._update_state
                except Exception:
                    pass
                self._update_state(None, None)

                self.ok = False; self.label = None; self.path = None

            def _update_state(self, s, a):
                try:
                    ok = bool((self.LabelBox.Text or '').strip()) and bool((self.PathBox.Text or '').strip())
                    if self.CreateBtn is not None: self.CreateBtn.IsEnabled = ok
                except Exception:
                    pass

            def _pick_file(self, s, a):
                p = _browse_for_file()
                if p and self.PathBox is not None:
                    self.PathBox.Text = p

            def _pick_folder(self, s, a):
                p = _browse_for_folder()
                if p and self.PathBox is not None:
                    self.PathBox.Text = p

            def _ok(self, s, a):
                label = (self.LabelBox.Text or '').strip()
                path  = (self.PathBox.Text  or '').strip()
                if not label or not path:
                    _alert('Please enter both a button name and a path/URL.', 'Missing Info'); return
                self.label = label; self.path = path; self.ok = True; self.Close()

            def _cancel(self, s, a):
                self.ok = False; self.Close()

        if not os.path.isfile(xaml_path):
            _alert('ui.xaml not found beside script.py', 'UI Error'); return None, None, False
        dlg = Dialog(xaml_path); dlg.ShowDialog()
        return dlg.label, dlg.path, dlg.ok
    except Exception as e:
        _alert('Could not open UI.\n{}'.format(e), 'UI Error'); return None, None, False

# -------- Ribbon helpers (fallback injection only if reload fails) --------
def _panel_and_tab(tab_title, panel_titles):
    try:
        import clr
        clr.AddReference('AdWindows')
        from Autodesk.Windows import ComponentManager
        rc = ComponentManager.Ribbon
        if rc is None: return None, None
        ttitle = (tab_title or '').strip().lower()
        tab = None
        for t in rc.Tabs:
            if ((t.Title or '').strip().lower() == ttitle) or ((getattr(t, 'Id', '') or '').strip().lower() == ttitle):
                tab = t; break
        if tab is None: return None, None
        # try each desired title; if none found, fall back to first panel
        for want in panel_titles:
            w = (want or '').strip().lower()
            for p in tab.Panels:
                title = getattr(p.Source, 'Title', '') or ''
                if title.strip().lower() == w:
                    return tab, p
        # fallback
        return tab, tab.Panels[0] if tab.Panels.Count > 0 else (tab, None)
    except Exception:
        return None, None

def _panel_has_button(panel, label):
    try:
        for it in list(panel.Source.Items):
            if (getattr(it, 'Text', None) or '').strip().lower() == (label or '').strip().lower():
                return True
    except Exception:
        pass
    return False

def _inject_ribbon_button(tab_title, desired_panel_title, label, btn_dir):
    """If reload fails, add a visible temporary button directly to the CustomBar panel."""
    try:
        import clr
        clr.AddReference('AdWindows'); clr.AddReference('PresentationFramework'); clr.AddReference('WindowsBase')
        from Autodesk.Windows import RibbonButton, RibbonItemSize
        from System import Uri
        from System.Windows.Input import ICommand
        from System.Windows.Media.Imaging import BitmapImage
    except Exception:
        return False

    panel_try = [desired_panel_title]
    tab, panel = _panel_and_tab(tab_title, panel_try)
    if panel is None:
        return False

    if _panel_has_button(panel, label):
        return True

    class _OpenLinkCommand(ICommand):
        def __init__(self, bdir): self._bdir = bdir
        def add_CanExecuteChanged(self, h): pass
        def remove_CanExecuteChanged(self, h): pass
        def CanExecute(self, p): return True
        def Execute(self, p):
            try:
                link_txt = os.path.join(self._bdir, 'link.txt')
                tgt = None
                with io.open(link_txt, 'r', encoding='utf-8') as f:
                    for raw in f:
                        s = raw.strip()
                        if not s or s.startswith('#'): continue
                        if s.lower().startswith('url='): s = s[4:].strip()
                        tgt = os.path.expandvars(os.path.expanduser(s)); break
                if not tgt:
                    _alert('Missing/empty link.txt.\nPut a URL or file path on the first non-empty line.', 'Open Link'); return
                is_url = (tgt.lower().startswith('http://') or tgt.lower().startswith('https://') or tgt.lower().startswith('mailto:'))
                if not is_url and not os.path.isabs(tgt):
                    tgt = os.path.abspath(os.path.join(self._bdir, tgt))
                os.startfile(tgt)
            except Exception as e:
                _alert('Failed to open:\n{}\n\n{}'.format(tgt, e), 'Open Link')

    btn = RibbonButton(); btn.Text = label; btn.ShowText = True; btn.Size = RibbonItemSize.Large
    btn.CommandHandler = _OpenLinkCommand(btn_dir)

    icon_path = os.path.join(btn_dir, 'icon.png')
    if os.path.isfile(icon_path):
        try:
            uri = Uri('file:///' + icon_path.replace('\\', '/')); img = BitmapImage(uri)
            btn.LargeImage = img; btn.Image = img
        except Exception:
            pass

    try:
        panel.Source.Items.Add(btn); return True
    except Exception:
        return False

# -------- your proven pyRevit auto-reload --------
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

# -------- main --------
def main():
    this_btn_dir   = os.path.dirname(os.path.abspath(__file__))   # ...\<This>.pushbutton
    this_panel_dir = os.path.dirname(this_btn_dir)                # ...\<Panel>.panel
    this_tab_dir   = os.path.dirname(this_panel_dir)              # ...\<Tab>.tab
    tab_title      = os.path.splitext(os.path.basename(this_tab_dir))[0]  # e.g. 'USPS' from 'USPS.tab'
    xaml_path      = os.path.join(this_btn_dir, 'ui.xaml')

    # Inputs
    label, target, ok = _get_inputs_with_ui(xaml_path)
    if not ok:
        return

    # Decide pulldown by target kind
    kind             = _classify_target(target)            # 'url' | 'file' | 'folder'
    pulldown_title   = _pulldown_title_for(kind)           # e.g. 'CustomBarURL'

    # Ensure parent panel exists
    parent_panel_dir = os.path.join(this_tab_dir, PARENT_PANEL_TITLE + '.panel')
    if not os.path.isdir(parent_panel_dir):
        os.makedirs(parent_panel_dir)

    # Ensure pulldown exists
    dest_pulldown_dir = os.path.join(parent_panel_dir, pulldown_title + '.pulldown')
    if not os.path.isdir(dest_pulldown_dir):
        os.makedirs(dest_pulldown_dir)

    # Unique button folder under that pulldown
    base    = _safe_name(label)
    btn_dir = os.path.join(dest_pulldown_dir, base + '.pushbutton')
    if os.path.exists(btn_dir):
        i = 2
        while os.path.exists(os.path.join(dest_pulldown_dir, u'{} ({}).pushbutton'.format(base, i))):
            i += 1
        btn_dir = os.path.join(dest_pulldown_dir, u'{} ({}).pushbutton'.format(base, i))
    os.makedirs(btn_dir)

    # Write files
    with io.open(os.path.join(btn_dir, 'script.py'), 'w', encoding='utf-8') as f:
        f.write(OPEN_LINK_SCRIPT)
    with io.open(os.path.join(btn_dir, 'link.txt'), 'w', encoding='utf-8') as f:
        f.write((target or u'').strip() + u'\n')

    # Icon based on kind
    icon_src = _pick_icon_path(this_btn_dir, kind)
    if icon_src:
        try:
            shutil.copy2(icon_src, os.path.join(btn_dir, 'icon.png'))
        except Exception:
            pass

    # Reload first to materialize the (possibly new) pulldown and button
    if not _autoreload_pyrevit():
        # If reload failed, temporarily inject into the parent CustomBar panel (visible immediately)
        injected = _inject_ribbon_button(tab_title, PARENT_PANEL_TITLE, label, btn_dir)
        if not injected:
            _alert(u'Created:\n{}\n\nCould not auto-inject or auto-reload. Use pyRevit ➜ Reload to refresh.'
                   .format(btn_dir), title='Created')

if __name__ == '__main__':
    main()
