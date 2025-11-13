# -*- coding: utf-8 -*-
# Modeless "Keynote Finder" window:
# - Reads all keynotes from current project's keynote table
# - Shows them in a searchable list
# - Owned by Revit, so it stays in front of Revit but not above other apps

from Autodesk.Revit.DB import KeynoteTable
from pyrevit import revit, forms
from System.Windows.Interop import WindowInteropHelper
import os


class KeynoteFinder(forms.WPFWindow):
    def __init__(self, xaml_path):
        forms.WPFWindow.__init__(self, xaml_path)

        # Make this window owned by Revit's main window:
        # When Revit is active, this window stays in front of it.
        # When you switch to another app (PDF, Bluebeam, etc.), that app can cover it.
        helper = WindowInteropHelper(self)
        helper.Owner = __revit__.MainWindowHandle

        self.doc = revit.doc
        self.all_items = self._load_keynotes()

        # Bind data to list
        self.KeynoteList.ItemsSource = self.all_items

        # Hook up events
        self.SearchBox.TextChanged += self.on_search_changed
        self.KeynoteList.MouseDoubleClick += self.on_item_double_click

    # ---------- helper for sorting ----------

    def _sort_item(self, item):
        """
        item looks like 'KEY | TEXT'
        We sort by the leading integer part of KEY, so 2 < 10 < 100, etc.
        """
        prefix = item.split('|', 1)[0].strip()  # '10', '2', '08 11 13', etc.
        num_str = ''
        for ch in prefix:
            if ch.isdigit():
                num_str += ch
            else:
                break

        if num_str:
            try:
                return (int(num_str), prefix)
            except:
                pass

        # Fallback: big number so non-numeric keys go to the end, then by text
        return (999999999, prefix)

    # ---------- data loading ----------

    def _load_keynotes(self):
        """Return list of display strings 'KEY | TEXT' from the current project's keynote table."""
        ktable = KeynoteTable.GetKeynoteTable(self.doc)
        if not ktable:
            forms.alert('No keynote table found in this project.\nCheck Keynoting Settings.')
            return []

        kentries = ktable.GetKeyBasedTreeEntries()

        items = []
        for entry in kentries:
            # Some nodes are just folders; skip anything without key or text
            try:
                key = entry.Key
                text = entry.KeynoteText
            except:
                continue

            if not key or not text:
                continue

            items.append(u"{0} | {1}".format(key, text))

        # numeric-ish sort so 2 < 10 < 100, etc.
        return sorted(items, key=self._sort_item)

    # ---------- search + UI behavior ----------

    def on_search_changed(self, sender, args):
        """Filter list when user types."""
        filter_text = self.SearchBox.Text.strip().lower()
        if not filter_text:
            self.KeynoteList.ItemsSource = self.all_items
            return

        filtered = [
            item for item in self.all_items
            if filter_text in item.lower()
        ]
        self.KeynoteList.ItemsSource = filtered

    def on_item_double_click(self, sender, args):
        """Double-click copies keynote code to clipboard (optional helper)."""
        selected = self.KeynoteList.SelectedItem
        if not selected:
            return

        key = selected.split('|', 1)[0].strip()
        try:
            from System.Windows import Clipboard
            Clipboard.SetText(key)
            forms.toast(
                'Keynote "{}" copied to clipboard.'.format(key),
                title='Keynote Finder'
            )
        except:
            forms.alert(
                'Keynote: {}\n\n(Could not access clipboard.)'.format(key),
                title='Keynote Finder'
            )


def main():
    script_dir = os.path.dirname(__file__)
    xaml_path = os.path.join(script_dir, 'keynotefinder.xaml')

    win = KeynoteFinder(xaml_path)
    # modeless window so it stays while you work in Revit
    win.Show()


if __name__ == '__main__':
    main()
