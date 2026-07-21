# -*- coding: utf-8 -*-
"""TAA Export Combined PDF, Revit 2025.4 / pyRevit."""

__title__ = "Export\nCombined PDF"
__author__ = "TAA Archi"

from pyrevit import revit, DB, forms, script

import clr
import json
import os

clr.AddReference("System.Collections")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")
clr.AddReference("System.Windows.Forms")

from System.Collections.Generic import List
from System.Windows import MessageBox
from System.Windows.Forms import FolderBrowserDialog, DialogResult

output = script.get_output()

SCRIPT_DIR = os.path.dirname(__file__)
XAML_FILE = os.path.join(SCRIPT_DIR, "ExportDialog.xaml")
RULES_DIR = os.path.join(os.environ.get("APPDATA", SCRIPT_DIR), "TaaExportPDF")
RULES_FILE = os.path.join(RULES_DIR, "naming_rules.json")

QUALITY_MAP = {
    "72 DPI (Draft)": DB.PDFExportQualityType.DPI72,
    "144 DPI (Screen)": DB.PDFExportQualityType.DPI144,
    "300 DPI (Print)": DB.PDFExportQualityType.DPI300,
    "600 DPI (High Detail)": DB.PDFExportQualityType.DPI600,
}

COLOR_MAP = {
    "Color": DB.ColorDepthType.Color,
    "Grayscale": DB.ColorDepthType.GrayScale,
    "Black and White": DB.ColorDepthType.BlackLine,
}


def read_rules():
    if not os.path.exists(RULES_FILE):
        return {}
    try:
        with open(RULES_FILE, "r") as f:
            return json.load(f)
    except Exception:
        return {}


def write_rules(rules):
    if not os.path.exists(RULES_DIR):
        os.makedirs(RULES_DIR)
    with open(RULES_FILE, "w") as f:
        json.dump(rules, f, indent=2)


def get_all_sheets(doc):
    sheets = DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet) \
        .WhereElementIsNotElementType()
    return sorted([s for s in sheets if not s.IsPlaceholder],
                  key=lambda s: s.SheetNumber)


def get_sheet_parameters(sheet):
    names = []
    for p in sheet.Parameters:
        if p.Definition and p.Definition.Name:
            names.append(p.Definition.Name)
    return sorted(set(names))


def get_param_value(element, name):
    p = element.LookupParameter(name)
    if not p or not p.HasValue:
        return None
    if p.StorageType == DB.StorageType.String:
        return p.AsString()
    if p.StorageType == DB.StorageType.Integer:
        return str(p.AsInteger())
    if p.StorageType == DB.StorageType.Double:
        return p.AsValueString() or str(p.AsDouble())
    if p.StorageType == DB.StorageType.ElementId:
        return str(p.AsElementId().IntegerValue)
    return None


def sanitize_filename(name):
    for c in ['<', '>', ':', '"', '/', '\\', '|', '?', '*']:
        name = name.replace(c, "_")
    return name.strip().rstrip('.')


def get_native_sets(doc):
    result = {}
    for item in DB.FilteredElementCollector(doc).OfClass(DB.ViewSheetSet):
        result[item.Name] = [v.Id.IntegerValue for v in item.Views
                             if isinstance(v, DB.ViewSheet)]
    return result


def save_native_set(doc, name, sheets):
    pm = doc.PrintManager
    pm.PrintRange = DB.PrintRange.Select
    view_set = DB.ViewSet()
    for sheet in sheets:
        view_set.Insert(sheet)

    setting = pm.ViewSheetSetting
    existing = None
    for item in DB.FilteredElementCollector(doc).OfClass(DB.ViewSheetSet):
        if item.Name == name:
            existing = item
            break

    if existing:
        setting.CurrentViewSheetSet = existing
        setting.CurrentViewSheetSet.Views = view_set
        setting.Save()
    else:
        setting.CurrentViewSheetSet.Views = view_set
        setting.SaveAs(name)


def delete_native_set(doc, name):
    for item in DB.FilteredElementCollector(doc).OfClass(DB.ViewSheetSet):
        if item.Name == name:
            doc.Delete(item.Id)
            return True
    return False


class SheetItem(object):
    def __init__(self, sheet):
        self.sheet = sheet
        self.Number = sheet.SheetNumber
        self.Name = sheet.Name
        self.IsSelected = False


class ParamItem(object):
    def __init__(self, name):
        self.Name = name
        self.IsSelected = False


class ExportDialog(forms.WPFWindow):
    def __init__(self, doc, sheets):
        forms.WPFWindow.__init__(self, XAML_FILE)
        self.doc = doc
        self.sheet_items = [SheetItem(s) for s in sheets]
        self.param_selection_order = []
        self.output_folder = None
        self.export_result = None
        self._sheet_sets = {}
        self._rules = {}
        self._loading = False

        self.lb_sheets.ItemsSource = self.sheet_items

        builtins = ["[Project Number]", "[Project Name]", "[Client Name]",
                    "[Sheet Number Range]", "[Custom Text]"]
        names = builtins + (get_sheet_parameters(sheets[0]) if sheets else [])
        self.param_items = [ParamItem(n) for n in sorted(set(names))]
        self.lb_params.ItemsSource = self.param_items

        self.refresh_sheet_sets()
        self.refresh_rules()
        self.update_count()
        self.update_preview()

    def refresh_sheet_sets(self):
        self._sheet_sets = get_native_sets(self.doc)
        self.cb_sheet_sets.Items.Clear()
        self.cb_sheet_sets.Items.Add("(Select a saved sheet set...)")
        for name in sorted(self._sheet_sets):
            self.cb_sheet_sets.Items.Add("{} ({} sheets)".format(
                name, len(self._sheet_sets[name])))
        self.cb_sheet_sets.SelectedIndex = 0

    def refresh_rules(self):
        self._rules = read_rules()
        self.cb_rules.Items.Clear()
        self.cb_rules.Items.Add("(Select a naming rule...)")
        for name in sorted(self._rules):
            self.cb_rules.Items.Add(name)
        self.cb_rules.SelectedIndex = 0

    def sheet_set_changed(self, sender, args):
        if self._loading or self.cb_sheet_sets.SelectedIndex <= 0:
            return
        names = sorted(self._sheet_sets.keys())
        name = names[self.cb_sheet_sets.SelectedIndex - 1]
        ids = self._sheet_sets[name]
        for item in self.sheet_items:
            item.IsSelected = item.sheet.Id.IntegerValue in ids
        self.lb_sheets.Items.Refresh()
        self.update_count()
        self.update_preview()

    def rule_changed(self, sender, args):
        if self._loading or self.cb_rules.SelectedIndex <= 0:
            return
        name = sorted(self._rules.keys())[self.cb_rules.SelectedIndex - 1]
        rule = self._rules[name]
        self._loading = True
        try:
            self.txt_prefix.Text = rule.get("prefix", "")
            self.txt_suffix.Text = rule.get("suffix", "")
            self.txt_separator.Text = rule.get("separator", "_")
            self.param_selection_order = []
            wanted = rule.get("parameters", [])
            for item in self.param_items:
                item.IsSelected = item.Name in wanted
                if item.IsSelected:
                    self.param_selection_order.append(item)
            self.lb_params.Items.Refresh()
            self.cb_quality.SelectedIndex = max(0, [
                "72 DPI (Draft)", "144 DPI (Screen)",
                "300 DPI (Print)", "600 DPI (High Detail)"
            ].index(rule.get("quality", "300 DPI (Print)")))
            self.cb_color.SelectedIndex = max(0, [
                "Color", "Grayscale", "Black and White"
            ].index(rule.get("color", "Color")))
            self.chk_vector.IsChecked = rule.get("vector", True)
        finally:
            self._loading = False
        self.update_preview()

    def save_sheet_set_click(self, sender, args):
        selected = [x.sheet for x in self.sheet_items if x.IsSelected]
        if not selected:
            MessageBox.Show("Select at least one sheet first.")
            return
        name = forms.ask_for_string("Sheet set name:", "Save Sheet Set", "My Set")
        if name:
            with revit.Transaction("Save Sheet Set"):
                save_native_set(self.doc, name.strip(), selected)
            self.refresh_sheet_sets()

    def delete_sheet_set_click(self, sender, args):
        if not self._sheet_sets:
            MessageBox.Show("No native sheet sets found.")
            return
        name = forms.SelectFromList.show(sorted(self._sheet_sets),
                                         title="Delete Sheet Set",
                                         button_name="Delete")
        if name:
            with revit.Transaction("Delete Sheet Set"):
                delete_native_set(self.doc, name)
            self.refresh_sheet_sets()

    def save_rule_click(self, sender, args):
        if not self.param_selection_order:
            MessageBox.Show("Select at least one naming parameter first.")
            return
        name = forms.ask_for_string("Naming rule name:", "Save Naming Rule", "My Rule")
        if not name:
            return
        self._rules[name.strip()] = self.current_rule()
        write_rules(self._rules)
        self.refresh_rules()

    def delete_rule_click(self, sender, args):
        if not self._rules:
            MessageBox.Show("No naming rules found.")
            return
        name = forms.SelectFromList.show(sorted(self._rules),
                                         title="Delete Naming Rule",
                                         button_name="Delete")
        if name:
            del self._rules[name]
            write_rules(self._rules)
            self.refresh_rules()

    def sheet_checked(self, sender, args):
        self.update_count()
        self.update_preview()

    def param_checked(self, sender, args):
        item = sender.DataContext
        if item.IsSelected and item not in self.param_selection_order:
            self.param_selection_order.append(item)
        elif not item.IsSelected and item in self.param_selection_order:
            self.param_selection_order.remove(item)
        self.update_preview()

    def select_all_click(self, sender, args):
        for item in self.sheet_items:
            item.IsSelected = True
        self.lb_sheets.Items.Refresh()
        self.update_count()
        self.update_preview()

    def select_none_click(self, sender, args):
        for item in self.sheet_items:
            item.IsSelected = False
        self.lb_sheets.Items.Refresh()
        self.update_count()
        self.update_preview()

    def update_count(self):
        count = len([x for x in self.sheet_items if x.IsSelected])
        self.txt_sheet_count.Text = "{} of {} sheets selected".format(
            count, len(self.sheet_items))

    def update_preview_event(self, sender, args):
        self.update_preview()

    def update_preview(self):
        parts = self.build_name_parts(preview=True)
        all_parts = []
        if self.txt_prefix.Text.strip():
            all_parts.append(self.txt_prefix.Text.strip())
        all_parts.extend(parts)
        if self.txt_suffix.Text.strip():
            all_parts.append(self.txt_suffix.Text.strip())
        sep = self.txt_separator.Text or "_"
        name = sep.join(all_parts) if all_parts else "(select parameters)"
        self.txt_preview.Text = sanitize_filename(name) + ".pdf"

    def build_name_parts(self, preview=False):
        selected = [x for x in self.sheet_items if x.IsSelected]
        parts = []
        for item in self.param_selection_order:
            key = item.Name
            value = None
            if key == "[Project Number]":
                value = self.doc.ProjectInformation.Number
            elif key == "[Project Name]":
                value = self.doc.ProjectInformation.Name
            elif key == "[Client Name]":
                value = get_param_value(self.doc.ProjectInformation, "Client Name")
            elif key == "[Sheet Number Range]":
                if selected:
                    value = selected[0].Number
                    if len(selected) > 1:
                        value += "-" + selected[-1].Number
            elif key == "[Custom Text]":
                value = "CUSTOM"
            elif selected:
                value = get_param_value(selected[0].sheet, key)
            if value:
                parts.append(value)
            elif preview:
                parts.append("<missing:{}>".format(key))
        return parts

    def missing_parameters(self, selected):
        missing = []
        if not selected:
            return ["No sheet selected"]
        first = selected[0].sheet
        for item in self.param_selection_order:
            key = item.Name
            if key in ["[Custom Text]", "[Sheet Number Range]"]:
                continue
            if key == "[Project Number]":
                ok = bool(self.doc.ProjectInformation.Number)
            elif key == "[Project Name]":
                ok = bool(self.doc.ProjectInformation.Name)
            elif key == "[Client Name]":
                ok = bool(get_param_value(self.doc.ProjectInformation, "Client Name"))
            else:
                ok = bool(get_param_value(first, key))
            if not ok:
                missing.append(key)
        return missing

    def current_rule(self):
        q = self.cb_quality.SelectedItem.Content if self.cb_quality.SelectedItem else "300 DPI (Print)"
        c = self.cb_color.SelectedItem.Content if self.cb_color.SelectedItem else "Color"
        return {"prefix": self.txt_prefix.Text,
                "suffix": self.txt_suffix.Text,
                "separator": self.txt_separator.Text or "_",
                "parameters": [x.Name for x in self.param_selection_order],
                "quality": q, "color": c,
                "vector": bool(self.chk_vector.IsChecked)}

    def browse_click(self, sender, args):
        dlg = FolderBrowserDialog()
        dlg.Description = "Select output folder for PDF"
        if dlg.ShowDialog() == DialogResult.OK:
            self.output_folder = dlg.SelectedPath
            self.txt_folder.Text = self.output_folder

    def cancel_click(self, sender, args):
        self.Close()

    def export_click(self, sender, args):
        selected = [x for x in self.sheet_items if x.IsSelected]
        if not selected:
            MessageBox.Show("Select at least one sheet.")
            return
        if not self.param_selection_order:
            MessageBox.Show("Select at least one naming parameter.")
            return
        if not self.output_folder:
            MessageBox.Show("Choose an output folder.")
            return
        missing = self.missing_parameters(selected)
        if missing:
            MessageBox.Show("Missing or empty parameters:\n\n{}\n\nExport cancelled.".format(
                "\n".join(missing)))
            return
        parts = self.build_name_parts()
        all_parts = ([self.txt_prefix.Text.strip()] if self.txt_prefix.Text.strip() else [])
        all_parts += parts
        all_parts += ([self.txt_suffix.Text.strip()] if self.txt_suffix.Text.strip() else [])
        filename = sanitize_filename((self.txt_separator.Text or "_").join(all_parts))
        self.export_result = {"sheets": [x.sheet for x in selected],
                              "filename": filename,
                              "folder": self.output_folder,
                              "quality": QUALITY_MAP.get(self.cb_quality.SelectedItem.Content, DB.PDFExportQualityType.DPI300),
                              "color": COLOR_MAP.get(self.cb_color.SelectedItem.Content, DB.ColorDepthType.Color),
                              "vector": bool(self.chk_vector.IsChecked)}
        self.Close()


def export_pdf(doc, result):
    ids = List[DB.ElementId]()
    for sheet in result["sheets"]:
        ids.Add(sheet.Id)
    options = DB.PDFExportOptions()
    options.Combine = True
    options.FileName = result["filename"]
    options.ExportQuality = result["quality"]
    options.ColorDepth = result["color"]
    options.AlwaysUseRaster = not result["vector"]
    options.HideCropBoundaries = True
    options.HideReferencePlane = True
    options.HideScopeBoxes = True
    options.HideUnreferencedViewTags = True
    return doc.Export(result["folder"], ids, options)


def main():
    doc = revit.doc
    sheets = get_all_sheets(doc)
    if not sheets:
        forms.alert("No sheets found.", exitscript=True)
    dialog = ExportDialog(doc, sheets)
    dialog.ShowDialog()
    if not dialog.export_result:
        return
    result = dialog.export_result
    if export_pdf(doc, result):
        output.print_md("## PDF Exported")
        output.print_md("**File:** `{}\\{}.pdf`".format(result["folder"], result["filename"]))
    else:
        forms.alert("Export failed. Check the output folder.")


if __name__ == "__main__":
    main()
