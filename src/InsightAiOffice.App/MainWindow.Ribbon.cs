using System.IO;
using System.Windows;
using System.Windows.Input;
using Syncfusion.Windows.Controls.RichTextBoxAdv;
using InsightAiOffice.App.Helpers;

namespace InsightAiOffice.App;

public partial class MainWindow
{
    // ── Ribbon Format Commands ───────────────────────────────────

    private void FormatBold_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType == "word" && RichTextEditor.Selection != null)
        {
            var cf = RichTextEditor.Selection.CharacterFormat;
            if (cf != null) cf.Bold = cf.Bold != true;
        }
    }

    private void FormatItalic_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType == "word" && RichTextEditor.Selection != null)
        {
            var cf = RichTextEditor.Selection.CharacterFormat;
            if (cf != null) cf.Italic = cf.Italic != true;
        }
    }

    private void FormatUnderline_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType == "word" && RichTextEditor.Selection != null)
        {
            var cf = RichTextEditor.Selection.CharacterFormat;
            if (cf != null)
                cf.Underline = cf.Underline == Underline.Single ? Underline.None : Underline.Single;
        }
    }

    private void FindReplace_Click(object sender, RoutedEventArgs e)
    {
        StatusText.Text = LanguageManager.Get("Hint_FindReplace");
    }

    // ── Word Paragraph Commands ───────────────────────────────────

    private void WordAlignLeft_Click(object sender, RoutedEventArgs e)
    {
        if (RichTextEditor.Selection?.ParagraphFormat != null)
            RichTextEditor.Selection.ParagraphFormat.TextAlignment = TextAlignment.Left;
    }

    private void WordAlignCenter_Click(object sender, RoutedEventArgs e)
    {
        if (RichTextEditor.Selection?.ParagraphFormat != null)
            RichTextEditor.Selection.ParagraphFormat.TextAlignment = TextAlignment.Center;
    }

    private void WordAlignRight_Click(object sender, RoutedEventArgs e)
    {
        if (RichTextEditor.Selection?.ParagraphFormat != null)
            RichTextEditor.Selection.ParagraphFormat.TextAlignment = TextAlignment.Right;
    }

    // ── Excel Ribbon Commands ─────────────────────────────────────

    private void ExcelCut_Click(object sender, RoutedEventArgs e)
    {
        try { ApplicationCommands.Cut.Execute(null, Spreadsheet?.ActiveGrid); }
        catch (InvalidOperationException) { StatusText.Text = LanguageManager.Get("Hint_Cut"); }
    }

    private void ExcelCopy_Click(object sender, RoutedEventArgs e)
    {
        try { ApplicationCommands.Copy.Execute(null, Spreadsheet?.ActiveGrid); }
        catch (InvalidOperationException) { StatusText.Text = LanguageManager.Get("Hint_Copy"); }
    }

    private void ExcelPaste_Click(object sender, RoutedEventArgs e)
    {
        try { ApplicationCommands.Paste.Execute(null, Spreadsheet?.ActiveGrid); }
        catch (InvalidOperationException) { StatusText.Text = LanguageManager.Get("Hint_Paste"); }
    }

    private void ExcelBold_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var ws = Spreadsheet?.ActiveSheet;
            var range = Spreadsheet?.ActiveGrid?.SelectedRanges?.ActiveRange;
            if (ws == null || range == null) return;
            var r = ws.Range[range.Top, range.Left, range.Bottom, range.Right];
            r.CellStyle.Font.Bold = !r.CellStyle.Font.Bold;
            Spreadsheet?.ActiveGrid?.InvalidateCells();
        }
        catch (Exception ex) { StatusText.Text = $"{LanguageManager.Get("Error_Title")}: {ex.Message}"; }
    }

    private void ExcelItalic_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var ws = Spreadsheet?.ActiveSheet;
            var range = Spreadsheet?.ActiveGrid?.SelectedRanges?.ActiveRange;
            if (ws == null || range == null) return;
            var r = ws.Range[range.Top, range.Left, range.Bottom, range.Right];
            r.CellStyle.Font.Italic = !r.CellStyle.Font.Italic;
            Spreadsheet?.ActiveGrid?.InvalidateCells();
        }
        catch (Exception ex) { StatusText.Text = $"{LanguageManager.Get("Error_Title")}: {ex.Message}"; }
    }

    private void ExcelUnderline_Click(object sender, RoutedEventArgs e)
    {
        StatusText.Text = LanguageManager.Get("Hint_Underline");
    }

    private void ExcelAlignLeft_Click(object sender, RoutedEventArgs e) => SetExcelFormat("halign-left");
    private void ExcelAlignCenter_Click(object sender, RoutedEventArgs e) => SetExcelFormat("halign-center");
    private void ExcelAlignRight_Click(object sender, RoutedEventArgs e) => SetExcelFormat("halign-right");
    private void ExcelPercent_Click(object sender, RoutedEventArgs e) => SetExcelFormat("percent");
    private void ExcelComma_Click(object sender, RoutedEventArgs e) => SetExcelFormat("comma");

    private void SetExcelFormat(string action)
    {
        try
        {
            var ws = Spreadsheet?.ActiveSheet;
            var range = Spreadsheet?.ActiveGrid?.SelectedRanges?.ActiveRange;
            if (ws == null || range == null) return;
            var r = ws.Range[range.Top, range.Left, range.Bottom, range.Right];
            switch (action)
            {
                case "percent": r.NumberFormat = "0%"; break;
                case "comma": r.NumberFormat = "#,##0"; break;
                case "halign-left": r.CellStyle.HorizontalAlignment = (dynamic)1; break;
                case "halign-center": r.CellStyle.HorizontalAlignment = (dynamic)2; break;
                case "halign-right": r.CellStyle.HorizontalAlignment = (dynamic)3; break;
            }
            Spreadsheet?.ActiveGrid?.InvalidateCells();
        }
        catch (Exception ex) { StatusText.Text = $"{LanguageManager.Get("Error_Title")}: {ex.Message}"; }
    }

    // ── PPTX Ribbon Commands ──────────────────────────────────────

    private void PptxExtractText_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_currentDocPath)) return;
        var text = ExtractPptxContent(_currentDocPath);
        System.Windows.Clipboard.SetText(text.Length > 5000 ? text[..5000] : text);
        StatusText.Text = LanguageManager.Format("Hint_TextExtracted", text.Length);
    }

    // ── Ribbon Export Commands ─────────────────────────────────────

    private void ExportWord_Click(object sender, RoutedEventArgs e)
    {
        HideAllBackstages();
        if (_activeEditorType == "word")
        {
            var dialog = new Microsoft.Win32.SaveFileDialog { Filter = "Word|*.docx", DefaultExt = ".docx" };
            if (dialog.ShowDialog() == true)
            {
                using var stream = File.Create(dialog.FileName);
                RichTextEditor.Save(stream, FormatType.Docx);
                StatusText.Text = LanguageManager.Format("Hint_WordExported", Path.GetFileName(dialog.FileName));
            }
        }
        else
            StatusText.Text = LanguageManager.Get("Hint_ExportWordFirst");
    }

    private void ExportExcel_Click(object sender, RoutedEventArgs e)
    {
        HideAllBackstages();
        if (_activeEditorType == "excel")
        {
            var dialog = new Microsoft.Win32.SaveFileDialog { Filter = "Excel|*.xlsx", DefaultExt = ".xlsx" };
            if (dialog.ShowDialog() == true)
            {
                Spreadsheet.SaveAs(dialog.FileName);
                StatusText.Text = LanguageManager.Format("Hint_ExcelExported", Path.GetFileName(dialog.FileName));
            }
        }
        else
            StatusText.Text = LanguageManager.Get("Hint_ExportExcelFirst");
    }

    private void HideAllBackstages()
    {
        try { MainRibbon?.HideBackStage(); } catch (InvalidOperationException) { /* backstage not active */ }
        try { WordRibbon?.HideBackStage(); } catch (InvalidOperationException) { /* backstage not active */ }
        try { ExcelRibbon?.HideBackStage(); } catch (InvalidOperationException) { /* backstage not active */ }
        try { PptxRibbon?.HideBackStage(); } catch (InvalidOperationException) { /* backstage not active */ }
    }

    // ── Backstage / Help ──────────────────────────────────────────

    private void RibbonHelp_Click(object sender, RoutedEventArgs e)
    {
        Views.HelpWindow.ShowSection(this, "overview");
        StatusText.Text = LanguageManager.Get("Menu_Help");
    }

    private void BackstagePrint_Click(object sender, RoutedEventArgs e)
    {
        HideAllBackstages();
        if (_activeEditorType == "word")
            RichTextEditor.PrintDocument();
        else
            StatusText.Text = LanguageManager.Get("Hint_PrintWordOnly");
    }

    private void BackstageSettings_Click(object sender, RoutedEventArgs e)
    {
        HideAllBackstages();
        new Views.SettingsWindow(_aiService) { Owner = this }.ShowDialog();
    }

    private void BackstageLicense_Click(object sender, RoutedEventArgs e)
    {
        HideAllBackstages();
        ShowLicenseDialog();
        UpdatePlanBadge();
        UpdateLicenseBackstage();
    }

    private void LanguageRadio_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not System.Windows.Controls.RadioButton rb || rb.Tag is not string lang) return;
        Helpers.LanguageManager.SetLanguage(lang);
        ApplyLocalization();
        UpdatePlanBadge();
        UpdateLicenseBackstage();
    }

    private void InitializeLanguageRadioButtons()
    {
        var current = Helpers.LanguageManager.CurrentLanguage;
        LangJapanese.IsChecked = current == "ja";
        LangEnglish.IsChecked = current == "en";
    }

    private void BackstageCloseDoc_Click(object sender, RoutedEventArgs e)
    {
        HideAllBackstages();
        CloseAllEditors();

        FileNameLabel.Text = LanguageManager.Get("App_Tagline");
        FileTypeLabel.Text = "";
        _currentDocPath = "";
        StatusText.Text = LanguageManager.Get("Doc_Closed");
    }
}
