using System.IO;
using System.Windows;
using System.Windows.Input;
using Syncfusion.Windows.Controls.RichTextBoxAdv;
using InsightAiOffice.App.Helpers;

namespace InsightAiOffice.App;

public partial class MainWindow
{
    // ── Undo / Redo ─────────────────────────────────────────────

    private void WordUndo_Click(object sender, RoutedEventArgs e)
    {
        try { ApplicationCommands.Undo.Execute(null, RichTextEditor); }
        catch { StatusText.Text = LanguageManager.Get("Hint_UndoUnavailable"); }
    }

    private void WordRedo_Click(object sender, RoutedEventArgs e)
    {
        try { ApplicationCommands.Redo.Execute(null, RichTextEditor); }
        catch { StatusText.Text = LanguageManager.Get("Hint_RedoUnavailable"); }
    }

    private void ExcelUndo_Click(object sender, RoutedEventArgs e)
    {
        try { ApplicationCommands.Undo.Execute(null, Spreadsheet?.ActiveGrid); }
        catch { StatusText.Text = LanguageManager.Get("Hint_UndoUnavailable"); }
    }

    private void ExcelRedo_Click(object sender, RoutedEventArgs e)
    {
        try { ApplicationCommands.Redo.Execute(null, Spreadsheet?.ActiveGrid); }
        catch { StatusText.Text = LanguageManager.Get("Hint_RedoUnavailable"); }
    }

    // ── Helpers ────────────────────────────────────────────────────

    /// <summary>Word 文字書式を安全に適用する共通ヘルパー</summary>
    private void ApplyWordCharFormat(Action<SelectionCharacterFormat> action)
    {
        if (_activeEditorType != "word") return;
        var cf = RichTextEditor.Selection?.CharacterFormat;
        if (cf != null) action(cf);
    }

    /// <summary>Excel 選択範囲に対する操作を安全に実行する共通ヘルパー</summary>
    private void WithExcelRange(Action<dynamic> action)
    {
        try
        {
            var ws = Spreadsheet?.ActiveSheet;
            var sel = Spreadsheet?.ActiveGrid?.SelectedRanges?.ActiveRange;
            if (ws == null || sel == null) return;
            var r = ws.Range[sel.Top, sel.Left, sel.Bottom, sel.Right];
            action(r);
            Spreadsheet?.ActiveGrid?.InvalidateCells();
        }
        catch (Exception ex) { StatusText.Text = $"{LanguageManager.Get("Error_Title")}: {ex.Message}"; }
    }

    // ── Ribbon Format Commands ───────────────────────────────────

    private void FormatBold_Click(object sender, RoutedEventArgs e) =>
        ApplyWordCharFormat(cf => cf.Bold = cf.Bold != true);

    private void FormatItalic_Click(object sender, RoutedEventArgs e) =>
        ApplyWordCharFormat(cf => cf.Italic = cf.Italic != true);

    private void FormatUnderline_Click(object sender, RoutedEventArgs e) =>
        ApplyWordCharFormat(cf => cf.Underline = cf.Underline == Underline.Single ? Underline.None : Underline.Single);

    private void FindReplace_Click(object sender, RoutedEventArgs e)
    {
        StatusText.Text = LanguageManager.Get("Hint_FindReplace");
    }

    private void FormatStrikethrough_Click(object sender, RoutedEventArgs e) =>
        ApplyWordCharFormat(cf => cf.StrikeThrough = cf.StrikeThrough != StrikeThrough.SingleStrike
            ? StrikeThrough.SingleStrike : StrikeThrough.None);

    private void WordBulletList_Click(object sender, RoutedEventArgs e)
    {
        StatusText.Text = LanguageManager.Get("Hint_BulletList");
    }

    private void WordInsertImage_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "word") return;
        var dlg = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "Images|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tiff"
        };
        if (dlg.ShowDialog(this) == true)
        {
            try
            {
                using var stream = File.OpenRead(dlg.FileName);
                RichTextEditor.Selection?.InsertPicture(stream);
                StatusText.Text = LanguageManager.Format("Hint_ImageInserted", Path.GetFileName(dlg.FileName));
            }
            catch (Exception ex)
            {
                StatusText.Text = $"{LanguageManager.Get("Error_Title")}: {ex.Message}";
            }
        }
    }

    // ── Word Paragraph Commands ───────────────────────────────────

    private void WordAlignLeft_Click(object sender, RoutedEventArgs e) => SetWordAlignment(TextAlignment.Left);
    private void WordAlignCenter_Click(object sender, RoutedEventArgs e) => SetWordAlignment(TextAlignment.Center);
    private void WordAlignRight_Click(object sender, RoutedEventArgs e) => SetWordAlignment(TextAlignment.Right);

    private void SetWordAlignment(TextAlignment alignment)
    {
        if (RichTextEditor.Selection?.ParagraphFormat != null)
            RichTextEditor.Selection.ParagraphFormat.TextAlignment = alignment;
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

    private void ExcelBold_Click(object sender, RoutedEventArgs e) =>
        WithExcelRange(r => r.CellStyle.Font.Bold = !r.CellStyle.Font.Bold);

    private void ExcelItalic_Click(object sender, RoutedEventArgs e) =>
        WithExcelRange(r => r.CellStyle.Font.Italic = !r.CellStyle.Font.Italic);

    private void ExcelUnderline_Click(object sender, RoutedEventArgs e)
    {
        StatusText.Text = LanguageManager.Get("Hint_Underline");
    }

    private void ExcelAlignLeft_Click(object sender, RoutedEventArgs e) => SetExcelFormat("halign-left");
    private void ExcelAlignCenter_Click(object sender, RoutedEventArgs e) => SetExcelFormat("halign-center");
    private void ExcelAlignRight_Click(object sender, RoutedEventArgs e) => SetExcelFormat("halign-right");
    private void ExcelPercent_Click(object sender, RoutedEventArgs e) => SetExcelFormat("percent");
    private void ExcelComma_Click(object sender, RoutedEventArgs e) => SetExcelFormat("comma");

    private void ExcelMergeCells_Click(object sender, RoutedEventArgs e) =>
        WithExcelRange(r => { if (r.MergeArea != null) r.UnMerge(); else r.Merge(); });

    private void ExcelWrapText_Click(object sender, RoutedEventArgs e) =>
        WithExcelRange(r => r.CellStyle.WrapText = !r.CellStyle.WrapText);

    private void ExcelBorders_Click(object sender, RoutedEventArgs e) =>
        WithExcelRange(r => { r.BorderAround(); r.BorderInside(); StatusText.Text = LanguageManager.Get("Hint_BordersApplied"); });

    private void ExcelFreezePanes_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var ws = Spreadsheet?.ActiveSheet;
            var sel = Spreadsheet?.ActiveGrid?.SelectedRanges?.ActiveRange;
            if (ws == null || sel == null) { StatusText.Text = LanguageManager.Get("Hint_SelectCellToFreeze"); return; }
            if (sel.Top > 1 || sel.Left > 1)
            {
                ws.Range[sel.Top, sel.Left].FreezePanes();
                Spreadsheet?.ActiveGrid?.InvalidateCells();
                StatusText.Text = LanguageManager.Get("Hint_FrozenPanes");
            }
            else
                StatusText.Text = LanguageManager.Get("Hint_SelectCellToFreeze");
        }
        catch (Exception ex) { StatusText.Text = $"{LanguageManager.Get("Error_Title")}: {ex.Message}"; }
    }

    private void ExcelCurrency_Click(object sender, RoutedEventArgs e) => SetExcelFormat("currency");
    private void ExcelDecimalInc_Click(object sender, RoutedEventArgs e) => SetExcelFormat("dec-inc");
    private void ExcelDecimalDec_Click(object sender, RoutedEventArgs e) => SetExcelFormat("dec-dec");

    private void SetExcelFormat(string action) => WithExcelRange(r =>
    {
        switch (action)
        {
            case "percent": r.NumberFormat = "0%"; break;
            case "comma": r.NumberFormat = "#,##0"; break;
            case "currency": r.NumberFormat = "¥#,##0"; break;
            case "dec-inc":
                var curFmt = r.NumberFormat ?? "0";
                var dotIdx = curFmt.IndexOf('.');
                var decimals = dotIdx >= 0 && dotIdx + 1 < curFmt.Length ? curFmt.Length - dotIdx - 1 + 1 : 1;
                r.NumberFormat = "0." + new string('0', decimals);
                break;
            case "dec-dec":
                var curFmt2 = r.NumberFormat ?? "0";
                var dotIdx2 = curFmt2.IndexOf('.');
                var decLen = dotIdx2 >= 0 && dotIdx2 + 1 < curFmt2.Length ? curFmt2.Length - dotIdx2 - 1 : 0;
                r.NumberFormat = decLen > 1 ? "0." + new string('0', decLen - 1) : "0";
                break;
            case "halign-left": r.CellStyle.HorizontalAlignment = (dynamic)1; break;
            case "halign-center": r.CellStyle.HorizontalAlignment = (dynamic)2; break;
            case "halign-right": r.CellStyle.HorizontalAlignment = (dynamic)3; break;
        }
    });

    // ── PPTX Ribbon Commands ──────────────────────────────────────

    private void PptxExtractText_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_currentDocPath)) return;
        var text = ExtractPptxContent(_currentDocPath);
        System.Windows.Clipboard.SetText(text.Length > 5000 ? text[..5000] : text);
        StatusText.Text = LanguageManager.Format("Hint_TextExtracted", text.Length);
    }

    private void WordExportPdf_Click(object sender, RoutedEventArgs e)
    {
        if (RichTextEditor?.Document == null) return;
        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "PDF|*.pdf",
            DefaultExt = ".pdf",
            FileName = Path.GetFileNameWithoutExtension(_currentDocPath) + ".pdf",
        };
        if (dialog.ShowDialog() == true)
        {
            try
            {
                // Syncfusion DocIO でWord→PDF変換
                using var docStream = new MemoryStream();
                RichTextEditor.Save(docStream, Syncfusion.Windows.Controls.RichTextBoxAdv.FormatType.Docx);
                docStream.Position = 0;

                using var wordDoc = new Syncfusion.DocIO.DLS.WordDocument(docStream, Syncfusion.DocIO.FormatType.Docx);
                var converter = new Syncfusion.DocToPDFConverter.DocToPDFConverter();
                var pdfDoc = converter.ConvertToPDF(wordDoc);
                pdfDoc.Save(dialog.FileName);
                pdfDoc.Close(true);
                converter.Dispose();

                StatusText.Text = LanguageManager.Format("Hint_PdfExported", Path.GetFileName(dialog.FileName));
            }
            catch (Exception ex) { StatusText.Text = $"PDF: {ex.Message}"; }
        }
    }

    private void PptxExportPdf_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_currentDocPath)) return;
        var dialog = new Microsoft.Win32.SaveFileDialog { Filter = "PDF|*.pdf", DefaultExt = ".pdf" };
        if (dialog.ShowDialog() == true)
        {
            try
            {
                Services.PptxService.ConvertToPdf(_currentDocPath, dialog.FileName);
                StatusText.Text = LanguageManager.Format("Hint_PdfExported", Path.GetFileName(dialog.FileName));
            }
            catch (Exception ex) { StatusText.Text = $"PDF: {ex.Message}"; }
        }
    }

    private void PptxAddSlide_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_currentDocPath)) return;
        var idx = PptxSlideList.SelectedIndex;
        if (idx < 0) idx = _pptxFullSlides.Count - 1;
        try
        {
            Services.PptxService.AddSlide(_currentDocPath, idx);
            ReloadPptxViewer();
            StatusText.Text = LanguageManager.Get("Pptx_SlideAdded");
        }
        catch (Exception ex) { StatusText.Text = $"{ex.Message}"; }
    }

    private void PptxDuplicateSlide_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_currentDocPath) || PptxSlideList.SelectedIndex < 0) return;
        try
        {
            Services.PptxService.DuplicateSlide(_currentDocPath, PptxSlideList.SelectedIndex);
            var selectIdx = PptxSlideList.SelectedIndex + 1;
            ReloadPptxViewer();
            if (selectIdx < PptxSlideList.Items.Count)
                PptxSlideList.SelectedIndex = selectIdx;
            StatusText.Text = LanguageManager.Get("Pptx_SlideDuplicated");
        }
        catch (Exception ex) { StatusText.Text = $"{ex.Message}"; }
    }

    private void PptxDeleteSlide_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_currentDocPath) || PptxSlideList.SelectedIndex < 0) return;
        if (_pptxFullSlides.Count <= 1)
        {
            StatusText.Text = LanguageManager.Get("Pptx_CannotDeleteLast");
            return;
        }
        try
        {
            var idx = PptxSlideList.SelectedIndex;
            Services.PptxService.DeleteSlide(_currentDocPath, idx);
            ReloadPptxViewer();
            if (idx >= PptxSlideList.Items.Count)
                idx = PptxSlideList.Items.Count - 1;
            if (idx >= 0)
                PptxSlideList.SelectedIndex = idx;
            StatusText.Text = LanguageManager.Get("Pptx_SlideDeleted");
        }
        catch (Exception ex) { StatusText.Text = $"{ex.Message}"; }
    }

    private void PptxMoveSlideUp_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_currentDocPath) || PptxSlideList.SelectedIndex <= 0) return;
        try
        {
            var idx = PptxSlideList.SelectedIndex;
            Services.PptxService.MoveSlide(_currentDocPath, idx, idx - 1);
            ReloadPptxViewer();
            PptxSlideList.SelectedIndex = idx - 1;
        }
        catch (Exception ex) { StatusText.Text = $"{ex.Message}"; }
    }

    private void PptxMoveSlideDown_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_currentDocPath) || PptxSlideList.SelectedIndex < 0) return;
        if (PptxSlideList.SelectedIndex >= _pptxFullSlides.Count - 1) return;
        try
        {
            var idx = PptxSlideList.SelectedIndex;
            Services.PptxService.MoveSlide(_currentDocPath, idx, idx + 1);
            ReloadPptxViewer();
            PptxSlideList.SelectedIndex = idx + 1;
        }
        catch (Exception ex) { StatusText.Text = $"{ex.Message}"; }
    }

    private void ReloadPptxViewer()
    {
        if (string.IsNullOrEmpty(_currentDocPath)) return;
        var selectedIdx = PptxSlideList.SelectedIndex;
        OpenPptxViewer(_currentDocPath, Path.GetFileName(_currentDocPath));
        if (selectedIdx >= 0 && selectedIdx < PptxSlideList.Items.Count)
            PptxSlideList.SelectedIndex = selectedIdx;
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
                try
                {
                    using var stream = File.Create(dialog.FileName);
                    RichTextEditor.Save(stream, FormatType.Docx);
                    StatusText.Text = LanguageManager.Format("Hint_WordExported", Path.GetFileName(dialog.FileName));
                }
                catch (Exception ex)
                {
                    StatusText.Text = $"{LanguageManager.Get("Error_Title")}: {ex.Message}";
                }
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
                try
                {
                    Spreadsheet.SaveAs(dialog.FileName);
                    StatusText.Text = LanguageManager.Format("Hint_ExcelExported", Path.GetFileName(dialog.FileName));
                }
                catch (Exception ex)
                {
                    StatusText.Text = $"{LanguageManager.Get("Error_Title")}: {ex.Message}";
                }
            }
        }
        else
            StatusText.Text = LanguageManager.Get("Hint_ExportExcelFirst");
    }

    private void HideAllBackstages()
    {
        try { MainRibbon?.HideBackStage(); } catch { /* backstage not active */ }
        try { WordRibbon?.HideBackStage(); } catch { /* backstage not active */ }
        try { ExcelRibbon?.HideBackStage(); } catch { /* backstage not active */ }
        try { PptxRibbon?.HideBackStage(); } catch { /* backstage not active */ }
        try { PdfRibbon?.HideBackStage(); } catch { /* backstage not active */ }
    }

    // ── Backstage / Help ──────────────────────────────────────────

    private void RibbonHelp_Click(object sender, RoutedEventArgs e)
    {
        Views.HelpWindow.ShowSection(this, "overview");
        StatusText.Text = LanguageManager.Get("Menu_Help");
    }

    private void BackstageOpen_Click(object sender, RoutedEventArgs e)
    {
        HideAllBackstages();
        if (DataContext is ViewModels.MainViewModel vm)
            vm.OpenDocumentCommand.Execute(null);
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
        new Views.SettingsWindow(_chatVm.AiService) { Owner = this }.ShowDialog();
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

        // Defer editor cleanup to allow backstage animation to complete,
        // preventing Syncfusion InvalidOperationException on ribbon switch.
        var pathToClose = _currentDocPath;
        Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background, () =>
        {
            if (!string.IsNullOrEmpty(pathToClose))
                CloseTab(pathToClose);
            else
                CloseAllEditors();
        });
    }
}
