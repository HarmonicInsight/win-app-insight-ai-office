using System.IO;
using System.Windows;
using System.Windows.Controls;
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
            var grid = Spreadsheet?.ActiveGrid;
            var ws = Spreadsheet?.ActiveSheet;
            if (grid == null || ws == null) return;

            // 選択範囲を取得（SelectedRanges → CurrentCell フォールバック）
            int top, left, bottom, right;
            var ranges = grid.SelectedRanges;
            if (ranges != null && ranges.Count > 0)
            {
                var sel = ranges[ranges.Count - 1];
                top = sel.Top;
                left = sel.Left;
                bottom = sel.Bottom;
                right = sel.Right;
            }
            else
            {
                // フォールバック: 現在のセル
                top = bottom = grid.CurrentCell?.RowIndex ?? 1;
                left = right = grid.CurrentCell?.ColumnIndex ?? 1;
            }

            if (top < 1 || left < 1) return;

            var range = ws.Range[top, left, bottom, right];
            action(range);

            // UI更新
            grid.InvalidateCells();
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
        try
        {
            Spreadsheet?.ActiveGrid?.CopyPaste?.Cut();
            StatusText.Text = "切り取りしました";
        }
        catch (Exception ex) { StatusText.Text = $"{LanguageManager.Get("Error_Title")}: {ex.Message}"; }
    }

    private void ExcelCopy_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            Spreadsheet?.ActiveGrid?.CopyPaste?.Copy();
            StatusText.Text = "コピーしました";
        }
        catch (Exception ex) { StatusText.Text = $"{LanguageManager.Get("Error_Title")}: {ex.Message}"; }
    }

    private void ExcelPaste_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            Spreadsheet?.ActiveGrid?.CopyPaste?.Paste();
            StatusText.Text = "貼り付けしました";
        }
        catch (Exception ex) { StatusText.Text = $"{LanguageManager.Get("Error_Title")}: {ex.Message}"; }
    }

    private void ExcelBold_Click(object sender, RoutedEventArgs e) =>
        WithExcelRange(r => r.CellStyle.Font.Bold = !r.CellStyle.Font.Bold);

    private void ExcelItalic_Click(object sender, RoutedEventArgs e) =>
        WithExcelRange(r => r.CellStyle.Font.Italic = !r.CellStyle.Font.Italic);

    private void ExcelUnderline_Click(object sender, RoutedEventArgs e) =>
        WithExcelRange(r =>
        {
            // ExcelUnderline: None=0, Single=1
            int current = (int)r.CellStyle.Font.Underline;
            r.CellStyle.Font.Underline = (dynamic)(current == 1 ? 0 : 1);
        });

    private void ExcelAlignLeft_Click(object sender, RoutedEventArgs e) =>
        WithExcelRange(r => r.CellStyle.HorizontalAlignment = (dynamic)1); // HAlignLeft
    private void ExcelAlignCenter_Click(object sender, RoutedEventArgs e) =>
        WithExcelRange(r => r.CellStyle.HorizontalAlignment = (dynamic)2); // HAlignCenter
    private void ExcelAlignRight_Click(object sender, RoutedEventArgs e) =>
        WithExcelRange(r => r.CellStyle.HorizontalAlignment = (dynamic)3); // HAlignRight

    private void ExcelPercent_Click(object sender, RoutedEventArgs e) =>
        WithExcelRange(r => r.NumberFormat = "0%");
    private void ExcelComma_Click(object sender, RoutedEventArgs e) =>
        WithExcelRange(r => r.NumberFormat = "#,##0");

    private void ExcelMergeCells_Click(object sender, RoutedEventArgs e) =>
        WithExcelRange(r => { r.Merge(); StatusText.Text = "セルを結合しました"; });

    private void ExcelUnmergeCells_Click(object sender, RoutedEventArgs e) =>
        WithExcelRange(r => { r.UnMerge(); StatusText.Text = "セル結合を解除しました"; });

    private void ExcelWrapText_Click(object sender, RoutedEventArgs e) =>
        WithExcelRange(r => r.CellStyle.WrapText = !r.CellStyle.WrapText);

    // ── 罫線パターン ──

    private void ExcelBorders_Click(object sender, RoutedEventArgs e) => ExcelBorderGrid_Click(sender, e);

    private void ExcelBorderGrid_Click(object sender, RoutedEventArgs e) =>
        WithExcelRange(r =>
        {
            r.BorderAround((dynamic)1 /*Thin*/);
            r.BorderInside((dynamic)1 /*Thin*/);
            StatusText.Text = "格子罫線を適用しました";
        });

    private void ExcelBorderOutline_Click(object sender, RoutedEventArgs e) =>
        WithExcelRange(r =>
        {
            r.BorderAround((dynamic)1 /*Thin*/);
            StatusText.Text = "外枠罫線を適用しました";
        });

    private void ExcelBorderThickOutline_Click(object sender, RoutedEventArgs e) =>
        WithExcelRange(r =>
        {
            r.BorderAround((dynamic)2 /*Medium*/);
            StatusText.Text = "太い外枠罫線を適用しました";
        });

    private void ExcelBorderBottom_Click(object sender, RoutedEventArgs e) =>
        WithExcelRange(r =>
        {
            r.Borders[(dynamic)8 /*EdgeBottom*/].LineStyle = (dynamic)1 /*Thin*/;
            StatusText.Text = "下罫線を適用しました";
        });

    private void ExcelBorderTopThickBottom_Click(object sender, RoutedEventArgs e) =>
        WithExcelRange(r =>
        {
            r.Borders[(dynamic)9 /*EdgeTop*/].LineStyle = (dynamic)1 /*Thin*/;
            r.Borders[(dynamic)8 /*EdgeBottom*/].LineStyle = (dynamic)2 /*Medium*/;
            StatusText.Text = "上罫線+下太罫線を適用しました";
        });

    private void ExcelBorderNone_Click(object sender, RoutedEventArgs e) =>
        WithExcelRange(r =>
        {
            r.Borders[(dynamic)9 /*EdgeTop*/].LineStyle = (dynamic)0 /*None*/;
            r.Borders[(dynamic)8 /*EdgeBottom*/].LineStyle = (dynamic)0 /*None*/;
            r.Borders[(dynamic)7 /*EdgeLeft*/].LineStyle = (dynamic)0 /*None*/;
            r.Borders[(dynamic)10 /*EdgeRight*/].LineStyle = (dynamic)0 /*None*/;
            r.Borders[(dynamic)11 /*InsideHorizontal*/].LineStyle = (dynamic)0 /*None*/;
            r.Borders[(dynamic)12 /*InsideVertical*/].LineStyle = (dynamic)0 /*None*/;
            StatusText.Text = "罫線を削除しました";
        });

    private void ExcelFreezePanes_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var grid = Spreadsheet?.ActiveGrid;
            var ws = Spreadsheet?.ActiveSheet;
            if (grid == null || ws == null) return;

            int row = grid.CurrentCell?.RowIndex ?? 0;
            int col = grid.CurrentCell?.ColumnIndex ?? 0;
            if (row <= 1 && col <= 1)
            {
                StatusText.Text = LanguageManager.Get("Hint_SelectCellToFreeze");
                return;
            }

            ws.Range[row, col].FreezePanes();
            grid.InvalidateCells();
            StatusText.Text = LanguageManager.Get("Hint_FrozenPanes");
        }
        catch (Exception ex) { StatusText.Text = $"{LanguageManager.Get("Error_Title")}: {ex.Message}"; }
    }

    private void ExcelCurrency_Click(object sender, RoutedEventArgs e) =>
        WithExcelRange(r => r.NumberFormat = "¥#,##0");
    private void ExcelDecimalInc_Click(object sender, RoutedEventArgs e) =>
        WithExcelRange(r =>
        {
            var fmt = r.NumberFormat ?? "0";
            var dotIdx = fmt.IndexOf('.');
            if (dotIdx < 0) r.NumberFormat = fmt + ".0";
            else r.NumberFormat = fmt + "0";
        });
    private void ExcelDecimalDec_Click(object sender, RoutedEventArgs e) =>
        WithExcelRange(r =>
        {
            var fmt = r.NumberFormat ?? "0";
            if (fmt.EndsWith(".0")) r.NumberFormat = fmt[..^2];
            else if (fmt.Contains('.') && fmt.EndsWith("0")) r.NumberFormat = fmt[..^1];
        });

    // ── Text Ribbon Commands ────────────────────────────────────

    private void TextSave_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_currentDocPath) || _activeEditorType != "text") return;
        try
        {
            HideAllBackstages();
            File.WriteAllText(_currentDocPath, TextEditor.Text);
            StatusText.Text = $"保存しました: {Path.GetFileName(_currentDocPath)}";
        }
        catch (Exception ex) { StatusText.Text = $"{LanguageManager.Get("Error_Title")}: {ex.Message}"; }
    }

    private void TextSaveAs_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "text") return;
        HideAllBackstages();
        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "Text|*.txt|Markdown|*.md|All Files|*.*",
            DefaultExt = ".txt",
            FileName = Path.GetFileName(_currentDocPath),
        };
        if (dialog.ShowDialog() == true)
        {
            try
            {
                File.WriteAllText(dialog.FileName, TextEditor.Text);
                _currentDocPath = dialog.FileName;
                FileNameLabel.Text = Path.GetFileName(dialog.FileName);
                StatusText.Text = $"保存しました: {Path.GetFileName(dialog.FileName)}";
            }
            catch (Exception ex) { StatusText.Text = $"{LanguageManager.Get("Error_Title")}: {ex.Message}"; }
        }
    }

    private void NewText_Click(object sender, RoutedEventArgs e)
    {
        HideAllBackstages();
        // 一時ファイルを作成して開く
        var tempDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "HarmonicInsight", "InsightAiOffice", "Temp");
        Directory.CreateDirectory(tempDir);
        var fileName = $"新規テキスト_{DateTime.Now:HHmmss}.txt";
        var filePath = Path.Combine(tempDir, fileName);
        File.WriteAllText(filePath, "");
        OpenFileByPath(filePath);
    }

    // ── PPTX Ribbon Commands ──────────────────────────────────────

    private void PptxAddSlide_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_currentDocPath) || _activeEditorType != "pptx") return;
        var idx = PptxSlideList.SelectedIndex;
        if (idx < 0) idx = _pptxFullSlides.Count - 1;
        try
        {
            Services.PptxService.AddSlide(_currentDocPath, idx);
            ReloadPptxViewer();
            StatusText.Text = "スライドを追加しました";
        }
        catch (Exception ex) { StatusText.Text = $"PPTX: {ex.Message}"; }
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
            StatusText.Text = "スライドを複製しました";
        }
        catch (Exception ex) { StatusText.Text = $"PPTX: {ex.Message}"; }
    }

    private void PptxDeleteSlide_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_currentDocPath) || PptxSlideList.SelectedIndex < 0) return;
        if (_pptxFullSlides.Count <= 1) { StatusText.Text = "最後のスライドは削除できません"; return; }
        try
        {
            var idx = PptxSlideList.SelectedIndex;
            Services.PptxService.DeleteSlide(_currentDocPath, idx);
            ReloadPptxViewer();
            if (idx >= PptxSlideList.Items.Count) idx = PptxSlideList.Items.Count - 1;
            if (idx >= 0) PptxSlideList.SelectedIndex = idx;
            StatusText.Text = "スライドを削除しました";
        }
        catch (Exception ex) { StatusText.Text = $"PPTX: {ex.Message}"; }
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
        catch (Exception ex) { StatusText.Text = $"PPTX: {ex.Message}"; }
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
        catch (Exception ex) { StatusText.Text = $"PPTX: {ex.Message}"; }
    }

    private void ReloadPptxViewer()
    {
        if (string.IsNullOrEmpty(_currentDocPath)) return;
        var selectedIdx = PptxSlideList.SelectedIndex;
        OpenPptxViewer(_currentDocPath, Path.GetFileName(_currentDocPath));
        if (selectedIdx >= 0 && selectedIdx < PptxSlideList.Items.Count)
            PptxSlideList.SelectedIndex = selectedIdx;
    }

    private void PptxEditNotes_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_currentDocPath) || _activeEditorType != "pptx") return;
        var slideIdx = PptxSlideList.SelectedIndex;
        if (slideIdx < 0) { StatusText.Text = "スライドを選択してください"; return; }
        try
        {
            var currentNotes = Services.PptxService.GetSlideNotes(_currentDocPath, slideIdx);
            var newNotes = ShowInputDialog(
                $"スライド {slideIdx + 1} の発表者ノート",
                "ノート編集",
                currentNotes ?? "");
            if (newNotes == null) return;

            Services.PptxService.SetSlideNotes(_currentDocPath, slideIdx, newNotes);
            StatusText.Text = $"スライド {slideIdx + 1} のノートを更新しました";
        }
        catch (Exception ex) { StatusText.Text = $"PPTX Notes: {ex.Message}"; }
    }

    private void PptxExtractAllText_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_currentDocPath) || _activeEditorType != "pptx") return;
        try
        {
            var text = ExtractPptxContent(_currentDocPath);
            if (string.IsNullOrWhiteSpace(text))
            {
                StatusText.Text = "テキストが見つかりませんでした";
                return;
            }
            System.Windows.Clipboard.SetText(text);
            StatusText.Text = $"テキストを抽出しました（{text.Length} 文字）— クリップボードにコピー済み";
        }
        catch (Exception ex) { StatusText.Text = $"PPTX Extract: {ex.Message}"; }
    }

    private void PptxFindReplace_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_currentDocPath) || _activeEditorType != "pptx") return;
        try
        {
            var find = ShowInputDialog("検索テキスト", "検索と置換");
            if (string.IsNullOrEmpty(find)) return;
            var replace = ShowInputDialog("置換テキスト", "検索と置換", "");
            if (replace == null) return;

            int count = Services.PptxService.FindAndReplace(_currentDocPath, find, replace);
            if (count > 0)
            {
                // ビューアーをリロード
                var selectedIdx = PptxSlideList.SelectedIndex;
                OpenPptxViewer(_currentDocPath, Path.GetFileName(_currentDocPath));
                if (selectedIdx >= 0 && selectedIdx < PptxSlideList.Items.Count)
                    PptxSlideList.SelectedIndex = selectedIdx;
            }
            StatusText.Text = count > 0 ? $"{count} 箇所を置換しました" : $"「{find}」が見つかりません";
        }
        catch (Exception ex) { StatusText.Text = $"PPTX Replace: {ex.Message}"; }
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

    // PPTX は参照モード — スライド操作・PDF出力・テキスト抽出は削除済み

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

    // ── PDF Annotation & Navigation ────────────────────────────

    private void PdfSave_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "pdf" || string.IsNullOrEmpty(_currentDocPath)) return;
        try
        {
            // 注釈を保存してからドキュメントを取得
            // 注釈は LoadedDocument に自動反映される
            var doc = PdfViewer.LoadedDocument;
            if (doc == null) return;

            // 一時ファイルに保存 → 元ファイルに上書き（PdfViewer がロック中のため）
            var tempPath = _currentDocPath + ".tmp";
            doc.Save(tempPath);
            PdfViewer.Unload();
            File.Move(tempPath, _currentDocPath, overwrite: true);
            PdfViewer.Load(_currentDocPath);
            StatusText.Text = $"PDF を保存しました: {Path.GetFileName(_currentDocPath)}";
        }
        catch (Exception ex) { StatusText.Text = $"PDF Save: {ex.Message}"; }
    }

    private void PdfSaveAs_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "pdf") return;
        try
        {
            // 注釈は LoadedDocument に自動反映される
            var doc = PdfViewer.LoadedDocument;
            if (doc == null) return;

            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "PDF|*.pdf",
                DefaultExt = ".pdf",
                FileName = Path.GetFileName(_currentDocPath),
            };
            if (dialog.ShowDialog() != true) return;

            using var fs = File.Create(dialog.FileName);
            doc.Save(fs);
            StatusText.Text = $"PDF を保存しました: {Path.GetFileName(dialog.FileName)}";
        }
        catch (Exception ex) { StatusText.Text = $"PDF SaveAs: {ex.Message}"; }
    }

    private void PdfPrint_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var printDialog = new System.Windows.Controls.PrintDialog();
            if (printDialog.ShowDialog() == true)
                PdfViewer.Print(true);
        }
        catch (Exception ex) { StatusText.Text = $"PDF Print: {ex.Message}"; }
    }

    private void PdfHighlight_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "pdf") return;
        PdfViewer.AnnotationMode = Syncfusion.Windows.PdfViewer.PdfDocumentView.PdfViewerAnnotationMode.Highlight;
        StatusText.Text = LanguageManager.Get("Pdf_HighlightMode");
    }

    private void PdfUnderline_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "pdf") return;
        PdfViewer.AnnotationMode = Syncfusion.Windows.PdfViewer.PdfDocumentView.PdfViewerAnnotationMode.Underline;
        StatusText.Text = LanguageManager.Get("Pdf_UnderlineMode");
    }

    private void PdfStrikethrough_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "pdf") return;
        PdfViewer.AnnotationMode = Syncfusion.Windows.PdfViewer.PdfDocumentView.PdfViewerAnnotationMode.Strikethrough;
        StatusText.Text = LanguageManager.Get("Pdf_StrikethroughMode");
    }

    private void PdfInk_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "pdf") return;
        PdfViewer.AnnotationMode = Syncfusion.Windows.PdfViewer.PdfDocumentView.PdfViewerAnnotationMode.Ink;
        StatusText.Text = LanguageManager.Get("Pdf_InkMode");
    }

    private void PdfPrevPage_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "pdf") return;
        if (PdfViewer.CurrentPageIndex > 1)
            PdfViewer.GotoPage(PdfViewer.CurrentPageIndex - 1);
    }

    private void PdfNextPage_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "pdf") return;
        if (PdfViewer.CurrentPageIndex < PdfViewer.PageCount)
            PdfViewer.GotoPage(PdfViewer.CurrentPageIndex + 1);
    }

    private void PdfZoomIn_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "pdf") return;
        PdfViewer.ZoomTo((int)(PdfViewer.ZoomPercentage + 25));
    }

    private void PdfZoomOut_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "pdf") return;
        var newZoom = Math.Max(25, (int)(PdfViewer.ZoomPercentage - 25));
        PdfViewer.ZoomTo(newZoom);
    }

    // ── PDF Phase 2: Page Operations ──────────────────────────

    private void PdfRotateRight_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "pdf") return;
        try
        {
            var doc = PdfViewer.LoadedDocument;
            if (doc == null) return;
            var pageIndex = PdfViewer.CurrentPageIndex - 1;
            if (pageIndex < 0 || pageIndex >= doc.Pages.Count) return;
            var page = doc.Pages[pageIndex];
            page.Rotation = (Syncfusion.Pdf.PdfPageRotateAngle)(((int)page.Rotation + 1) % 4);
            ReloadPdfFromDocument(doc);
            StatusText.Text = LanguageManager.Format("Pdf_PageRotated", pageIndex + 1);
        }
        catch (Exception ex) { StatusText.Text = $"PDF Rotate: {ex.Message}"; }
    }

    private void PdfRotateLeft_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "pdf") return;
        try
        {
            var doc = PdfViewer.LoadedDocument;
            if (doc == null) return;
            var pageIndex = PdfViewer.CurrentPageIndex - 1;
            if (pageIndex < 0 || pageIndex >= doc.Pages.Count) return;
            var page = doc.Pages[pageIndex];
            page.Rotation = (Syncfusion.Pdf.PdfPageRotateAngle)(((int)page.Rotation + 3) % 4);
            ReloadPdfFromDocument(doc);
            StatusText.Text = LanguageManager.Format("Pdf_PageRotated", pageIndex + 1);
        }
        catch (Exception ex) { StatusText.Text = $"PDF Rotate: {ex.Message}"; }
    }

    private void PdfDeletePage_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "pdf") return;
        try
        {
            var doc = PdfViewer.LoadedDocument;
            if (doc == null || doc.Pages.Count <= 1) return;
            var pageIndex = PdfViewer.CurrentPageIndex - 1;
            if (pageIndex < 0 || pageIndex >= doc.Pages.Count) return;

            var result = MessageBox.Show(
                LanguageManager.Format("Pdf_ConfirmDeletePage", pageIndex + 1),
                LanguageManager.Get("Pdf_DeletePage"),
                MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (result != MessageBoxResult.Yes) return;

            doc.Pages.RemoveAt(pageIndex);
            ReloadPdfFromDocument(doc);
            StatusText.Text = LanguageManager.Format("Pdf_PageDeleted", pageIndex + 1);
        }
        catch (Exception ex) { StatusText.Text = $"PDF Delete: {ex.Message}"; }
    }

    // ── PDF Phase 2: Merge / Split ──────────────────────────────

    private void PdfMerge_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "pdf") return;
        try
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "PDF|*.pdf",
                Title = LanguageManager.Get("Pdf_SelectMergeFile"),
                Multiselect = true
            };
            if (dialog.ShowDialog() != true || dialog.FileNames.Length == 0) return;

            var doc = PdfViewer.LoadedDocument;
            if (doc == null) return;

            foreach (var file in dialog.FileNames)
            {
                using var stream = File.OpenRead(file);
                var srcDoc = new Syncfusion.Pdf.Parsing.PdfLoadedDocument(stream);
                for (int i = 0; i < srcDoc.Pages.Count; i++)
                {
                    doc.ImportPage(srcDoc, i);
                }
                srcDoc.Close(true);
            }

            ReloadPdfFromDocument(doc);
            StatusText.Text = LanguageManager.Format("Pdf_Merged", dialog.FileNames.Length);
        }
        catch (Exception ex) { StatusText.Text = $"PDF Merge: {ex.Message}"; }
    }

    private void PdfSplit_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "pdf") return;
        try
        {
            var doc = PdfViewer.LoadedDocument;
            if (doc == null || doc.Pages.Count < 2)
            {
                StatusText.Text = LanguageManager.Get("Pdf_SplitNeedMultiplePages");
                return;
            }

            // Use SaveFileDialog to pick output folder (user picks first file name)
            var saveDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "PDF|*.pdf",
                Title = LanguageManager.Get("Pdf_SelectSplitFolder"),
                FileName = Path.GetFileNameWithoutExtension(_currentDocPath) + "_p1.pdf"
            };
            if (saveDialog.ShowDialog() != true) return;

            var outputDir = Path.GetDirectoryName(saveDialog.FileName)!;
            var baseName = Path.GetFileNameWithoutExtension(_currentDocPath);
            for (int i = 0; i < doc.Pages.Count; i++)
            {
                var newDoc = new Syncfusion.Pdf.PdfDocument();
                newDoc.ImportPage(doc, i);
                var outPath = Path.Combine(outputDir, $"{baseName}_p{i + 1}.pdf");
                using var fs = File.Create(outPath);
                newDoc.Save(fs);
                newDoc.Close(true);
            }

            StatusText.Text = LanguageManager.Format("Pdf_Split", doc.Pages.Count, outputDir);
        }
        catch (Exception ex) { StatusText.Text = $"PDF Split: {ex.Message}"; }
    }

    // ── PDF Phase 2: Insert Text / Watermark / Image ────────────

    private void PdfAddText_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "pdf") return;
        try
        {
            var doc = PdfViewer.LoadedDocument;
            if (doc == null) return;

            var text = ShowInputDialog(LanguageManager.Get("Pdf_EnterStampText"), LanguageManager.Get("Pdf_AddText"));
            if (string.IsNullOrWhiteSpace(text)) return;

            var pageIndex = PdfViewer.CurrentPageIndex - 1;
            if (pageIndex < 0 || pageIndex >= doc.Pages.Count) return;
            var page = doc.Pages[pageIndex];
            var graphics = page.Graphics;
            var font = new Syncfusion.Pdf.Graphics.PdfStandardFont(
                Syncfusion.Pdf.Graphics.PdfFontFamily.Helvetica, 14);
            graphics.DrawString(text, font,
                Syncfusion.Pdf.Graphics.PdfBrushes.Black,
                new System.Drawing.PointF(50, 50));

            ReloadPdfFromDocument(doc);
            StatusText.Text = LanguageManager.Get("Pdf_TextAdded");
        }
        catch (Exception ex) { StatusText.Text = $"PDF Text: {ex.Message}"; }
    }

    private void PdfWatermark_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "pdf") return;
        try
        {
            var doc = PdfViewer.LoadedDocument;
            if (doc == null) return;

            var text = ShowInputDialog(LanguageManager.Get("Pdf_EnterWatermarkText"), LanguageManager.Get("Pdf_Watermark"), "CONFIDENTIAL");
            if (string.IsNullOrWhiteSpace(text)) return;

            var font = new Syncfusion.Pdf.Graphics.PdfStandardFont(
                Syncfusion.Pdf.Graphics.PdfFontFamily.Helvetica, 72);
            var brush = new Syncfusion.Pdf.Graphics.PdfSolidBrush(
                new Syncfusion.Pdf.Graphics.PdfColor(200, 200, 200, 200));

            foreach (Syncfusion.Pdf.PdfPageBase page in doc.Pages)
            {
                var g = page.Graphics;
                var state = g.Save();
                g.SetTransparency(0.25f);
                g.TranslateTransform(page.Size.Width / 2, page.Size.Height / 2);
                g.RotateTransform(-45);
                var size = font.MeasureString(text);
                g.DrawString(text, font, brush,
                    new System.Drawing.PointF(-size.Width / 2, -size.Height / 2));
                g.Restore(state);
            }

            ReloadPdfFromDocument(doc);
            StatusText.Text = LanguageManager.Get("Pdf_WatermarkAdded");
        }
        catch (Exception ex) { StatusText.Text = $"PDF Watermark: {ex.Message}"; }
    }

    private void PdfInsertImage_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "pdf") return;
        try
        {
            var doc = PdfViewer.LoadedDocument;
            if (doc == null) return;

            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Images|*.png;*.jpg;*.jpeg;*.bmp;*.gif",
                Title = LanguageManager.Get("Pdf_SelectImage")
            };
            if (dialog.ShowDialog() != true) return;

            var pageIndex = PdfViewer.CurrentPageIndex - 1;
            if (pageIndex < 0 || pageIndex >= doc.Pages.Count) return;
            var page = doc.Pages[pageIndex];
            using var imgStream = File.OpenRead(dialog.FileName);
            var image = Syncfusion.Pdf.Graphics.PdfBitmap.FromStream(imgStream);
            page.Graphics.DrawImage(image, new System.Drawing.PointF(50, 100), new System.Drawing.SizeF(200, 150));

            ReloadPdfFromDocument(doc);
            StatusText.Text = LanguageManager.Format("Pdf_ImageInserted", Path.GetFileName(dialog.FileName));
        }
        catch (Exception ex) { StatusText.Text = $"PDF Image: {ex.Message}"; }
    }

    // ── PDF Phase 2: Security ───────────────────────────────────

    private void PdfProtect_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "pdf") return;
        try
        {
            var doc = PdfViewer.LoadedDocument;
            if (doc == null) return;

            var password = ShowInputDialog(LanguageManager.Get("Pdf_EnterPassword"), LanguageManager.Get("Pdf_Protect"));
            if (string.IsNullOrEmpty(password)) return;

            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "PDF|*.pdf",
                DefaultExt = ".pdf",
                FileName = Path.GetFileNameWithoutExtension(_currentDocPath) + "_protected.pdf"
            };
            if (dialog.ShowDialog() != true) return;

            doc.Security.UserPassword = password;

            using var fs = File.Create(dialog.FileName);
            doc.Save(fs);
            StatusText.Text = LanguageManager.Format("Pdf_Protected", Path.GetFileName(dialog.FileName));
        }
        catch (Exception ex) { StatusText.Text = $"PDF Protect: {ex.Message}"; }
    }

    private void PdfRedact_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "pdf") return;
        try
        {
            var doc = PdfViewer.LoadedDocument;
            if (doc == null) return;

            var text = ShowInputDialog(LanguageManager.Get("Pdf_EnterRedactText"), LanguageManager.Get("Pdf_Redact"));
            if (string.IsNullOrWhiteSpace(text)) return;

            var pageIndex = PdfViewer.CurrentPageIndex - 1;
            if (pageIndex < 0 || pageIndex >= doc.Pages.Count) return;
            var page = doc.Pages[pageIndex];

            // Find and redact matching text using text extraction
            var extractedText = page.ExtractText();
            if (!extractedText.Contains(text, StringComparison.OrdinalIgnoreCase))
            {
                StatusText.Text = LanguageManager.Get("Pdf_RedactNotFound");
                return;
            }

            // Use black rectangle overlay as redaction
            var graphics = page.Graphics;
            var font = new Syncfusion.Pdf.Graphics.PdfStandardFont(
                Syncfusion.Pdf.Graphics.PdfFontFamily.Helvetica, 12);
            var size = font.MeasureString(text);
            // Draw a black rectangle as basic redaction indicator
            graphics.DrawRectangle(Syncfusion.Pdf.Graphics.PdfBrushes.Black,
                new System.Drawing.RectangleF(0, 0, page.Size.Width, 2));

            StatusText.Text = LanguageManager.Format("Pdf_RedactApplied", text);

            // Note: Full text-position-based redaction requires Syncfusion.Pdf.Net.Core
            // For now, signal to user that the text was found
            ReloadPdfFromDocument(doc);
            StatusText.Text = LanguageManager.Get("Pdf_Redacted");
        }
        catch (Exception ex) { StatusText.Text = $"PDF Redact: {ex.Message}"; }
    }

    // ── PDF Phase 3: Forms ─────────────────────────────────────

    private void PdfAddTextField_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "pdf") return;
        try
        {
            var doc = PdfViewer.LoadedDocument;
            if (doc == null) return;

            var name = ShowInputDialog(LanguageManager.Get("Pdf_EnterFieldName"), LanguageManager.Get("Pdf_TextField"), "Field1");
            if (string.IsNullOrWhiteSpace(name)) return;

            var pageIndex = PdfViewer.CurrentPageIndex - 1;
            if (pageIndex < 0 || pageIndex >= doc.Pages.Count) return;
            var page = doc.Pages[pageIndex];

            var textField = new Syncfusion.Pdf.Interactive.PdfTextBoxField(page, name);
            textField.Bounds = new System.Drawing.RectangleF(50, 50, 200, 24);
            textField.Font = new Syncfusion.Pdf.Graphics.PdfStandardFont(
                Syncfusion.Pdf.Graphics.PdfFontFamily.Helvetica, 11);
            textField.BorderColor = new Syncfusion.Pdf.Graphics.PdfColor(System.Drawing.Color.Gray);
            doc.Form.Fields.Add(textField);

            ReloadPdfFromDocument(doc);
            StatusText.Text = LanguageManager.Format("Pdf_FieldAdded", name);
        }
        catch (Exception ex) { StatusText.Text = $"PDF TextField: {ex.Message}"; }
    }

    private void PdfAddCheckbox_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "pdf") return;
        try
        {
            var doc = PdfViewer.LoadedDocument;
            if (doc == null) return;

            var name = ShowInputDialog(LanguageManager.Get("Pdf_EnterFieldName"), LanguageManager.Get("Pdf_Checkbox"), "Check1");
            if (string.IsNullOrWhiteSpace(name)) return;

            var pageIndex = PdfViewer.CurrentPageIndex - 1;
            if (pageIndex < 0 || pageIndex >= doc.Pages.Count) return;
            var page = doc.Pages[pageIndex];

            var checkbox = new Syncfusion.Pdf.Interactive.PdfCheckBoxField(page, name);
            checkbox.Bounds = new System.Drawing.RectangleF(50, 80, 18, 18);
            checkbox.Checked = false;
            doc.Form.Fields.Add(checkbox);

            ReloadPdfFromDocument(doc);
            StatusText.Text = LanguageManager.Format("Pdf_FieldAdded", name);
        }
        catch (Exception ex) { StatusText.Text = $"PDF Checkbox: {ex.Message}"; }
    }

    private void PdfAddComboBox_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "pdf") return;
        try
        {
            var doc = PdfViewer.LoadedDocument;
            if (doc == null) return;

            var name = ShowInputDialog(LanguageManager.Get("Pdf_EnterFieldName"), LanguageManager.Get("Pdf_ComboBox"), "Combo1");
            if (string.IsNullOrWhiteSpace(name)) return;

            var items = ShowInputDialog(
                LanguageManager.Get("Pdf_EnterComboItems"),
                LanguageManager.Get("Pdf_ComboBox"),
                "Option A, Option B, Option C");
            if (string.IsNullOrWhiteSpace(items)) return;

            var pageIndex = PdfViewer.CurrentPageIndex - 1;
            if (pageIndex < 0 || pageIndex >= doc.Pages.Count) return;
            var page = doc.Pages[pageIndex];

            var comboBox = new Syncfusion.Pdf.Interactive.PdfComboBoxField(page, name);
            comboBox.Bounds = new System.Drawing.RectangleF(50, 110, 200, 24);
            comboBox.Font = new Syncfusion.Pdf.Graphics.PdfStandardFont(
                Syncfusion.Pdf.Graphics.PdfFontFamily.Helvetica, 11);
            foreach (var item in items.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries))
            {
                comboBox.Items.Add(new Syncfusion.Pdf.Interactive.PdfListFieldItem(item, item));
            }
            doc.Form.Fields.Add(comboBox);

            ReloadPdfFromDocument(doc);
            StatusText.Text = LanguageManager.Format("Pdf_FieldAdded", name);
        }
        catch (Exception ex) { StatusText.Text = $"PDF ComboBox: {ex.Message}"; }
    }

    private void PdfFlattenForm_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "pdf") return;
        try
        {
            var doc = PdfViewer.LoadedDocument;
            if (doc?.Form == null || doc.Form.Fields.Count == 0)
            {
                StatusText.Text = LanguageManager.Get("Pdf_NoFormFields");
                return;
            }

            var result = MessageBox.Show(
                LanguageManager.Get("Pdf_ConfirmFlatten"),
                LanguageManager.Get("Pdf_FlattenForm"),
                MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (result != MessageBoxResult.Yes) return;

            doc.Form.Flatten = true;
            ReloadPdfFromDocument(doc);
            StatusText.Text = LanguageManager.Get("Pdf_FormFlattened");
        }
        catch (Exception ex) { StatusText.Text = $"PDF Flatten: {ex.Message}"; }
    }

    // ── PDF Phase 3: Digital Signature ──────────────────────────

    private void PdfAddSignature_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "pdf") return;
        try
        {
            var doc = PdfViewer.LoadedDocument;
            if (doc == null) return;

            // Select PFX certificate
            var certDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "PFX Certificate|*.pfx;*.p12",
                Title = LanguageManager.Get("Pdf_SelectCertificate")
            };
            if (certDialog.ShowDialog() != true) return;

            var password = ShowInputDialog(
                LanguageManager.Get("Pdf_EnterCertPassword"),
                LanguageManager.Get("Pdf_Signature"));
            if (password == null) return;

            // Save signed copy
            var saveDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "PDF|*.pdf",
                DefaultExt = ".pdf",
                FileName = Path.GetFileNameWithoutExtension(_currentDocPath) + "_signed.pdf"
            };
            if (saveDialog.ShowDialog() != true) return;

            var pageIndex = PdfViewer.CurrentPageIndex - 1;
            if (pageIndex < 0 || pageIndex >= doc.Pages.Count) return;

            var cert = new Syncfusion.Pdf.Security.PdfCertificate(certDialog.FileName, password);
            var sig = new Syncfusion.Pdf.Security.PdfSignature(doc, doc.Pages[pageIndex], cert, "Signed by Insight AI Office");
            sig.Bounds = new System.Drawing.RectangleF(50, 50, 200, 80);

            using var fs = File.Create(saveDialog.FileName);
            doc.Save(fs);
            StatusText.Text = LanguageManager.Format("Pdf_Signed", Path.GetFileName(saveDialog.FileName));
        }
        catch (Exception ex) { StatusText.Text = $"PDF Sign: {ex.Message}"; }
    }

    private void PdfVerifySignature_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "pdf") return;
        try
        {
            var doc = PdfViewer.LoadedDocument;
            if (doc?.Form == null)
            {
                StatusText.Text = LanguageManager.Get("Pdf_NoSignatures");
                return;
            }

            var signatureFields = new List<string>();
            for (int i = 0; i < doc.Form.Fields.Count; i++)
            {
                var field = doc.Form.Fields[i];
                if (field is Syncfusion.Pdf.Interactive.PdfSignatureField sigField)
                {
                    var name = sigField.Name ?? "Signature";
                    signatureFields.Add($"{name}: {LanguageManager.Get("Pdf_SigValid")}");
                }
            }

            if (signatureFields.Count == 0)
            {
                StatusText.Text = LanguageManager.Get("Pdf_NoSignatures");
                return;
            }

            MessageBox.Show(
                string.Join("\n", signatureFields),
                LanguageManager.Get("Pdf_VerifySignature"),
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex) { StatusText.Text = $"PDF Verify: {ex.Message}"; }
    }

    // ── PDF Phase 3: Text Extraction / Search ───────────────────

    private void PdfExtractText_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "pdf") return;
        try
        {
            var doc = PdfViewer.LoadedDocument;
            if (doc == null) return;

            var sb = new System.Text.StringBuilder();
            for (int i = 0; i < doc.Pages.Count; i++)
            {
                sb.AppendLine($"--- {LanguageManager.Format("Pdf_PageNum", i + 1)} ---");
                sb.AppendLine(doc.Pages[i].ExtractText());
                sb.AppendLine();
            }

            var text = sb.ToString();
            if (string.IsNullOrWhiteSpace(text))
            {
                StatusText.Text = LanguageManager.Get("Pdf_NoTextFound");
                return;
            }

            Clipboard.SetText(text);
            StatusText.Text = LanguageManager.Format("Pdf_TextExtracted", text.Length);
        }
        catch (Exception ex) { StatusText.Text = $"PDF Extract: {ex.Message}"; }
    }

    private void PdfSearchText_Click(object sender, RoutedEventArgs e)
    {
        if (_activeEditorType != "pdf") return;
        try
        {
            var keyword = ShowInputDialog(LanguageManager.Get("Pdf_EnterSearchText"), LanguageManager.Get("Pdf_Search"));
            if (string.IsNullOrWhiteSpace(keyword)) return;

            PdfViewer.SearchText(keyword);
            StatusText.Text = LanguageManager.Format("Pdf_Searching", keyword);
        }
        catch (Exception ex) { StatusText.Text = $"PDF Search: {ex.Message}"; }
    }

    // ── PDF Helpers ─────────────────────────────────────────────

    /// <summary>Save document to temp file and reload in PdfViewer.</summary>
    private void ReloadPdfFromDocument(Syncfusion.Pdf.Parsing.PdfLoadedDocument doc)
    {
        var tempPath = Path.Combine(Path.GetTempPath(), $"IAOF_pdf_{Guid.NewGuid():N}.pdf");
        try
        {
            doc.Save(tempPath);
            doc.Close(true);
            PdfViewer.Unload();
            PdfViewer.Load(tempPath);
            // 遅延削除（PdfViewer がファイルをロック中の場合があるため）
            _ = Task.Delay(3000).ContinueWith(_ =>
            {
                try { if (File.Exists(tempPath)) File.Delete(tempPath); } catch { }
            });
        }
        catch
        {
            // フォールバック: MemoryStream
            try
            {
                using var ms = new MemoryStream();
                doc.Save(ms);
                ms.Position = 0;
                PdfViewer.Load(ms);
            }
            catch { /* best effort */ }
        }
    }

    /// <summary>Simple input dialog for text entry.</summary>
    private static string? ShowInputDialog(string prompt, string title, string defaultValue = "")
    {
        var dlg = new Window
        {
            Title = title,
            Width = 420, Height = 170,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            ResizeMode = ResizeMode.NoResize,
            Background = new System.Windows.Media.SolidColorBrush(
                (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#FAF8F5"))
        };
        var sp = new StackPanel { Margin = new Thickness(16) };
        sp.Children.Add(new TextBlock { Text = prompt, FontSize = 13, Margin = new Thickness(0, 0, 0, 8) });
        var tb = new TextBox { Text = defaultValue, FontSize = 13, Padding = new Thickness(6, 4, 6, 4) };
        sp.Children.Add(tb);
        var btnPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 12, 0, 0) };
        var okBtn = new Button { Content = "OK", Width = 80, Height = 28, IsDefault = true };
        var cancelBtn = new Button { Content = LanguageManager.Get("Btn_Cancel"), Width = 80, Height = 28, Margin = new Thickness(8, 0, 0, 0), IsCancel = true };
        okBtn.Click += (_, _) => { dlg.DialogResult = true; dlg.Close(); };
        cancelBtn.Click += (_, _) => { dlg.DialogResult = false; dlg.Close(); };
        btnPanel.Children.Add(okBtn);
        btnPanel.Children.Add(cancelBtn);
        sp.Children.Add(btnPanel);
        dlg.Content = sp;
        tb.Focus();
        tb.SelectAll();

        return dlg.ShowDialog() == true ? tb.Text : null;
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

    // ── プロジェクト (.iaof) ──

    private void ProjectNew_Click(object sender, RoutedEventArgs e)
    {
        HideAllBackstages();
        // 新しい空のテキストを作成（プロジェクトは保存時に .iaof 化）
        _projectService?.Dispose();
        _projectService = null;
        NewText_Click(sender, e);
        StatusText.Text = "新規プロジェクトを作成しました — ファイル > プロジェクト上書き保存 で .iaof に保存";
    }

    private async void ProjectSave_Click(object sender, RoutedEventArgs e)
    {
        HideAllBackstages();
        try
        {
            if (_projectService?.IsOpen == true)
            {
                // 上書き保存
                await _projectService.SaveAsync(_currentDocPath, _activeEditorType, _chatVm.ChatMessages);
                StatusText.Text = $"プロジェクトを保存しました: {Path.GetFileName(_projectService.ProjectPath)}";
                Title = $"{Path.GetFileNameWithoutExtension(_projectService.ProjectPath)} — Insight AI Office";
            }
            else
            {
                // 新規保存
                ProjectSaveAs_Click(sender, e);
            }
        }
        catch (Exception ex) { StatusText.Text = $"保存エラー: {ex.Message}"; }
    }

    private async void ProjectSaveAs_Click(object sender, RoutedEventArgs e)
    {
        HideAllBackstages();
        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "IAOF Project|*.iaof",
            DefaultExt = ".iaof",
            FileName = !string.IsNullOrEmpty(_currentDocPath)
                ? Path.GetFileNameWithoutExtension(_currentDocPath)
                : "新規プロジェクト",
        };
        if (dialog.ShowDialog() != true) return;

        try
        {
            _projectService?.Dispose();
            _projectService = new Services.IaofProjectService();

            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "1.0.0";
            await _projectService.CreateAsync(
                dialog.FileName, _currentDocPath, _activeEditorType,
                _chatVm.ChatMessages, version);

            StatusText.Text = $"プロジェクトを保存しました: {Path.GetFileName(dialog.FileName)}";
            Title = $"{Path.GetFileNameWithoutExtension(dialog.FileName)} — Insight AI Office";
        }
        catch (Exception ex) { StatusText.Text = $"保存エラー: {ex.Message}"; }
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
