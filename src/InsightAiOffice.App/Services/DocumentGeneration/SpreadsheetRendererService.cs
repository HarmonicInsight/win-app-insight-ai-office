using System;
using System.Globalization;
using System.Linq;
using ClosedXML.Excel;
using InsightAiOffice.App.Services.DocumentGeneration;

namespace InsightAiOffice.App.Services.DocumentGeneration;

/// <summary>
/// SpreadsheetStructure（AI 生成 JSON）→ .xlsx レンダリング
///
/// プレミアム Ivory &amp; Gold テーマ適用。
/// 法人向けに「おっ」と思われるプロフェッショナルな見た目を実現。
/// </summary>
public static class SpreadsheetRendererService
{
    // カラーはテーマから動的取得
    private static XLColor ToXL(System.Drawing.Color c) => XLColor.FromArgb(c.R, c.G, c.B);
    private static readonly XLColor PositiveGreen = XLColor.FromHtml("#16A34A");
    private static readonly XLColor NegativeRed = XLColor.FromHtml("#DC2626");

    private const string FontName = "Yu Gothic UI";
    private const double FontSizeHeader = 11;
    private const double FontSizeBody = 10;
    private const double FontSizeTitle = 14;

    public static string Render(SpreadsheetStructure spec, string outputPath, string? themeName = null)
    {
        var t = DocumentColorTheme.FromName(themeName);
        var GoldPrimary = ToXL(t.Primary);
        var GoldDark = ToXL(t.PrimaryDark);
        var HeaderBg = ToXL(t.TableHeaderBg);
        var HeaderText = ToXL(t.TableHeaderText);
        var SubHeaderBg = ToXL(t.SummaryBg);
        var StripeBg = ToXL(t.TableStripeBg);
        var BorderLight = ToXL(t.BorderColor);
        var TextPrimary = ToXL(t.TextPrimary);
        var TextSecondary = ToXL(t.TextSecondary);
        using var workbook = new XLWorkbook();

        foreach (var sheet in spec.Sheets)
        {
            var ws = workbook.Worksheets.Add(
                string.IsNullOrWhiteSpace(sheet.Name) ? "Sheet1" : sheet.Name);

            int startRow = 1;

            // ── タイトル行（シート名をタイトルとして表示）──
            if (!string.IsNullOrWhiteSpace(sheet.Name) && sheet.Headers.Count > 0)
            {
                var titleRange = ws.Range(1, 1, 1, Math.Max(sheet.Headers.Count, 1));
                titleRange.Merge();
                var titleCell = ws.Cell(1, 1);
                titleCell.Value = sheet.Name;
                titleCell.Style.Font.FontName = FontName;
                titleCell.Style.Font.FontSize = FontSizeTitle;
                titleCell.Style.Font.Bold = true;
                titleCell.Style.Font.FontColor = GoldPrimary;
                titleCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                titleCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                ws.Row(1).Height = 30;

                // タイトル下のゴールドライン
                var lineRange = ws.Range(2, 1, 2, Math.Max(sheet.Headers.Count, 1));
                lineRange.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                lineRange.Style.Border.BottomBorderColor = GoldPrimary;
                ws.Row(2).Height = 4;

                startRow = 3;
            }

            int headerRow = startRow;

            // ── ヘッダー行（ゴールド背景 + 白文字）──
            if (sheet.Headers.Count > 0)
            {
                for (int c = 0; c < sheet.Headers.Count; c++)
                {
                    var cell = ws.Cell(headerRow, c + 1);
                    cell.Value = sheet.Headers[c];
                    cell.Style.Font.FontName = FontName;
                    cell.Style.Font.FontSize = FontSizeHeader;
                    cell.Style.Font.Bold = true;
                    cell.Style.Font.FontColor = HeaderText;
                    cell.Style.Fill.BackgroundColor = HeaderBg;
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    cell.Style.Border.BottomBorderColor = GoldDark;
                }
                ws.Row(headerRow).Height = 28;
                startRow = headerRow + 1;
            }

            // ── データ行（ストライプ + 数値自動書式）──
            int dataRow = startRow;
            foreach (var rowData in sheet.Rows)
            {
                bool isStripe = (dataRow - startRow) % 2 == 1;

                for (int c = 0; c < rowData.Count; c++)
                {
                    var cell = ws.Cell(dataRow, c + 1);
                    var value = rowData[c];

                    // 数値検出 + 書式設定
                    if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var num))
                    {
                        cell.Value = num;
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                        // パーセント検出
                        if (value.Contains('%'))
                        {
                            cell.Value = num / 100.0;
                            cell.Style.NumberFormat.Format = "0.0%";
                        }
                        // 通貨検出（大きな数値は桁区切り）
                        else if (Math.Abs(num) >= 1000)
                        {
                            cell.Style.NumberFormat.Format = "#,##0";
                        }
                        else if (num != Math.Floor(num))
                        {
                            cell.Style.NumberFormat.Format = "#,##0.0";
                        }

                        // 正負で色分け（ヘッダーに「前年比」「増減」等を含む場合）
                        if (sheet.Headers.Count > c)
                        {
                            var header = sheet.Headers[c];
                            if (header.Contains("前年") || header.Contains("増減") || header.Contains("差") || header.Contains("変化"))
                            {
                                if (num > 0)
                                    cell.Style.Font.FontColor = PositiveGreen;
                                else if (num < 0)
                                    cell.Style.Font.FontColor = NegativeRed;
                            }
                        }
                    }
                    else
                    {
                        cell.Value = value;
                    }

                    // 共通スタイル
                    cell.Style.Font.FontName = FontName;
                    cell.Style.Font.FontSize = FontSizeBody;
                    cell.Style.Font.FontColor = TextPrimary;
                    cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    cell.Style.Border.BottomBorder = XLBorderStyleValues.Hair;
                    cell.Style.Border.BottomBorderColor = BorderLight;

                    // ストライプ背景
                    if (isStripe)
                        cell.Style.Fill.BackgroundColor = StripeBg;
                }

                ws.Row(dataRow).Height = 22;
                dataRow++;
            }

            // ── 合計行（データがすべて数値の列のみ）──
            if (sheet.Rows.Count >= 3 && sheet.Headers.Count > 0)
            {
                var totalRow = dataRow;
                bool hasTotals = false;

                for (int c = 0; c < sheet.Headers.Count; c++)
                {
                    var colValues = sheet.Rows.Select(r => c < r.Count ? r[c] : "").ToList();
                    bool allNumeric = colValues.All(v =>
                        double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out _));

                    if (allNumeric && c > 0) // 最初の列（ラベル列）はスキップ
                    {
                        var cell = ws.Cell(totalRow, c + 1);
                        var sum = colValues.Sum(v =>
                            double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out var n) ? n : 0);
                        cell.Value = sum;
                        cell.Style.Font.Bold = true;
                        cell.Style.NumberFormat.Format = "#,##0";
                        hasTotals = true;
                    }

                    var totalCell = ws.Cell(totalRow, c + 1);
                    totalCell.Style.Font.FontName = FontName;
                    totalCell.Style.Font.FontSize = FontSizeBody;
                    totalCell.Style.Font.FontColor = TextPrimary;
                    totalCell.Style.Font.Bold = true;
                    totalCell.Style.Fill.BackgroundColor = SubHeaderBg;
                    totalCell.Style.Border.TopBorder = XLBorderStyleValues.Double;
                    totalCell.Style.Border.TopBorderColor = GoldPrimary;
                    totalCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                }

                if (hasTotals)
                {
                    ws.Cell(totalRow, 1).Value = "合計";
                    ws.Row(totalRow).Height = 24;
                }
            }

            // ── 列幅調整（最小10, 最大40）──
            ws.Columns().AdjustToContents();
            foreach (var col in ws.ColumnsUsed())
            {
                var w = col.Width;
                col.Width = Math.Max(10, Math.Min(w + 3, 40));
            }

            // ── 印刷設定 ──
            ws.PageSetup.PrintAreas.Clear();
            ws.PageSetup.PageOrientation = XLPageOrientation.Landscape;
            ws.PageSetup.FitToPages(1, 0);
            ws.PageSetup.ShowGridlines = false;

            // ── ウィンドウ枠の固定（ヘッダー行まで）──
            if (sheet.Headers.Count > 0)
            {
                ws.SheetView.FreezeRows(headerRow);
            }
        }

        workbook.SaveAs(outputPath);
        return outputPath;
    }
}
