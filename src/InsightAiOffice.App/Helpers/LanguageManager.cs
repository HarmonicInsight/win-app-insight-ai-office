namespace InsightAiOffice.App.Helpers;

public static class LanguageManager
{
    private static readonly Dictionary<string, Dictionary<string, string>> s_resources = new()
    {
        ["ja"] = new()
        {
            // App
            ["App_Title"] = "Insight AI Office",
            ["App_Tagline"] = "AIの民主化：誰もが業務でAIを使えるように",

            // Welcome
            ["Welcome_Tagline"] = "AI の民主化：誰もが業務で AI を使えるように",
            ["Welcome_EditTitle"] = "オフィスを編集する",
            ["Welcome_EditDesc"] = "Word・Excel を直接編集\nPowerPoint・PDF を表示・操作",
            ["Welcome_CreateTitle"] = "新しい資料を作る",
            ["Welcome_CreateDesc"] = "AI が要約・分析・翻訳して\nレポート・グラフ・表を自動生成",
            ["Welcome_OpenLabel"] = "ファイルを開いて編集する",
            ["Welcome_OpenDesc"] = ".docx / .xlsx / .pptx / .pdf を開いて AI と一緒に作業",
            ["Welcome_ChatLabel"] = "AI コンシェルジュで資料を作成",
            ["Welcome_ChatDescNew"] = "要約・分析・翻訳 → レポート・グラフ・表を自動生成",
            ["Welcome_Start"] = "はじめに",
            ["Welcome_OpenChat"] = "AI チャットを開く",
            ["Welcome_ChatDesc"] = "ドキュメント生成・分析・質問応答",
            ["Welcome_AiGenerate"] = "AI 文書生成",
            ["Welcome_AiGenerateDesc"] = "プロンプトからドキュメントを自動作成",
            ["Welcome_DragHint"] = "ファイルをドラッグ＆ドロップでも開けます",

            // Menu / Ribbon
            ["Menu_File"] = "ファイル",
            ["Menu_Home"] = "ホーム",
            ["Menu_AI"] = "AI",
            ["Menu_Prompt"] = "プロンプト",
            ["Menu_View"] = "表示",
            ["Menu_Help"] = "ヘルプ",

            // File operations
            ["File_New"] = "新規作成",
            ["File_Open"] = "開く",
            ["File_Save"] = "保存",
            ["File_SaveAs"] = "名前を付けて保存",
            ["File_SaveOverwrite"] = "上書き保存",
            ["File_Print"] = "印刷",
            ["File_Close"] = "閉じる",
            ["File_CloseDoc"] = "ドキュメントを閉じる",
            ["File_ProjectSave"] = "プロジェクト保存 (.iaof)",
            ["File_Settings"] = "設定",
            ["File_Export"] = "エクスポート",

            // Format
            ["Format_Font"] = "フォント",
            ["Format_Bold"] = "太字",
            ["Format_Italic"] = "斜体",
            ["Format_Underline"] = "下線",
            ["Format_Paragraph"] = "段落",
            ["Format_AlignLeft"] = "左揃え",
            ["Format_AlignCenter"] = "中央揃え",
            ["Format_AlignRight"] = "右揃え",
            ["Format_Strikethrough"] = "取り消し線",
            ["Format_Edit"] = "編集",
            ["Format_FindReplace"] = "検索と置換",
            ["Format_UndoRedo"] = "元に戻す",
            ["Format_Undo"] = "元に戻す",
            ["Format_Redo"] = "やり直し",
            ["Format_Insert"] = "挿入",
            ["Format_InsertImage"] = "画像挿入",
            ["Format_BulletList"] = "箇条書き",

            // Excel
            ["Excel_Clipboard"] = "クリップボード",
            ["Excel_Paste"] = "貼り付け",
            ["Excel_Cut"] = "切り取り",
            ["Excel_Copy"] = "コピー",
            ["Excel_Alignment"] = "配置",
            ["Excel_Number"] = "数値",
            ["Excel_Comma"] = "桁区切り",
            ["Excel_Currency"] = "通貨",
            ["Excel_DecInc"] = "小数点以下+",
            ["Excel_DecDec"] = "小数点以下−",
            ["Excel_Cells"] = "セル",
            ["Excel_MergeCells"] = "セル結合",
            ["Excel_WrapText"] = "折り返し",
            ["Excel_Borders"] = "罫線",
            ["Excel_View"] = "表示",
            ["Excel_FreezePanes"] = "枠の固定",

            // PPTX
            ["Pptx_SlideInfo"] = "スライド情報",
            ["Pptx_ExtractText"] = "テキスト抽出",
            ["Pptx_Proofread"] = "校正",
            ["Pptx_SlideOps"] = "スライド操作",
            ["Pptx_Add"] = "追加",
            ["Pptx_Duplicate"] = "複製",
            ["Pptx_Delete"] = "削除",
            ["Pptx_MoveUp"] = "上へ",
            ["Pptx_MoveDown"] = "下へ",
            ["Pptx_ExportPdf"] = "PDF出力",
            ["Pptx_SlideAdded"] = "スライドを追加しました",
            ["Pptx_SlideDuplicated"] = "スライドを複製しました",
            ["Pptx_SlideDeleted"] = "スライドを削除しました",
            ["Pptx_CannotDeleteLast"] = "最後のスライドは削除できません",

            // Panes
            ["Pane_Chat"] = "AI コンシェルジュ",
            ["Pane_Prompt"] = "プロンプト管理",
            ["Pane_Reference"] = "参考資料",
            ["Pane_PromptRef"] = "プロンプト・資料",
            ["Pane_AiSettings"] = "AI 設定",

            // AI
            ["AI_Chat"] = "AI チャット",
            ["AI_Summarize"] = "要約",
            ["AI_Proofread"] = "校正",
            ["AI_Analyze"] = "分析",
            ["AI_DataAnalysis"] = "データ分析",
            ["AI_FormulaHelp"] = "数式ヘルプ",
            ["AI_CrossAnalysis"] = "横断分析",
            ["AI_GenerateDoc"] = "AI 文書生成",
            ["AI_GenerateReport"] = "AI レポート",
            ["AI_InsertToDoc"] = "ドキュメントに挿入",
            ["AI_Copy"] = "コピー",

            // Reference panel
            ["Ref_Title"] = "参考資料",
            ["Ref_AddFile"] = "ファイル追加",
            ["Ref_AddFolder"] = "フォルダ追加",
            ["Ref_EmptyHint"] = "ファイルをドラッグ＆ドロップ\nまたは下のボタンで追加",
            ["Ref_FileFilter"] = "対応ファイル|*.pdf;*.docx;*.doc;*.xlsx;*.xls;*.txt;*.csv|すべてのファイル|*.*",
            ["Ref_Added"] = "参考資料を追加: {0}",
            ["Ref_Removed"] = "参考資料を削除: {0}",
            ["Ref_Error"] = "参考資料の追加に失敗しました",
            ["Ref_DragDropHint"] = "参考資料をドラッグ＆ドロップ\nまたは「＋」ボタンで追加",

            // Chat
            ["Chat_Send"] = "送信",
            ["Chat_ClearHistory"] = "チャット履歴をクリア",
            ["Chat_Presets"] = "プロンプト集",
            ["Chat_PromptEditor"] = "プロンプトエディタ",
            ["Chat_Help"] = "AI コンシェルジュのヘルプ",
            ["Chat_Settings"] = "AI 設定",
            ["Chat_Close"] = "閉じる",
            ["Chat_Cancel"] = "キャンセル",
            ["Chat_Placeholder"] = "ドキュメントについて質問...",
            ["Chat_WelcomeMessage"] = "こんにちは！AI コンシェルジュです。\nWord・Excel・PowerPoint の編集サポート、\n要約・校正・分析など、お気軽にどうぞ。",
            ["Chat_CopyMessage"] = "コピー",
            ["Chat_RetryMessage"] = "再実行",

            // Prompt Editor
            ["PE_Title"] = "プロンプトエディタ",
            ["PE_Presets"] = "プリセット",
            ["PE_Custom"] = "マイプロンプト",
            ["PE_Name"] = "名前",
            ["PE_Category"] = "カテゴリ",
            ["PE_Model"] = "モデル",
            ["PE_PromptText"] = "プロンプトテキスト",
            ["PE_Save"] = "保存",
            ["PE_Execute"] = "✓ AI実行",
            ["PE_ExecuteTooltip"] = "このプロンプトでAIを実行し、結果をチャットに表示",
            ["PE_Add"] = "新規プロンプト追加",
            ["PE_Delete"] = "削除",
            ["PE_NewPrompt"] = "新しいプロンプト",
            ["PE_Export"] = "エクスポート",
            ["PE_Import"] = "インポート",
            ["PE_ExportTooltip"] = "プロンプトをJSONファイルにエクスポート",
            ["PE_ImportTooltip"] = "JSONファイルからプロンプトをインポート",
            ["PE_ExportEmpty"] = "エクスポートするプロンプトがありません。",
            ["PE_ExportSuccess"] = "{0} 件のプロンプトをエクスポートしました。",
            ["PE_ImportEmpty"] = "インポートできるプロンプトがありませんでした。",
            ["PE_ImportSuccess"] = "{0} 件のプロンプトをインポートしました。（{1} 件は重複スキップ）",
            ["PE_Customize"] = "マイプロンプト登録",
            ["PE_CustomizeTooltip"] = "プリセットを修正してマイプロンプトとして登録",
            ["PE_CustomSuffix"] = "マイプロンプト",

            // Backstage
            ["BS_Open"] = "開く",
            ["BS_Recent"] = "最近使ったファイル",
            ["BS_RecentEmpty"] = "まだファイルを開いていません",
            ["BS_ProjectSave"] = "プロジェクト保存 (.iaof)",
            ["BS_Settings"] = "設定",
            ["BS_Language"] = "言語 / Language",
            ["BS_License"] = "ライセンス",
            ["BS_CloseDoc"] = "ドキュメントを閉じる",

            // Window
            ["Window_Minimize"] = "最小化",
            ["Window_Maximize"] = "最大化",
            ["Window_Close"] = "閉じる",

            // License
            ["License_Title"] = "ライセンス",
            ["License_Management"] = "ライセンス管理",
            ["License_CurrentPlan"] = "現在のプラン",
            ["License_OpenManager"] = "ライセンス管理を開く",

            // Status
            ["Status_Ready"] = "準備完了",
            ["Status_Loading"] = "読み込み中...",

            // AI Status
            ["AI_Thinking"] = "AI が考えています...",
            ["AI_Responding"] = "AI 応答中...",
            ["AI_Processing"] = "AI 処理中...",
            ["AI_GeneratingDoc"] = "AI 文書生成中...",
            ["AI_GeneratingReport"] = "AI レポート生成中...",
            ["AI_CrossAnalyzing"] = "AI クロスフォーマット分析中...",
            ["AI_ConfigRequired"] = "API キーが設定されていません。AI タブ → AI 設定 から設定してください。",
            ["AI_ConfigShort"] = "API キーが設定されていません。AI 設定から設定してください。",
            ["AI_ConfigDone"] = "AI 設定完了 — {0}",
            ["AI_ConfigNotSet"] = "API キー未設定",
            ["AI_Timeout"] = "AI 応答がタイムアウトしました",
            ["AI_Error"] = "AI エラー",
            ["AI_InsertedToDoc"] = "AI 応答をドキュメントに挿入しました",
            ["AI_CopiedToClipboard"] = "クリップボードにコピーしました",
            ["AI_CopiedFallback"] = "クリップボードにコピーしました（Word 以外では直接挿入できません）",
            ["AI_ToolExecuted"] = "準備完了（{0} 件の操作を実行）",
            ["AI_PresetUpdated"] = "プロンプトプリセットを更新しました",
            ["AI_InputTheme"] = "テーマを入力してください...",
            ["AI_DocGenerated"] = "AI 文書を生成しました — 編集して保存できます",
            ["AI_ReportGenerated"] = "AI レポートを生成しました — 編集して保存できます",

            // Document
            ["Doc_OpenFirst"] = "ドキュメントを開いてから実行してください",
            ["Doc_OpenExcelFirst"] = "Excel データを開いてからレポート生成を実行してください",
            ["Doc_NeedBoth"] = "ドキュメントと参考資料の両方が必要です",
            ["Doc_ChatCleared"] = "チャット履歴をクリアしました",
            ["Doc_Unsupported"] = "未対応のファイル形式: {0}",
            ["Doc_Loaded"] = "{0} を読み込みました",
            ["Doc_ProjectNotFound"] = "プロジェクトファイル内にドキュメントが見つかりません",
            ["Doc_ProjectOpened"] = "プロジェクト {0} を開きました",
            ["Doc_ProjectSaved"] = "プロジェクト保存: {0}",
            ["Doc_SaveFirst"] = "ドキュメントを開いてからプロジェクト保存してください",
            ["Doc_RefAdded"] = "参考資料を追加しました（{0} 件）",
            ["Doc_RefRemoved"] = "参考資料を削除しました",
            ["Doc_RefUsed"] = "参考資料 {0} 件を AI コンテキストに使用",
            ["Doc_NoValidFiles"] = "対応ファイルが見つかりませんでした",
            ["Doc_Delete"] = "削除",
            ["Doc_Expire"] = "有効期限: {0:yyyy年MM月dd日}",
            ["Doc_Adding"] = "追加中: {0}...",
            ["Doc_AddError"] = "追加エラー: {0}",

            // Ribbon hints
            ["Hint_FindReplace"] = "検索と置換: Ctrl+H で使用できます",
            ["Hint_Cut"] = "切り取り: Ctrl+X を使用してください",
            ["Hint_Copy"] = "コピー: Ctrl+C を使用してください",
            ["Hint_Paste"] = "貼り付け: Ctrl+V を使用してください",
            ["Hint_Underline"] = "下線: セルの書式設定から設定できます",
            ["Hint_ExportWordFirst"] = "Word ドキュメントを開いている状態でエクスポートできます",
            ["Hint_ExportExcelFirst"] = "Excel ファイルを開いている状態でエクスポートできます",
            ["Hint_PrintWordOnly"] = "印刷: Word ドキュメントのみ対応",
            ["Hint_TextExtracted"] = "テキスト抽出完了 — クリップボードにコピーしました（{0} 文字）",
            ["Hint_WordExported"] = "Word エクスポート: {0}",
            ["Hint_ExcelExported"] = "Excel エクスポート: {0}",
            ["Hint_PdfExported"] = "PDF エクスポート: {0}",
            ["Hint_BulletList"] = "箇条書き: リボンから書式を選択してください",
            ["Hint_ImageInserted"] = "画像挿入: {0}",
            ["Hint_UndoUnavailable"] = "元に戻す操作はありません",
            ["Hint_RedoUnavailable"] = "やり直す操作はありません",
            ["Hint_BordersApplied"] = "罫線を適用しました",
            ["Hint_FrozenPanes"] = "ウィンドウ枠を固定しました",
            ["Hint_SelectCellToFreeze"] = "固定する位置のセルを選択してください",
            ["Hint_ClearAll"] = "すべての参考資料を削除しました",
            ["Doc_Closed"] = "ドキュメントを閉じました",

            // Settings
            ["Settings_AiSettings"] = "AI 設定",
            ["Settings_AiSettingsDesc"] = "AI プロバイダー・モデル・API キーを設定",
            ["Settings_OpenAiSettings"] = "AI 設定を開く",

            // Errors
            ["Error_Title"] = "エラー",
            ["Error_StartupError"] = "起動エラー",
            ["Error_StartupFailed"] = "アプリの起動に失敗しました: {0}",
            ["Error_Unexpected"] = "予期しないエラーが発生しました: {0}",
        },
        ["en"] = new()
        {
            // App
            ["App_Title"] = "Insight AI Office",
            ["App_Tagline"] = "The New Office — Work with AI",

            // Welcome
            ["Welcome_Tagline"] = "Democratizing AI: Empowering Everyone to Use AI at Work",
            ["Welcome_EditTitle"] = "Edit Office Documents",
            ["Welcome_EditDesc"] = "Edit Word & Excel directly\nView PowerPoint & PDF",
            ["Welcome_CreateTitle"] = "Create New Materials",
            ["Welcome_CreateDesc"] = "AI summarizes, analyzes & translates\nAuto-generates reports, charts & tables",
            ["Welcome_OpenLabel"] = "Open & Edit Files",
            ["Welcome_OpenDesc"] = "Open .docx / .xlsx / .pptx / .pdf and work with AI",
            ["Welcome_ChatLabel"] = "Create with AI Concierge",
            ["Welcome_ChatDescNew"] = "Summarize, analyze, translate → auto-generate reports, charts & tables",
            ["Welcome_Start"] = "Get Started",
            ["Welcome_OpenChat"] = "Open AI Chat",
            ["Welcome_ChatDesc"] = "Document generation, analysis & Q&A",
            ["Welcome_AiGenerate"] = "AI Document Generation",
            ["Welcome_AiGenerateDesc"] = "Auto-create documents from prompts",
            ["Welcome_DragHint"] = "You can also drag & drop files to open them",

            // Menu / Ribbon
            ["Menu_File"] = "File",
            ["Menu_Home"] = "Home",
            ["Menu_AI"] = "AI",
            ["Menu_Prompt"] = "Prompt",
            ["Menu_View"] = "View",
            ["Menu_Help"] = "Help",

            // File operations
            ["File_New"] = "New",
            ["File_Open"] = "Open",
            ["File_Save"] = "Save",
            ["File_SaveAs"] = "Save As",
            ["File_SaveOverwrite"] = "Save",
            ["File_Print"] = "Print",
            ["File_Close"] = "Close",
            ["File_CloseDoc"] = "Close Document",
            ["File_ProjectSave"] = "Save Project (.iaof)",
            ["File_Settings"] = "Settings",
            ["File_Export"] = "Export",

            // Format
            ["Format_Font"] = "Font",
            ["Format_Bold"] = "Bold",
            ["Format_Italic"] = "Italic",
            ["Format_Underline"] = "Underline",
            ["Format_Paragraph"] = "Paragraph",
            ["Format_AlignLeft"] = "Align Left",
            ["Format_AlignCenter"] = "Center",
            ["Format_AlignRight"] = "Align Right",
            ["Format_Strikethrough"] = "Strikethrough",
            ["Format_Edit"] = "Edit",
            ["Format_FindReplace"] = "Find & Replace",
            ["Format_UndoRedo"] = "Undo",
            ["Format_Undo"] = "Undo",
            ["Format_Redo"] = "Redo",
            ["Format_Insert"] = "Insert",
            ["Format_InsertImage"] = "Insert Image",
            ["Format_BulletList"] = "Bullet List",

            // Excel
            ["Excel_Clipboard"] = "Clipboard",
            ["Excel_Paste"] = "Paste",
            ["Excel_Cut"] = "Cut",
            ["Excel_Copy"] = "Copy",
            ["Excel_Alignment"] = "Alignment",
            ["Excel_Number"] = "Number",
            ["Excel_Comma"] = "Comma Style",
            ["Excel_Currency"] = "Currency",
            ["Excel_DecInc"] = "Increase Decimal",
            ["Excel_DecDec"] = "Decrease Decimal",
            ["Excel_Cells"] = "Cells",
            ["Excel_MergeCells"] = "Merge Cells",
            ["Excel_WrapText"] = "Wrap Text",
            ["Excel_Borders"] = "Borders",
            ["Excel_View"] = "View",
            ["Excel_FreezePanes"] = "Freeze Panes",

            // PPTX
            ["Pptx_SlideInfo"] = "Slide Info",
            ["Pptx_ExtractText"] = "Extract Text",
            ["Pptx_Proofread"] = "Proofread",
            ["Pptx_SlideOps"] = "Slide Operations",
            ["Pptx_Add"] = "Add",
            ["Pptx_Duplicate"] = "Duplicate",
            ["Pptx_Delete"] = "Delete",
            ["Pptx_MoveUp"] = "Move Up",
            ["Pptx_MoveDown"] = "Move Down",
            ["Pptx_ExportPdf"] = "Export PDF",
            ["Pptx_SlideAdded"] = "Slide added",
            ["Pptx_SlideDuplicated"] = "Slide duplicated",
            ["Pptx_SlideDeleted"] = "Slide deleted",
            ["Pptx_CannotDeleteLast"] = "Cannot delete the last slide",

            // Panes
            ["Pane_Chat"] = "AI Concierge",
            ["Pane_Prompt"] = "Prompt Manager",
            ["Pane_Reference"] = "References",
            ["Pane_PromptRef"] = "Prompts & References",
            ["Pane_AiSettings"] = "AI Settings",

            // AI
            ["AI_Chat"] = "AI Chat",
            ["AI_Summarize"] = "Summarize",
            ["AI_Proofread"] = "Proofread",
            ["AI_Analyze"] = "Analyze",
            ["AI_DataAnalysis"] = "Data Analysis",
            ["AI_FormulaHelp"] = "Formula Help",
            ["AI_CrossAnalysis"] = "Cross Analysis",
            ["AI_GenerateDoc"] = "AI Generate Doc",
            ["AI_GenerateReport"] = "AI Report",
            ["AI_InsertToDoc"] = "Insert to Document",
            ["AI_Copy"] = "Copy",

            // Reference panel
            ["Ref_Title"] = "References",
            ["Ref_AddFile"] = "Add File",
            ["Ref_AddFolder"] = "Add Folder",
            ["Ref_EmptyHint"] = "Drag & drop files\nor use buttons below",
            ["Ref_FileFilter"] = "Supported Files|*.pdf;*.docx;*.doc;*.xlsx;*.xls;*.txt;*.csv|All Files|*.*",
            ["Ref_Added"] = "Reference added: {0}",
            ["Ref_Removed"] = "Reference removed: {0}",
            ["Ref_Error"] = "Failed to add reference",
            ["Ref_DragDropHint"] = "Drag & drop reference files\nor click + to add",

            // Chat
            ["Chat_Send"] = "Send",
            ["Chat_ClearHistory"] = "Clear Chat History",
            ["Chat_Presets"] = "Presets",
            ["Chat_PromptEditor"] = "Prompt Editor",
            ["Chat_Help"] = "AI Concierge Help",
            ["Chat_Settings"] = "AI Settings",
            ["Chat_Close"] = "Close",
            ["Chat_Cancel"] = "Cancel",
            ["Chat_Placeholder"] = "Ask about your document...",
            ["Chat_WelcomeMessage"] = "Hello! I'm your AI Concierge.\nI can help with editing, summarizing,\nproofreading, and analyzing your documents.",
            ["Chat_CopyMessage"] = "Copy",
            ["Chat_RetryMessage"] = "Retry",

            // Prompt Editor
            ["PE_Title"] = "Prompt Editor",
            ["PE_Presets"] = "Presets",
            ["PE_Custom"] = "My Prompts",
            ["PE_Name"] = "Name",
            ["PE_Category"] = "Category",
            ["PE_Model"] = "Model",
            ["PE_PromptText"] = "Prompt Text",
            ["PE_Save"] = "Save",
            ["PE_Execute"] = "✓ Run AI",
            ["PE_ExecuteTooltip"] = "Run AI with this prompt and show results in chat",
            ["PE_Add"] = "Add New Prompt",
            ["PE_Delete"] = "Delete",
            ["PE_NewPrompt"] = "New Prompt",
            ["PE_Export"] = "Export",
            ["PE_Import"] = "Import",
            ["PE_ExportTooltip"] = "Export prompts to a JSON file",
            ["PE_ImportTooltip"] = "Import prompts from a JSON file",
            ["PE_ExportEmpty"] = "No prompts to export.",
            ["PE_ExportSuccess"] = "{0} prompts exported.",
            ["PE_ImportEmpty"] = "No prompts found to import.",
            ["PE_ImportSuccess"] = "{0} prompts imported. ({1} duplicates skipped)",
            ["PE_Customize"] = "Save as My Prompt",
            ["PE_CustomizeTooltip"] = "Save this preset as a custom prompt for editing",
            ["PE_CustomSuffix"] = "Custom",

            // Backstage
            ["BS_Open"] = "Open",
            ["BS_Recent"] = "Recent Files",
            ["BS_RecentEmpty"] = "No files opened yet",
            ["BS_ProjectSave"] = "Save Project (.iaof)",
            ["BS_Settings"] = "Settings",
            ["BS_Language"] = "Language",
            ["BS_License"] = "License",
            ["BS_CloseDoc"] = "Close Document",

            // Window
            ["Window_Minimize"] = "Minimize",
            ["Window_Maximize"] = "Maximize",
            ["Window_Close"] = "Close",

            // License
            ["License_Title"] = "License",
            ["License_Management"] = "License Management",
            ["License_CurrentPlan"] = "Current Plan",
            ["License_OpenManager"] = "Open License Manager",

            // Status
            ["Status_Ready"] = "Ready",
            ["Status_Loading"] = "Loading...",

            // AI Status
            ["AI_Thinking"] = "AI is thinking...",
            ["AI_Responding"] = "AI responding...",
            ["AI_Processing"] = "AI processing...",
            ["AI_GeneratingDoc"] = "Generating document...",
            ["AI_GeneratingReport"] = "Generating report...",
            ["AI_CrossAnalyzing"] = "Cross-format analysis...",
            ["AI_ConfigRequired"] = "API key is not configured. Go to AI tab → AI Settings to set up.",
            ["AI_ConfigShort"] = "API key is not configured. Please set it in AI Settings.",
            ["AI_ConfigDone"] = "AI configured — {0}",
            ["AI_ConfigNotSet"] = "API key not set",
            ["AI_Timeout"] = "AI response timed out",
            ["AI_Error"] = "AI Error",
            ["AI_InsertedToDoc"] = "AI response inserted into document",
            ["AI_CopiedToClipboard"] = "Copied to clipboard",
            ["AI_CopiedFallback"] = "Copied to clipboard (direct insertion is only available for Word)",
            ["AI_ToolExecuted"] = "Ready ({0} operations executed)",
            ["AI_PresetUpdated"] = "Prompt presets updated",
            ["AI_InputTheme"] = "Enter a theme...",
            ["AI_DocGenerated"] = "AI document generated — ready to edit and save",
            ["AI_ReportGenerated"] = "AI report generated — ready to edit and save",

            // Document
            ["Doc_OpenFirst"] = "Please open a document first",
            ["Doc_OpenExcelFirst"] = "Please open Excel data before generating a report",
            ["Doc_NeedBoth"] = "Both a document and reference materials are required",
            ["Doc_ChatCleared"] = "Chat history cleared",
            ["Doc_Unsupported"] = "Unsupported file format: {0}",
            ["Doc_Loaded"] = "{0} loaded",
            ["Doc_ProjectNotFound"] = "No document found in the project file",
            ["Doc_ProjectOpened"] = "Project {0} opened",
            ["Doc_ProjectSaved"] = "Project saved: {0}",
            ["Doc_SaveFirst"] = "Please open a document before saving as project",
            ["Doc_RefAdded"] = "Reference materials added ({0} items)",
            ["Doc_RefRemoved"] = "Reference material removed",
            ["Doc_RefUsed"] = "{0} reference(s) used as AI context",
            ["Doc_NoValidFiles"] = "No supported files found",
            ["Doc_Delete"] = "Delete",
            ["Doc_Expire"] = "Expires: {0:MMM dd, yyyy}",
            ["Doc_Adding"] = "Adding: {0}...",
            ["Doc_AddError"] = "Error adding: {0}",

            // Ribbon hints
            ["Hint_FindReplace"] = "Find & Replace: Use Ctrl+H",
            ["Hint_Cut"] = "Cut: Use Ctrl+X",
            ["Hint_Copy"] = "Copy: Use Ctrl+C",
            ["Hint_Paste"] = "Paste: Use Ctrl+V",
            ["Hint_Underline"] = "Underline: Available in cell formatting options",
            ["Hint_ExportWordFirst"] = "Open a Word document first to export",
            ["Hint_ExportExcelFirst"] = "Open an Excel file first to export",
            ["Hint_PrintWordOnly"] = "Print: Available for Word documents only",
            ["Hint_TextExtracted"] = "Text extracted — copied to clipboard ({0} characters)",
            ["Hint_WordExported"] = "Word exported: {0}",
            ["Hint_ExcelExported"] = "Excel exported: {0}",
            ["Hint_PdfExported"] = "PDF exported: {0}",
            ["Hint_BulletList"] = "Bullet list: Select a format from the ribbon",
            ["Hint_ImageInserted"] = "Image inserted: {0}",
            ["Hint_UndoUnavailable"] = "Nothing to undo",
            ["Hint_RedoUnavailable"] = "Nothing to redo",
            ["Hint_BordersApplied"] = "Borders applied",
            ["Hint_FrozenPanes"] = "Panes frozen",
            ["Hint_SelectCellToFreeze"] = "Select a cell to freeze panes at",
            ["Hint_ClearAll"] = "All reference materials removed",
            ["Doc_Closed"] = "Document closed",

            // Settings
            ["Settings_AiSettings"] = "AI Settings",
            ["Settings_AiSettingsDesc"] = "Configure AI provider, model, and API key",
            ["Settings_OpenAiSettings"] = "Open AI Settings",

            // Errors
            ["Error_Title"] = "Error",
            ["Error_StartupError"] = "Startup Error",
            ["Error_StartupFailed"] = "Failed to start the application: {0}",
            ["Error_Unexpected"] = "An unexpected error occurred: {0}",
        },
    };

    public static string CurrentLanguage { get; private set; } = "ja";

    public static void SetLanguage(string lang)
    {
        CurrentLanguage = s_resources.ContainsKey(lang) ? lang : "ja";
    }

    public static string Get(string key)
    {
        if (s_resources.TryGetValue(CurrentLanguage, out var dict) && dict.TryGetValue(key, out var value))
            return value;
        if (s_resources.TryGetValue("ja", out var jaDict) && jaDict.TryGetValue(key, out var jaValue))
            return jaValue;
        return key;
    }

    public static string Format(string key, params object[] args)
    {
        return string.Format(Get(key), args);
    }

    // Static properties for x:Static binding in XAML DataTemplates
    public static string InsertToDocLabel => Get("AI_InsertToDoc");
    public static string CopyLabel => Get("AI_Copy");
}
