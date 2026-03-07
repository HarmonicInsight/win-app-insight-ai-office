using System.Windows;
using InsightCommon.AI;
using InsightCommon.Addon;
using InsightCommon.License;
using InsightAiOffice.App.ViewModels;

namespace InsightAiOffice.App;

public partial class MainWindow : Window
{
    private readonly InsightLicenseManager _licenseManager;
    private readonly AiService _aiService;
    private readonly PromptPresetService _presetService;
    private readonly ReferenceMaterialsService _referenceService;
    private readonly ChatPanelViewModel _chatVm;
    private bool _isLeftPanelOpen;
    private bool _isRightPanelOpen;
    private string _activeEditorType = ""; // "word", "excel", "pptx"
    private string _currentDocPath = "";

    public MainWindow(
        InsightLicenseManager licenseManager,
        AiService aiService,
        PromptPresetService presetService,
        ReferenceMaterialsService referenceService)
    {
        InitializeComponent();
        _licenseManager = licenseManager;
        _aiService = aiService;
        _presetService = presetService;
        _referenceService = referenceService;

        _chatVm = new ChatPanelViewModel(aiService, presetService, referenceService);
        _chatVm.GetDocumentContent = ExtractDocumentContent;
        _chatVm.SetStatusText = msg => StatusText.Text = msg;
        _chatVm.OnBeforeSend = OnBeforeChatSend;
        _chatVm.SetToolExecutor(GetToolExecutor());
        ChatPanel.DataContext = _chatVm;

        ChatPanel.HelpRequested += (_, _) => Views.HelpWindow.ShowSection(this, "ai-assistant");
        ChatPanel.CloseRequested += (_, _) => CloseRightPanel_Click(this, new RoutedEventArgs());
        ChatPanel.PromptEditorRequested += (_, _) => ChatPromptEditor_Click(this, new RoutedEventArgs());
        ChatPanel.InsertToDocumentRequested += InsertAiResponseText;
        ChatPanel.CopyResponseRequested += CopyAiResponseText;

        UpdatePlanBadge();
        UpdateLicenseBackstage();

        var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        VersionLabel.Text = $"v{version?.Major}.{version?.Minor}.{version?.Build}";
        VersionText.Text = $"v{version?.Major}.{version?.Minor}.{version?.Build}";

        Loaded += OnWindowLoaded;
    }

    private void OnWindowLoaded(object sender, RoutedEventArgs e)
    {
        InsightCommon.Theme.SyncfusionInitializer.ApplyTheme(
            MainRibbon, WordRibbon, ExcelRibbon, PptxRibbon, Spreadsheet);

        InsightCommon.UI.InsightScaleManager.Instance.ApplyToWindow(this);

        InitializeLanguageRadioButtons();
        ApplyLocalization();
    }

    private void ApplyLocalization()
    {
        var L = Helpers.LanguageManager.Get;

        // Title bar
        FileNameLabel.Text = L("App_Tagline");
        ChatToggleBtn.ToolTip = L("Pane_Chat");
        System.Windows.Automation.AutomationProperties.SetName(ChatToggleBtn, L("Pane_Chat"));
        PromptToggleBtn.ToolTip = L("Pane_Prompt");
        PromptToggleText.Text = L("Menu_Prompt");
        RefToggleBtn.ToolTip = L("Pane_Reference");
        RefToggleText.Text = L("Pane_Reference");
        MinimizeBtn.ToolTip = L("Window_Minimize");
        System.Windows.Automation.AutomationProperties.SetName(MinimizeBtn, L("Window_Minimize"));
        MaximizeButton.ToolTip = L("Window_Maximize");
        System.Windows.Automation.AutomationProperties.SetName(MaximizeButton, L("Window_Maximize"));
        CloseBtn.ToolTip = L("Window_Close");
        System.Windows.Automation.AutomationProperties.SetName(CloseBtn, L("Window_Close"));

        // Reference panel
        RefHeaderText.Text = L("Pane_Reference");
        RefAddFolderBtn.ToolTip = L("Ref_AddFolder");
        RefAddFileBtn.ToolTip = L("Ref_AddFile");
        RefCloseBtn.ToolTip = L("Window_Close");
        RefEmptyHint.Text = L("Ref_DragDropHint");

        // Chat panel
        _chatVm.RefreshLocalization();
        _chatVm.LoadPresetGroups();
        ChatPanel.RefreshLocalization();

        // Backstage
        MainRibbon.BackStageHeader = L("Menu_File");
        BS_Open.Header = L("BS_Open");
        BS_ProjectSave.Header = L("BS_ProjectSave");
        BS_Settings.Header = L("BS_Settings");
        BS_LanguageTab.Header = L("BS_Language");
        BS_LanguageTitle.Text = L("BS_Language");
        BS_LicenseTab.Header = L("BS_License");
        LicenseTitleText.Text = L("License_Title");
        CurrentPlanLabel.Text = L("License_CurrentPlan");
        LicenseOpenBtn.Content = L("License_OpenManager");
        BS_CloseDoc.Header = L("BS_CloseDoc");

        // Welcome panel
        WelcomeTagline.Text = L("Welcome_Tagline");
        WelcomeStartLabel.Text = L("Welcome_Start");
        WelcomeOpenLabel.Text = L("File_Open");
        WelcomeOpenDesc.Text = ".docx / .xlsx / .pptx / .iaof";
        WelcomeChatLabel.Text = L("Welcome_OpenChat");
        WelcomeChatDesc.Text = L("Welcome_ChatDesc");
        WelcomeGenLabel.Text = L("Welcome_AiGenerate");
        WelcomeGenDesc.Text = L("Welcome_AiGenerateDesc");
        WelcomeDragHint.Text = L("Welcome_DragHint");

        // Status bar
        StatusText.Text = L("Status_Ready");

        // Ribbon localization (Syncfusion properties don't support DynamicResource)
        ApplyRibbonLocalization();
    }

    private void ApplyRibbonLocalization()
    {
        var L = Helpers.LanguageManager.Get;

        // ── Default Ribbon ──
        MainRibbon.BackStageHeader = L("Menu_File");
        ApplyTab(MainRibbon, 0, L("Menu_Home"), new[]
        {
            (L("Menu_File"), new[] { L("File_Open") }),
            ("AI", new[] { L("AI_Chat"), L("AI_GenerateDoc"), L("AI_CrossAnalysis"), L("Pane_AiSettings") }),
            (L("Menu_Help"), new[] { L("Menu_Help") }),
        });
        ApplyBackstage(MainRibbon, L("File_Open"), L("File_ProjectSave"), L("File_Settings"), null, L("License_Title"));

        // ── Word Ribbon ──
        WordRibbon.BackStageHeader = L("Menu_File");
        ApplyTab(WordRibbon, 0, L("Menu_Home"), new[]
        {
            (L("Menu_File"), new[] { L("File_Open"), L("File_Save"), L("File_SaveAs") }),
            (L("File_Export"), new[] { "Word" }),
            (L("Format_Font"), new[] { L("Format_Bold"), L("Format_Italic"), L("Format_Underline") }),
            (L("Format_Paragraph"), new[] { L("Format_AlignLeft"), L("Format_AlignCenter"), L("Format_AlignRight") }),
            (L("Format_Edit"), new[] { L("Format_FindReplace") }),
            (L("File_Print"), new[] { L("File_Print") }),
        });
        ApplyTab(WordRibbon, 1, L("Menu_AI"), new[]
        {
            ("AI", new[] { L("AI_Chat"), L("AI_Summarize"), L("AI_Proofread"), L("AI_Analyze"), L("AI_CrossAnalysis") }),
            (L("Pane_PromptRef"), new[] { L("Pane_Prompt"), L("Pane_Reference") }),
            (L("Pane_AiSettings"), new[] { L("Pane_AiSettings") }),
        });
        ApplyBackstage(WordRibbon, L("File_Open"), L("File_SaveOverwrite"), L("File_Print"), null, L("File_CloseDoc"));

        // ── Excel Ribbon ──
        ExcelRibbon.BackStageHeader = L("Menu_File");
        ApplyTab(ExcelRibbon, 0, L("Menu_Home"), new[]
        {
            (L("Menu_File"), new[] { L("File_Open"), L("File_Save") }),
            (L("Excel_Clipboard"), new[] { L("Excel_Paste"), L("Excel_Cut"), L("Excel_Copy") }),
            (L("Format_Font"), new[] { L("Format_Bold"), L("Format_Italic"), L("Format_Underline") }),
            (L("Excel_Alignment"), new[] { L("Format_AlignLeft"), L("Format_AlignCenter"), L("Format_AlignRight") }),
            (L("Excel_Number"), new[] { "%", L("Excel_Comma") }),
        });
        ApplyTab(ExcelRibbon, 1, L("Menu_AI"), new[]
        {
            ("AI", new[] { L("AI_Chat"), L("AI_DataAnalysis"), L("AI_FormulaHelp"), L("AI_CrossAnalysis"), L("AI_GenerateReport") }),
            (L("Pane_PromptRef"), new[] { L("Pane_Prompt"), L("Pane_Reference") }),
            (L("Pane_AiSettings"), new[] { L("Pane_AiSettings") }),
        });
        ApplyBackstage(ExcelRibbon, L("File_Open"), L("File_SaveAs"), null, L("File_CloseDoc"));

        // ── PPTX Ribbon ──
        PptxRibbon.BackStageHeader = L("Menu_File");
        ApplyTab(PptxRibbon, 0, L("Menu_Home"), new[]
        {
            (L("Menu_File"), new[] { L("File_Open") }),
            (L("Pptx_SlideInfo"), new[] { L("Pptx_ExtractText") }),
        });
        ApplyTab(PptxRibbon, 1, L("Menu_AI"), new[]
        {
            ("AI", new[] { L("AI_Chat"), L("AI_Summarize"), L("AI_Analyze"), L("AI_CrossAnalysis") }),
            (L("Pane_PromptRef"), new[] { L("Pane_Prompt"), L("Pane_Reference") }),
            (L("Pane_AiSettings"), new[] { L("Pane_AiSettings") }),
        });
        ApplyBackstage(PptxRibbon, L("File_Open"), null, L("File_CloseDoc"));
    }

    private static void ApplyTab(Syncfusion.Windows.Tools.Controls.Ribbon ribbon, int tabIndex,
        string caption, (string header, string[] labels)[] bars)
    {
        if (tabIndex >= ribbon.Items.Count ||
            ribbon.Items[tabIndex] is not Syncfusion.Windows.Tools.Controls.RibbonTab tab) return;
        tab.Caption = caption;
        for (var i = 0; i < bars.Length && i < tab.Items.Count; i++)
        {
            if (tab.Items[i] is not Syncfusion.Windows.Tools.Controls.RibbonBar bar) continue;
            bar.Header = bars[i].header;
            for (var j = 0; j < bars[i].labels.Length && j < bar.Items.Count; j++)
            {
                if (bar.Items[j] is Syncfusion.Windows.Tools.Controls.RibbonButton btn)
                    btn.Label = bars[i].labels[j];
            }
        }
    }

    private static void ApplyBackstage(Syncfusion.Windows.Tools.Controls.Ribbon ribbon, params string?[] headers)
    {
        if (ribbon.BackStage is not Syncfusion.Windows.Tools.Controls.Backstage backstage) return;
        var idx = 0;
        foreach (var item in backstage.Items)
        {
            if (idx >= headers.Length) break;
            switch (item)
            {
                case Syncfusion.Windows.Tools.Controls.BackStageCommandButton cmd:
                    if (headers[idx] != null) cmd.Header = headers[idx]!;
                    idx++;
                    break;
                case Syncfusion.Windows.Tools.Controls.BackstageTabItem tab:
                    if (headers[idx] != null) tab.Header = headers[idx]!;
                    idx++;
                    break;
                case Syncfusion.Windows.Tools.Controls.BackStageSeparator:
                    if (headers[idx] == null) idx++;
                    break;
            }
        }
    }

    // ── Ribbon Switching ──────────────────────────────────────────

    private void SwitchRibbon(string editorType)
    {
        DefaultRibbonPanel.Visibility = Visibility.Collapsed;
        WordRibbonPanel.Visibility = Visibility.Collapsed;
        ExcelRibbonPanel.Visibility = Visibility.Collapsed;
        PptxRibbonPanel.Visibility = Visibility.Collapsed;

        switch (editorType)
        {
            case "word":
                WordRibbonPanel.Visibility = Visibility.Visible;
                break;
            case "excel":
                ExcelRibbonPanel.Visibility = Visibility.Visible;
                break;
            case "pptx":
                PptxRibbonPanel.Visibility = Visibility.Visible;
                break;
            default:
                DefaultRibbonPanel.Visibility = Visibility.Visible;
                break;
        }
    }
}

