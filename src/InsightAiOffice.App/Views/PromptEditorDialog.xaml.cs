using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using InsightCommon.AI;
using Microsoft.Win32;

namespace InsightAiOffice.App.Views;

public partial class PromptEditorDialog : Window
{
    private readonly PromptPresetService _presetService;
    private bool _isPresetSelected;
    private UserPromptPreset? _selectedPreset;
    private readonly HashSet<AiProviderType> _configuredProviders;
    private readonly bool _isEnterprise;

    public bool HasChanges { get; private set; }
    public string? ExecutePromptText { get; private set; }
    public string? ExecuteModelId { get; private set; }

    public PromptEditorDialog(PromptPresetService presetService, AiProviderConfig? providerConfig = null, bool isEnterprise = false)
    {
        InitializeComponent();
        _presetService = presetService;
        _isEnterprise = isEnterprise;

        _configuredProviders = BuildConfiguredProviders(providerConfig);
        ApplyLocalization();
        RefreshCategoryComboBox();
        BuildTree();

        // ENT以外は組織プリセット機能を非表示
        if (!_isEnterprise)
        {
            AddToOrgButton.Visibility = Visibility.Collapsed;
            ExportButton.Visibility = Visibility.Collapsed;
            ImportButton.Visibility = Visibility.Collapsed;
        }
    }

    private static HashSet<AiProviderType> BuildConfiguredProviders(AiProviderConfig? config)
    {
        var set = new HashSet<AiProviderType>();
        if (config == null) return set;
        foreach (var provider in AiModelRegistry.GetAvailableProviders())
        {
            if (!string.IsNullOrEmpty(config.GetApiKey(provider)))
                set.Add(provider);
        }
        return set;
    }

    private void ApplyLocalization()
    {
        var L = Helpers.LanguageManager.Get;
        Title = L("PE_Title");
        PromptListLabel.Text = L("PE_Presets");
        AddButton.ToolTip = L("PE_Add");
        NameLabel.Text = L("PE_Name");
        CategoryLabel.Text = L("PE_Category");
        PromptTextLabel.Text = L("PE_PromptText");
        SaveButton.Content = L("PE_Save");
        ExecuteButton.Content = L("PE_Execute");
        ExecuteButton.ToolTip = L("PE_ExecuteTooltip");
        DeleteButton.Content = L("PE_Delete");
        CustomizeButton.Content = L("PE_Customize");
        CustomizeButton.ToolTip = L("PE_CustomizeTooltip");
        ExportButton.ToolTip = L("PE_ExportTooltip");
        ImportButton.ToolTip = L("PE_ImportTooltip");
    }

    private void BuildTree() => BuildTree("");

    private void BuildTree(string filter)
    {
        PromptTree.Items.Clear();
        var L = Helpers.LanguageManager.Get;

        var allPresets = _presetService.LoadAll();

        // 検索フィルタ
        if (!string.IsNullOrWhiteSpace(filter))
        {
            var terms = filter.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            allPresets = allPresets.Where(p =>
                terms.All(t =>
                    (p.Name?.Contains(t, StringComparison.OrdinalIgnoreCase) ?? false) ||
                    (p.Category?.Contains(t, StringComparison.OrdinalIgnoreCase) ?? false) ||
                    (p.Description?.Contains(t, StringComparison.OrdinalIgnoreCase) ?? false)))
                .ToList();
        }

        // マイプロンプト（ユーザー作成・組織以外）
        var customHeader = new TreeViewItem
        {
            Header = "📁 " + L("PE_Custom"),
            IsExpanded = true,
            FontWeight = FontWeights.Bold,
            FontSize = 13,
        };
        var custom = allPresets
            .Where(p => !p.Id.StartsWith("builtin_", StringComparison.Ordinal) && p.Group != "org")
            .ToList();
        if (custom.Count == 0)
        {
            customHeader.Items.Add(new TreeViewItem
            {
                Header = "＋ ボタンで追加できます",
                FontStyle = FontStyles.Italic,
                FontSize = 10,
                Foreground = System.Windows.Media.Brushes.Gray,
                IsEnabled = false,
            });
        }
        else
        {
            foreach (var preset in custom)
            {
                var label = preset.Name;
                if (preset.IsDefault) label = "\u2605 " + label;
                customHeader.Items.Add(new TreeViewItem
                {
                    Header = (preset.Icon ?? "📝") + " " + label,
                    Tag = preset,
                    FontWeight = FontWeights.Normal,
                    FontSize = 11,
                });
            }
        }
        PromptTree.Items.Add(customHeader);

        // 自組織プリセット（部長・リテラシー高い人がキュレーション）
        var orgPresets = allPresets
            .Where(p => p.Group == "org")
            .ToList();
        var orgHeader = new TreeViewItem
        {
            Header = "🏢 自組織プリセット" + (orgPresets.Count > 0 ? $"（{orgPresets.Count}件）" : ""),
            IsExpanded = true,
            FontWeight = FontWeights.Bold,
            FontSize = 13,
            Foreground = new System.Windows.Media.SolidColorBrush(
                (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#7C3AED")),
        };
        if (orgPresets.Count == 0)
        {
            orgHeader.Items.Add(new TreeViewItem
            {
                Header = "ビルトインから「組織に追加」で登録",
                FontStyle = FontStyles.Italic,
                FontSize = 10,
                Foreground = System.Windows.Media.Brushes.Gray,
                IsEnabled = false,
            });
        }
        else
        {
            foreach (var preset in orgPresets)
            {
                orgHeader.Items.Add(new TreeViewItem
                {
                    Header = (preset.Icon ?? "📝") + " " + preset.Name,
                    Tag = preset,
                    FontWeight = FontWeights.Normal,
                    FontSize = 11,
                });
            }
        }
        PromptTree.Items.Add(orgHeader);

        PromptTree.Items.Add(new Separator { Margin = new Thickness(4, 6, 4, 6) });

        // Built-in presets
        var builtIn = allPresets.Where(p => p.Id.StartsWith("builtin_", StringComparison.Ordinal)).ToList();

        if (_isSceneView)
            BuildSceneView(builtIn, filter);
        else
            BuildFunctionView(builtIn, filter);
    }

    private void BuildSceneView(List<UserPromptPreset> builtIn, string filter)
    {
        var sceneOrder = new[] { "相談・壁打ち", "コンサルタント", "システム部門", "総務部向け", "経理部向け", "人事部向け", "営業部向け", "経営企画", "法務部向け", "教育・研修", "レビュー", "翻訳", "全般" };
        var sceneIcons = new Dictionary<string, string>
        {
            ["相談・壁打ち"] = "💭", ["コンサルタント"] = "💼", ["システム部門"] = "💻",
            ["総務部向け"] = "🏢", ["経理部向け"] = "💰", ["人事部向け"] = "👤",
            ["営業部向け"] = "📈", ["経営企画"] = "🎯", ["法務部向け"] = "⚖️",
            ["教育・研修"] = "📚", ["レビュー"] = "🔍", ["翻訳"] = "🌐", ["全般"] = "📁",
        };

        var sceneGroups = builtIn
            .GroupBy(InferSceneTag)
            .OrderBy(g =>
            {
                var i = Array.IndexOf(sceneOrder, g.Key);
                return i >= 0 ? i : 99;
            });

        foreach (var group in sceneGroups)
        {
            var icon = sceneIcons.GetValueOrDefault(group.Key, "📁");
            var sceneItem = new TreeViewItem
            {
                Header = $"{icon} {group.Key}（{group.Count()}件）",
                IsExpanded = !string.IsNullOrWhiteSpace(filter),
                FontWeight = FontWeights.Bold,
                FontSize = 12,
            };
            foreach (var preset in group.OrderBy(p => p.Name))
            {
                sceneItem.Items.Add(new TreeViewItem
                {
                    Header = (preset.Icon ?? "📝") + " " + preset.Name,
                    Tag = preset,
                    FontWeight = FontWeights.Normal,
                    FontSize = 11,
                });
            }
            PromptTree.Items.Add(sceneItem);
        }
    }

    private void BuildFunctionView(List<UserPromptPreset> builtIn, string filter)
    {
        var grouped = builtIn
            .GroupBy(p =>
            {
                var cat = p.Category ?? "";
                var colonIdx = cat.IndexOf(':');
                if (colonIdx > 0 && (cat.StartsWith("📊") || cat.StartsWith("📄") || cat.StartsWith("📗") || cat.StartsWith("📑") || cat.StartsWith("🔄")))
                    return cat[..colonIdx].Trim();
                return "Office 共通";
            })
            .OrderBy(g => g.Key switch
            {
                "Office 共通" => 0,
                _ when g.Key.StartsWith("📄") => 1,
                _ when g.Key.StartsWith("📗") => 2,
                _ when g.Key.StartsWith("📊") => 3,
                _ => 4,
            });

        foreach (var productGroup in grouped)
        {
            var productItem = new TreeViewItem
            {
                Header = "📋 " + productGroup.Key,
                IsExpanded = !string.IsNullOrWhiteSpace(filter),
                FontWeight = FontWeights.Bold,
                FontSize = 12,
            };

            var subGroups = productGroup
                .GroupBy(p =>
                {
                    var cat = p.Category ?? "";
                    var colonIdx = cat.IndexOf(':');
                    return colonIdx > 0 ? cat[(colonIdx + 1)..].Trim() : cat;
                })
                .OrderBy(g => g.Key);

            foreach (var sub in subGroups)
            {
                var subItem = new TreeViewItem
                {
                    Header = sub.Key,
                    IsExpanded = !string.IsNullOrWhiteSpace(filter),
                    FontWeight = FontWeights.Normal,
                    FontSize = 11,
                };
                foreach (var preset in sub)
                {
                    subItem.Items.Add(new TreeViewItem
                    {
                        Header = (preset.Icon ?? "📝") + " " + preset.Name,
                        Tag = preset,
                        FontWeight = FontWeights.Normal,
                        FontSize = 11,
                    });
                }
                productItem.Items.Add(subItem);
            }
            PromptTree.Items.Add(productItem);
        }
    }

    private void PromptTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
    {
        if (e.NewValue is not TreeViewItem tvi || tvi.Tag is not UserPromptPreset preset)
        {
            _isPresetSelected = false;
            _selectedPreset = null;
            EditorPanel.IsEnabled = false;
            DeleteButton.IsEnabled = false;
            CustomizeButton.IsEnabled = false;
            AddToOrgButton.IsEnabled = false;
            return;
        }

        _selectedPreset = preset;
        EditorPanel.IsEnabled = true;
        NameBox.Text = preset.Name;
        PromptBox.Text = preset.SystemPrompt;
        CategoryCombo.Text = preset.Category ?? "";
        ExecuteButton.IsEnabled = true;

        // 組織に追加ボタン: ビルトインまたはマイプロンプト（既に組織でないもの）で有効
        AddToOrgButton.IsEnabled = preset.Group != "org";

        if (preset.Id.StartsWith("builtin_", StringComparison.Ordinal))
        {
            _isPresetSelected = true;
            NameBox.IsReadOnly = true;
            PromptBox.IsReadOnly = true;
            CategoryCombo.IsEnabled = false;
            SaveButton.IsEnabled = false;
            DeleteButton.IsEnabled = false;
            CustomizeButton.IsEnabled = true;
        }
        else
        {
            _isPresetSelected = false;
            NameBox.IsReadOnly = false;
            PromptBox.IsReadOnly = false;
            CategoryCombo.IsEnabled = true;
            SaveButton.IsEnabled = true;
            DeleteButton.IsEnabled = true; // マイプロンプト・組織プリセットとも削除可能
            CustomizeButton.IsEnabled = false;
        }
    }

    private void Add_Click(object sender, RoutedEventArgs e)
    {
        var newPreset = new UserPromptPreset
        {
            Id = PromptPresetService.GenerateId(),
            Name = Helpers.LanguageManager.Get("PE_NewPrompt"),
            SystemPrompt = "",
            Category = "",
        };
        _presetService.Add(newPreset);
        HasChanges = true;
        BuildTree();
        SelectPresetInTree(newPreset.Id);
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedPreset == null || _isPresetSelected) return;

        var updated = new UserPromptPreset
        {
            Id = _selectedPreset.Id,
            Name = NameBox.Text.Trim(),
            SystemPrompt = PromptBox.Text.Trim(),
            Description = _selectedPreset.Description,
            Author = _selectedPreset.Author,
            Category = (CategoryCombo.Text ?? "").Trim(),
            Mode = "check",
            IsDefault = _selectedPreset.IsDefault,
            IsPinned = _selectedPreset.IsPinned,
            UsageCount = _selectedPreset.UsageCount,
            LastUsedAt = _selectedPreset.LastUsedAt,
            CreatedAt = _selectedPreset.CreatedAt,
            ModifiedAt = DateTime.Now,
        };

        _presetService.Update(_selectedPreset.Id, updated);
        _selectedPreset = updated;
        HasChanges = true;
        RefreshCategoryComboBox();

        var savedId = updated.Id;
        BuildTree();
        SelectPresetInTree(savedId);
    }

    private void Delete_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedPreset == null || _isPresetSelected) return;

        _presetService.Remove(_selectedPreset.Id);
        _selectedPreset = null;
        HasChanges = true;
        RefreshCategoryComboBox();
        EditorPanel.IsEnabled = false;
        BuildTree();
    }

    private void AddToOrg_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedPreset == null) return;

        // 既に組織プリセットなら何もしない
        if (_selectedPreset.Group == "org") return;

        // ビルトインの場合は複製して組織に追加
        var orgPreset = new UserPromptPreset
        {
            Id = PromptPresetService.GenerateId(),
            Name = _selectedPreset.Name,
            SystemPrompt = _selectedPreset.SystemPrompt,
            Description = _selectedPreset.Description,
            Category = _selectedPreset.Category,
            Icon = _selectedPreset.Icon,
            Mode = _selectedPreset.Mode,
            Group = "org",
            Author = "組織管理者",
        };
        _presetService.Add(orgPreset);
        HasChanges = true;
        BuildTree(_searchText);
        SelectPresetInTree(orgPreset.Id);
    }

    private void Customize_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedPreset == null) return;

        var customSuffix = Helpers.LanguageManager.Get("PE_CustomSuffix");
        var newPreset = new UserPromptPreset
        {
            Id = PromptPresetService.GenerateId(),
            Name = _selectedPreset.Name + " (" + customSuffix + ")",
            SystemPrompt = PromptBox.Text.Trim(),
            Description = _selectedPreset.Description,
            Author = "",
            Category = _selectedPreset.Category,
            Mode = _selectedPreset.Mode,
        };
        _presetService.Add(newPreset);
        HasChanges = true;
        RefreshCategoryComboBox();
        BuildTree();
        SelectPresetInTree(newPreset.Id);
    }

    private void Execute_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedPreset == null) return;

        var text = _isPresetSelected ? _selectedPreset.SystemPrompt : PromptBox.Text.Trim();

        if (!_isPresetSelected)
            Save_Click(sender, e);

        ExecutePromptText = text;
        ExecuteModelId = _selectedPreset.ModelId ?? ClaudeModels.DefaultModel;
        _presetService.IncrementUsage(_selectedPreset.Id);
        Close();
    }

    private void SelectPresetInTree(string id)
    {
        static bool TrySelect(ItemCollection items, string targetId)
        {
            foreach (object item in items)
            {
                if (item is not TreeViewItem tvi) continue;
                if (tvi.Tag is UserPromptPreset p && p.Id == targetId)
                {
                    tvi.IsSelected = true;
                    tvi.BringIntoView();
                    return true;
                }
                if (tvi.Items.Count > 0 && TrySelect(tvi.Items, targetId))
                {
                    tvi.IsExpanded = true;
                    return true;
                }
            }
            return false;
        }
        TrySelect(PromptTree.Items, id);
    }

    // ── Export / Import ──

    private void Export_Click(object sender, RoutedEventArgs e)
    {
        // 組織プリセットのみエクスポート（配布用）
        var allPresets = _presetService.LoadAll();
        var custom = allPresets.Where(p => p.Group == "org").ToList();
        if (custom.Count == 0)
        {
            MessageBox.Show(
                Helpers.LanguageManager.Get("PE_ExportEmpty"),
                Helpers.LanguageManager.Get("PE_Export"),
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var dlg = new SaveFileDialog
        {
            Filter = "JSON|*.json",
            FileName = $"prompts_{DateTime.Now:yyyyMMdd_HHmmss}.json",
            Title = Helpers.LanguageManager.Get("PE_Export"),
        };
        if (dlg.ShowDialog(this) != true) return;

        try
        {
            PromptPresetService.Export(custom, dlg.FileName);
            MessageBox.Show(
                Helpers.LanguageManager.Format("PE_ExportSuccess", custom.Count),
                Helpers.LanguageManager.Get("PE_Export"),
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, Helpers.LanguageManager.Get("PE_Export"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void Import_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Filter = "JSON|*.json",
            Title = Helpers.LanguageManager.Get("PE_Import"),
        };
        if (dlg.ShowDialog(this) != true) return;

        try
        {
            var imported = PromptPresetService.Import(dlg.FileName);
            if (imported.Count == 0)
            {
                MessageBox.Show(
                    Helpers.LanguageManager.Get("PE_ImportEmpty"),
                    Helpers.LanguageManager.Get("PE_Import"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var allPresets = _presetService.LoadAll();
            var existingIds = new HashSet<string>(allPresets.Select(p => p.Id));
            var added = 0;
            foreach (var item in imported)
            {
                if (string.IsNullOrEmpty(item.Id))
                    item.Id = PromptPresetService.GenerateId();
                if (!existingIds.Contains(item.Id))
                {
                    _presetService.Add(item);
                    existingIds.Add(item.Id);
                    added++;
                }
            }

            if (added > 0)
            {
                HasChanges = true;
                RefreshCategoryComboBox();
                BuildTree();
            }

            MessageBox.Show(
                Helpers.LanguageManager.Format("PE_ImportSuccess", added, imported.Count - added),
                Helpers.LanguageManager.Get("PE_Import"),
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, Helpers.LanguageManager.Get("PE_Import"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // ── Search & View Mode ──

    private string _searchText = "";
    private bool _isSceneView;
    private System.Windows.Threading.DispatcherTimer? _searchDebounce;

    private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        _searchText = SearchBox.Text.Trim();
        SearchPlaceholder.Visibility = string.IsNullOrEmpty(_searchText)
            ? Visibility.Visible : Visibility.Collapsed;

        // デバウンス: 200ms後にツリー再構築
        if (_searchDebounce == null)
        {
            _searchDebounce = new System.Windows.Threading.DispatcherTimer { Interval = TimeSpan.FromMilliseconds(200) };
            _searchDebounce.Tick += (_, _) => { _searchDebounce.Stop(); BuildTree(_searchText); };
        }
        _searchDebounce.Stop();
        _searchDebounce.Start();
    }

    private void ViewMode_Changed(object sender, RoutedEventArgs e)
    {
        // InitializeComponent 中は _presetService がまだ null
        if (_presetService == null) return;
        _isSceneView = ViewByScene?.IsChecked == true;
        BuildTree(_searchText);
    }

    /// <summary>
    /// カテゴリ・名前から業務シーン（部署タグ）を自動推定する。
    /// チュートリアルの分類と統一。
    /// </summary>
    private static string InferSceneTag(UserPromptPreset preset)
    {
        var name = (preset.Name ?? "").ToLowerInvariant();
        var cat = (preset.Category ?? "").ToLowerInvariant();
        var desc = (preset.Description ?? "").ToLowerInvariant();
        var all = name + " " + cat + " " + desc;

        // 明示的な Group があればそれを使う
        if (!string.IsNullOrWhiteSpace(preset.Group)) return preset.Group;

        // 部署・シーン推定
        if (all.Contains("議事録") || all.Contains("会議") || all.Contains("社内通知") || all.Contains("案内文"))
            return "総務部向け";
        if (all.Contains("売上") || all.Contains("経費") || all.Contains("請求") || all.Contains("仕訳") || all.Contains("決算")
            || all.Contains("予算") || all.Contains("月次") || all.Contains("入出金") || all.Contains("財務"))
            return "経理部向け";
        if (all.Contains("採用") || all.Contains("人事") || all.Contains("応募") || all.Contains("候補者")
            || all.Contains("内定") || all.Contains("勤怠") || all.Contains("給与"))
            return "人事部向け";
        if (all.Contains("営業") || all.Contains("提案書") || all.Contains("見積") || all.Contains("商談")
            || all.Contains("顧客") || all.Contains("案件"))
            return "営業部向け";
        if (all.Contains("プレゼン") || all.Contains("経営") || all.Contains("戦略") || all.Contains("企画")
            || all.Contains("ストーリー"))
            return "経営企画";
        if (all.Contains("レビュー") || all.Contains("校正") || all.Contains("赤入れ") || all.Contains("比較")
            || all.Contains("差分") || all.Contains("品質"))
            return "レビュー";
        if (all.Contains("翻訳") || all.Contains("英訳") || all.Contains("和訳"))
            return "翻訳";
        if (all.Contains("相談") || all.Contains("壁打ち") || all.Contains("ブレイン"))
            return "相談・壁打ち";
        if (all.Contains("契約") || all.Contains("コンプライアンス") || all.Contains("法務"))
            return "法務部向け";
        if (all.Contains("マニュアル") || all.Contains("手順書") || all.Contains("研修"))
            return "教育・研修";
        if (cat.Contains("コンサルタント") || all.Contains("ワークショップ")
            || all.Contains("ヒアリング結果") || all.Contains("面談準備"))
            return "コンサルタント";
        if (cat.Contains("システム") || all.Contains("要件定義") || all.Contains("テストケース")
            || all.Contains("移行計画") || all.Contains("障害報告") || all.Contains("api仕様")
            || all.Contains("rfp") || all.Contains("セキュリティチェック"))
            return "システム部門";

        return "全般";
    }

    // ── Font Size Control ──

    private void FontSizeUp_Click(object sender, RoutedEventArgs e)
    {
        var size = Math.Min(PromptBox.FontSize + 1, 24);
        PromptBox.FontSize = size;
        FontSizeLabel.Text = size.ToString();
    }

    private void FontSizeDown_Click(object sender, RoutedEventArgs e)
    {
        var size = Math.Max(PromptBox.FontSize - 1, 9);
        PromptBox.FontSize = size;
        FontSizeLabel.Text = size.ToString();
    }

    private void RefreshCategoryComboBox()
    {
        var allPresets = _presetService.LoadAll();
        var categories = allPresets
            .Select(p => p.Category)
            .Where(c => !string.IsNullOrWhiteSpace(c))
            .Distinct()
            .OrderBy(c => c)
            .ToList();
        CategoryCombo.ItemsSource = categories;
    }

    private sealed record ModelItem(string ModelId, string DisplayName, bool IsProviderConfigured);
}
