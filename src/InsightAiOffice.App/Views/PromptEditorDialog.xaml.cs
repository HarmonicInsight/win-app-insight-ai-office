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

    public bool HasChanges { get; private set; }
    public string? ExecutePromptText { get; private set; }
    public string? ExecuteModelId { get; private set; }

    public PromptEditorDialog(PromptPresetService presetService, AiProviderConfig? providerConfig = null)
    {
        InitializeComponent();
        _presetService = presetService;

        _configuredProviders = BuildConfiguredProviders(providerConfig);

        var models = AiModelRegistry.GetModelsWithCapability(AiCapability.Chat)
            .Select(m => new ModelItem(
                m.Id,
                m.DisplayName,
                _configuredProviders.Contains(m.Provider)))
            .ToList();
        ModelCombo.ItemsSource = models;
        ModelCombo.DisplayMemberPath = "DisplayName";
        ModelCombo.SelectedValuePath = "ModelId";
        ModelCombo.SelectedValue = ClaudeModels.DefaultModel;

        ApplyLocalization();
        RefreshCategoryComboBox();
        BuildTree();
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
        EditorTitle.Text = L("PE_Title");
        NameLabel.Text = L("PE_Name");
        CategoryLabel.Text = L("PE_Category");
        ModelLabel.Text = L("PE_Model");
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

    private void BuildTree()
    {
        PromptTree.Items.Clear();
        var L = Helpers.LanguageManager.Get;

        var allPresets = _presetService.LoadAll();

        // My Prompts (custom) — displayed first
        var customHeader = new TreeViewItem
        {
            Header = L("PE_Custom"),
            IsExpanded = true,
            FontWeight = FontWeights.SemiBold,
            FontSize = 12,
        };
        var custom = allPresets.Where(p => !p.Id.StartsWith("builtin_", StringComparison.Ordinal)).ToList();
        foreach (var preset in custom)
        {
            var label = preset.Name;
            if (preset.IsDefault) label = "\u2605 " + label;
            if (preset.IsPinned) label = "\uD83D\uDCCC " + label;
            customHeader.Items.Add(new TreeViewItem
            {
                Header = label,
                Tag = preset,
                FontWeight = FontWeights.Normal,
                FontSize = 11,
            });
        }
        PromptTree.Items.Add(customHeader);

        // Built-in presets grouped by product → subcategory
        var builtIn = allPresets.Where(p => p.Id.StartsWith("builtin_", StringComparison.Ordinal)).ToList();

        // カテゴリを「製品グループ : サブカテゴリ」に分離
        // 例: "📊 Slide: 品質レビュー" → group="📊 Slide", sub="品質レビュー"
        // 例: "分析・要約" → group="Office 共通", sub="分析・要約"
        var grouped = builtIn
            .GroupBy(p =>
            {
                var cat = p.Category ?? "";
                var colonIdx = cat.IndexOf(':');
                if (colonIdx > 0 && (cat.StartsWith("📊") || cat.StartsWith("📄") || cat.StartsWith("📗")))
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
                Header = productGroup.Key,
                IsExpanded = false,
                FontWeight = FontWeights.SemiBold,
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
                    IsExpanded = false,
                    FontWeight = FontWeights.Normal,
                    FontSize = 11,
                };
                foreach (var preset in sub)
                {
                    subItem.Items.Add(new TreeViewItem
                    {
                        Header = preset.Name,
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
            return;
        }

        _selectedPreset = preset;
        EditorPanel.IsEnabled = true;
        NameBox.Text = preset.Name;
        PromptBox.Text = preset.SystemPrompt;
        CategoryCombo.Text = preset.Category ?? "";
        ExecuteButton.IsEnabled = true;

        var modelId = preset.ModelId;
        if (string.IsNullOrEmpty(modelId))
            modelId = ClaudeModels.DefaultModel;
        ModelCombo.SelectedValue = modelId;
        if (ModelCombo.SelectedItem == null)
            ModelCombo.SelectedValue = ClaudeModels.DefaultModel;

        if (preset.Id.StartsWith("builtin_", StringComparison.Ordinal))
        {
            _isPresetSelected = true;
            NameBox.IsReadOnly = true;
            PromptBox.IsReadOnly = true;
            CategoryCombo.IsEnabled = false;
            ModelCombo.IsEnabled = false;
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
            ModelCombo.IsEnabled = true;
            SaveButton.IsEnabled = true;
            DeleteButton.IsEnabled = true;
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

        var savedModelId = ModelCombo.SelectedValue as string ?? ClaudeModels.DefaultModel;
        var updated = new UserPromptPreset
        {
            Id = _selectedPreset.Id,
            Name = NameBox.Text.Trim(),
            SystemPrompt = PromptBox.Text.Trim(),
            Description = _selectedPreset.Description,
            Author = _selectedPreset.Author,
            Category = (CategoryCombo.Text ?? "").Trim(),
            Mode = "check",
            ModelId = savedModelId,
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

    private void Customize_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedPreset == null) return;

        var customSuffix = Helpers.LanguageManager.Get("PE_CustomSuffix");
        var selectedModelId = ModelCombo.SelectedValue as string ?? ClaudeModels.DefaultModel;
        var newPreset = new UserPromptPreset
        {
            Id = PromptPresetService.GenerateId(),
            Name = _selectedPreset.Name + " (" + customSuffix + ")",
            SystemPrompt = _selectedPreset.SystemPrompt,
            Description = _selectedPreset.Description,
            Author = "",
            Category = _selectedPreset.Category,
            Mode = _selectedPreset.Mode,
            ModelId = selectedModelId,
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
        ExecuteModelId = ModelCombo.SelectedValue as string ?? ClaudeModels.DefaultModel;
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
        var allPresets = _presetService.LoadAll();
        var custom = allPresets.Where(p => !p.Id.StartsWith("builtin_", StringComparison.Ordinal)).ToList();
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
