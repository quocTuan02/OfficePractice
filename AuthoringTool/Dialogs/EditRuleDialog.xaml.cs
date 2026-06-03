using System.Windows;
using System.Windows.Controls;
using CommonLibrary.Models;

namespace AuthoringTool.Dialogs;

public partial class EditRuleDialog : Window
{
    private ScoringRule _rule = new();

    public EditRuleDialog()
    {
        InitializeComponent();
        CboRuleType.SelectedIndex = 0;
    }

    public void SetRule(ScoringRule rule)
    {
        _rule = rule;
        TxtDescription.Text = rule.Description;
        TxtPoints.Text = rule.Points.ToString();

        var typeIndex = (int)rule.RuleType;
        CboRuleType.SelectedIndex = typeIndex;

        var p = rule.Parameters;
        switch (rule.RuleType)
        {
            case RuleType.CellValue:
                TxtCellAddress.Text = p.GetValueOrDefault("CellAddress", "");
                TxtExpectedValue.Text = p.GetValueOrDefault("ExpectedValue", "");
                SetComboByContent(CboCellCompare, p.GetValueOrDefault("CompareType", "Equals"));
                break;
            case RuleType.Formula:
                TxtFormulaCell.Text = p.GetValueOrDefault("CellAddress", "");
                TxtExpectedFormula.Text = p.GetValueOrDefault("ExpectedFormula", "");
                break;
            case RuleType.Formatting:
                TxtFmtCellAddress.Text = p.GetValueOrDefault("CellAddress", "");
                SetComboByContent(CboFmtProperty, p.GetValueOrDefault("Property", "Bold"));
                TxtFmtExpectedValue.Text = p.GetValueOrDefault("ExpectedValue", "");
                break;
            case RuleType.ObjectExistence:
                SetComboByContent(CboObjectType, p.GetValueOrDefault("ObjectType", "Chart"));
                TxtObjectName.Text = p.GetValueOrDefault("ObjectName", "");
                break;
            case RuleType.RangeComparison:
                TxtRangeAddress.Text = p.GetValueOrDefault("RangeAddress", "");
                TxtRangeReference.Text = p.GetValueOrDefault("Reference", "");
                TxtRangeTolerance.Text = p.GetValueOrDefault("Tolerance", "0");
                break;
        }
    }

    public ScoringRule GetResult()
    {
        var rule = new ScoringRule
        {
            RuleId = _rule.RuleId,
            TaskId = _rule.TaskId,
            Description = TxtDescription.Text.Trim(),
            Points = int.TryParse(TxtPoints.Text, out var p) ? p : 0,
            RuleType = (RuleType)CboRuleType.SelectedIndex,
            Parameters = BuildParameters()
        };
        return rule;
    }

    public void ApplyResult(ScoringRule target)
    {
        target.Description = TxtDescription.Text.Trim();
        target.Points = int.TryParse(TxtPoints.Text, out var p) ? p : 0;
        target.RuleType = (RuleType)CboRuleType.SelectedIndex;
        target.Parameters = BuildParameters();
    }

    private Dictionary<string, string> BuildParameters()
    {
        var ruleType = (RuleType)CboRuleType.SelectedIndex;
        return ruleType switch
        {
            RuleType.CellValue => new Dictionary<string, string>
            {
                ["CellAddress"] = TxtCellAddress.Text.Trim(),
                ["ExpectedValue"] = TxtExpectedValue.Text.Trim(),
                ["CompareType"] = (CboCellCompare.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "Equals"
            },
            RuleType.Formula => new Dictionary<string, string>
            {
                ["CellAddress"] = TxtFormulaCell.Text.Trim(),
                ["ExpectedFormula"] = TxtExpectedFormula.Text.Trim()
            },
            RuleType.Formatting => new Dictionary<string, string>
            {
                ["CellAddress"] = TxtFmtCellAddress.Text.Trim(),
                ["Property"] = (CboFmtProperty.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "Bold",
                ["ExpectedValue"] = TxtFmtExpectedValue.Text.Trim()
            },
            RuleType.ObjectExistence => new Dictionary<string, string>
            {
                ["ObjectType"] = (CboObjectType.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "Chart",
                ["ObjectName"] = TxtObjectName.Text.Trim()
            },
            RuleType.RangeComparison => new Dictionary<string, string>
            {
                ["RangeAddress"] = TxtRangeAddress.Text.Trim(),
                ["Reference"] = TxtRangeReference.Text.Trim(),
                ["Tolerance"] = TxtRangeTolerance.Text.Trim()
            },
            _ => new Dictionary<string, string>()
        };
    }

    private void RuleType_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (!IsLoaded) return;

        PanelCellValue.Visibility = Visibility.Collapsed;
        PanelFormula.Visibility = Visibility.Collapsed;
        PanelFormatting.Visibility = Visibility.Collapsed;
        PanelObjectExistence.Visibility = Visibility.Collapsed;
        PanelRangeComparison.Visibility = Visibility.Collapsed;

        var visible = CboRuleType.SelectedIndex switch
        {
            0 => PanelCellValue,
            1 => PanelFormula,
            2 => PanelFormatting,
            3 => PanelObjectExistence,
            4 => PanelRangeComparison,
            _ => PanelCellValue
        };
        visible.Visibility = Visibility.Visible;
    }

    private static void SetComboByContent(ComboBox combo, string content)
    {
        foreach (ComboBoxItem item in combo.Items)
        {
            if (item.Content?.ToString() == content)
            {
                combo.SelectedItem = item;
                return;
            }
        }
        combo.SelectedIndex = 0;
    }

    private void Ok_Click(object sender, RoutedEventArgs e) => DialogResult = true;
    private void Cancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;
}
