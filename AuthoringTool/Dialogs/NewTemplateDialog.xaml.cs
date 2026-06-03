using System.Windows;
using CommonLibrary.Models;

namespace AuthoringTool.Dialogs;

public partial class NewTemplateDialog : Window
{
    public string TemplateName { get; private set; } = string.Empty;
    public string TemplateDescription { get; private set; } = string.Empty;
    public OfficeType SelectedOfficeType { get; private set; } = OfficeType.Excel;

    public NewTemplateDialog() => InitializeComponent();

    private void Ok_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(TxtName.Text))
        {
            MessageBox.Show("Vui lòng nhập tên template.", "Thiếu thông tin",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            TxtName.Focus();
            return;
        }

        TemplateName = TxtName.Text.Trim();
        SelectedOfficeType = CboType.SelectedIndex switch
        {
            1 => OfficeType.Word,
            2 => OfficeType.PowerPoint,
            _ => OfficeType.Excel
        };

        DialogResult = true;
    }

    private void Cancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;
}
