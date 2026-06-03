using System.Windows;

namespace AuthoringTool.Dialogs;

public partial class NewTestDialog : Window
{
    public string TestName { get; private set; } = string.Empty;
    public string TestDescription { get; private set; } = string.Empty;
    public int TotalTimeMinutes { get; private set; } = 60;

    public NewTestDialog() => InitializeComponent();

    private void Ok_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(TxtName.Text))
        {
            MessageBox.Show("Vui lòng nhập tên bài thi.", "Thiếu thông tin",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            TxtName.Focus();
            return;
        }

        TestName = TxtName.Text.Trim();
        TestDescription = TxtDesc.Text.Trim();
        TotalTimeMinutes = int.TryParse(TxtTime.Text, out var t) ? t : 60;
        DialogResult = true;
    }

    private void Cancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;
}
