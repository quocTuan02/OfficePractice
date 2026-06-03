using System.Windows;
using ScoringEngine.Models;

namespace MainApp
{
    public partial class ScoringResultWindow : Window
    {
        public ScoringResultWindow(TaskScoreResult result)
        {
            InitializeComponent();

            TaskLabel.Text     = $"Task {result.TaskNumber} – {result.TaskObjective}";
            ScoreText.Text     = $"{result.TotalPoints}/{result.MaxPoints}";
            SummaryText.Text   = result.AllPassed ? "Hoàn thành xuất sắc!" : "Cần cải thiện";
            SubSummaryText.Text = result.Summary;
            ResultsList.ItemsSource = result.RuleResults;
        }

        private void Close_Click(object sender, RoutedEventArgs e) => Close();
    }
}
