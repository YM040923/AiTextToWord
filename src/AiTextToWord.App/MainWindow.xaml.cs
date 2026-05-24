using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;

namespace AiTextToWord.App;

public sealed partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }

    private void InputTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        StatusText.Text = string.IsNullOrWhiteSpace(InputTextBox.Text) ? string.Empty : "Preview will be available after conversion is wired.";
    }

    private void Paste_Click(object sender, RoutedEventArgs e)
    {
        StatusText.Text = "Paste support will be wired in the UI task.";
    }

    private void Clear_Click(object sender, RoutedEventArgs e)
    {
        InputTextBox.Text = string.Empty;
        PreviewPanel.Children.Clear();
        StatusText.Text = string.Empty;
    }

    private void ExportWord_Click(object sender, RoutedEventArgs e)
    {
        StatusText.Text = "Export support will be wired in the UI task.";
    }
}
