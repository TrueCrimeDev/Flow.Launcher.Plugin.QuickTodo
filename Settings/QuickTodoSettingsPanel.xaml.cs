using System.Windows;
using System.Windows.Controls;

namespace Flow.Launcher.Plugin.QuickTodo.Settings;

public partial class QuickTodoSettingsPanel : UserControl
{
    private readonly SettingsViewModel _viewModel;

    public QuickTodoSettingsPanel(SettingsViewModel viewModel)
    {
        InitializeComponent();
        _viewModel = viewModel;
        DataContext = viewModel;
    }

    private void AddCategory_Click(object sender, RoutedEventArgs e)
    {
        _viewModel.AddCategory();
    }

    private void RemoveCategory_Click(object sender, RoutedEventArgs e)
    {
        if (CategoryList.SelectedItem is string selected)
            _viewModel.RemoveCategory(selected);
    }
}
