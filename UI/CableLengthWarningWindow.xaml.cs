using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;

namespace Connections.UI
{
    public partial class CableLengthWarningWindow : Window
    {
        public CableLengthWarningWindow(IEnumerable<string> warnings)
        {
            InitializeComponent();
            WarningsListBox.ItemsSource = warnings;
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Header_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
                DragMove();
        }
    }
}
