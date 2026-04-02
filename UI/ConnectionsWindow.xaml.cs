using Autodesk.Revit.UI;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;

namespace Connections.UI
{
    public partial class ConnectionsWindow : Window
    {
        #region Constants / PInvoke

        private const string ConfigFilePath    = @"C:\ProgramData\RK Tools\Connections\config.json";
        private const string WindowLeftKey     = "ConnectionsWindow.Left";
        private const string WindowTopKey      = "ConnectionsWindow.Top";
        private const string WindowWidthKey    = "ConnectionsWindow.Width";
        private const string WindowHeightKey   = "ConnectionsWindow.Height";
        private const string PanelKey          = "ConnectionsWindow.Panel";
        private const string ParamNameKey      = "ConnectionsWindow.ParamName";
        private const string ParamValueKey     = "ConnectionsWindow.ParamValue";
        private const string ConnectModeKey      = "ConnectionsWindow.ConnectIndividually";
        private const string ConnectionLimitKey   = "ConnectionsWindow.ConnectionLimit";
        private const string MaxCableLengthKey    = "ConnectionsWindow.MaxCableLength";

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        #endregion

        #region Fields

        private readonly WindowResizer _windowResizer;
        private bool _isDarkMode = true;
        private bool _isDataLoaded;
        private readonly UIApplication _uiApplication;
        private readonly Services.Revit.RevitExternalEventService _externalEventService;
        private List<PanelItem> _allPanels = new List<PanelItem>();
        private int _sessionConnectionCount;
        private readonly List<Autodesk.Revit.DB.ElementId> _highlightedElementIds = new List<Autodesk.Revit.DB.ElementId>();

        #endregion

        public ConnectionsWindow(UIApplication app, Services.Revit.RevitExternalEventService externalEventService)
        {
            _uiApplication       = app;
            _externalEventService = externalEventService;

            InitializeComponent();

            DataContext = this;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            DeferWindowShow();

            _windowResizer = new WindowResizer(this);
            Closed += MainWindow_Closed;
            MouseLeftButtonUp += Window_MouseLeftButtonUp;

            LoadThemeState();
            LoadWindowState();
            LoadPanels();
            LoadSavedState();

            _isDataLoaded = true;
            TryShowWindow();
        }

        #region Panel Loading

        private void LoadPanels()
        {
            try
            {
                var doc = _uiApplication.ActiveUIDocument?.Document;
                if (doc == null) return;

                var collector = new Autodesk.Revit.DB.FilteredElementCollector(doc)
                    .OfCategory(Autodesk.Revit.DB.BuiltInCategory.OST_ElectricalEquipment)
                    .WhereElementIsNotElementType();

                _allPanels = collector
                    .Cast<Autodesk.Revit.DB.FamilyInstance>()
                    .Select(fi => new PanelItem(fi.Name, fi.Id))
                    .OrderBy(p => p.Name)
                    .ToList();

                PanelComboBox.ItemsSource = _allPanels;
                PanelComboBox.DisplayMemberPath = "Name";
            }
            catch (Exception ex)
            {
                App.Logger?.Error("Failed to load panels", ex);
            }
        }

        private void PanelComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            UpdateExistingConnectionsCount();
            SaveState();
        }

        private void UpdateExistingConnectionsCount()
        {
            ExistingConnectionsText.Text = string.Empty;
            try
            {
                if (!(PanelComboBox.SelectedItem is PanelItem selectedPanel)) return;

                var doc = _uiApplication.ActiveUIDocument?.Document;
                if (doc == null) return;

                int count = new Autodesk.Revit.DB.FilteredElementCollector(doc)
                    .OfClass(typeof(Autodesk.Revit.DB.Electrical.ElectricalSystem))
                    .Cast<Autodesk.Revit.DB.Electrical.ElectricalSystem>()
                    .Count(sys => sys.BaseEquipment?.Id == selectedPanel.ElementId);

                ExistingConnectionsText.Text = $"Existing connections: {count}";
            }
            catch { }
        }

        #endregion

        #region Connection Mode Toggle

        private void ConnectIndividuallyCheck_Checked(object sender, RoutedEventArgs e)
        {
            if (ConnectCombinedCheck != null) ConnectCombinedCheck.IsChecked = false;
            SaveState();
        }

        private void ConnectCombinedCheck_Checked(object sender, RoutedEventArgs e)
        {
            if (ConnectIndividuallyCheck != null) ConnectIndividuallyCheck.IsChecked = false;
            SaveState();
        }

        #endregion

        #region Select & Connect Button

        private void SelectConnectButton_Click(object sender, RoutedEventArgs e)
        {
            // Validate panel selection
            if (!(PanelComboBox.SelectedItem is PanelItem selectedPanel))
            {
                ResultText.Text = "Please select a valid panel.";
                return;
            }

            // Check connection limit
            int limit = GetConnectionLimit();
            if (limit > 0 && _sessionConnectionCount >= limit)
            {
                ResultText.Text = $"Connection limit reached ({limit}). Clear the counter to continue.";
                return;
            }

            string paramName  = CircuitParamNameBox.Text?.Trim();
            string paramValue = CircuitParamValueBox.Text;
            bool connectIndividually = ConnectIndividuallyCheck.IsChecked == true;
            double maxCableLength = GetMaxCableLength();
            int connectionLimit  = GetConnectionLimit();

            // Persist state
            SaveState();

            ResultText.Text = string.Empty;
            SelectConnectButton.IsEnabled = false;

            // Hide the window so Revit's selection mode is not obstructed
            this.Hide();

            var request = new Services.Revit.ConnectToPanelRequest(
                selectedPanel.ElementId,
                connectIndividually,
                paramName,
                paramValue,
                maxCableLength,
                connectionLimit,
                (result) =>
                {
                    this.Show();
                    this.Activate();
                    ResultText.Text = result;
                    SelectConnectButton.IsEnabled = true;

                    // Parse success count from result and update session counter
                    UpdateSessionCounter(result);

                    // Refresh existing connections count
                    UpdateExistingConnectionsCount();

                    // Show cable length warning popup if any warnings are present
                    ShowCableLengthWarningsIfAny(result);
                },
                (ids) =>
                {
                    foreach (var id in ids)
                        if (!_highlightedElementIds.Contains(id))
                            _highlightedElementIds.Add(id);
                    ClearWarningsButton.Visibility = Visibility.Visible;
                });

            _externalEventService.Raise(request);
        }

        #endregion

        #region Session Counter

        private void UpdateSessionCounter(string result)
        {
            // Parse "X succeeded" from the result text
            var match = System.Text.RegularExpressions.Regex.Match(result, @"(\d+)\s+succeeded");
            if (match.Success && int.TryParse(match.Groups[1].Value, out int count) && count > 0)
            {
                _sessionConnectionCount += count;
                UpdateSessionCounterDisplay();
            }
        }

        private void UpdateSessionCounterDisplay()
        {
            int limit = GetConnectionLimit();
            SessionCounterText.Text = limit > 0
                ? $"Connections: {_sessionConnectionCount} / {limit}"
                : $"Connections: {_sessionConnectionCount}";
        }

        private int GetConnectionLimit()
        {
            if (int.TryParse(ConnectionLimitBox.Text, out int limit) && limit > 0)
                return limit;
            return 0;
        }

        private void ClearCounterButton_Click(object sender, RoutedEventArgs e)
        {
            _sessionConnectionCount = 0;
            UpdateSessionCounterDisplay();
        }

        private void ClearWarningsButton_Click(object sender, RoutedEventArgs e)
        {
            if (_highlightedElementIds.Count == 0) return;

            var ids = _highlightedElementIds.ToList();
            var request = new Services.Revit.ClearWarningOverridesRequest(ids, () =>
            {
                _highlightedElementIds.Clear();
                ClearWarningsButton.Visibility = Visibility.Collapsed;
            });
            _externalEventService.Raise(request);
        }

        private void ShowCableLengthWarningsIfAny(string result)
        {
            var warnings = new List<string>();
            foreach (var line in result.Split('\n'))
            {
                var trimmed = line.Trim();
                if (trimmed.StartsWith("⚠"))
                    warnings.Add(trimmed.TrimStart('⚠').Trim());
            }

            if (warnings.Count > 0)
            {
                var popup = new CableLengthWarningWindow(warnings)
                {
                    Owner = this
                };
                popup.Show();
            }
        }

        private void ConnectionLimitBox_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            e.Handled = !int.TryParse(e.Text, out _);
        }

        private void NumericOnly_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            e.Handled = !int.TryParse(e.Text, out _) && !double.TryParse(e.Text,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out _);
        }

        private double GetMaxCableLength()
        {
            if (double.TryParse(MaxCableLengthBox.Text,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out double v) && v > 0)
                return v;
            return 0;
        }

        #endregion

        #region State Persistence

        private void LoadSavedState()
        {
            try
            {
                var config = LoadConfig();

                if (config.TryGetValue(ParamNameKey, out var rawName) && rawName is string s && !string.IsNullOrEmpty(s))
                    CircuitParamNameBox.Text = s;
                if (config.TryGetValue(ParamValueKey, out var rawVal) && rawVal is string v && !string.IsNullOrEmpty(v))
                    CircuitParamValueBox.Text = v;
                if (config.TryGetValue(ConnectionLimitKey, out var rawLimit) && rawLimit != null)
                {
                    var limitStr = rawLimit.ToString();
                    if (!string.IsNullOrEmpty(limitStr))
                        ConnectionLimitBox.Text = limitStr;
                }
                if (config.TryGetValue(MaxCableLengthKey, out var rawCableLen) && rawCableLen != null)
                {
                    var cableStr = rawCableLen.ToString();
                    if (!string.IsNullOrEmpty(cableStr))
                        MaxCableLengthBox.Text = cableStr;
                }
                if (TryGetBool(config, ConnectModeKey, out bool isIndividual))
                {
                    ConnectIndividuallyCheck.IsChecked = isIndividual;
                    ConnectCombinedCheck.IsChecked = !isIndividual;
                }
                if (config.TryGetValue(PanelKey, out var rawPanel) && rawPanel is string panelName && !string.IsNullOrEmpty(panelName))
                {
                    var match = _allPanels.FirstOrDefault(p => p.Name == panelName);
                    if (match != null)
                        PanelComboBox.SelectedItem = match;
                    else
                        PanelComboBox.Text = panelName;
                }
            }
            catch { }
        }

        private void SaveState()
        {
            if (!_isDataLoaded) return;
            try
            {
                var cfg = LoadConfig();
                cfg[ParamNameKey]   = CircuitParamNameBox.Text ?? string.Empty;
                cfg[ParamValueKey]  = CircuitParamValueBox.Text ?? string.Empty;
                cfg[ConnectModeKey]      = ConnectIndividuallyCheck.IsChecked == true;
                cfg[ConnectionLimitKey]  = ConnectionLimitBox.Text ?? "0";
                cfg[MaxCableLengthKey]   = MaxCableLengthBox.Text ?? "0";

                string panelName = (PanelComboBox.SelectedItem as PanelItem)?.Name ?? string.Empty;
                cfg[PanelKey] = panelName ?? string.Empty;

                SaveConfig(cfg);
            }
            catch { }
        }

        #endregion

        #region Window chrome / resize handlers

        private void MainWindow_Closed(object sender, EventArgs e)
        {
            SaveWindowState();
            SaveState();
        }

        private void TitleBar_Loaded(object sender, RoutedEventArgs e) { }

        private void LeftEdge_MouseEnter(object sender, MouseEventArgs e)         => Cursor = Cursors.SizeWE;
        private void RightEdge_MouseEnter(object sender, MouseEventArgs e)        => Cursor = Cursors.SizeWE;
        private void BottomEdge_MouseEnter(object sender, MouseEventArgs e)       => Cursor = Cursors.SizeNS;
        private void Edge_MouseLeave(object sender, MouseEventArgs e)             => Cursor = Cursors.Arrow;
        private void BottomLeftCorner_MouseEnter(object sender, MouseEventArgs e) => Cursor = Cursors.SizeNESW;
        private void BottomRightCorner_MouseEnter(object sender, MouseEventArgs e)=> Cursor = Cursors.SizeNWSE;

        private void Window_MouseMove(object sender, MouseEventArgs e)                               => _windowResizer.ResizeWindow(e);
        private void Window_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)                 => _windowResizer.StopResizing();
        private void LeftEdge_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)             => _windowResizer.StartResizing(e, ResizeDirection.Left);
        private void RightEdge_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)            => _windowResizer.StartResizing(e, ResizeDirection.Right);
        private void BottomEdge_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)           => _windowResizer.StartResizing(e, ResizeDirection.Bottom);
        private void BottomLeftCorner_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)     => _windowResizer.StartResizing(e, ResizeDirection.BottomLeft);
        private void BottomRightCorner_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)    => _windowResizer.StartResizing(e, ResizeDirection.BottomRight);

        private void Window_PreviewMouseDown(object sender, MouseButtonEventArgs e) { }
        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)           { }
        private void Window_PreviewKeyUp(object sender, KeyEventArgs e)             { }

        #endregion

        #region Window Startup

        private void DeferWindowShow()
        {
            Opacity = 0;
            Loaded += ConnectionsWindow_Loaded;
        }

        private void ConnectionsWindow_Loaded(object sender, RoutedEventArgs e)
        {
            TryShowWindow();
        }

        private void TryShowWindow()
        {
            if (!_isDataLoaded) return;
            Opacity = 1;
        }

        #endregion

        #region Theme

        private ResourceDictionary _currentThemeDictionary;

        private void LoadTheme()
        {
            try
            {
                var themeUri = new Uri(_isDarkMode
                    ? "pack://application:,,,/Connections;component/UI/Themes/DarkTheme.xaml"
                    : "pack://application:,,,/Connections;component/UI/Themes/LightTheme.xaml",
                    UriKind.Absolute);

                var newDict = new ResourceDictionary { Source = themeUri };

                if (_currentThemeDictionary != null)
                    Resources.MergedDictionaries.Remove(_currentThemeDictionary);

                Resources.MergedDictionaries.Add(newDict);
                _currentThemeDictionary = newDict;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading theme: {ex.Message}");
            }
        }

        private void ToggleTheme_Click(object sender, RoutedEventArgs e)
        {
            _isDarkMode = ThemeToggleButton.IsChecked == true;
            LoadTheme();
            SaveThemeState();

            var icon = ThemeToggleButton?.Template?.FindName("ThemeToggleIcon", ThemeToggleButton)
                       as MaterialDesignThemes.Wpf.PackIcon;
            if (icon != null)
            {
                icon.Kind = _isDarkMode
                    ? MaterialDesignThemes.Wpf.PackIconKind.ToggleSwitchOffOutline
                    : MaterialDesignThemes.Wpf.PackIconKind.ToggleSwitchOutline;
            }
        }

        private void LoadThemeState()
        {
            try
            {
                var config = LoadConfig();
                if (TryGetBool(config, "IsDarkMode", out bool isDark))
                    _isDarkMode = isDark;
            }
            catch { }

            if (ThemeToggleButton != null)
            {
                ThemeToggleButton.IsChecked = _isDarkMode;
                var icon = ThemeToggleButton.Template?.FindName("ThemeToggleIcon", ThemeToggleButton)
                           as MaterialDesignThemes.Wpf.PackIcon;
                if (icon != null)
                {
                    icon.Kind = _isDarkMode
                        ? MaterialDesignThemes.Wpf.PackIconKind.ToggleSwitchOffOutline
                        : MaterialDesignThemes.Wpf.PackIconKind.ToggleSwitchOutline;
                }
            }

            LoadTheme();
        }

        private void SaveThemeState()
        {
            try
            {
                var config = LoadConfig();
                config["IsDarkMode"] = _isDarkMode;
                SaveConfig(config);
            }
            catch { }
        }

        #endregion

        #region Window State

        private void LoadWindowState()
        {
            try
            {
                var config   = LoadConfig();
                bool hasLeft  = TryGetDouble(config, WindowLeftKey,   out double left);
                bool hasTop   = TryGetDouble(config, WindowTopKey,    out double top);
                bool hasWidth = TryGetDouble(config, WindowWidthKey,  out double width);
                bool hasHeight= TryGetDouble(config, WindowHeightKey, out double height);

                bool hasSize = hasWidth && hasHeight && width > 0 && height > 0;
                bool hasPos  = hasLeft  && hasTop   && !double.IsNaN(left) && !double.IsNaN(top);

                if (!hasSize && !hasPos) return;

                WindowStartupLocation = WindowStartupLocation.Manual;

                if (hasSize)
                {
                    Width  = Math.Max(MinWidth,  width);
                    Height = Math.Max(MinHeight, height);
                }

                if (hasPos)
                {
                    Left = left;
                    Top  = top;
                }
            }
            catch { }
        }

        private void SaveWindowState()
        {
            try
            {
                var config = LoadConfig();
                var bounds = WindowState == WindowState.Normal
                    ? new Rect(Left, Top, Width, Height)
                    : RestoreBounds;

                config[WindowLeftKey]   = bounds.Left;
                config[WindowTopKey]    = bounds.Top;
                config[WindowWidthKey]  = bounds.Width;
                config[WindowHeightKey] = bounds.Height;

                SaveConfig(config);
            }
            catch { }
        }

        #endregion

        #region Config Helpers

        private Dictionary<string, object> LoadConfig()
        {
            try
            {
                if (File.Exists(ConfigFilePath))
                {
                    var json   = File.ReadAllText(ConfigFilePath);
                    var config = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                    if (config != null) return config;
                }
            }
            catch { }

            return new Dictionary<string, object>();
        }

        private void SaveConfig(Dictionary<string, object> config)
        {
            var dir = Path.GetDirectoryName(ConfigFilePath);
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            File.WriteAllText(ConfigFilePath, JsonConvert.SerializeObject(config, Formatting.Indented));
        }

        private static bool TryGetBool(Dictionary<string, object> config, string key, out bool value)
        {
            value = false;
            if (!config.TryGetValue(key, out var raw) || raw == null) return false;

            if (raw is bool boolVal)                                   { value = boolVal; return true; }
            if (raw is Newtonsoft.Json.Linq.JToken t && t.Type == Newtonsoft.Json.Linq.JTokenType.Boolean)
                                                                       { value = (bool)t; return true; }
            if (raw is string s && bool.TryParse(s, out var parsed))   { value = parsed; return true; }

            return false;
        }

        private static bool TryGetDouble(Dictionary<string, object> config, string key, out double value)
        {
            value = 0;
            if (!config.TryGetValue(key, out var raw) || raw == null) return false;

            switch (raw)
            {
                case double d:   value = d;         return true;
                case float  f:   value = f;         return true;
                case decimal m:  value = (double)m; return true;
                case long   l:   value = l;         return true;
                case int    i:   value = i;         return true;
                case Newtonsoft.Json.Linq.JToken t when t.Type == Newtonsoft.Json.Linq.JTokenType.Float
                    || t.Type == Newtonsoft.Json.Linq.JTokenType.Integer:
                    value = (double)t; return true;
                case string s when double.TryParse(s, NumberStyles.Float,
                    CultureInfo.InvariantCulture, out var p):
                    value = p; return true;
            }

            return false;
        }

        #endregion
    }

    public sealed class PanelItem
    {
        public string Name { get; }
        public Autodesk.Revit.DB.ElementId ElementId { get; }

        public PanelItem(string name, Autodesk.Revit.DB.ElementId elementId)
        {
            Name = name;
            ElementId = elementId;
        }

        public override string ToString() => Name;
    }
}
