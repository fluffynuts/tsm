using System.Collections.Concurrent;
using System.Data;
using System.Diagnostics;
using NStack;
using PeanutButter.INI;
using PeanutButter.Utils;
using PeanutButter.WindowsServiceManagement;
using PeanutButter.WindowsServiceManagement.Exceptions;
using Terminal.Gui;
using Attribute = Terminal.Gui.Attribute;

// ReSharper disable VirtualMemberCallInConstructor

namespace services
{
    public class MainView : Window
    {
        private const int ID_COLUMN = 0;
        private const int NAME_COLUMN = 1;
        private const int STATE_COLUMN = 2;
        private const int AUTO_REFRESH_INTERVAL_MS = 100;
        private const int BACKGROUND_REFRESH_INTERVAL_MS = 5000;
        public TableView ServicesList { get; set; }
        public DataTable FilteredServices { get; set; }
        public DataTable AllServices { get; set; }
        public Label SearchLabel { get; set; }
        public TextField SearchBox { get; set; }

        private static readonly ConcurrentDictionary<string, int> FilteredServiceIdToRowMap = new();
        private static readonly ConcurrentDictionary<string, int> ServiceIdToRowMap = new();

        private static readonly ColorScheme DefaultColorScheme = new()
        {
            Normal = new Attribute(Color.Green, Color.Black),
            HotNormal = new Attribute(Color.BrightGreen, Color.Black),
            Focus = new Attribute(Color.White, Color.Black),
            HotFocus = new Attribute(Color.BrightYellow, Color.Blue),
            Disabled = new Attribute(Color.Gray, Color.Black)
        };

        private const string SECTION_DEFAULT_COLOR = "Color.Defaults";
        private const string SECTION_STATE_COLOR = "Color.States";

        private ColorScheme _runningScheme = DefaultColorScheme;
        private ColorScheme _stoppedScheme = DefaultColorScheme;
        private ColorScheme _pausedScheme = DefaultColorScheme;
        private ColorScheme _pendingScheme = DefaultColorScheme;

        private Color DefaultRunningColor = Color.BrightYellow;
        private Color DefaultStoppedColor = Color.BrightRed;
        private Color DefaultPausedColor = Color.Cyan;
        private Color DefaultIntermediateColor = Color.DarkGray;


        private void LoadColorScheme()
        {
            ColorScheme = DefaultColorScheme;
            if (_config.HasSection(SECTION_DEFAULT_COLOR))
            {
                var section = _config.GetSection(SECTION_DEFAULT_COLOR);
                ColorScheme = new()
                {
                    Normal = MakeAttribute(section, nameof(ColorScheme.Normal), DefaultColorScheme.Normal),
                    HotNormal = MakeAttribute(section, nameof(ColorScheme.HotNormal), DefaultColorScheme.HotNormal),
                    Focus = MakeAttribute(section, nameof(ColorScheme.Focus), DefaultColorScheme.Focus),
                    HotFocus = MakeAttribute(section, nameof(ColorScheme.HotFocus), DefaultColorScheme.Focus),
                    Disabled = MakeAttribute(section, nameof(ColorScheme.Disabled), DefaultColorScheme.Disabled)
                };
            }

            var defaultBackground = DefaultColorScheme.Normal.Background;
            _runningScheme = MakeScheme(DefaultRunningColor, defaultBackground);
            _stoppedScheme = MakeScheme(DefaultStoppedColor, defaultBackground);
            _pausedScheme = MakeScheme(DefaultPausedColor, defaultBackground);
            _pendingScheme = MakeScheme(DefaultIntermediateColor, defaultBackground);


            if (_config.HasSection(SECTION_STATE_COLOR))
            {
                var runningColor =
                    _config.GetValue(SECTION_STATE_COLOR, $"{ServiceState.Running}", $"{Color.BrightGreen}");
                var stoppedColor =
                    _config.GetValue(SECTION_STATE_COLOR, $"{ServiceState.Stopped}", $"{Color.BrightRed}");
                var pausedColor = _config.GetValue(SECTION_STATE_COLOR, $"{ServiceState.Paused}", $"{Color.Cyan}");
                var pendingColor = _config.GetValue(SECTION_STATE_COLOR, "pending", $"{Color.DarkGray}");

                _runningScheme = MakeScheme(runningColor, defaultBackground);
                _stoppedScheme = MakeScheme(stoppedColor, defaultBackground);
                _pausedScheme = MakeScheme(pausedColor, defaultBackground);
                _pendingScheme = MakeScheme(pendingColor, defaultBackground);
            }
        }

        private Attribute MakeAttribute(
            IDictionary<string, string> section,
            string group,
            Attribute defaults
        )
        {
            var fg = section.TryGetValue($"{group}.fg", out var userFg) &&
                Enum.TryParse<Color>(userFg, out var fgColor)
                    ? fgColor
                    : defaults.Foreground;
            var bg = section.TryGetValue($"{group}.bg", out var userBg) &&
                Enum.TryParse<Color>(userBg, out var bgColor)
                    ? bgColor
                    : defaults.Background;
            return new Attribute(fg, bg);
        }

        private ColorScheme MakeScheme(
            string foreground,
            Color background
        )
        {
            if (!Enum.TryParse<Color>(foreground, out var fgColor))
            {
                return DefaultColorScheme;
            }

            return MakeScheme(fgColor, background);
        }

        private static ColorScheme MakeScheme(
            Color foreground,
            Color background
        )
        {
            return new ColorScheme()
            {
                Normal = new Attribute(foreground, background),
                HotNormal = new Attribute(foreground, background),
                Focus = new Attribute(foreground, background),
                HotFocus = new Attribute(foreground, background),
                Disabled = new Attribute(foreground, background)
            };
        }

        private IINIFile LoadConfig()
        {
            if (IniFilePath is null)
            {
                return new INIFile();
            }

            _isFirstRun = !File.Exists(IniFilePath);
            return new INIFile(IniFilePath);
        }

        private string? IniFilePath
        {
            get
            {
                if (_iniFilePath is not null)
                {
                    return _iniFilePath;
                }

                var userProfileFolder = Environment.GetEnvironmentVariable("USERPROFILE");
                if (userProfileFolder is null)
                {
                    return null;
                }

                return _iniFilePath ??= Path.Combine(userProfileFolder, "tsm.ini");
            }
        }

        private string? _iniFilePath;

        private readonly IINIFile _config;

        public MainView(Toplevel top)
        {
            _config = LoadConfig();
            LoadColorScheme();
            PersistConfig();

            FilteredServices = new DataTable();
            _stateDataColumn = new DataColumn() { ColumnName = "State" };
            FilteredServices.Columns.Add(new DataColumn() { ColumnName = "Id" });
            FilteredServices.Columns.Add(new DataColumn() { ColumnName = "Name" });
            FilteredServices.Columns.Add(_stateDataColumn);
            AllServices = new DataTable();
            AllServices.Columns.Add(new DataColumn() { ColumnName = "Id" });
            AllServices.Columns.Add(new DataColumn() { ColumnName = "Name" });
            AllServices.Columns.Add(new DataColumn() { ColumnName = "State" });


            StatusBar = new StatusBar()
            {
                Y = Pos.Bottom(top),
                Width = Dim.Percent(100)
            };
            top.Add(StatusBar);
            SearchLabel = new Label()
            {
                Y = 0,
                X = 0,
                Width = 8,
                Text = "Filter:",
            };
            Add(SearchLabel);
            SearchBox = new TextField()
            {
                Y = 0,
                X = 8,
                Width = Dim.Fill(),
            };
            SearchBox.TextChanged += ApplyFilter;
            SearchBox.Enter += _ => ServicesList?.SetFocus();
            Add(SearchBox);

            ServicesList = new TableView()
            {
                Y = 1,
                Width = Dim.Fill(),
                Height = Dim.Fill(1),
                MultiSelect = true,
                FullRowSelect = true,
                Table = FilteredServices,
                Style =
                {
                    AlwaysShowHeaders = true,
                    ShowVerticalCellLines = true,
                    ShowVerticalHeaderLines = true,
                    ColumnStyles = new Dictionary<DataColumn, TableView.ColumnStyle>()
                    {
                        [_stateDataColumn] = new()
                        {
                            ColorGetter = StateColumnColorGetter,
                        }
                    }
                }
            };
            Add(ServicesList);
            UiLock = new SemaphoreSlim(1);
            ClearStatusBar();
            Refresh(this);
            ServicesList.KeyDown += KeyPressHandler;
            ServicesList.SetFocus();
        }

        private void PersistConfig()
        {
            if (IniFilePath is null)
            {
                return;
            }

            // always rewrite the file, in case new settings are added - then it's
            // easy for the user to re-configure
            if (!_config.HasSection(SECTION_DEFAULT_COLOR))
            {
                _config.AddSection(SECTION_DEFAULT_COLOR);
            }

            var section = _config.GetSection(SECTION_DEFAULT_COLOR);
            section[$"{nameof(ColorScheme.Normal)}.fg"] = $"{ColorScheme.Normal.Foreground}";
            section[$"{nameof(ColorScheme.Normal)}.bg"] = $"{ColorScheme.Normal.Background}";
            section[$"{nameof(ColorScheme.HotNormal)}.fg"] = $"{ColorScheme.HotNormal.Foreground}";
            section[$"{nameof(ColorScheme.HotNormal)}.bg"] = $"{ColorScheme.HotNormal.Background}";
            section[$"{nameof(ColorScheme.Focus)}.fg"] = $"{ColorScheme.Focus.Foreground}";
            section[$"{nameof(ColorScheme.Focus)}.bg"] = $"{ColorScheme.Focus.Background}";
            section[$"{nameof(ColorScheme.HotFocus)}.fg"] = $"{ColorScheme.HotFocus.Foreground}";
            section[$"{nameof(ColorScheme.HotFocus)}.bg"] = $"{ColorScheme.HotFocus.Background}";
            section[$"{nameof(ColorScheme.Disabled)}.fg"] = $"{ColorScheme.Disabled.Foreground}";
            section[$"{nameof(ColorScheme.Disabled)}.bg"] = $"{ColorScheme.Disabled.Background}";

            if (!_config.HasSection(SECTION_STATE_COLOR))
            {
                _config.AddSection(SECTION_STATE_COLOR);
            }

            section = _config.GetSection(SECTION_STATE_COLOR);
            section[$"{ServiceState.Running}"] = $"{_runningScheme.Normal.Foreground}";
            section[$"{ServiceState.Paused}"] = $"{_pausedScheme.Normal.Foreground}";
            section[$"{ServiceState.Stopped}"] = $"{_stoppedScheme.Normal.Foreground}";
            section["pending"] = $"{_pendingScheme.Normal.Foreground}";

            _config.Persist();
        }

        private ColorScheme StateColumnColorGetter(TableView.CellColorGetterArgs args)
        {
            var cellValue = $"{args.CellValue}";
            if (cellValue.Contains("Pend") || cellValue.Contains("*"))
            {
                return _pendingScheme;
            }

            switch (args.CellValue)
            {
                case "Stopped":
                    return _stoppedScheme;
                case "Running":
                    return _runningScheme;
                case "Paused":
                    return _pausedScheme;
            }

            return DefaultColorScheme;
        }

        public SemaphoreSlim UiLock { get; }

        protected override void Dispose(bool disposing)
        {
            StopBackgroundRefresh();
            base.Dispose(disposing);
        }

        DateTime _deferBackgroundRefreshUntil = DateTime.MinValue;

        private void ApplyFilter(ustring? obj)
        {
            Application.DoEvents();
            _deferBackgroundRefreshUntil = DateTime.Now.AddSeconds(1);
            Application.MainLoop.Invoke(() =>
            {
                using var _ = new AutoLocker(UiLock);
                ServicesList.SetNeedsDisplay();
                var search = SearchBox.Text?.ToString()?.Split(" ", StringSplitOptions.RemoveEmptyEntries) ??
                    Array.Empty<string>();
                lock (FilteredServices)
                {
                    FilteredServices.Clear();
                    FilteredServiceIdToRowMap.Clear();
                    ServiceIdToRowMap.Clear();
                    for (var i = 0; i < AllServices.Rows.Count; i++)
                    {
                        var row = AllServices.Rows[i];
                        var serviceId = ServiceIdFor(row);
                        var serviceName = ServiceNameFor(row);
                        var serviceState = ServiceStateFor(row);
                        var stringData = $"{serviceId} {serviceName} {serviceState}";
                        ServiceIdToRowMap[serviceId] = i;
                        if (search.Length == 0 ||
                            search.All(s => stringData.Contains(s, StringComparison.OrdinalIgnoreCase)))
                        {
                            FilteredServices.Rows.Add(DataRowFor(row));
                            var index = FilteredServices.Rows.Count - 1;
                            FilteredServiceIdToRowMap[serviceId] = index;
                        }
                    }
                }

                if (ServicesList.SelectedRow < 0)
                {
                    ServicesList.SelectedRow = 0;
                }

                var selected = FindSelectedServiceName(this);
                if (selected is not null)
                {
                    UpdateStatusBarForSelectedService(this, selected);
                }
            });
        }

        private void ClearStatusBar()
        {
            while (StatusBar.Items.Length > 0)
            {
                StatusBar.RemoveItem(0);
            }
        }

        private void ResetStatusBar()
        {
            ClearStatusBar();
            StatusBar.AddItemAt(0,
                new(Key.F5, "F5 Refresh", () => Refresh(this))
            );
            if (!string.IsNullOrWhiteSpace(SearchBox.Text.ToString()))
            {
                StatusBar.AddItemAt(1,
                    new(Key.F6, "F6 Start all", StartAll)
                );
                StatusBar.AddItemAt(2,
                    new(Key.F7, "F7 Restart all", RestartAll)
                );
                StatusBar.AddItemAt(3,
                    new(Key.F8, "F8 Stop all", StopAll)
                );
            }
        }

        private string[] CurrentlyFilteredServices()
        {
            return FilteredServiceIdToRowMap.Keys.ToArray();
        }

        private void StartAll()
        {
            var services = CurrentlyFilteredServices();
            foreach (var service in services)
            {
                RunWithService(service, util =>
                {
                    if (util.State is ServiceState.Running or ServiceState.StartPending)
                    {
                        return;
                    }

                    DoStart(util);
                });
            }
        }

        private void RestartAll()
        {
            var services = CurrentlyFilteredServices();
            foreach (var service in services)
            {
                RunWithService(service, DoRestart);
            }
        }

        private void StopAll()
        {
            var services = CurrentlyFilteredServices();
            foreach (var service in services)
            {
                RunWithService(service, util =>
                {
                    if (util.State is ServiceState.Stopped)
                    {
                        return;
                    }

                    DoStop(util);
                });
            }
        }

        private static IWindowsServiceUtil? FindSelectedService(
            MainView view,
            bool refresh = true
        )
        {
            try
            {
                var serviceName = FindSelectedServiceName(view);
                return serviceName is null
                    ? null
                    : ServiceUtilFor(serviceName, refresh);
            }
            catch
            {
                return null;
            }
        }

        private static string? FindSelectedServiceName(MainView view)
        {
            try
            {
                DataRow? data;
                lock (view.FilteredServices)
                {
                    if (view.ServicesList.SelectedRow < 0 ||
                        view.ServicesList.SelectedRow > view.FilteredServices.Rows.Count)
                    {
                        return null;
                    }

                    data = view.FilteredServices.Rows[view.ServicesList.SelectedRow];
                }

                return data[0] as string;
            }
            catch
            {
                return null;
            }
        }

        private static void UpdateStatusBarForSelectedService(
            MainView mainView,
            string? serviceName
        )
        {
            mainView.ResetStatusBar();
            if (serviceName is null)
            {
                return;
            }

            var util = ServiceUtilFor(serviceName);
            if (util is null)
            {
                return;
            }

            switch (util.State)
            {
                case ServiceState.Unknown:
                    break;
                case ServiceState.NotFound:
                    break;
                case ServiceState.Stopped:
                    InsertStart(mainView);
                    break;
                case ServiceState.StartPending:
                    break;
                case ServiceState.StopPending:
                    break;
                case ServiceState.Running:
                    InsertStop(mainView);
                    InsertRestart(mainView);
                    break;
                case ServiceState.ContinuePending:
                    break;
                case ServiceState.PausePending:
                    break;
                case ServiceState.Paused:
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }

            AppendUninstall(mainView);
        }

        private static void RefreshServiceStateUntilReached(
            MainView mainView,
            string serviceName,
            ServiceState expected,
            bool isFinalState = true
        )
        {
            mainView.ResetStatusBar();
            var current = ServiceState.Unknown;
            // Don't poll forever, just in case
            for (var i = 0; i < 120; i++)
            {
                Trace.WriteLine($"refresh '{serviceName}' until '{expected}' (current: '{current}')");
                if (current == expected)
                {
                    UpdateServiceStateInList(mainView, serviceName, isIntermediateRefresh: !isFinalState, current);
                    DoFinalUpdate();
                    return;
                }

                current = RefreshServiceStatus(mainView, serviceName, isIntermediateRefresh: true);
                Thread.Sleep(AUTO_REFRESH_INTERVAL_MS);
            }

            RefreshServiceStatus(mainView, serviceName, isIntermediateRefresh: false);
            Application.MainLoop.Invoke(() =>
            {
                MessageBox.ErrorQuery(
                    "Unable to change service state",
                    $"Service [{serviceName}] did not respond to the request",
                    defaultButton: 0,
                    "Ok"
                );
            });
            DoFinalUpdate();

            // TODO: get sad?
            void DoFinalUpdate()
            {
                Application.MainLoop.Invoke(() =>
                {
                    mainView.StatusBar.SetNeedsDisplay();
                    UpdateStatusBarForSelectedService(mainView, serviceName);
                });
            }
        }

        private static ServiceState RefreshServiceStatus(
            MainView mainView,
            string serviceName,
            bool isIntermediateRefresh
        )
        {
            var util = ServiceUtilFor(serviceName);
            if (util is null)
            {
                return ServiceState.Unknown;
            }

            var state = util.State;
            if (ServiceIdToRowMap.TryGetValue(serviceName, out var r))
            {
                var current = $"{mainView.AllServices.Rows[r][STATE_COLUMN]}";
                var update = $"{util.State}";
                if (current == update)
                {
                    Trace.WriteLine($"{serviceName} still has state {update}");
                    return util.State;
                }


                mainView.AllServices.Rows[r][STATE_COLUMN] = $"{state}";
            }


            return UpdateServiceStateInList(mainView, serviceName, isIntermediateRefresh, state);
        }

        private static ServiceState UpdateServiceStateInList(
            MainView mainView,
            string serviceName,
            bool isIntermediateRefresh,
            ServiceState state
        )
        {
            if (!FilteredServiceIdToRowMap.TryGetValue(serviceName, out var row))
            {
                return state;
            }

            lock (mainView.FilteredServices)
            {
                if (mainView.FilteredServices.Rows.Count <= row)
                {
                    return state;
                }
            }

            Application.MainLoop.Invoke(() =>
            {
                try
                {
                    mainView.ServicesList.SetNeedsDisplay();
                    var stringState = isIntermediateRefresh
                        ? $"{state}*"
                        : $"{state}";
                    Trace.WriteLine($"Update status: {row} {serviceName} {state}");
                    lock (mainView.FilteredServices)
                    {
                        mainView.FilteredServices.Rows[row][STATE_COLUMN] = stringState;
                    }

                    mainView.StatusBar.SetNeedsDisplay();
                    UpdateStatusBarForSelectedService(mainView, FindSelectedServiceName(mainView));
                }
                catch
                {
                    // suppress
                }
            });

            return state;
        }


        private static void InsertStop(MainView mainView)
        {
            mainView.StatusBar.AddItemAt(
                0,
                new(
                    Key.F3, "F3 Stop", mainView.StopSelected
                )
            );
        }

        private static void AppendUninstall(MainView mainView)
        {
            mainView.StatusBar.AddItemAt(
                mainView.StatusBar.Items.Length,
                new(
                    Key.F12, "F12 Uninstall", mainView.UninstallSelected
                )
            );
        }

        private static void InsertStart(MainView mainView)
        {
            mainView.StatusBar.AddItemAt(
                0,
                new(
                    Key.F1, "F1 Start", mainView.StartSelected
                )
            );
        }

        private static void InsertRestart(MainView mainView)
        {
            mainView.StatusBar.AddItemAt(
                0,
                new(
                    Key.F2, "F2 Restart", mainView.RestartSelected
                )
            );
        }

        public void RunWithSelectedService(Action<IWindowsServiceUtil> toRun)
        {
            Task.Run(() =>
            {
                var util = FindSelectedService(this);
                if (util is null)
                {
                    return;
                }

                try
                {
                    toRun(util);
                }
                catch (Exception ex)
                {
                    Application.MainLoop.Invoke(() =>
                    {
                        MessageBox.ErrorQuery(
                            "Service control failed",
                            ex.Message,
                            defaultButton: 0,
                            "Ok"
                        );
                        RefreshServiceStatus(this, util.ServiceName, false);
                    });
                }
            });
        }

        public void RunWithService(string serviceName, Action<IWindowsServiceUtil> toRun)
        {
            Task.Run(() =>
            {
                var util = ServiceUtilFor(serviceName);
                if (util is null)
                {
                    return;
                }

                try
                {
                    toRun(util);
                }
                catch (Exception ex)
                {
                    Application.MainLoop.Invoke(() =>
                    {
                        MessageBox.ErrorQuery(
                            "Service control failed",
                            ex.Message,
                            defaultButton: 0,
                            "Ok"
                        );
                        RefreshServiceStatus(this, serviceName, false);
                    });
                }
            });
        }

        private void StartSelected()
        {
            RunWithSelectedService(DoStart);
        }

        private void StopSelected()
        {
            RunWithSelectedService(DoStop);
        }

        private void UninstallSelected()
        {
            RunWithSelectedService(DoUninstall);
        }

        private void RestartSelected()
        {
            RunWithSelectedService(DoRestart);
        }

        private void DoUninstall(IWindowsServiceUtil svc)
        {
            int answer = 0;
            Application.MainLoop.Invoke(() =>
            {
                answer = MessageBox.Query(
                    "Confirm uninstall",
                    $"You're about to uninstall '{svc.ServiceName}' ({svc.DisplayName})?",
                    defaultButton: 1,
                    "Ok",
                    "Cancel"
                );
                if (answer == 1)
                {
                    return;
                }

                svc.Uninstall();
                Refresh(this);
            });
        }

        private void DoStart(IWindowsServiceUtil svc)
        {
            svc.Start(wait: false);
            RefreshServiceStateUntilReached(
                this,
                svc.ServiceName,
                ServiceState.Running
            );
        }

        private void DoStop(IWindowsServiceUtil svc)
        {
            svc.Stop(wait: false);
            RefreshServiceStateUntilReached(
                this,
                svc.ServiceName,
                ServiceState.Stopped
            );
        }

        private void DoRestart(IWindowsServiceUtil svc)
        {
            if (svc.State != ServiceState.Stopped)
            {
                svc.Stop(wait: false);
                RefreshServiceStateUntilReached(
                    this,
                    svc.ServiceName,
                    ServiceState.Stopped,
                    isFinalState: false
                );
            }

            svc.Start(wait: false);
            RefreshServiceStateUntilReached(
                this,
                svc.ServiceName,
                ServiceState.Running,
                isFinalState: true
            );
        }


        private void KeyPressHandler(KeyEventEventArgs keyEvent)
        {
            var key = keyEvent.KeyEvent.Key;
            if (_keyHandlers.TryGetValue(key, out var handler))
            {
                handler(this);
                keyEvent.Handled = handler(this);
                return;
            }

            if (_lastKeyEvent != keyEvent)
            {
                var isalpha = (int) key >= (int) Key.A && (int) key <= (int) Key.z;
                var isnumeric = (int) key >= (int) Key.D0 && (int) key <= (int) Key.D9;
                var isspace = key == Key.Space;
                var isDash = keyEvent.KeyEvent.KeyValue == (int) '-';
                if (isalpha || isnumeric || isspace || isDash)
                {
                    SearchBox.Text += $"{(char) keyEvent.KeyEvent.KeyValue}";
                }
            }

            _lastKeyEvent = keyEvent;
            keyEvent.Handled = false;
        }

        private readonly Dictionary<Key, Func<MainView, bool>> _keyHandlers = new()
        {
            [Key.Esc] = ClearSearch,
            [Key.Backspace] = BackspaceOnSearch,
            [Key.CursorUp] = SelectedPrevious,
            [Key.CursorDown] = SelectedNext
        };

        private static bool BackspaceOnSearch(MainView arg)
        {
            var searchText = arg.SearchBox.Text.ToString() ?? "";
            if (searchText.Length == 0)
            {
                return true;
            }

            searchText = searchText.Substring(0, searchText.Length - 1);
            Application.MainLoop.Invoke(() =>
            {
                arg.SearchBox.SetNeedsDisplay();
                arg.SearchBox.Text = searchText;
            });
            return true;
        }

        private static bool SelectedNext(MainView arg)
        {
            var serviceName = arg.FindServiceAt(arg.ServicesList.SelectedRow + 1);
            UpdateStatusBarForSelectedService(arg, serviceName);
            return false;
        }

        private static bool SelectedPrevious(MainView arg)
        {
            var serviceName = arg.FindServiceAt(arg.ServicesList.SelectedRow - 1);
            UpdateStatusBarForSelectedService(arg, serviceName);
            return false;
        }

        private string? FindServiceAt(int idx)
        {
            try
            {
                return FilteredServices.Rows[idx][0] as string;
            }
            catch
            {
                return null;
            }
        }

        private static bool ClearSearch(MainView obj)
        {
            obj.SearchBox.Text = "";
            obj.ApplyFilter("");
            return true;
        }

        private static readonly ConcurrentDictionary<string, IWindowsServiceUtil> Services = new();
        private KeyEventEventArgs? _lastKeyEvent;
        private static CancellationTokenSource? _backgroundRefreshToken;
        private readonly DataColumn _stateDataColumn;
        private bool _isFirstRun;

        private static void UpdateTitle(MainView view, string title)
        {
            view.Title = $"{title}    (ctrl-Q to quit)";
        }

        private void StartBackgroundRefresh()
        {
            StopBackgroundRefresh();
            _backgroundRefreshToken = new CancellationTokenSource();
            var token = _backgroundRefreshToken.Token;
            Task.Factory.StartNew(() =>
            {
                var scm = new ServiceControlInterface();
                while (!token.IsCancellationRequested)
                {
                    if (DateTime.Now > _deferBackgroundRefreshUntil)
                    {
                        Trace.WriteLine("Starting background refresh");
                        var services = scm.ListAllServices();
                        Parallel.ForEach(services, service =>
                        {
                            try
                            {
                                RefreshServiceStatus(this, service, false);
                            }
                            catch (Exception ex)
                            {
                                Trace.WriteLine($"Unable to update service state for '{{service}}': {ex.Message}");
                            }
                        });
                    }

                    Sleep(BACKGROUND_REFRESH_INTERVAL_MS, token);
                }
            }, token);
        }


        private void Sleep(int ms, CancellationToken token)
        {
            while (!token.IsCancellationRequested && ms > 0)
            {
                var toSleep = ms > AUTO_REFRESH_INTERVAL_MS
                    ? AUTO_REFRESH_INTERVAL_MS
                    : ms;
                ms -= toSleep;
                Thread.Sleep(toSleep);
            }
        }

        private static void StopBackgroundRefresh()
        {
            try
            {
                _backgroundRefreshToken?.Cancel();
                _backgroundRefreshToken = null;
            }
            catch
            {
                // suppress
            }
        }

        private static void Refresh(MainView mainView)
        {
            using var _ = new AutoLocker(mainView.UiLock);
            mainView.AllServices.Clear();
            mainView.FilteredServices.Clear();
            UpdateTitle(mainView, "Services - Refreshing...");
            mainView.ClearStatusBar();
            mainView.SetNeedsDisplay();
            Application.MainLoop.Invoke(() =>
            {
                var scm = new ServiceControlInterface();
                var services = scm.ListAllServices().ToArray();
                var funcs = services.Select<string, Func<IWindowsServiceUtil?>>(
                    svc => () => ServiceUtilFor(svc)
                ).ToArray();
                var worker = new ParallelWorker<IWindowsServiceUtil?>(funcs);
                var results = worker.RunAll()
                    .Where(o => o.Result is not null)
                    .OrderBy(o => o.Result!.DisplayName);

                mainView.AllServices.Clear();
                foreach (var result in results)
                {
                    var util = result.Result;
                    mainView.AllServices.Rows.Add(DataRowFor(util!));
                }

                mainView.ApplyFilter(mainView.SearchBox.Text.ToString());
                UpdateTitle(mainView, $"Services - Last refreshed: {DateTime.Now}");
                mainView.StartBackgroundRefresh();
            });
        }

        private static IWindowsServiceUtil? ServiceUtilFor(
            string svc,
            bool refresh = true
        )
        {
            try
            {
                if (Services.TryGetValue(svc, out var cached))
                {
                    if (refresh)
                    {
                        cached.Refresh();
                    }

                    return cached;
                }

                var util = new WindowsServiceUtil(svc);
                Services.TryAdd(svc, util);
                return util;
            }
            catch (ServiceNotInstalledException)
            {
                return null;
            }
        }

        private static object[] DataRowFor(IWindowsServiceUtil svc)
        {
            return new object[] { svc.ServiceName, svc.DisplayName, $"{svc.State}" };
        }

        private static object[] DataRowFor(DataRow data)
        {
            return new[] { data[ID_COLUMN], data[NAME_COLUMN], data[STATE_COLUMN] };
        }

        private static string ServiceNameFor(DataRow row)
        {
            return (row[NAME_COLUMN] as string)!;
        }

        private static string ServiceIdFor(DataRow row)
        {
            return (row[ID_COLUMN] as string)!;
        }

        private static string ServiceStateFor(DataRow row)
        {
            return (row[STATE_COLUMN] as string)!;
        }
    }
}