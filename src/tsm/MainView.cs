using System.Collections.Concurrent;
using System.Data;
using System.Diagnostics;
using NStack;
using PeanutButter.Utils;
using PeanutButter.WindowsServiceManagement;
using PeanutButter.WindowsServiceManagement.Exceptions;
using Terminal.Gui;

namespace services
{
    public class MainView : Window
    {
        private const int ID_COLUMN = 0;
        private const int NAME_COLUMN = 1;
        private const int STATE_COLUMN = 2;
        private const int AUTO_REFRESH_INTERVAL_MS = 100;
        public TableView ServicesList { get; set; }
        public ColorScheme GreenOnBlack { get; set; }
        public DataTable FilteredServices { get; set; }
        public DataTable AllServices { get; set; }
        public Label SearchLabel { get; set; }
        public TextField SearchBox { get; set; }

        private static readonly ConcurrentDictionary<string, int> FilteredServiceIdToRowMap = new();
        private static readonly ConcurrentDictionary<string, int> ServiceIdToRowMap = new();


        public MainView(Toplevel top)
        {
            GreenOnBlack = new ColorScheme
            {
                Normal = new Terminal.Gui.Attribute(Color.Green, Color.Black),
                HotNormal = new Terminal.Gui.Attribute(Color.BrightGreen, Color.Black),
                Focus = new Terminal.Gui.Attribute(Color.Green, Color.Magenta),
                HotFocus = new Terminal.Gui.Attribute(Color.BrightGreen, Color.Magenta),
                Disabled = new Terminal.Gui.Attribute(Color.Gray, Color.Black)
            };
            ColorScheme = GreenOnBlack;

            FilteredServices = new DataTable();
            FilteredServices.Columns.Add(new DataColumn() { ColumnName = "Id" });
            FilteredServices.Columns.Add(new DataColumn() { ColumnName = "Name" });
            FilteredServices.Columns.Add(new DataColumn() { ColumnName = "State" });
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
                Width = Dim.Fill(0),
                ColorScheme = GreenOnBlack
            };
            SearchBox.TextChanged += ApplyFilter;
            SearchBox.Enter += args => ServicesList?.SetFocus();
            Add(SearchBox);

            ServicesList = new TableView()
            {
                Y = 1,
                Width = Dim.Fill(0),
                Height = Dim.Fill(1),
                MultiSelect = true,
                ColorScheme = GreenOnBlack,
                FullRowSelect = true,
                Table = FilteredServices,
                Style =
                {
                    AlwaysShowHeaders = true,
                    ShowVerticalCellLines = true,
                    ShowVerticalHeaderLines = true,
                }
            };
            Add(ServicesList);


            ClearStatusBar();
            Refresh(this);
            ServicesList.KeyDown += KeyPressHandler;
            ServicesList.SetFocus();
        }

        protected override void Dispose(bool disposing)
        {
            _shutdownInitiated = true;
            base.Dispose(disposing);
        }

        private void ApplyFilter(ustring? obj)
        {
            Application.DoEvents();
            Application.MainLoop.Invoke(() =>
            {
                ServicesList.SetNeedsDisplay();
                var search = SearchBox.Text?.ToString()?.Split(" ", StringSplitOptions.RemoveEmptyEntries) ?? Array.Empty<string>();
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
                    new(Key.F6, "F6 Start all", () => StartAll(this))
                );
                StatusBar.AddItemAt(2,
                    new(Key.F7, "F7 Restart all", () => RestartAll(this))
                );
                StatusBar.AddItemAt(3,
                    new(Key.F8, "F8 Stop all", () => StopAll(this))
                );
            }
        }

        private string[] CurrentlyFilteredServices()
        {
            return FilteredServiceIdToRowMap.Keys.ToArray();
        }

        private void StartAll(MainView mainView)
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

        private void RestartAll(MainView mainView)
        {
            var services = CurrentlyFilteredServices();
            foreach (var service in services)
            {
                RunWithService(service, DoRestart);
            }
        }

        private void StopAll(MainView mainView)
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
                DataRow? data = null;
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
                            ex.Message
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
            svc.Uninstall();
            Refresh(this);
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
            if (KeyHandlers.TryGetValue(key, out var handler))
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

        private readonly Dictionary<Key, Func<MainView, bool>> KeyHandlers = new()
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
        private bool _shutdownInitiated;
        private KeyEventEventArgs _lastKeyEvent;

        private static void UpdateTitle(MainView view, string title)
        {
            view.Title = $"{title}    (ctrl-Q to quit)";
        }

        private static void Refresh(MainView mainView)
        {
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