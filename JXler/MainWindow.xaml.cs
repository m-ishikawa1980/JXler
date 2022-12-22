using JXler.Libraries;
using JXler.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace JXler
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            var setting = Utils.GetSettings();
            dataGridJsonXls.ItemsSource = setting.JsonXlsHash;
            dataGridJsonXls.Items.Refresh();
        }
        
        private void dataGridJsonXls_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var fileNames = (string[])e.Data.GetData(DataFormats.FileDrop);
                var settings = Utils.GetSettings();

                foreach (var fileName in fileNames)
                {
                    var extension = Path.GetExtension(fileName);
                    var jsonToXls = new JsonXls();
                    switch (extension)
                    {
                        case ".xlsx":
                        case ".xlsm":
                            jsonToXls.XlsPath = Path.GetDirectoryName(fileName);
                            jsonToXls.XlsName = Path.GetFileName(fileName);
                            jsonToXls.No = settings.JsonXlsHash.Count > 0 ?
                                            settings.JsonXlsHash.Select(o => o.No).Max() + 1 :
                                            1;
                            jsonToXls.Action = Utils.GetExecAction(execAction: Utils.ExecAction.Rl);
                            settings.JsonXlsHash.Add(jsonToXls);
                            break;
                        case ".json":
                            jsonToXls.JsonPath = Path.GetDirectoryName(fileName);
                            jsonToXls.JsonName = Path.GetFileName(fileName);
                            jsonToXls.No = settings.JsonXlsHash.Count > 0 ?
                                            settings.JsonXlsHash.Select(o => o.No).Max() + 1 :
                                            1;
                            jsonToXls.Action = Utils.GetExecAction(execAction: Utils.ExecAction.Lr);
                            settings.JsonXlsHash.Add(jsonToXls);
                            break;
                        default:
                            WriteLogJsonXls(
                                msg: $"対象外のためスキップ {fileName}",
                                logLevel: LogLevel.Error);
                            break;
                    };
                }
                settings.SaveSettings();
                dataGridJsonXls.ItemsSource = settings.JsonXlsHash;
                dataGridJsonXls.Items.Refresh();
            }
        }

        private void Reload_Click(object sender, RoutedEventArgs e)
        {
            if (TabJsonXls.IsSelected)
            {
                var settings = Utils.GetSettings();
                dataGridJsonXls.ItemsSource = settings.JsonXlsHash;
                dataGridJsonXls.Items.Refresh();
            }
            else
            {
            }
        }

        private void ContextJsonXlsOpen(object sender, RoutedEventArgs e)
        {
            var rowIndex = dataGridJsonXls.SelectedIndex;
            if (rowIndex < 0)
            {
                GridJsonXls_Menu_Add.IsEnabled = true;
                GridJsonXls_Menu_Update.IsEnabled = false;
                GridJsonXls_Menu_Delete.IsEnabled = false;
                GridJsonXls_Menu_Move.IsEnabled = false;
                GridJsonXls_Menu_Copy.IsEnabled = false;
            }
            else
            {
                GridJsonXls_Menu_Add.IsEnabled = true;
                GridJsonXls_Menu_Update.IsEnabled = true;
                GridJsonXls_Menu_Delete.IsEnabled = true;
                GridJsonXls_Menu_Move.IsEnabled = true;
                GridJsonXls_Menu_Copy.IsEnabled = true;
            }
        }

        private void MoveXlsPath_Click(object sender, RoutedEventArgs e)
        {
            if (TabJsonXls.IsSelected)
            {
                var rowIndex = dataGridJsonXls.SelectedIndex;
                if (rowIndex >= 0)
                {
                    Process.Start(
                        "explorer.exe",
                        Utils.ComplementRelativeDir(
                            dir: Utils.GetSettings().JsonXlsHash[rowIndex].XlsPath)
                        );
                }
            }
        }

        private void MoveJsonPath_Click(object sender, RoutedEventArgs e)
        {
            if (TabJsonXls.IsSelected)
            {
                var rowIndex = dataGridJsonXls.SelectedIndex;
                if (rowIndex >= 0)
                {
                    Process.Start(
                        "explorer.exe",
                        Utils.ComplementRelativeDir(
                            dir: Utils.GetSettings().JsonXlsHash[rowIndex].JsonPath)
                        );
                }
            }
        }

        private void Add_Click(object sender, RoutedEventArgs e)
        {
            var settings = Utils.GetSettings();
            if (TabJsonXls.IsSelected)
            {
                var win = new SubWindowJsonXls();
                win.Owner = GetWindow(this);
                win.ShowDialog();
                if (win.Action == SubWindowJsonXls.ActionType.OK)
                {
                    win.value.No = settings.JsonXlsHash.Count > 0 ?
                                    settings.JsonXlsHash.Select(o => o.No).Max() + 1 :
                                    1;
                    win.value.Action = Utils.GetExecAction(execAction: Utils.ExecAction.Lr);
                    settings.JsonXlsHash.Add(win.value);
                    dataGridJsonXls.ItemsSource = settings.JsonXlsHash;
                    dataGridJsonXls.Items.Refresh();
                }
            }
            settings.SaveSettings();
        }

        private void Setting_Click(object sender, RoutedEventArgs e)
        {
            var win = new SettingWindow();
            win.Owner = GetWindow(this);
            win.ShowDialog();
        }

        private void Copy_Click(object sender, RoutedEventArgs e)
        {

            if (MessageBox.Show(owner: this,
                    messageBoxText: "コピーします",
                    caption: "",
                    button: MessageBoxButton.OKCancel) == MessageBoxResult.OK)
            {
                var settings = Utils.GetSettings();
                if (TabJsonXls.IsSelected)
                {
                    var rowIndex = dataGridJsonXls.SelectedIndex;
                    if (rowIndex >= 0)
                    {
                        var jsonToXls = (List<JsonXls>)dataGridJsonXls.ItemsSource;
                        jsonToXls[rowIndex].No = settings.JsonXlsHash.Count > 0 ?
                                                settings.JsonXlsHash.Select(o => o.No).Max() + 1 :
                                                1;
                        settings.JsonXlsHash.Add(jsonToXls[rowIndex]);
                        dataGridJsonXls.ItemsSource = settings.JsonXlsHash;
                        dataGridJsonXls.Items.Refresh();
                    }
                }
                settings.SaveSettings();
            }
        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show(owner: this,
                    messageBoxText: "削除します",
                    caption: "",
                    button: MessageBoxButton.OKCancel) == MessageBoxResult.OK)
            {
                var settings = Utils.GetSettings();
                if (TabJsonXls.IsSelected)
                {
                    var rowItems = dataGridJsonXls.SelectedItems;
                    var jsonToXlsList = settings.JsonXlsHash;
                    foreach (JsonXls rowItem in rowItems)
                    {
                        jsonToXlsList = jsonToXlsList.Where(o => o.No != rowItem.No).ToList();
                    }
                    settings.JsonXlsHash = jsonToXlsList.RenumberingJsonXlsList();
                    dataGridJsonXls.ItemsSource = settings.JsonXlsHash;
                    dataGridJsonXls.Items.Refresh();
                }
                settings.SaveSettings();
            }
        }

        private void Update_Click(object sender, RoutedEventArgs e)
        {
            if (TabJsonXls.IsSelected)
            {
                var rowIndex = dataGridJsonXls.SelectedIndex;
                if (rowIndex >= 0)
                {
                    var win = new SubWindowJsonXls(rowIndex);
                    win.Owner = GetWindow(this);
                    win.ShowDialog();
                    if (win.Action == SubWindowJsonXls.ActionType.OK)
                    {
                        var settings = Utils.GetSettings();
                        settings.JsonXlsHash[rowIndex] = win.value;
                        settings.SaveSettings();
                        dataGridJsonXls.ItemsSource = settings.JsonXlsHash;
                        dataGridJsonXls.Items.Refresh();
                    }
                }
            }
        }

        public void WriteLogJsonXls(string msg, LogLevel logLevel = LogLevel.Info)
        {
            WriteLog(
                listBox: logMessageJsonXls,
                msg: msg,
                logLevel: logLevel);
        }

        private void WriteLog(ListBox listBox, string msg, LogLevel logLevel)
        {
            var time = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            var line = $"{time}[{logLevel}] {msg}";
            listBox.Items.Add(new LogMessage { Level = logLevel, Message = line });
        }

        private void GridJsonXls_Button_Click(object sender, RoutedEventArgs e)
        {
            var execAction = ((Button)sender).Content.ToString();
            var settings = Utils.GetSettings();
            settings.JsonXlsHash
                .Where(o => o.No == int.Parse(((Button)sender).Tag.ToString()))
                .Select(o => o).FirstOrDefault().Action = ExecChenge(execAction);
            settings.SaveSettings();
            dataGridJsonXls.ItemsSource = settings.JsonXlsHash;
            dataGridJsonXls.Items.Refresh();
        }

        private string ExecChenge(string execAction)
        {
            switch (Utils.CheckExecAction(execAction: execAction))
            {
                case Utils.ExecAction.Lr:
                    return Utils.GetExecAction(execAction: Utils.ExecAction.Rl);
                case Utils.ExecAction.Rl:
                    return Utils.GetExecAction(execAction: Utils.ExecAction.None);
                case Utils.ExecAction.None:
                    return Utils.GetExecAction(execAction: Utils.ExecAction.Lr);
                default:
                    return execAction;
            }
        }

        private void Exec_Click(object sender, RoutedEventArgs e)
        {
            var win = new ConfirmWindow(this);
            win.Owner = GetWindow(this);
            win.ShowDialog();

            if (TabJsonXls.IsSelected)
            {
                switch (win.Action)
                {
                    case ConfirmWindow.ActionType.OK:
                        var settings = Utils.GetSettings();
                        settings.ExecPtn = win.value.execPtn;
                        settings.Path = win.value.Path;
                        settings.SaveSettings();
                        ConvertJsonXls.Convert(this);
                        break;
                    case ConfirmWindow.ActionType.Error:
                        WriteLogJsonXls(
                            msg: win.value.Message,
                            logLevel: LogLevel.Error);
                        break;
                }
            }
        }
    }
}
