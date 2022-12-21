using Microsoft.WindowsAPICodePack.Dialogs;
using JXler.Libraries;
using JXler.Models;
using System.IO;
using System.Windows;
using System.Runtime.InteropServices;
using System;
using System.Windows.Interop;

namespace JXler
{
    /// <summary>
    /// SubWindowJsonXls.xaml の相互作用ロジック
    /// </summary>
    public partial class SubWindowJsonXls : Window
    {
        [DllImport("user32.dll")]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        const int GWL_STYLE = -16;
        const int WS_SYSMENU = 0x80000;

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            IntPtr handle = new WindowInteropHelper(this).Handle;
            var style = GetWindowLong(handle, GWL_STYLE);
            style = style & (~WS_SYSMENU);
            SetWindowLong(handle, GWL_STYLE, style);
        }

        public JsonXls value = new JsonXls();
        public ActionType Action { get; set; }

        public enum ActionType
        {
            OK,
            Cancel
        }

        public SubWindowJsonXls(int rowNum = -1)
        {
            InitializeComponent();
            var settings = Utils.GetSettings();
            if (rowNum >= 0)
            {
                No.Text = settings.JsonXlsHash[rowNum].No.ToString();
                XlsPath.Text = settings.JsonXlsHash[rowNum].XlsPath;
                XlsName.Text = settings.JsonXlsHash[rowNum].XlsName;
                JsonPath.Text = settings.JsonXlsHash[rowNum].JsonPath;
                JsonName.Text = settings.JsonXlsHash[rowNum].JsonName;
                _Action.Text = settings.JsonXlsHash[rowNum].Action;
            }
            else
            {
                XlsPath.Text = settings.BasePath;
                JsonPath.Text = settings.BasePath;
            }
        }

        private void SelectJsonFolder_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new CommonOpenFileDialog();
            if (!string.IsNullOrEmpty(JsonPath.Text))
            {
                dlg.InitialDirectory = JsonPath.Text;
            }
            dlg.IsFolderPicker = true;
            var res = dlg.ShowDialog();
            if (res == CommonFileDialogResult.Ok)
            {
                JsonPath.Text = dlg.FileName;
            }
        }

        private void SelectJsonFile_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new CommonOpenFileDialog();
            if (!string.IsNullOrEmpty(JsonPath.Text))
            {
                dlg.InitialDirectory = JsonPath.Text;
            }
            var res = dlg.ShowDialog();
            if (res == CommonFileDialogResult.Ok)
            {

                var selectPath = Path.GetDirectoryName(dlg.FileName);
                if (JsonPath.Text != selectPath)
                {
                    JsonPath.Text = selectPath;
                }
                JsonName.Text = dlg.FileAsShellObject.Name;
            }
        }

        private void SelectWbFolder_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new CommonOpenFileDialog();
            if (!string.IsNullOrEmpty(XlsPath.Text))
            {
                dlg.InitialDirectory = XlsPath.Text;
            }
            dlg.IsFolderPicker = true;
            var res = dlg.ShowDialog();
            if (res == CommonFileDialogResult.Ok)
            {
                XlsPath.Text = dlg.FileName;
            }
        }

        private void SelectWbFile_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new CommonOpenFileDialog();
            if (!string.IsNullOrEmpty(XlsPath.Text))
            {
                dlg.InitialDirectory = XlsPath.Text;
            }
            var res = dlg.ShowDialog();
            if (res == CommonFileDialogResult.Ok)
            {

                var selectPath = Path.GetDirectoryName(dlg.FileName);
                if (XlsPath.Text != selectPath)
                {
                    XlsPath.Text = selectPath;
                }
                XlsName.Text = dlg.FileAsShellObject.Name;
            }
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            value.No = !string.IsNullOrEmpty(No.Text) ? int.Parse(No.Text) : 0;
            value.JsonPath = JsonPath.Text;
            value.JsonName = JsonName.Text;
            value.XlsPath = XlsPath.Text;
            value.XlsName = XlsName.Text;
            value.Action = _Action.Text;
            Action = ActionType.OK;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Action = ActionType.Cancel;
            Close();
        }
    }
}
