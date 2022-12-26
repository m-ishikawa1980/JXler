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
    /// SubWindowApiXls.xaml の相互作用ロジック
    /// </summary>
    public partial class SubWindowApiXls : Window
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

        public ApiXls value = new ApiXls();
        public ActionType Action { get; set; }

        public enum ActionType
        {
            OK,
            Cancel
        }

        public SubWindowApiXls(int rowNum = -1)
        {
            InitializeComponent();
            var settings = Utils.GetSettings();
            if (rowNum >= 0)
            {
                No.Text = settings.ApiXlsHash[rowNum].No.ToString();
                ReqPath.Text = settings.ApiXlsHash[rowNum].ReqPath;
                ReqPath.Text = settings.ApiXlsHash[rowNum].ReqPath;
                ResPath.Text = settings.ApiXlsHash[rowNum].ResPath;
                ResPath.Text = settings.ApiXlsHash[rowNum].ResPath;
                _Action.Text = settings.ApiXlsHash[rowNum].Action;
            }
            else
            {
                ReqPath.Text = settings.BasePath;
                ResPath.Text = settings.BasePath;
            }
        }

        private void SelectReqFolder_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new CommonOpenFileDialog();
            if (!string.IsNullOrEmpty(ReqPath.Text))
            {
                dlg.InitialDirectory = ReqPath.Text;
            }
            dlg.IsFolderPicker = true;
            var res = dlg.ShowDialog();
            if (res == CommonFileDialogResult.Ok)
            {
                ReqPath.Text = dlg.FileName;
            }
        }

        private void SelectReqFile_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new CommonOpenFileDialog();
            if (!string.IsNullOrEmpty(ReqPath.Text))
            {
                dlg.InitialDirectory = ReqPath.Text;
            }
            var res = dlg.ShowDialog();
            if (res == CommonFileDialogResult.Ok)
            {

                var selectPath = Path.GetDirectoryName(dlg.FileName);
                if (ReqPath.Text != selectPath)
                {
                    ReqPath.Text = selectPath;
                }
                ReqName.Text = dlg.FileAsShellObject.Name;
            }
        }

        private void SelectResFolder_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new CommonOpenFileDialog();
            if (!string.IsNullOrEmpty(ResPath.Text))
            {
                dlg.InitialDirectory = ResPath.Text;
            }
            dlg.IsFolderPicker = true;
            var res = dlg.ShowDialog();
            if (res == CommonFileDialogResult.Ok)
            {
                ResPath.Text = dlg.FileName;
            }
        }

        private void SelectResFile_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new CommonOpenFileDialog();
            if (!string.IsNullOrEmpty(ResPath.Text))
            {
                dlg.InitialDirectory = ResPath.Text;
            }
            var res = dlg.ShowDialog();
            if (res == CommonFileDialogResult.Ok)
            {

                var selectPath = Path.GetDirectoryName(dlg.FileName);
                if (ResPath.Text != selectPath)
                {
                    ResPath.Text = selectPath;
                }
                ResName.Text = dlg.FileAsShellObject.Name;
            }
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            value.No = !string.IsNullOrEmpty(No.Text) ? int.Parse(No.Text) : 0;
            value.ReqPath = ReqPath.Text;
            value.ReqName = ReqName.Text;
            value.ResPath = ResPath.Text;
            value.ResName = ResName.Text;
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
