using Microsoft.WindowsAPICodePack.Dialogs;
using JXler.Libraries;
using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace JXler
{
    /// <summary>
    /// SettingWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class SettingWindow : Window
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

        public SettingWindow()
        {
            InitializeComponent();
            var settings = Utils.GetSettings();
            BasePath.Text = settings.BasePath;
            IndexSheet.Text = settings.IndexSheet;
            CellText.Text = settings.XlsCellStyles.Text;
            CellDateTime.Text = settings.XlsCellStyles.DateTime;
            CellInt.Text = settings.XlsCellStyles.NumberInt;
            CellFloat.Text = settings.XlsCellStyles.NumberFloat;
            if (settings.OutputXlsLink)
            {
                CheckTrue.IsChecked = true;
            }
            else
            {
                CheckFalse.IsChecked = true;
            }
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            var settings = Utils.GetSettings();
            settings.BasePath = BasePath.Text;
            settings.IndexSheet = IndexSheet.Text;
            settings.XlsCellStyles.Text = CellText.Text;
            settings.XlsCellStyles.DateTime = CellDateTime.Text;
            settings.XlsCellStyles.NumberInt = CellInt.Text;
            settings.XlsCellStyles.NumberFloat = CellFloat.Text;
            settings.OutputXlsLink = CheckTrue.IsChecked == true ? true : false;
            settings.SaveSettings();
            Close();
        }

        private void SelectBaseFolder_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new CommonOpenFileDialog();
            dlg.IsFolderPicker = true;
            var res = dlg.ShowDialog();
            if (res == CommonFileDialogResult.Ok)
            {
                BasePath.Text = dlg.FileName;
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
