using JXler.Libraries;
using JXler.Models;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;

namespace JXler
{
    /// <summary>
    /// ConfirmWindow.xaml の相互作用ロジック
    /// </summary>


    public partial class ConfirmWindow : Window
    {
        public Confirm value = new Confirm();

        public ActionType Action { get; set; }

        public enum ActionType
        {
            OK,
            Cancel,
            Error
        }

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
            int style = GetWindowLong(handle, GWL_STYLE);
            style = style & (~WS_SYSMENU);
            SetWindowLong(handle, GWL_STYLE, style);
        }

        public ConfirmWindow(MainWindow parentWindow)
        {
            InitializeComponent();

            var settings = Utils.GetSettings();
            switch (settings.ExecPtn)
            {
                case ExecPtn.SameInput:
                    SameInput.IsChecked = true;
                    break;
                case ExecPtn.SpecifyPath:
                    SpecifyPath.IsChecked = true;
                    break;
                case ExecPtn.SetIndividually:
                    SetIndividually.IsChecked = true;
                    break;
            }
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            var selectItem = ExecPtn.Unselected;
            foreach (RadioButton radioButton in radio.Children)
            {
                if (radioButton.IsChecked == true)
                {
                    Enum.TryParse(radioButton.Name, out selectItem);
                }
            }
            value.execPtn = selectItem;
            if (value.execPtn == ExecPtn.SpecifyPath)
            {
                var path = SelectFolder();
                if(path == null)
                {
                    return;
                }
                value.Path = path;
            }
            Action = string.IsNullOrEmpty(value.Message) ? ActionType.OK : ActionType.Error;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Action = ActionType.Cancel;
            Close();
        }

        private string SelectFolder()
        {
            var settings = Utils.GetSettings();
            var dlg = new CommonOpenFileDialog();
            if (!string.IsNullOrEmpty(settings.Path))
            {
                dlg.InitialDirectory = settings.Path;
            }
            dlg.IsFolderPicker = true;
            var res = dlg.ShowDialog();
            if (res == CommonFileDialogResult.Ok)
            {
                SpecifyPath.IsChecked = true;
                return dlg.FileName;
            }
            return null;
        }
    }
}
