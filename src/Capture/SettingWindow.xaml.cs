using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Capture
{
    /// <summary>
    /// SettingWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class SettingWindow : Window
    {
        UserSetting setting = UserSetting.LoadSetting();
        public SettingWindow()
        {
            InitializeComponent();

            txtTop.Text = setting.ClipSizeTop.ToString();
            txtLeft.Text = setting.ClipSizeLeft.ToString();
            txtBottom.Text = setting.ClipSizeBottom.ToString();
            txtRight.Text = setting.ClipSizeRight.ToString();

            txtSave.Text = setting.SaveDirOrg;
        }

        private void btnSelect_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
            fbd.Description = "フォルダを指定してください。";
            fbd.RootFolder = Environment.SpecialFolder.Desktop;
            fbd.ShowNewFolderButton = true;

            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.txtSave.Text = fbd.SelectedPath;
            }
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            int top = 0;
            int left = 0;
            int bottom = 0;
            int right = 0;

            if (int.TryParse(txtTop.Text, out top))
                setting.ClipSizeTop = top;
            if (int.TryParse(txtLeft.Text, out left))
                setting.ClipSizeLeft = left;
            if (int.TryParse(txtBottom.Text, out bottom))
                setting.ClipSizeBottom = bottom;
            if (int.TryParse(txtRight.Text, out right))
                setting.ClipSizeRight = right;

            //補正
            if (setting.ClipSizeTop >= 0)
                setting.ClipSizeTop = -1;
            if (setting.ClipSizeLeft >= 0)
                setting.ClipSizeLeft = -1;
            if (setting.ClipSizeBottom <= 0)
                setting.ClipSizeBottom = 1;
            if (setting.ClipSizeRight <= 0)
                setting.ClipSizeRight = 1;

            setting.SaveDir = txtSave.Text;
            setting.Save();

            this.DialogResult = true;
            this.Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}
