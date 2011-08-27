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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Capture
{
    /// <summary>
    /// CaptureListItem.xaml の相互作用ロジック
    /// </summary>
    public partial class CaptureListItem : UserControl
    {
        CaptureEntry entry;
        public CaptureListItem()
        {
            InitializeComponent();
            this.Loaded += new RoutedEventHandler(CaptureListItem_Loaded);
        }

        void CaptureListItem_Loaded(object sender, RoutedEventArgs e)
        {
            entry = (CaptureEntry)this.DataContext;
            if (entry.SelectedIndex == 0)
            {
                border1.BorderThickness = new Thickness(2.0);
                border2.BorderThickness = new Thickness(0.0);
            }
            else
            {
                border1.BorderThickness = new Thickness(0.0);
                border2.BorderThickness = new Thickness(2.0);
            }

            entry.LoadImage();

            image2.Source = entry.ClipImage;
            image1.Source = entry.WindowImage;


            txtName2.Text = string.IsNullOrEmpty(entry.ClickedName) ? entry.ClickedControlType : entry.ClickedName;

            if (entry.WindowHandle != entry.ClickedWindowHandle)
                txtName1.Text = entry.WindowName; // string.IsNullOrEmpty(entry.WindowName) ? entry.WindowControlType : entry.WindowName;

            //同じなら表示しない
            txtName1.Visibility = (txtName2.Text == txtName1.Text) ? System.Windows.Visibility.Collapsed : System.Windows.Visibility.Visible;
        }

        private void image1_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            border1.BorderThickness = new Thickness(2.0);
            border2.BorderThickness = new Thickness(0.0);
            entry.SelectedIndex = 0;
        }

        private void image2_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            border1.BorderThickness = new Thickness(0.0);
            border2.BorderThickness = new Thickness(2.0);
            entry.SelectedIndex = 1;
        }
    }
}
