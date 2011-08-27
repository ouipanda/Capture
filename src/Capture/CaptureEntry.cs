using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Windows.Media.Imaging;

namespace Capture
{
    public class CaptureEntry
    {
        public int No { get; set; }
        public string PathToWindowImage { get; set; }
        public string PathToClipImage { get; set; }
        public System.Windows.Rect ClipImageRect { get; set; }

        public System.Windows.Point Cursor { get; set; }
        public MouseMessage MouseMessage { get; set; }

        public int ClickedWindowHandle { get; set; }
        public System.Windows.Rect ClickedRect { get; set; }

        public int WindowHandle { get; set; }
        public System.Windows.Rect WindowRect { get; set; }

        public string ClickedName { get; set; }
        public string ClickedControlType { get; set; }
        public string WindowName { get; set; }
        public string WindowControlType { get; set; }

        public int SelectedIndex { get; set; }

        System.Windows.Media.Imaging.BitmapImage _windowImage;
        System.Windows.Media.Imaging.BitmapImage _clipImage;

        public BitmapImage WindowImage { get { return _windowImage; } }
        public BitmapImage ClipImage { get { return _clipImage; } }

        public void LoadImage()
        {
            _windowImage = MakeBmp(PathToWindowImage);
            _clipImage = MakeBmp(PathToClipImage);
        }

        System.Windows.Media.Imaging.BitmapImage MakeBmp(string path)
        {
            System.IO.FileStream fs = new System.IO.FileStream(path, System.IO.FileMode.Open);
            byte[] abyData = new byte[fs.Length];
            fs.Read(abyData, 0, (int)fs.Length);
            fs.Close();
            System.IO.MemoryStream ms = new System.IO.MemoryStream(abyData);

            System.Windows.Media.Imaging.BitmapImage bi = new System.Windows.Media.Imaging.BitmapImage();
            bi.BeginInit();
            bi.StreamSource = ms;
            bi.EndInit();
            return bi;
        }

        public string Caption
        {
            get
            {
                string name1 = null;
                string name2 = string.IsNullOrEmpty(ClickedName) ? ClickedControlType : ClickedName;

                if (WindowHandle != ClickedWindowHandle)
                    name1 = WindowName;

                if (name1 == name2)
                    return name2;


                string mouse = "クリック";
                if (MouseMessage == Capture.MouseMessage.RDown)
                    mouse = "右クリック";

                return name1 + "\n[" + name2 + "] " + mouse;
            }
        }

        public override string ToString()
        {
            return "";
        }
    }
}
