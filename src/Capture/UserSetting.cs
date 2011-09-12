using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Capture
{
    public class UserSetting
    {
        static UserSetting _setting;
        Uri u1;

        public static UserSetting LoadSetting()
        {
            if (_setting == null)
            {
                _setting = new UserSetting();

                if (Properties.Settings.Default.IsUpgrated)
                {
                    //アセンブリバージョンが変更されたので設定情報をアップグレードする
                    Properties.Settings.Default.Upgrade();
                    Properties.Settings.Default.IsUpgrated = false;
                    _setting.Save();
                }

            }

            return _setting;
        }

        private UserSetting()
        {
            u1 = new Uri(GetStartupPath());

            /*
            this.SaveDir = System.IO.Path.Combine(GetStartupPath(), "data");
            if (!System.IO.Directory.Exists(this.SaveDir))
                System.IO.Directory.CreateDirectory(this.SaveDir);

            this.ExcelTemplate = System.IO.Path.Combine(GetStartupPath(), "Template.xls");
            */



            //Uri u2 = new Uri(u1, "data");
            //
        }

        public void Save()
        {
            Properties.Settings.Default.Save();
        }



        public string SaveDir
        {
            get
            {
                Uri u2 = new Uri(u1, Properties.Settings.Default.save_dir);
                return u2.LocalPath;
            }
            set
            {
                Properties.Settings.Default.save_dir = value;
            }
        }

        public string SaveDirOrg
        {
            get
            {
                return Properties.Settings.Default.save_dir;
            }
        }

        public string ExcelTemplate
        {
            get
            {
                Uri u2 = new Uri(u1, Properties.Settings.Default.excel_template);
                return u2.LocalPath;
            }
            set
            {
                Properties.Settings.Default.excel_template = value;
            }
        }

        public int ClipSizeTop
        {
            get
            {
                return Properties.Settings.Default.size_top;
            }
            set
            {
                Properties.Settings.Default.size_top = value;
            }
        }
        public int ClipSizeLeft
        {
            get
            {
                return Properties.Settings.Default.size_left;
            }
            set
            {
                Properties.Settings.Default.size_left = value;
            }
        }
        public int ClipSizeBottom
        {
            get
            {
                return Properties.Settings.Default.size_bottom;
            }
            set
            {
                Properties.Settings.Default.size_bottom = value;
            }
        }
        public int ClipSizeRight
        {
            get
            {
                return Properties.Settings.Default.size_right;
            }
            set
            {
                Properties.Settings.Default.size_right = value;
            }
        }

        /// <summary>
        /// 初期画像 0:ウィンドウ 1:カーソル
        /// </summary>
        public int InitialImage
        {
            get
            {
                return Properties.Settings.Default.initial_image;
            }
            set
            {
                Properties.Settings.Default.initial_image = value;
            }
        }

        public bool Topmost
        {
            get { return Properties.Settings.Default.topmost; }
            set { Properties.Settings.Default.topmost = value; }
        }

        public string GetStartupPath()
        {
            return System.IO.Path.GetDirectoryName(Environment.GetCommandLineArgs()[0])+ "\\";
        }
    }
}
