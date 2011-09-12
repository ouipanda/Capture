using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Data;
using System.Reflection;
using System.Windows.Media.Imaging;

namespace Capture
{
    public class ExcelFile : IDisposable
    {
        public List<string> Errors;

        string _saveto;
        object _app;
        object _workbooks;
        object _workbook;
        object _sheet;
        //object _range;

        object mark_template;
        object text_template;

        object[] Parameters;

        public ExcelFile(string filename, string saveto)
        {
            Errors = new List<string>();

            _saveto = saveto;

            _app = CreateObject("Excel.Application");

            object version = _app.GetType().InvokeMember("Version", BindingFlags.GetProperty, null, _app, null);


            //可視
            Parameters = new Object[] { true };
            _app.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, _app, Parameters);

            _workbooks = _app.GetType().InvokeMember("Workbooks", BindingFlags.GetProperty, null, _app, null);
            Parameters = new Object[] { filename };
            _workbooks.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, _workbooks, Parameters);


            _workbook = _app.GetType().InvokeMember("ActiveWorkbook", BindingFlags.GetProperty, null, _app, null);


            this.Save();

            return;

            /*
            Parameters = new Object[] { "A1" };
            _range = _sheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, _sheet, Parameters);

            Parameters = new Object[] { "ほげ" };
            _range.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, _range, Parameters);
            
            
            
            Parameters = new Object[] { 5, 47.25, 48, 228, 112.5 };
            _shapes.GetType().InvokeMember("AddShape", BindingFlags.InvokeMethod, null, _shapes, Parameters);

            Parameters = new Object[] { 1, 70.5, 214.5, 220.5, 93.75 };
            object _shape = _shapes.GetType().InvokeMember("AddShape", BindingFlags.InvokeMethod, null, _shapes, Parameters);

            */

            /*
            //2010で動いた
            object _textFrame2 = _shape.GetType().InvokeMember("TextFrame2", BindingFlags.GetProperty, null, _shape, null);
            object _textRange = _textFrame2.GetType().InvokeMember("TextRange", BindingFlags.GetProperty, null, _textFrame2, null);
            object _characters = _textRange.GetType().InvokeMember("Characters", BindingFlags.GetProperty, null, _textRange, null);
            Parameters = new Object[] { "ほげら" };
            _characters.GetType().InvokeMember("Text", BindingFlags.SetProperty, null, _characters, Parameters);
            */

            /*
            //2003で動いた(2010でも動く)
            object _textFrame2 = _shape.GetType().InvokeMember("TextFrame", BindingFlags.GetProperty, null, _shape, null);
            Parameters = new Object[] { System.Type.Missing, System.Type.Missing };
            object _characters = _textFrame2.GetType().InvokeMember("Characters", BindingFlags.InvokeMethod, null, _textFrame2, Parameters);
            Parameters = new Object[] { "ほげら2" };
            _characters.GetType().InvokeMember("Text", BindingFlags.SetProperty, null, _characters, Parameters);
            */

            /*
            object _characters = _shape.GetType().InvokeMember("Characters", BindingFlags.GetProperty, null, _shape, null);
            Parameters = new Object[] { "ほげら" };
            _characters.GetType().InvokeMember("Text", BindingFlags.SetProperty, null, _characters, Parameters);
            */

            /*
            object _textFrame2 = _shape.GetType().InvokeMember("TextFrame2", BindingFlags.GetProperty, null, _shape, null);
            object _textRange = _textFrame2.GetType().InvokeMember("TextRange", BindingFlags.GetProperty, null, _textFrame2, null);
            object _characters = _textRange.GetType().InvokeMember("Characters", BindingFlags.GetProperty, null, _textRange, null);
            Parameters = new Object[] { "ほげら" };
            _characters.GetType().InvokeMember("Text", BindingFlags.SetProperty, null, _characters, Parameters);
            */

            /*
            object _pictures = _sheet.GetType().InvokeMember("Pictures", BindingFlags.GetProperty, null, _sheet, null);
            Parameters = new Object[] { @"C:\Users\ouipanda\Pictures\file.png" };
            object _picture = _pictures.GetType().InvokeMember("Insert", BindingFlags.InvokeMethod, null, _pictures, Parameters);


            Parameters = new Object[] { 100 };
            _picture.GetType().InvokeMember("Top", BindingFlags.SetProperty, null, _picture, Parameters);
            _picture.GetType().InvokeMember("Left", BindingFlags.SetProperty, null, _picture, Parameters);


            Parameters = new Object[] { @"C:\Users\ouipanda\Pictures\file.png" };
            _picture = _pictures.GetType().InvokeMember("Insert", BindingFlags.InvokeMethod, null, _pictures, Parameters);

            Parameters = new Object[] { 200 };
            _picture.GetType().InvokeMember("Top", BindingFlags.SetProperty, null, _picture, Parameters);
            _picture.GetType().InvokeMember("Left", BindingFlags.SetProperty, null, _picture, Parameters);
            */



            /*

                ActiveSheet.Pictures.Insert("C:\Users\ouipanda\Pictures\file.png").Select
    ActiveSheet.Pictures.Insert("C:\Users\ouipanda\Pictures\folder2.png").Select
    Selection.ShapeRange.IncrementLeft 9.75
    Selection.ShapeRange.IncrementTop 33.75
    ActiveSheet.Pictures.Insert("C:\Users\ouipanda\Pictures\icon.png").Select
    Selection.ShapeRange.IncrementLeft 10.5
    Selection.ShapeRange.IncrementTop 80.25 
             
             * 
             * 
             * 
             * 
             * 
MsgBox msoShapeRoundedRectangle 5
MsgBox msoShapeRectangle 1
             * 
    ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 47.25, 48, 228, 112.5). _
        Select
    ActiveSheet.Shapes.AddShape(msoShapeRectangle, 70.5, 214.5, 220.5, 93.75). _
        Select
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "test"
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 4). _
        ParagraphFormat
        .FirstLineIndent = 0
        .Alignment = msoAlignLeft
    End With
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 4).Font
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorLight1
        .Fill.ForeColor.TintAndShade = 0
        .Fill.ForeColor.Brightness = 0
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 11
        .Name = "+mn-lt"
    End With

             */
        }

        /// <summary>
        /// Excel内の座標比率に直す。
        /// だいたい1.33で割ると画像のピクセル数と同じくらいになった
        /// </summary>
        /// <param name="v"></param>
        /// <returns></returns>
        double w(double v)
        {
            return (v / 1.33);
            //return (int)(v / 1.3);
        }
        double h(double v)
        {
            return (v / 1.33);
            //return (int)(v / 1.3);
        }


        void AddPicture(double x, double y, string path, BitmapImage bi)
        {
            x = w(x);
            y = h(y);

            object _pictures = _sheet.GetType().InvokeMember("Pictures", BindingFlags.GetProperty, null, _sheet, null);
            Parameters = new Object[] { path };
            object _picture = _pictures.GetType().InvokeMember("Insert", BindingFlags.InvokeMethod, null, _pictures, Parameters);

            Parameters = new Object[] { x };
            _picture.GetType().InvokeMember("Left", BindingFlags.SetProperty, null, _picture, Parameters);
            Parameters = new Object[] { y };
            _picture.GetType().InvokeMember("Top", BindingFlags.SetProperty, null, _picture, Parameters);

            /*
            //100%になるようにしたいけど…
            object shaperange = _picture.GetType().InvokeMember("ShapeRange", BindingFlags.GetProperty, null, _picture, null);

            Parameters = new Object[] { 1 };
            shaperange.GetType().InvokeMember("ScaleHeight", BindingFlags.SetProperty, null, shaperange, Parameters);

            MarshalReleaseComObject(ref shaperange);
            */
            MarshalReleaseComObject(ref _picture);
            MarshalReleaseComObject(ref _pictures);
        }

        void AddShape(double x, double y, double width, double height)
        {
            x = w(x);
            y = h(y);
            width = w(width);
            height = h(height);

            object newshape = mark_template.GetType().InvokeMember("Duplicate", BindingFlags.InvokeMethod, null, mark_template, null);

            Parameters = new Object[] { x };
            newshape.GetType().InvokeMember("Left", BindingFlags.SetProperty, null, newshape, Parameters);
            Parameters = new Object[] { y };
            newshape.GetType().InvokeMember("Top", BindingFlags.SetProperty, null, newshape, Parameters);

            Parameters = new Object[] { width };
            newshape.GetType().InvokeMember("Width", BindingFlags.SetProperty, null, newshape, Parameters);
            Parameters = new Object[] { height };
            newshape.GetType().InvokeMember("Height", BindingFlags.SetProperty, null, newshape, Parameters);

            MarshalReleaseComObject(ref newshape);
        }

        void AddTextbox(double x, double y, int width, int height, string text)
        {
            x = w(x);
            y = h(y);
            //width = w(width);
            //height = h(height);

            object newtextbox = text_template.GetType().InvokeMember("Duplicate", BindingFlags.InvokeMethod, null, text_template, null);

            Parameters = new Object[] { x };
            newtextbox.GetType().InvokeMember("Left", BindingFlags.SetProperty, null, newtextbox, Parameters);
            Parameters = new Object[] { y };
            newtextbox.GetType().InvokeMember("Top", BindingFlags.SetProperty, null, newtextbox, Parameters);

            object _textFrame2 = newtextbox.GetType().InvokeMember("TextFrame", BindingFlags.GetProperty, null, newtextbox, null);
            Parameters = new Object[] { System.Type.Missing, System.Type.Missing };
            object _characters = _textFrame2.GetType().InvokeMember("Characters", BindingFlags.InvokeMethod, null, _textFrame2, Parameters);
            Parameters = new Object[] { text };
            _characters.GetType().InvokeMember("Text", BindingFlags.SetProperty, null, _characters, Parameters);

            MarshalReleaseComObject(ref _characters);
            MarshalReleaseComObject(ref _textFrame2);
            MarshalReleaseComObject(ref newtextbox);
        }


        public void Write(IEnumerable<CaptureEntry> list)
        {
            _sheet = _workbook.GetType().InvokeMember("ActiveSheet", BindingFlags.GetProperty, null, _workbook, null);

            object _shapes = _sheet.GetType().InvokeMember("Shapes", BindingFlags.GetProperty, null, _sheet, null);

            //赤枠テンプレート
            Parameters = new Object[] { "CaptureMark1" };
            mark_template = _shapes.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, _shapes, Parameters);

            //TextBoxテンプレート
            Parameters = new Object[] { "CaptureText1" };
            text_template = _shapes.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, _shapes, Parameters);

            string errors = "";
            int pict_left_margin = 10;
            int pict_top_margin = 40;
            int text_left_margin = 10;
            int text_top_margin = 10;
            int top = 10;

            foreach (var entry in list)
            {
                try
                {

                    string path = null;
                
                    BitmapImage bi;
                    int shape_x = (int)(entry.ClickedRect.X - entry.WindowRect.X);
                    int shape_y = (int)(entry.ClickedRect.Y - entry.WindowRect.Y);
                    int width = (int)entry.ClickedRect.Width;
                    int height = (int)entry.ClickedRect.Height;
                    if (entry.SelectedIndex == 0)
                    {
                        //ウィンドウ画像
                        path = entry.PathToWindowImage;
                        bi = entry.WindowImage;

                        //クリップ領域より対象が大きい場合は調整する
                        if (entry.WindowImage.Height < height + shape_y)
                            height = (int)(entry.WindowImage.Height - shape_y);
                        if (entry.WindowImage.Width < width + shape_x)
                            width = (int)(entry.WindowImage.Width - shape_x);
                    }
                    else
                    {
                        //クリップ画像
                        path = entry.PathToClipImage;
                        bi = entry.ClipImage;

                        shape_x = (int)(entry.ClickedRect.X - entry.ClipImageRect.X);
                        shape_y = (int)(entry.ClickedRect.Y - entry.ClipImageRect.Y);

                        //クリップ領域より対象が大きい場合は調整する
                        if (shape_x < 0)
                            shape_x = 0;
                        if (shape_y < 0)
                            shape_y = 0;
                        if (entry.ClipImage.Height < height + shape_y)
                            height = (int)(entry.ClipImage.Height - shape_y);
                        if (entry.ClipImage.Width < width + shape_x)
                            width = (int)(entry.ClipImage.Width - shape_x);

                    }
                
                    /*
                    string tmp = @"pict  {0} {1} {2} {3}
    shape {4} {5} {6} {7}";
                    tmp = string.Format(tmp, pict_left_margin, top, bi.PixelWidth, bi.PixelHeight,
                        10 + shape_x, top + shape_y, width, height);
                    */

                    AddPicture(pict_left_margin, top, path, bi);
                    AddShape(pict_left_margin + shape_x, top + shape_y, width, height);

                    //AddTextbox(text_left_margin, top + text_top_margin, 100, 100, tmp);
                    AddTextbox(text_left_margin, top + text_top_margin, 100, 100, entry.Caption);

                    top += (int)(bi.PixelHeight + pict_top_margin);

                }
                catch (Exception ex)
                {
                    errors += ex.Message + "\n" + ex.StackTrace + "\n";
                }


            }


            mark_template.GetType().InvokeMember("Delete", BindingFlags.InvokeMethod, null, mark_template, null);
            text_template.GetType().InvokeMember("Delete", BindingFlags.InvokeMethod, null, text_template, null);

            MarshalReleaseComObject(ref mark_template);
            MarshalReleaseComObject(ref text_template);
            MarshalReleaseComObject(ref _shapes);
            MarshalReleaseComObject(ref _sheet);

            this.Save();

            //可視
            Parameters = new Object[] { true };
            _app.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, _app, Parameters);

            if(!string.IsNullOrEmpty(errors))
                throw new ApplicationException(errors);

        }


        void Save()
        {

            Parameters = new Object[] { false };
            _app.GetType().InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, _app, Parameters);

            Parameters = new Object[] { _saveto };
            _workbook.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, _workbook, Parameters);

            Parameters = new Object[] { true };
            _app.GetType().InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, _app, Parameters);
        }

        public void Close()
        {
            if (_workbook != null)
                MarshalReleaseComObject(ref _workbook);

            if (_workbooks != null)
                MarshalReleaseComObject(ref _workbooks);

            if (_app != null)
                MarshalReleaseComObject(ref _app);

            _workbook = null;
            _workbooks = null;
            _app = null;
        }

        public void Dispose()
        {
            this.Close();
        }



        public static object CreateObject(string progId, string serverName)
        {
            Type t;
            if (serverName == null || serverName.Length == 0)
                t = Type.GetTypeFromProgID(progId);
            else
                t = Type.GetTypeFromProgID(progId, serverName, true);
            return Activator.CreateInstance(t);
        }

        public static object CreateObject(string progId)
        {
            return CreateObject(progId, null);
        }

        private void MarshalReleaseComObject(ref object objCom)
        {
            try
            {
                int i = 1;
                if (objCom != null && System.Runtime.InteropServices.Marshal.IsComObject(objCom))
                {
                    do
                    {
                        i = System.Runtime.InteropServices.Marshal.ReleaseComObject(objCom);
                    } while (i > 0);
                }
            }
            finally
            {
                objCom = null;
            }
        }
    }
    
}
