using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using Office;
using Excel;
using AddInDesignerObjects;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

//[assembly:CommandClass(typeof(xl2cad_cad.zDBText))]
//namespace xl2cad_cad
//{
//    public class zDBText
//    {
//        [CommandMethod("zdbtext")]
//        public void zdbtext()
//        {

//        }
//    }
//}


namespace xl2cad_wps
{
    [ComVisible(true)]
    public class Ribbon : IDTExtensibility2, IRibbonExtensibility
    {
        public static Excel.Application app;
        public Autodesk.AutoCAD.Interop.AcadApplication AcadApp;
        public Autodesk.AutoCAD.Interop.AcadDocument AcadDoc;

        public void OnAddInsUpdate(ref Array custom)
        {
            throw new NotImplementedException();
        }

        public void OnBeginShutdown(ref Array custom)
        {
            throw new NotImplementedException();
        }

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            //把当前ETApplication赋值给app
            app = (Excel.Application)Application;
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            throw new NotImplementedException();
        }

        public void OnStartupComplete(ref Array custom)
        {
            throw new NotImplementedException();
        }

        public string GetCustomUI(string RibbonID)
        {
            return xl2cad.Resource1.RibbonXML;
        }

        //Ribbon界面的回调函数，响应事件
        public void OnButton1Pressed(IRibbonControl control)
        {
            //打开cad
            try
            {
                AcadApp = (Autodesk.AutoCAD.Interop.AcadApplication)System.Runtime.InteropServices.Marshal.GetActiveObject("AutoCAD.Application");
                AcadDoc = AcadApp.ActiveDocument;

            }
            catch
            {
                OpenFileDialog op = new OpenFileDialog();
                op.Filter = "CAD文件(*.dwg)|*.dwg|CAD图形文件(*.dxf)|*.dxf";
                op.Title = "打开CAD文件";
                op.ShowDialog();

                string filePath = op.FileName;
                if (filePath == "")
                {
                    MessageBox.Show("选择CAD文件无效!", "文件无效!");
                    System.Windows.Forms.Application.Exit();
                }
                AcadApp = new Autodesk.AutoCAD.Interop.AcadApplication();
                AcadDoc = AcadApp.Documents.Open(filePath, null, null);
            }
            AcadApp.Application.Visible = true;
            //使CAD程序跳到在最前面，需要添加引用“Microsoft.VisualBasic”
            Microsoft.VisualBasic.Interaction.AppActivate(AcadApp.Caption);

            AcadDoc.SendCommand("ztfPickPoint ");

        }

        public void OnButton2Pressed(IRibbonControl control)
        {

        }
    }
}

