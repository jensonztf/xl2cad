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
using Autodesk.AutoCAD.Colors;
using xl2cad;


namespace xl2cad_cad
{
    public class zDBText
    {
        [CommandMethod("zdbtext")]
        public void zdbtext()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage("请选取一行单行文字作为表格第一行\n"
                           + "每列文字的字体与此列第一行文字相同。");
            //单行文字的过滤器
            FilterType[] filter = new FilterType[1];
            filter[0] = FilterType.Text;
            //选取第一行DBText
            DBObjectCollection EntityCollection = GetFirstRow(filter);
            //

            using (Transaction transaction = db.TransactionManager.StartTransaction())
            {
                try
                {
                    foreach (DBObject obj in EntityCollection)
                    {
                        Entity ent = (Entity)transaction.GetObject(obj.ObjectId, OpenMode.ForWrite, true);
                        ent.Color = Color.FromColorIndex(ColorMethod.ByAci, 2);
                    }
                    transaction.Commit();
                }
                catch
                {
                    ed.WriteMessage("Error！");
                }
                finally
                {
                    transaction.Dispose();
                }
            }


        }

        public static DBObjectCollection GetFirstRow(FilterType[] filter)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            Entity entity = null;
            DBObjectCollection entityCollection = new DBObjectCollection();
            PromptSelectionOptions selops = new PromptSelectionOptions();
            //建立选择的过滤器内容
            TypedValue[] filList = new TypedValue[filter.Length + 2];
            filList[0] = new TypedValue((int)DxfCode.Operator, "<or");
            filList[filter.Length + 1] = new TypedValue((int)DxfCode.Operator, "or>");
            for (int i = 0; i < filter.Length; i++)
            {
                filList[i + 1] = new TypedValue((int)DxfCode.Start, filter[i].ToString());
            }
            //建立过滤器
            SelectionFilter f = new SelectionFilter(filList);
            //按照过滤器进行选择
            PromptSelectionResult ents = ed.GetSelection(selops, f);
            if (ents.Status == PromptStatus.OK)
            {
                using (Transaction transaction = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        SelectionSet SS = ents.Value;
                        foreach (ObjectId id in SS.GetObjectIds())
                        {
                            entity = (Entity)transaction.GetObject(id, OpenMode.ForWrite, true);
                            if (entity != null)
                            {
                                entityCollection.Add(entity);
                            }
                        }
                        transaction.Commit();
                    }
                    catch
                    {
                        ed.WriteMessage("Error ");
                    }
                    finally
                    {
                        transaction.Dispose();
                    }
                }
            }
            return entityCollection;
        }
    }

    public enum FilterType
    {
        Curve, Dimension, Polyline, BlockRef, Circle, Line, Arc, Text, MText, Polyline3d, Surface,
        Region, Solid3d, Hatch, Helix, DBPoint
    }
}


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

        //Ribbon界面的回调函数，响应事件，将Excel表格数据写入CAD
        public void OnButton1Pressed(IRibbonControl control)
        {
            //选取Excel表格数据
            System.Data.DataTable dt = new System.Data.DataTable();


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

            //让CAD自动执行netload命令加载程序集DLL
            AcadDoc.SendCommand("(command \"_netload\" \"" + @"C:\\Windows\\System\\xl2cad.dll" + "\") ");
            AcadDoc.SendCommand("zdbtext ");

        }

        public void OnButton2Pressed(IRibbonControl control)
        {

        }
    }
}