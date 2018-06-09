using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.IO;
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
using ADOX;
using System.Data.OleDb;

namespace xl2cad_cad
{
    public class zDBText
    {
        public System.Data.DataTable dt;

        [CommandMethod("zdbtext")]
        public void zdbtext()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage("请选取一行单行文字作为表格第一行\n"
                           + "每列文字的字体与此列第一行文字相同。");

            //从数据库读取文件
            try
            {
                OleDbConnection con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data source=C:\\Windows\\System\\xl2cad.mdb");
                con.Open();
                string sql = "select * from excel2cad";
                OleDbDataAdapter da = new OleDbDataAdapter(sql, con);
                System.Data.DataSet ds = new System.Data.DataSet();
                dt = new System.Data.DataTable();
                da.Fill(ds, "excel2cad");
                dt = ds.Tables[0];
                con.Close();

            }
            catch (SystemException ex)
            {
                ed.WriteMessage(ex.ToString());
            }

            //选取第一行DBText
            //单行文字的过滤器
            FilterType[] filter = new FilterType[1];
            filter[0] = FilterType.Text;
            DBObjectCollection EntityCollection = GetFirstRow(filter);

            //写入数据
            //判断选取的列数和数据库的列数是否相同
            //按照选好的行距，拷贝一行，修改这一行的数据，循环
            //提交

            int rowCountofSource = dt.Rows.Count;
            int columnCountofSource = dt.Columns.Count;

            //必须是同一行的才能算入选中的列数
            //先取第一个的Y值
            int columnCountofSelect = 0;
            DBText firstText = EntityCollection[EntityCollection.Count - 1] as DBText;
            double Y = firstText.AlignmentPoint.Y;
            foreach (DBObject dbo in EntityCollection)
            {
                DBText dbt = dbo as DBText;
                if (dbt.AlignmentPoint.Y == Y)
                {
                    columnCountofSelect = columnCountofSelect + 1;
                }
            }
            double dis = 0;
            if (columnCountofSource == columnCountofSelect)
            {
                //获取行距
                PromptDoubleResult pdr = ed.GetDistance("请输入行距\n");
                if (pdr.Status == PromptStatus.OK)
                {
                    dis = pdr.Value;
                }
                //拷贝、修改
                using (Transaction transaction = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = transaction.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord modelSpace = transaction.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    try
                    {
                        for (int i = 0; i < rowCountofSource; i++)
                        {
                            for (int j = 0; j < EntityCollection.Count; j++)
                            {
                                Entity ent = EntityCollection[EntityCollection.Count - j - 1] as Entity;
                                Entity newEnt = ent.Clone() as Entity;
                                newEnt.TransformBy(Matrix3d.Displacement(new Vector3d(0, (-1 * dis * i), 0)));
                                DBText newtext = newEnt as DBText;
                                newtext.TextString = dt.Rows[i][j].ToString();
                                modelSpace.AppendEntity(newtext);
                                transaction.AddNewlyCreatedDBObject(newtext, true);
                            }
                        }
                        transaction.Commit();
                        ed.WriteMessage("\n已成功写入！");
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

                //删除原始第一行用来选择的文字
                using (Transaction transaction = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = transaction.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord modelSpace = transaction.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    try
                    {
                        foreach (DBObject dbo in EntityCollection)
                        {
                            Entity ent = transaction.GetObject(dbo.ObjectId, OpenMode.ForWrite, true) as Entity;
                            ent.Erase(true);
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
            else
            {
                ed.WriteMessage(String.Format("Excel表头({0}个)与选取的个数({1}个)不符！", columnCountofSource, columnCountofSelect));
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
        public static Excel.Workbook wbk;
        public static Excel.Worksheet wsh;
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
            wbk = app.ActiveWorkbook;
            wsh = (Worksheet)wbk.ActiveSheet;
            Excel.Range rngLeftTop = null;
            Excel.Range rngRightButtom = null;
            rngLeftTop = (Excel.Range)app.InputBox("点击左上角单元格", Type: 8);
            rngRightButtom = (Excel.Range)app.InputBox("点击右下角单元格", Type: 8);
            object[,] data = (object[,])wsh.Range[rngLeftTop.Address + ":" + rngRightButtom.Address].Value2;
            //把数据导入Access数据库
            try
            {
                AccessDataBase.WriteDB(data);
                MessageBox.Show("选定的数据读取完毕，点击确定导入CAD");
            }
            catch (SystemException ex)
            {
                MessageBox.Show(ex.ToString());
            }

            //选定的数据读取完毕，点击确定导入CAD

            if (rngLeftTop != null & rngRightButtom != null)
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
                    }
                    else
                   ｛
                        AcadApp = new Autodesk.AutoCAD.Interop.AcadApplication();
                        AcadDoc = AcadApp.Documents.Open(filePath, null, null);
                    ｝
                }
                AcadApp.Application.Visible = true;
                //使CAD程序跳到在最前面，需要添加引用“Microsoft.VisualBasic”
                Microsoft.VisualBasic.Interaction.AppActivate(AcadApp.Caption);

                //让CAD自动执行netload命令加载程序集DLL,如果注册表加载方法无效的话
                AcadDoc.SendCommand("(command \"_netload\" \"" + @"C:\\Windows\\System\\xl2cad.dll" + "\") ");


                AcadDoc.SendCommand("zdbtext ");
            }
            else
            {
                MessageBox.Show("没有选择数据！");
            }
        }

        public void OnButton2Pressed(IRibbonControl control)
        {

        }
    }

    public class AccessDataBase
    {
        public static void WriteDB(object[,] data)
        {
            string filePath = @"C:\Windows\System\xl2cad.mdb";
            string mdbCommand = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Jet OLEDB:Engine Type=5";

            //创建数据库，有就删除重建，没有则新建
            ADOX.CatalogClass cat = new CatalogClass();
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
            cat.Create(mdbCommand);

            //连接数据库
            ADODB.Connection cn = new ADODB.Connection();
            cn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath, null, null, -1);
            cat = null;
            cat = new CatalogClass();
            cat.ActiveConnection = cn;

            //创建数据表
            ADOX.TableClass table = new TableClass();
            string tablename = "excel2cad";
            table.ParentCatalog = cat;
            table.Name = tablename;
            cat.Tables.Append(table);

            //将数据写入数据表
            //Excel.WorkSheet.Range导出的object[obj1,obj2]中,obj1为行，obj2为列
            //因此先按obj2的个数建立字段，在按obj1的个数一行行填入数据
            object[,] datas = data;
            int rowCount = 0;
            int columnCount = 0;
            rowCount = datas.GetUpperBound(0);      //第一维obj1的最大值，行数
            columnCount = datas.GetUpperBound(1);   //第二维obj2的最大值，列数

            //建立字段
            for (int i = 1; i <= columnCount; i++)
            {
                ADOX.ColumnClass col = null;
                col = new ADOX.ColumnClass();
                col.ParentCatalog = cat;
                col.Properties["Jet OLEDB:Allow Zero Length"].Value = true;
                col.Name = "Value" + i;
                table.Columns.Append(col, ADOX.DataTypeEnum.adVarChar, 25);
            }

            //按行填入数据
            object ra = null;
            ADODB.Recordset rs = new ADODB.Recordset();
            for (int i = 1; i <= rowCount; i++)
            {
                //构造按行写入的sql语句
                string sql1 = String.Format("INSERT INTO excel2cad (");
                string sql2 = String.Format(") VALUES (");
                string strValue = null;
                for (int j = 1; j <= columnCount; j++)
                {
                    if (data[i, j] == null)
                    {
                        strValue = datas[i, j] as string;
                    }
                    else
                    {
                        strValue = datas[i, j].ToString();
                    }

                    sql1 += String.Format("Value{0},", j);
                    sql2 += String.Format("'{0}',", strValue);
                }
                string sql = sql1 + sql2 + ")";
                //将 ,) 替换为 ）
                sql = sql.Replace(",)", ")");
                rs = cn.Execute(sql, out ra, -1);
            }
            cn.Close();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(cn);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(cat);

        }

        public static object[,] ReadDB()
        {
            object[,] data = null;

            return data;
        }

    }
}
