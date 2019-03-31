using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Text;
using System.IO;
using Office;
using System.Windows.Forms;
using Excel;
using Autodesk;
using AddInDesignerObjects;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;
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
                OleDbConnection con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data source=D:\\xl2cad.mdb");
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
                                //(-1 * dis * i)就是向下逐行添加，改为(1 * dis * i)就是向上，以后有需求可以增加判断
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
    public class Select
    {
        [CommandMethod("zselect")]
        public void zselect()
        {
            //Excel.Application app;
            //Excel.Workbook wbk;
            //Excel.Worksheet wsh;
            Database db = HostApplicationServices.WorkingDatabase;
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage("请选取矩形\n");
            Autodesk.AutoCAD.Interop.AcadApplication AcadApp = (Autodesk.AutoCAD.Interop.AcadApplication)System.Runtime.InteropServices.Marshal.GetActiveObject("AutoCAD.Application");
            Autodesk.AutoCAD.Interop.AcadDocument acaddoc = AcadApp.ActiveDocument;
            Entity entity = null;
            FilterType[] filter = new FilterType[1];
            //选取矩形之过滤条件
            filter[0] = FilterType.LWPolyline;
            //设置矩形框选中的表格的直线之过滤条件
            TypedValue[] tv = new TypedValue[1];
            tv.SetValue(new TypedValue((int)DxfCode.Start, "Line,LWPolyline"), 0);
            SelectionFilter sf = new SelectionFilter(tv);
            //设置矩形框选中的表格中文字之过滤条件
            TypedValue[] tv1 = new TypedValue[1];
            tv1.SetValue(new TypedValue((int)DxfCode.Start, "Text"), 0);
            SelectionFilter sf1 = new SelectionFilter(tv1);
            //形框选中的表格
            try
            {
                Dictionary<ObjectId, Entity> rectangleDic = GetRectangle(filter);
                //获取选中了的矩形
                Dictionary<ObjectId, Entity>.ValueCollection valcollect = rectangleDic.Values;
                DBObjectCollection rectangleCollection = new DBObjectCollection();
                //为每个矩形框中的内容生成表格
                List<midTable> tables = new List<midTable>();
                foreach (Entity ent in valcollect)
                {
                    rectangleCollection.Add(ent);
                }
                int count = rectangleCollection.Count;
                //可同时计算多个矩形框选中的表格
                for (int i = 0; i < count; i++)
                {
                    #region 获取计算表格单元格
                    int j = i + 1;
                    ed.WriteMessage("\n第" + j + "个矩形选择:\n");
                    Polyline pl = rectangleCollection[i] as Polyline;
                    Point3dCollection p3d = new Point3dCollection();
                    //ed.WriteMessage("\n1:" + pl.GetPoint3dAt(0).X + "  " + pl.GetPoint3dAt(0).Y);
                    //ed.WriteMessage("\n2:" + pl.GetPoint3dAt(1).X + "  " + pl.GetPoint3dAt(1).Y);
                    //ed.WriteMessage("\n3:" + pl.GetPoint3dAt(2).X + "  " + pl.GetPoint3dAt(2).Y);
                    //ed.WriteMessage("\n4:" + pl.GetPoint3dAt(3).X + "  " + pl.GetPoint3dAt(3).Y);
                    p3d.Add(pl.GetPoint3dAt(0));
                    p3d.Add(pl.GetPoint3dAt(1));
                    p3d.Add(pl.GetPoint3dAt(2));
                    p3d.Add(pl.GetPoint3dAt(3));
                    //获取选中的矩形框选选中的直线
                    DBObjectCollection lineCollection = new DBObjectCollection();
                    PromptSelectionResult ents = ed.SelectCrossingPolygon(p3d, sf);
                    if (ents.Status == PromptStatus.OK)
                    {
                        using (Transaction transaction = db.TransactionManager.StartTransaction())
                        {
                            SelectionSet SS = ents.Value;
                            foreach (ObjectId id in SS.GetObjectIds())
                            {
                                entity = (Entity)transaction.GetObject(id, OpenMode.ForRead, true);
                                //SelectCrossingPolygon（）会把框选用的矩形也包含
                                //要把这个矩形排除掉
                                bool isContain = rectangleDic.ContainsKey(id);
                                if (entity != null && !isContain)
                                {
                                    lineCollection.Add(entity);
                                    //ed.WriteMessage(entity.GetRXClass().DxfName + "\n");
                                }
                            }
                            ed.WriteMessage(lineCollection.Count.ToString());
                            transaction.Commit();
                        }
                    }
                    else
                    {
                        ed.WriteMessage("什么都没选到。");
                    }
                    List<Curve> HorizonLine = new List<Curve>();
                    List<Curve> VerticalLine = new List<Curve>();
                    foreach (Curve cur in lineCollection)
                    {
                        if (cur.StartPoint.X == cur.EndPoint.X)
                        {
                            VerticalLine.Add(cur);
                        }
                        else if (cur.StartPoint.Y == cur.EndPoint.Y)
                        {
                            HorizonLine.Add(cur);
                        }
                        else
                        {
                            ed.WriteMessage("表格里有斜线！" + "\n");
                            break;
                        }
                    }
                    //水平线从上往下按Y坐标由大到小排序
                    HorizonLine.Sort(new HorizontalLinePointComparer());
                    //垂直线从左往右按X坐标由小到大排序
                    VerticalLine.Sort(new VerticalLinePointComparer());
                    //ed.WriteMessage("\n水平线：\n-------------------\n");
                    //for (int a = 0; a < HorizonLine.Count;a++ )
                    //{
                    //    Curve cur = HorizonLine[a] as Curve;
                    //    ed.WriteMessage(cur.GetType().ToString());
                    //    ed.WriteMessage("  Y坐标为：" + cur.StartPoint.Y + "\n");
                    //}
                    //ed.WriteMessage("\n垂直线：\n-------------------\n");
                    //for (int b = 0; b < VerticalLine.Count; b++)
                    //{
                    //    Curve cur = VerticalLine[b] as Curve;
                    //    ed.WriteMessage(cur.GetType().ToString());
                    //    ed.WriteMessage("  X坐标为：" + cur.StartPoint.X + "\n");
                    //}
                    //划分单元格
                    //记录所有直线的交点
                    //交点的记录格式为 “1-1，[]”,1-1表示第一行第一列，[0]为X的坐标值,[1]为Y的坐标值
                    string pointName = null;
                    Dictionary<string, double[]> IntersectionPoints = new Dictionary<string, double[]>();
                    int horizonCount = HorizonLine.Count;
                    int verticalCount = VerticalLine.Count;
                    for (int m = 0; m < horizonCount; m++)
                    {
                        for (int n = 0; n < verticalCount; n++)
                        {
                            pointName = String.Format("{0}-{1}", m + 1, n + 1);
                            double[] intersectionpoint = new double[2];
                            Point3dCollection p = new Point3dCollection();
                            HorizonLine[m].IntersectWith(VerticalLine[n], Intersect.OnBothOperands, p, 0, 0);
                            if (p.Count == 1)
                            {
                                intersectionpoint[0] = p[0].X;
                                intersectionpoint[1] = p[0].Y;
                                IntersectionPoints.Add(pointName, intersectionpoint);
                            }
                        }
                    }
                    //foreach (string s in IntersectionPoints.Keys)
                    //{
                    //    ed.WriteMessage("\nName:" + s + " 交点坐标X:" + IntersectionPoints[s][0] + " 交点坐标Y:" + IntersectionPoints[s][1]);
                    //}
                    #endregion

                    #region 获取文字并填入单元格
                    //获取选中的矩形框选选中的文字
                    DBObjectCollection textCollection = new DBObjectCollection();
                    PromptSelectionResult ents1 = ed.SelectCrossingPolygon(p3d, sf1);
                    if (ents1.Status == PromptStatus.OK)
                    {
                        using (Transaction transaction = db.TransactionManager.StartTransaction())
                        {
                            SelectionSet SS = ents1.Value;
                            foreach (ObjectId id in SS.GetObjectIds())
                            {
                                entity = (Entity)transaction.GetObject(id, OpenMode.ForRead, true);
                                if (entity != null)
                                {
                                    textCollection.Add(entity);
                                    ed.WriteMessage("\n" + entity.GetRXClass().DxfName);
                                }
                            }
                            ed.WriteMessage("\n" + textCollection.Count.ToString());
                            transaction.Commit();
                        }
                    }
                    //保存选中数据的数组object[pointName,text]
                    object[,] table = new object[HorizonLine.Count + 1, VerticalLine.Count + 1];
                    string tablename = "table" + j;
                    //为每个文字找到对应的单元格，如果有多个文字在一个单元格范围内，则合并
                    //合并的操作在写EXCEL单元格的时候进行，加空格Append在原有文字的后面
                    //通过Text.AlignmentPoint来定位
                    foreach (Entity ent_text in textCollection)
                    {
                        DBText db_text = ent_text as DBText;
                        string pointName1 = "";
                        string x_text = "";
                        string y_text = "";
                        int x_index = 0;
                        int y_index = 0;
                        //先查找AlignmentPoint.X在"List<Curve> HorizonLine"中行的位置
                        for (int aa = 0; aa < HorizonLine.Count; aa++)
                        {
                            if (db_text.AlignmentPoint.Y > HorizonLine[aa].StartPoint.Y)
                            {
                                y_text = (aa + 1).ToString();
                                y_index = aa;
                                break;
                            }
                            else
                            {
                                if (aa == (HorizonLine.Count - 1))
                                {
                                    y_text = (aa + 2).ToString();
                                    y_index = aa + 1;
                                }
                            }
                        }
                        //再查找AlignmentPoint.X在"List<Curve> VerticalLine"中列的位置
                        for (int bb = 0; bb < VerticalLine.Count; bb++)
                        {
                            if (db_text.AlignmentPoint.X < VerticalLine[bb].StartPoint.X)
                            {
                                x_text = (bb + 1).ToString();
                                x_index = bb;
                                break;
                            }
                            else
                            {
                                if (bb == (VerticalLine.Count - 1))
                                {
                                    x_text = (bb + 2).ToString();
                                    x_index = bb + 1;
                                }
                            }
                        }
                        pointName1 = y_text + "-" + x_text;
                        //如果单元格内有多个文字，则将其内容拼在一起
                        table[y_index, x_index] += db_text.TextString;
                        ed.WriteMessage("\n" + db_text.TextString + " 位置为" + pointName1);
                    }

                    #endregion

                    #region  将数据写入midTable

                    midTable midtable = new midTable();
                    midtable.Name = tablename;
                    midtable.data = table;
                    tables.Add(midtable);

                    #endregion
                }

                //写入Access数据库
                xl2cad_wps.AccessDataBase.WriteDB(tables);
                //提示读取数据完毕，点击确定后跳到EXCEL导入数据
                MessageBox.Show("选定的数据读取完毕，点击确定导入Excel");
                //xl2cad_wps.AccessDataBase.ReadDB

                #region 将数据写入Excel

                //app = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("ket.Application");
                //wbk = app.ActiveWorkbook;
                //wsh = (Worksheet)wbk.ActiveSheet;
                ////使excel程序跳到在最前面，需要添加引用“Microsoft.VisualBasic”
                //Microsoft.VisualBasic.Interaction.AppActivate(app.Caption);
                string fileName1 = "";
                SaveFileDialog saveFile = new SaveFileDialog();
                saveFile.Filter = ("Excel 文件(*.xls)|*.xls");//指定文件后缀名为Excel 文件。
                if (saveFile.ShowDialog() == DialogResult.OK)
                {
                    fileName1 = saveFile.FileName;
                    if (System.IO.File.Exists(fileName1))
                    {
                        System.IO.File.Delete(fileName1);//如果文件存在删除文件。
                    }
                    int index = fileName1.LastIndexOf("//");//获取最后一个/的索引
                    fileName1 = fileName1.Substring(index + 1);//获取excel名称(新建表的路径相对于SaveFileDialog的路径)
                }
                foreach(midTable table in tables)
                {
                    xl2cad_wps.AccessDataBase.ReadDB(table.Name, fileName1);
                }

                #endregion

            }
            catch (System.Exception ex)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(ex.Message);
            }
            finally
            {
            }
        }
        private Dictionary<ObjectId, Entity> GetRectangle(FilterType[] filter)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            Entity entity = null;
            //矩形选择框的ObjectID和Entity组成的字典，选择矩形框的时候记录好对应关系，方便以后区分
            Dictionary<ObjectId, Entity> rectangleDic = new Dictionary<ObjectId, Entity>();
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
                                rectangleDic.Add(id, entity);
                            }
                        }
                        transaction.Commit();
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage(ex.Message);
                    }
                    finally
                    {
                        transaction.Dispose();
                    }
                }
            }
            return rectangleDic;
        }
    }
    public class HorizontalLinePointComparer : IComparer<Curve>
    {
        public int Compare(Curve x, Curve y)
        {
            if (x == null && y == null) return 0;
            if (x == null) return -1;
            if (y == null) return 1;
            double diff = x.StartPoint.Y - y.StartPoint.Y;
            if (diff > 0) return -1;
            if (diff < 0) return 1;
            return 0;
        }
    }
    public class VerticalLinePointComparer : IComparer<Curve>
    {
        public int Compare(Curve x, Curve y)
        {
            if (x == null && y == null) return 0;
            if (x == null) return -1;
            if (y == null) return 1;
            double diff = x.StartPoint.X - y.StartPoint.X;
            if (diff > 0) return 1;
            if (diff < 0) return -1;
            return 0;
        }
    }
    public class midTable
    {
        public object[,] data;
        public string Name;
    }
    public enum FilterType
    {
        Curve, Dimension, Polyline, LWPolyline, BlockRef, Circle, Line, Arc, Text, MText, Polyline3d, Surface,
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
                        MessageBox.Show("选择CAD文件无效!!", "文件无效!!");
                    }
                    else
                    {
                        AcadApp = new Autodesk.AutoCAD.Interop.AcadApplication();
                        AcadDoc = AcadApp.Documents.Open(filePath, null, null);
                    }
                }
                AcadApp.Application.Visible = true;
                //使CAD程序跳到在最前面，需要添加引用“Microsoft.VisualBasic”
                Microsoft.VisualBasic.Interaction.AppActivate(AcadApp.Caption);
                //让CAD自动执行netload命令加载程序集DLL,如果注册表加载方法无效的话
                AcadDoc.SendCommand("(command \"_netload\" \"" + @"C:\\Windows\\System\\xl2cad\\xl2cad.dll" + "\") ");
                AcadDoc.SendCommand("zdbtext ");
            }
            else
            {
                MessageBox.Show("没有选择数据！");
            }
        }
        public void OnButton2Pressed(IRibbonControl control)
        {
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
                    MessageBox.Show("选择CAD文件无效!!", "文件无效!!");
                }
                else
                {
                    AcadApp = new Autodesk.AutoCAD.Interop.AcadApplication();
                    AcadDoc = AcadApp.Documents.Open(filePath, null, null);
                }
            }
            AcadApp.Application.Visible = true;
            //使CAD程序跳到在最前面，需要添加引用“Microsoft.VisualBasic”
            Microsoft.VisualBasic.Interaction.AppActivate(AcadApp.Caption);
            //让CAD自动执行netload命令加载程序集DLL,如果注册表加载方法无效的话
            AcadDoc.SendCommand("(command \"_netload\" \"" + @"C:\\Windows\\System\\xl2cad\\xl2cad.dll" + "\") ");
            AcadDoc.SendCommand("zselect ");
        }
    }
    public class AccessDataBase
    {
        public static void WriteDB(object[,] data)
        {
            string filePath = @"D:\xl2cad.mdb";
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

        public static void WriteDB(List<xl2cad_cad.midTable> tables)
        {
            string filePath = @"D:\xl2cad.mdb";
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
            foreach (xl2cad_cad.midTable midtable in tables)
            {
                ADOX.TableClass table = new TableClass();
                string tablename = midtable.Name;
                table.ParentCatalog = cat;
                table.Name = tablename;
                cat.Tables.Append(table);
                //将数据写入数据表
                //Excel.WorkSheet.Range导出的object[obj1,obj2]中,obj1为行，obj2为列
                //因此先按obj2的个数建立字段，在按obj1的个数一行行填入数据
                object[,] datas = midtable.data;
                int rowCount = 0;
                int columnCount = 0;
                rowCount = datas.GetUpperBound(0);      //第一维obj1的最大值，行数
                columnCount = datas.GetUpperBound(1);   //第二维obj2的最大值，列数
                //建立字段
                for (int i = 0; i <= columnCount; i++)
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
                for (int i = 0; i <= rowCount; i++)
                {
                    //构造按行写入的sql语句
                    string sql1 = String.Format("INSERT INTO {0} (", tablename);
                    string sql2 = String.Format(") VALUES (");
                    string strValue = null;
                    for (int j = 0; j <= columnCount; j++)
                    {
                        if (datas[i, j] == null)
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
            }
            cn.Close();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(cn);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(cat);
        }

        public static void ReadDB(string tableName, string fileName)
        {
            string filePath = @"D:\xl2cad.mdb";
            if (File.Exists(filePath))
            {
                //select * into 建立 新的表。
                //[[Excel 8.0;database= excel名].[sheet名] 如果是新建sheet表不能加$,如果向sheet里插入数据要加$.　
                //sheet最多存储65535条数据。
                object ra = null;
                string sql = "select top 65535 * into [Excel 8.0;database=" + fileName + "].[" + tableName + "] from " + tableName;
                ADODB.Connection cn = new ADODB.Connection();
                cn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath, null, null, -1);
                cn.Execute(sql, out ra, -1);
                cn.Close();
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(cn);
            }
        }
    }
}