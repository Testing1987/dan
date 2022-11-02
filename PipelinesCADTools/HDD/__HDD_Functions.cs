using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;




using Microsoft.Office.Interop.Excel;
using System.Data;

namespace Alignment_mdi
{
    class Functions
    {

        public static bool is_dan_popescu()
        {
            if (Environment.UserName.ToUpper() == "POP70694")
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public static bool is_richard_pangburn()
        {
            if (Environment.UserName.ToUpper() == "PAN71158")
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static bool is_hector_morales()
        {
            if (Environment.UserName.ToUpper() == "MOR72937")
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public static bool is_eric_st_germain()
        {
            if (Environment.UserName.ToUpper() == "STG46680")
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static bool is_alex_dumais()
        {
            if (Environment.UserName.ToUpper() == "DUM64749")
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static bool is_joe_lynskey()
        {
            if (Environment.UserName.ToUpper() == "LYN69372")
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        public static bool is_monica_forgarty()
        {
            if (Environment.UserName.ToUpper() == "WHI45143")
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        public static bool is_eli_barboza()
        {
            if (Environment.UserName.ToUpper() == "BAR55261")
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        public static ObjectId GetObjectId(Database db, string handle)
        {
            try
            {
                return db.GetObjectId(false, new Handle(Convert.ToInt64(handle)), 0);
            }
            catch (System.Exception EX)
            {
                //MessageBox.Show(EX.Message + "\r\nObject ID not present in the drawing database");
                return ObjectId.Null;
            }

        }

        static public bool IsNumeric(string s)
        {
            double myNum = 0;
            if (double.TryParse(s, out myNum))
            {
                if (s.Contains(",")) return false;
                return true;
            }
            else
            {
                return false;
            }
        }

        static public Worksheet Get_active_worksheet_from_Excel()
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application Excel1;
                Microsoft.Office.Interop.Excel.Workbook Workbook1;
                Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                if (Excel1 == null) return null;
                Workbook1 = Excel1.ActiveWorkbook;
                return Workbook1.ActiveSheet;
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                return null;
            }



        }

        static public Worksheet Get_NEW_worksheet_from_Excel()
        {
            Microsoft.Office.Interop.Excel.Application Excel1;
            Microsoft.Office.Interop.Excel.Workbook Workbook1;
            try
            {
                Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            }
            catch (System.Exception ex)
            {
                Excel1 = new Microsoft.Office.Interop.Excel.Application();
            }

            try
            {
                Excel1.Visible = true;
                Excel1.Workbooks.Add();
                Workbook1 = Excel1.ActiveWorkbook;
                return Workbook1.ActiveSheet;
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                return null;
            }


        }



        static public int Get_no_of_workbooks_from_Excel()
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application Excel1;
                Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                if (Excel1 == null) return 0;
                return Excel1.Workbooks.Count;
            }
            catch (System.Exception ex)
            {
                return 0;
            }
        }

        static public void Kill_excel()
        {
            List<int> ProcessID = Functions.GetAllExcelProcessID();
            if (ProcessID.Count > 0)
            {
                foreach (int Id in ProcessID)
                {
                    try
                    {
                        System.Diagnostics.Process proc = System.Diagnostics.Process.GetProcessById(Id);
                        // Microsoft.Office.Interop.Excel.Application Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                        try
                        {
                            try
                            {
                                if (System.Diagnostics.Process.GetProcessById(Id).MainWindowTitle.ToString() == "")
                                {
                                    proc.Kill();

                                }
                            }
                            catch (System.InvalidOperationException ex)
                            {

                            }
                        }
                        catch (System.ComponentModel.Win32Exception ex)
                        {

                        }

                        //MessageBox.Show(Process.GetProcessById(Id).MainWindowHandle.ToString());
                        //  
                    }
                    catch (System.ArgumentException ex)
                    {

                    }
                }


            }
        }

        static public Workbook Get_Existing_workbook_from_Excel(string name1, string name2)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application Excel1;

                Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                if (Excel1 == null) return null;

                foreach (Microsoft.Office.Interop.Excel.Workbook Workbook1 in Excel1.Workbooks)
                {
                    bool exista1 = false;
                    bool exista2 = false;
                    foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook1.Worksheets)
                    {
                        if (Wx.Name.ToUpper() == name1.ToUpper())
                        {
                            exista1 = true;
                        }
                        if (Wx.Name.ToUpper() == name2.ToUpper())
                        {
                            exista2 = true;
                        }
                    }
                    if (exista1 == true && exista2 == true)
                    {
                        return Workbook1;
                    }
                }
                return null;
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                return null;
            }
        }

        static public List<int> GetAllExcelProcessID()
        {
            List<int> ProcessID = new List<int>();

            List<System.Diagnostics.Process> currentExcelProcessList = System.Diagnostics.Process.GetProcessesByName("EXCEL").ToList();
            foreach (var item in currentExcelProcessList)
            {
                ProcessID.Add(item.Id);
            }

            return ProcessID;
        }

        public static bool is_civil_3d()
        {
            System.Reflection.Assembly[] assemblies = AppDomain.CurrentDomain.GetAssemblies();

            foreach (System.Reflection.Assembly asm in assemblies)
            {
                System.Reflection.AssemblyName asmName = asm.GetName();
                if (asmName.Name == "AeccDbMgd")
                {
                    return true;
                }
                //FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(asm.Location);
                //string name = asmName.Name;
                //Version asmV = asmName.Version;
                //string fileV = fvi.FileVersion;
                //string prodV = fvi.ProductVersion;
            }
            return false;
        }


        int GetApplicationExcelProcessID(List<int> ProcessID1, List<int> ProcessID2)
        {
            foreach (var processid in ProcessID2)
            {
                if (!ProcessID1.Contains(processid))
                {
                    return processid;
                }
            }
            return -1;
        }

        void KillExcel(int ProcessID)
        {
            System.Diagnostics.Process process = System.Diagnostics.Process.GetProcessById(ProcessID);
            process.Kill();
        }

        static public void Close_excel_processes()
        {
            try
            {
                foreach (System.Diagnostics.Process proc in System.Diagnostics.Process.GetProcessesByName("utorrent"))
                {
                    proc.Kill();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }





        }

        public void closeOpenedFile(string file_name)
        {
            //Excel Application Object
            Microsoft.Office.Interop.Excel.Application oExcelApp;
            //Get reference to Excel.Application from the ROT.
            if (System.Diagnostics.Process.GetProcessesByName("EXCEL").Count() > 0)
            {
                oExcelApp = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");

                foreach (Microsoft.Office.Interop.Excel.Workbook WB in oExcelApp.Workbooks)
                {
                    //MessageBox.Show(WB.FullName);
                    if (WB.Name == file_name)
                    {
                        WB.Save();
                        WB.Close();
                        //oExcelApp.Quit();
                    }
                }
            }
        }

        static public string extrage_station_din_text_de_la_inceputul_textului(string string1)
        {


            string station_string = "";

            if (string1.Contains("+") == true)
            {
                for (int i = 0; i < string1.Length; ++i)
                {
                    String Litera = string1.Substring(i, 1);

                    switch (Litera)
                    {

                        case ".":
                            if (i == station_string.Length)
                            {
                                station_string = station_string + Litera;
                            }
                            break;
                        case "0":
                            if (i == station_string.Length)
                            {
                                station_string = station_string + Litera;
                            }
                            break;
                        case "1":
                            if (i == station_string.Length)
                            {
                                station_string = station_string + Litera;
                            }
                            break;
                        case "2":
                            if (i == station_string.Length)
                            {
                                station_string = station_string + Litera;
                            }
                            break;
                        case "3":
                            if (i == station_string.Length)
                            {
                                station_string = station_string + Litera;
                            }
                            break;
                        case "4":
                            if (i == station_string.Length)
                            {
                                station_string = station_string + Litera;
                            }
                            break;
                        case "5":
                            if (i == station_string.Length)
                            {
                                station_string = station_string + Litera;
                            }
                            break;
                        case "6":
                            if (i == station_string.Length)
                            {
                                station_string = station_string + Litera;
                            }
                            break;
                        case "7":
                            if (i == station_string.Length)
                            {
                                station_string = station_string + Litera;
                            }
                            break;
                        case "8":
                            if (i == station_string.Length)
                            {
                                station_string = station_string + Litera;
                            }
                            break;
                        case "9":
                            if (i == station_string.Length)
                            {
                                station_string = station_string + Litera;
                            }
                            break;
                        case "+":
                            if (i == station_string.Length)
                            {
                                station_string = station_string + Litera;
                            }
                            break;
                        case "-":
                            if (i == 0)
                            {
                                station_string = station_string + Litera;
                            }
                            break;

                        default:
                            break;

                    }
                }
            }


            return station_string;

        }

        static public void Creaza_layer(string Layername, short Culoare, bool Plot)
        {
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.LayerTable LayerTable1;
                        LayerTable1 = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        if (LayerTable1.Has(Layername) == true)
                        {
                            LayerTable1.UpgradeOpen();
                            LayerTableRecord new_layer = Trans1.GetObject(LayerTable1[Layername], OpenMode.ForWrite) as LayerTableRecord;
                            if (new_layer != null)
                            {
                                new_layer.IsPlottable = Plot;

                            }
                        }

                        if (LayerTable1.Has(Layername) == false)
                        {
                            LayerTableRecord new_layer = new Autodesk.AutoCAD.DatabaseServices.LayerTableRecord();
                            new_layer.Name = Layername;
                            new_layer.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, Culoare);
                            new_layer.IsPlottable = Plot;
                            LayerTable1.Add(new_layer);
                            Trans1.AddNewlyCreatedDBObject(new_layer, true);

                        }

                        Trans1.Commit();
                    }
                }


            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }



        static public void Creaza_layer_on_database(Database Database1, string Layername, short Culoare, bool Plot)
        {
            try
            {

                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = Database1.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.LayerTable LayerTable1;
                    LayerTable1 = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(Database1.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                    if (LayerTable1.Has(Layername) == true)
                    {
                        LayerTable1.UpgradeOpen();
                        LayerTableRecord new_layer = Trans1.GetObject(LayerTable1[Layername], OpenMode.ForWrite) as LayerTableRecord;
                        if (new_layer != null)
                        {
                            new_layer.IsPlottable = Plot;
                            Trans1.Commit();
                        }
                    }

                    if (LayerTable1.Has(Layername) == false)
                    {
                        LayerTable1.UpgradeOpen();
                        LayerTableRecord new_layer = new Autodesk.AutoCAD.DatabaseServices.LayerTableRecord();
                        new_layer.Name = Layername;
                        new_layer.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, Culoare);
                        new_layer.IsPlottable = Plot;
                        LayerTable1.Add(new_layer);
                        Trans1.AddNewlyCreatedDBObject(new_layer, true);
                        Trans1.Commit();
                    }


                    Trans1.Dispose();

                }



            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        static public String verify_layername_from_combobox_different_database(Database Database1, System.Windows.Forms.ComboBox Combo_layername)
        {

            string Layer_name = "0";
            if (Combo_layername.Text != "")
            {

                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = Database1.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.LayerTable Layer_table = Trans1.GetObject(Database1.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as LayerTable;

                    if (Layer_table.Has(Combo_layername.Text) == true)
                    {
                        Layer_name = Combo_layername.Text;
                    }
                    Trans1.Dispose();
                }




            }


            return Layer_name;
        }

        static public void Incarca_existing_layers_to_combobox__different_database(Database Database1, System.Windows.Forms.ComboBox Combo_layer)
        {

            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = Database1.TransactionManager.StartTransaction())
            {
                Autodesk.AutoCAD.DatabaseServices.LayerTable Layer_table = Trans1.GetObject(Database1.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as LayerTable;

                Combo_layer.Items.Clear();
                foreach (ObjectId Layer_id in Layer_table)
                {
                    LayerTableRecord Layer1 = (LayerTableRecord)Trans1.GetObject(Layer_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                    string Name_of_layer = Layer1.Name;
                    if (Name_of_layer.Contains("|") == false & Name_of_layer.Contains("$") == false)
                    {
                        Combo_layer.Items.Add(Name_of_layer);
                    }
                }
                Combo_layer.SelectedIndex = 0;
                Trans1.Dispose();
            }
        }

        static public Point3dCollection Intersect_on_both_operands(Curve Curba1, Curve Curba2)
        {
            Point3dCollection Col_int = new Point3dCollection();
            Point3dCollection Col_int_on_both_operands = new Point3dCollection();
            Point3dCollection Col_int_on_both_operands_DUPLICATE = new Point3dCollection();

            Curba1.IntersectWith(Curba2, Intersect.OnBothOperands, Col_int, IntPtr.Zero, IntPtr.Zero);

            if (Col_int.Count == 1) return Col_int;
            if (Col_int.Count == 0) return Col_int;

            if (Col_int.Count > 1)
            {
                if (Curba1 is Polyline & Curba2 is Polyline)
                {
                    for (int i = 0; i < Col_int.Count; ++i)
                    {
                        Point3d Pt1 = new Point3d();
                        Pt1 = Col_int[i];
                        try
                        {
                            double param_on_1 = Curba1.GetParameterAtPoint(Pt1);
                            double param_on_2 = Curba2.GetParameterAtPoint(Pt1);


                            if (Col_int_on_both_operands_DUPLICATE.Contains(new Point3d(Math.Round(Pt1.X, 4), Math.Round(Pt1.Y, 4), Math.Round(Pt1.Z, 4))) == false)
                            {
                                Col_int_on_both_operands.Add(Pt1);
                                Col_int_on_both_operands_DUPLICATE.Add(new Point3d(Math.Round(Pt1.X, 4), Math.Round(Pt1.Y, 4), Math.Round(Pt1.Z, 4)));
                            }
                        }
                        catch (System.Exception ex)
                        {
                        }
                    }
                    return Col_int_on_both_operands;
                }
                else
                {
                    return Col_int;
                }
            }
            else
            {
                return Col_int;
            }
        }


        static public Point3dCollection Intersect_with_extend(Curve Curba_extend_this, Curve Curba2)
        {
            Point3dCollection Col_int = new Point3dCollection();
            Point3dCollection Col_int_on_both_operands = new Point3dCollection();
            Point3dCollection Col_int_on_both_operands_DUPLICATE = new Point3dCollection();

            Curba_extend_this.IntersectWith(Curba2, Intersect.ExtendThis, Col_int, IntPtr.Zero, IntPtr.Zero);

            if (Col_int.Count == 1) return Col_int;
            if (Col_int.Count == 0) return Col_int;

            if (Col_int.Count > 1)
            {
                if (Curba_extend_this is Polyline & Curba2 is Polyline)
                {
                    for (int i = 0; i < Col_int.Count; ++i)
                    {
                        Point3d Pt1 = new Point3d();
                        Pt1 = Col_int[i];
                        try
                        {
                            double param_on_1 = Curba_extend_this.GetParameterAtPoint(Pt1);
                            double param_on_2 = Curba2.GetParameterAtPoint(Pt1);


                            if (Col_int_on_both_operands_DUPLICATE.Contains(new Point3d(Math.Round(Pt1.X, 4), Math.Round(Pt1.Y, 4), Math.Round(Pt1.Z, 4))) == false)
                            {
                                Col_int_on_both_operands.Add(Pt1);
                                Col_int_on_both_operands_DUPLICATE.Add(new Point3d(Math.Round(Pt1.X, 4), Math.Round(Pt1.Y, 4), Math.Round(Pt1.Z, 4)));
                            }
                        }
                        catch (System.Exception ex)
                        {
                        }
                    }
                    return Col_int_on_both_operands;
                }
                else
                {
                    return Col_int;
                }
            }
            else
            {
                return Col_int;
            }
        }





        static public Point3dCollection Intersect_with_extend_2d_3d(Polyline3d Poly3d, Polyline Poly2d)
        {
            Point3dCollection Col_int = new Point3dCollection();
            Point3dCollection Col_int_on_both_operands = new Point3dCollection();
            Point3dCollection Col_int_on_both_operands_DUPLICATE = new Point3dCollection();

            Polyline Poly1 = Build_2dpoly_from_3d(Poly3d);
            Poly1.Elevation = Poly2d.Elevation;

            Poly1.IntersectWith(Poly2d, Intersect.ExtendBoth, Col_int, IntPtr.Zero, IntPtr.Zero);

            if (Col_int.Count == 1) return Col_int;
            if (Col_int.Count == 0) return Col_int;

            if (Col_int.Count > 1)
            {

                for (int i = 0; i < Col_int.Count; ++i)
                {
                    Point3d Pt1 = new Point3d();
                    Pt1 = Col_int[i];
                    try
                    {
                        double param_on_1 = Poly1.GetParameterAtPoint(Pt1);
                        double param_on_2 = Poly2d.GetParameterAtPoint(Pt1);


                        if (Col_int_on_both_operands_DUPLICATE.Contains(new Point3d(Math.Round(Pt1.X, 4), Math.Round(Pt1.Y, 4), Math.Round(Pt1.Z, 4))) == false)
                        {
                            Col_int_on_both_operands.Add(Pt1);
                            Col_int_on_both_operands_DUPLICATE.Add(new Point3d(Math.Round(Pt1.X, 4), Math.Round(Pt1.Y, 4), Math.Round(Pt1.Z, 4)));
                        }
                    }
                    catch (System.Exception ex)
                    {
                    }
                }
                return Col_int_on_both_operands;


            }
            else
            {
                return Col_int;
            }
        }

        static public double GET_Bearing_rad(double x1, double y1, double x2, double y2)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;
            CoordinateSystem3d CurentUCS = CurentUCSmatrix.CoordinateSystem3d;
            Plane Planul_curent = new Plane(new Point3d(0, 0, 0), Vector3d.ZAxis);
            return new Point3d(x1, y1, 0).GetVectorTo(new Point3d(x2, y2, 0)).AngleOnPlane(Planul_curent);
        }

        static public double GET_deltaX_rad()
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;
            CoordinateSystem3d CurentUCS = CurentUCSmatrix.CoordinateSystem3d;
            Plane Planul_curent = new Plane(new Point3d(0, 0, 0), Vector3d.ZAxis);
            return CurentUCS.Xaxis.AngleOnPlane(Planul_curent);
        }

        static public double Get_deflection_angle_rad(double x1, double y1, double x2, double y2, double x3, double y3)
        {
            double a1 = x2 - x1;
            double b1 = y2 - y1;
            double a2 = x3 - x2;
            double b2 = y3 - y2;
            double Defl_DD = Math.Acos((a1 * a2 + b1 * b2) / (Math.Pow(a1 * a1 + b1 * b1, 0.5) * Math.Pow(a2 * a2 + b2 * b2, 0.5)));
            //return Defl_DD;

            Vector3d vector1 = new Point3d(x1, y1, 0).GetVectorTo(new Point3d(x2, y2, 0));
            Vector3d vector2 = new Point3d(x2, y2, 0).GetVectorTo(new Point3d(x3, y3, 0));
            return (vector2.GetAngleTo(vector1));


        }

        static public string Get_deflection_angle_dms(double x1, double y1, double x2, double y2, double x3, double y3)
        {


            double a1 = x2 - x1;
            double b1 = y2 - y1;
            double a2 = x3 - x2;
            double b2 = y3 - y2;
            double Defl_DD = 180 * Math.Acos((a1 * a2 + b1 * b2) / (Math.Pow(a1 * a1 + b1 * b1, 0.5) * Math.Pow(a2 * a2 + b2 * b2, 0.5))) / Math.PI;

            Vector3d vector1 = new Point3d(x1, y1, 0).GetVectorTo(new Point3d(x2, y2, 0));
            Vector3d vector2 = new Point3d(x2, y2, 0).GetVectorTo(new Point3d(x3, y3, 0));
            Defl_DD = (vector2.GetAngleTo(vector1)) * 180 / Math.PI;


            double Bearing1 = 180 * Functions.GET_Bearing_rad(x1, y1, x2, y2) / Math.PI;
            double Bearing2 = 180 * Functions.GET_Bearing_rad(x2, y2, x3, y3) / Math.PI;

            String Suffix1 = " ";


            if (Bearing1 < 180)
            {

                if (Bearing2 < Bearing1 + 180 && Bearing2 > Bearing1)
                {
                    Suffix1 = " LT";
                }
                else
                {
                    Suffix1 = " RT";
                }
            }
            else
            {
                if (Bearing2 < Bearing1 && Bearing2 > Bearing1 - 180)
                {
                    Suffix1 = " RT";
                }
                else
                {
                    Suffix1 = " LT";
                }
            }

            return Get_DMS(Defl_DD, 0) + Suffix1;



        }

        public static string Angle_left_right(Polyline Poly2D, Point3d Punct1)
        {
            String LT_RT = "";
            Point3d Point_on_poly = Poly2D.GetClosestPointTo(Punct1, Autodesk.AutoCAD.Geometry.Vector3d.ZAxis, true);
            Autodesk.AutoCAD.Geometry.Vector3d vector2 = Point_on_poly.GetVectorTo(Punct1);
            double Param1 = Poly2D.GetParameterAtPoint(Point_on_poly);
            Autodesk.AutoCAD.Geometry.Vector3d vector1;
            if (Param1 > 0)
            {
                if (Param1 == Poly2D.NumberOfVertices - 1)
                {
                    vector1 = Poly2D.GetPointAtParameter(Param1 - 1).GetVectorTo(Poly2D.GetPointAtParameter(Param1));
                }
                else
                {
                    vector1 = Poly2D.GetPointAtParameter(Math.Floor(Param1)).GetVectorTo(Poly2D.GetPointAtParameter(Math.Ceiling(Param1)));
                }
            }
            else
            {
                vector1 = Poly2D.GetPointAtParameter(0).GetVectorTo(Poly2D.GetPointAtParameter(1));
            }
            Plane Planul_curent = new Plane(new Point3d(0, 0, 0), Autodesk.AutoCAD.Geometry.Vector3d.ZAxis);
            double Bearing1 = (vector1.AngleOnPlane(Planul_curent)) * 180 / Math.PI;
            double Bearing2 = (vector2.AngleOnPlane(Planul_curent)) * 180 / Math.PI;
            double angle1 = (vector2.GetAngleTo(vector1)) * 180 / Math.PI;
            if (Bearing1 < 180)
            {
                if (Bearing2 < Bearing1 + 180 && Bearing2 > Bearing1)
                {
                    LT_RT = "LT.";
                }
                else
                {
                    LT_RT = "RT.";
                }
            }
            else
            {
                if (Bearing2 < Bearing1 & Bearing2 > Bearing1 - 180)
                {
                    LT_RT = "RT.";
                }
                else
                {
                    LT_RT = "LT.";
                }
            }
            return LT_RT;
        }

        static public string Get_Quadrant_bearing(double Radian1)
        {
            string Prefix1 = "N ";
            string Suffix1 = " E";
            double Quadrant1 = Math.PI / 2 - Radian1;
            if (Radian1 > Math.PI / 2 & Radian1 <= Math.PI)
            {
                Quadrant1 = Radian1 - Math.PI / 2;
                Suffix1 = " W";
            }
            if (Radian1 > Math.PI & Radian1 <= 3 * Math.PI / 2)
            {
                Quadrant1 = 3 * Math.PI / 2 - Radian1;
                Prefix1 = "S ";
                Suffix1 = " W";
            }
            if (Radian1 > 3 * Math.PI / 2)
            {
                Quadrant1 = Radian1 - 3 * Math.PI / 2;
                Prefix1 = "S ";
                Suffix1 = " E";
            }
            return Prefix1 + Get_DMS(Quadrant1 * 180 / Math.PI, 0) + Suffix1;
        }

        static public string Get_DMS(double Numar, int round_seconds)
        {

            bool Negative = false;
            if (Numar < 0)
            {
                Negative = true;
                Numar = -Numar;
            }
            int Degree1 = Convert.ToInt32(Math.Floor(Numar));



            int Minutes1 = Convert.ToInt32(Math.Floor((Numar - Convert.ToDouble(Degree1)) * 60));

            double rest1 = Convert.ToDouble(Degree1) + Convert.ToDouble(Minutes1) / 60;
            double Seconds1 = Math.Round((Numar - rest1) * 3600, round_seconds);



            if (Seconds1 == 60)
            {
                Minutes1 = Minutes1 + 1;
                Seconds1 = 0;
            }

            if (Minutes1 == 60)
            {
                Degree1 = Degree1 + 1;
                Minutes1 = 0;
            }

            string D = Degree1.ToString();

            if (Negative == true) D = "-" + D;

            string M = Minutes1.ToString();
            string S = Get_String_Rounded(Seconds1, round_seconds);

            if (M.Length == 1)
            {
                M = "0" + M;
            }

            if (Seconds1 < 10)
            {
                S = "0" + S;
            }

            char deg_symbol = (char)176;
            char sec_symbol = (char)34;

            return D + deg_symbol + M + "'" + S + sec_symbol;
        }

        public static void Add_to_clipboard_Data_table(System.Data.DataTable Data_table)
        {
            String sTR1 = "";

            if (Data_table.Rows.Count > 0)
            {

                for (int i = 0; i < Data_table.Columns.Count; ++i)
                {
                    if (i == 0)
                    {
                        sTR1 = Data_table.Columns[i].ColumnName;
                    }
                    else
                    {
                        sTR1 = sTR1 + "\t" + Data_table.Columns[i].ColumnName;
                    }
                }
                for (int i = 0; i < Data_table.Rows.Count; ++i)
                {
                    sTR1 = sTR1 + "\r\n";

                    for (int j = 0; j < Data_table.Columns.Count; ++j)
                    {
                        if (Data_table.Rows[i][j] != DBNull.Value)
                        {


                            if (j == 0)
                            {
                                sTR1 = sTR1 + Data_table.Rows[i][j].ToString();
                            }
                            else
                            {
                                sTR1 = sTR1 + "\t" + Data_table.Rows[i][j].ToString();
                            }
                        }
                        else
                        {
                            if (j == 0)
                            {
                                sTR1 = sTR1 + "";
                            }
                            else
                            {
                                sTR1 = sTR1 + "\t" + "";
                            }
                        }

                    }

                }

            }


            Clipboard.SetText(sTR1);
        }

        public static void Transfer_to_worksheet_Data_table(Worksheet W1, System.Data.DataTable Data_table, int Start_row, string format_cell)
        {


            W1.Columns["A:XX"].Delete();

            if (Data_table != null)
            {
                if (Data_table.Rows.Count > 0)
                {
                    int NrR = Data_table.Rows.Count;
                    int NrC = Data_table.Columns.Count;


                    Object[,] values = new object[NrR + 1, NrC];
                    for (int i = 0; i < NrR; ++i)
                    {
                        for (int j = 0; j < NrC; ++j)
                        {
                            if (Data_table.Rows[i][j] != DBNull.Value)
                            {
                                values[i + 1, j] = Data_table.Rows[i][j];
                            }
                        }
                    }
                    for (int j = 0; j < NrC; ++j)
                    {
                        values[0, j] = Data_table.Columns[j].ColumnName;
                    }


                    Microsoft.Office.Interop.Excel.Range range0 = W1.Range[W1.Columns[1], W1.Columns[NrC]];
                    range0.ClearContents();
                    range0.UnMerge();
                    range0.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                    range0.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                    range0.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                    range0.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;

                    Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[Start_row, 1], W1.Cells[NrR + Start_row, NrC]];
                    range1.Cells.NumberFormat = format_cell;
                    range1.Value2 = values;
                    Color_border_range_inside(range1, 0);

                }
            }
        }

        public static void Create_header_centerline_file(Worksheet W1, string Client, string Project, string Segment, bool CSF, string Version)
        {


            Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A1:B8"];


            Object[,] valuesH = new object[8, 2];

            valuesH[0, 0] = "CLIENT";
            valuesH[0, 1] = Client;
            valuesH[1, 0] = "PROJECT";
            valuesH[1, 1] = Project;
            valuesH[2, 0] = "SEGMENT";
            valuesH[2, 1] = Segment;
            valuesH[3, 0] = "VERSION";
            valuesH[3, 1] = Version;
            valuesH[4, 0] = "DATE CREATED";
            valuesH[4, 1] = DateTime.Now.ToString(new System.Globalization.CultureInfo("en-US"));
            valuesH[5, 0] = "USER ID";
            valuesH[5, 1] = Environment.UserName;
            valuesH[6, 0] = "Do not manually edit any of the table information below.";
            valuesH[7, 0] = "Do not add any columns to this table, also do not add any rows above row 10";
            range1.Value2 = valuesH;

            range1 = W1.Range["A1:B6"];
            Color_border_range_inside(range1, 46);

            if (CSF == false)
            {
                range1 = W1.Range["A7:R7"];
            }
            else
            {
                range1 = W1.Range["A7:T7"];
            }

            Color_border_range_outside(range1, 6);

            if (CSF == false)
            {
                range1 = W1.Range["A8:R8"];
            }
            else
            {
                range1 = W1.Range["A8:T8"];
            }

            Color_border_range_outside(range1, 3);

            if (CSF == false)
            {
                range1 = W1.Range["A9:R9"];
            }
            else
            {
                range1 = W1.Range["A9:T9"];
            }

            Color_border_range_inside(range1, 41);


            if (CSF == false)
            {
                range1 = W1.Range["C1:R6"];
            }
            else
            {
                range1 = W1.Range["C1:T6"];
            }
            range1.Merge();
            range1.MergeCells = true;
            range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            range1.Value2 = "Centerline";
            range1.Font.Name = "Arial Black";
            range1.Font.Size = 20;
            range1.Font.Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleSingle;
            Color_border_range_outside(range1, 0);
            if (CSF == false)
            {
                range1 = W1.Range["A9:R9"];
            }
            else
            {
                range1 = W1.Range["A9:T9"];

            }
            range1.Font.Color = 16777215;
            range1.Font.Bold = true;

        }





        public static string get_excel_column_letter(int intCol)
        {

            string columnString = "";
            decimal columnNumber = intCol;
            while (columnNumber > 0)
            {
                decimal currentLetterNumber = (columnNumber - 1) % 26;
                char currentLetter = (char)(currentLetterNumber + 65);
                columnString = currentLetter + columnString;
                columnNumber = (columnNumber - (currentLetterNumber + 1)) / 26;
            }
            return columnString;
        }

        public static void Create_header_block_attributes_file(Worksheet W1, string Client, string Project, string Segment, int nr_coloane)
        {
            string Last_coloana = get_excel_column_letter(nr_coloane);

            W1.Columns["A:XX"].Delete();

            Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A1:B6"];

            Object[,] valuesH = new object[6, 2];

            valuesH[0, 0] = "CLIENT";
            valuesH[0, 1] = Client;
            valuesH[1, 0] = "PROJECT";
            valuesH[1, 1] = Project;
            valuesH[2, 0] = "SEGMENT";
            valuesH[2, 1] = Segment;
            valuesH[3, 0] = "VERSION";
            valuesH[4, 0] = "DATE CREATED";
            valuesH[4, 1] = DateTime.Now.ToString(new System.Globalization.CultureInfo("en-US"));
            valuesH[5, 0] = "USER ID";
            valuesH[5, 1] = Environment.UserName;



            range1.Value2 = valuesH;

            Color_border_range_inside(range1, 46);




            Microsoft.Office.Interop.Excel.Range range3 = W1.Range["A7:" + Last_coloana + "7"];
            Color_border_range_inside(range3, 43);

            Microsoft.Office.Interop.Excel.Range range4 = W1.Range["A8:" + Last_coloana + "8"];
            range4.Font.Color = 16777215;
            range4.Font.Bold = true;
            Color_border_range_inside(range4, 41);


            Microsoft.Office.Interop.Excel.Range range5 = W1.Range["C1:" + Last_coloana + "6"];
            range5.Merge();
            range5.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            range5.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            range5.Value2 = "TBLK Attributes Table";
            range5.Font.Name = "Arial Black";
            range5.Font.Size = 20;
            range5.Font.Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleSingle;


            Color_border_range_outside(range5, 0);


        }

        public static void Color_border_range_inside(Microsoft.Office.Interop.Excel.Range range1, int cid)
        {

            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            range1.Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternNone;
            range1.Interior.TintAndShade = 0;
            range1.Interior.PatternTintAndShade = 0;
            if (cid != 0)
            {
                range1.Interior.ColorIndex = cid;
            }

        }

        private static void Color_border_range_outside(Microsoft.Office.Interop.Excel.Range range1, int cid)
        {


            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
            range1.Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternNone;
            range1.Interior.TintAndShade = 0;
            range1.Interior.PatternTintAndShade = 0;
            if (cid != 0)
            {
                range1.Interior.ColorIndex = cid;
            }
        }

        private static void Clear_formatting_range(Microsoft.Office.Interop.Excel.Range range1)
        {

            range1.Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternNone;
            range1.Interior.TintAndShade = 0;
            range1.Interior.PatternTintAndShade = 0;

        }

        public static System.Data.DataTable Creaza_centerline_datatable_structure()
        {

            string Col_MMid = "MMID";
            string Col_Type = "Type";
            string Col_x = "X";
            string Col_y = "Y";
            string Col_z = "Z";
            string Col_2DSta = "2DSta";
            string Col_3DSta = "3DSta";
            string Col_EqSta = "EqSta";
            string Col_BackSta = "BackSta";
            string Col_AheadSta = "AheadSta";
            string Col_DeflAng = "DeflAng";
            string Col_DeflAngDMS = "DeflAngDMS";
            string Col_Bearing = "Bearing";
            string Col_Distance = "Distance";
            string Col_DisplaySta = "DisplaySta";
            string Col_DisplayPI = "DisplayPI";
            string Col_DisplayProf = "DisplayProf";
            string Col_Symbol = "Symbol";

            System.Type type_MMid = typeof(string);
            System.Type type_Type = typeof(string);
            System.Type type_x = typeof(double);
            System.Type type_y = typeof(double);
            System.Type type_z = typeof(double);
            System.Type type_2DSta = typeof(double);
            System.Type type_3DSta = typeof(double);
            System.Type type_EqSta = typeof(double);
            System.Type type_BackSta = typeof(double);
            System.Type type_AheadSta = typeof(double);
            System.Type type_DeflAng = typeof(double);
            System.Type type_DeflAngDMS = typeof(string);
            System.Type type_Bearing = typeof(string);
            System.Type type_Distance = typeof(double);
            System.Type type_DisplaySta = typeof(double);
            System.Type type_DisplayPI = typeof(int);
            System.Type type_DisplayProf = typeof(int);
            System.Type type_Symbol = typeof(string);


            List<string> Lista1 = new List<string>();
            List<System.Type> Lista2 = new List<System.Type>();

            Lista1.Add(Col_MMid);
            Lista1.Add(Col_Type);
            Lista1.Add(Col_x);
            Lista1.Add(Col_y);
            Lista1.Add(Col_z);
            Lista1.Add(Col_2DSta);
            Lista1.Add(Col_3DSta);
            Lista1.Add(Col_EqSta);
            Lista1.Add(Col_BackSta);
            Lista1.Add(Col_AheadSta);
            Lista1.Add(Col_DeflAng);
            Lista1.Add(Col_DeflAngDMS);
            Lista1.Add(Col_Bearing);
            Lista1.Add(Col_Distance);
            Lista1.Add(Col_DisplaySta);
            Lista1.Add(Col_DisplayPI);
            Lista1.Add(Col_DisplayProf);
            Lista1.Add(Col_Symbol);

            Lista2.Add(type_MMid);
            Lista2.Add(type_Type);
            Lista2.Add(type_x);
            Lista2.Add(type_y);
            Lista2.Add(type_z);
            Lista2.Add(type_2DSta);
            Lista2.Add(type_3DSta);
            Lista2.Add(type_EqSta);
            Lista2.Add(type_BackSta);
            Lista2.Add(type_AheadSta);
            Lista2.Add(type_DeflAng);
            Lista2.Add(type_DeflAngDMS);
            Lista2.Add(type_Bearing);
            Lista2.Add(type_Distance);
            Lista2.Add(type_DisplaySta);
            Lista2.Add(type_DisplayPI);
            Lista2.Add(type_DisplayProf);
            Lista2.Add(type_Symbol);


            System.Data.DataTable Data_table_centerline = new System.Data.DataTable();

            for (int i = 0; i < Lista1.Count; ++i)
            {
                Data_table_centerline.Columns.Add(Lista1[i], Lista2[i]);
            }
            return Data_table_centerline;
        }







        public static System.Data.DataTable Build_Data_table_mat_point_from_excel(Microsoft.Office.Interop.Excel.Worksheet W1, int Start_row)
        {

            System.Data.DataTable dtm = Creaza_dt_mat_point_structure();
            string Col1 = "B";

            Range range2 = W1.Range[Col1 + Start_row.ToString() + ":" + Col1 + "30000"];
            object[,] values2 = new object[30000, 1];
            values2 = range2.Value2;


            bool is_data = false;
            for (int i = 1; i <= values2.Length; ++i)
            {
                object Valoare2 = values2[i, 1];
                if (Valoare2 != null)
                {
                    dtm.Rows.Add();
                    is_data = true;
                }
                else
                {
                    i = values2.Length + 1;
                }
            }


            if (is_data == false)
            {
                MessageBox.Show("no data found in the MATERIAL file");
                return dtm;
            }

            int NrR = dtm.Rows.Count;
            int NrC = dtm.Columns.Count;

            if (is_data == true)
            {
                Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[Start_row, 1], W1.Cells[NrR + Start_row - 1, NrC]];

                object[,] values = new object[NrR - 1, NrC - 1];

                values = range1.Value2;

                for (int i = 0; i < dtm.Rows.Count; ++i)
                {
                    for (int j = 0; j < dtm.Columns.Count; ++j)
                    {
                        object Valoare = values[i + 1, j + 1];
                        if (Valoare == null) Valoare = DBNull.Value;
                        dtm.Rows[i][j] = Valoare;
                    }
                }
            }
            return dtm;
        }

        public static System.Data.DataTable Creaza_dt_mat_lin_structure()
        {
            string dcol0 = "MMID";
            string dcol1 = "ItemNo";
            string dcol2 = "2DStaBeg";
            string dcol3 = "2DStaEnd";
            string dcol4 = "3DStaBeg";
            string dcol5 = "3DStaEnd";
            string dcol6 = "EqStaBeg";
            string dcol7 = "EqStaEnd";

            string dcol8 = "2D len";
            string dcol9 = "3D len";
            string dcol10 = "AltDesc";

            System.Data.DataTable dt1 = new System.Data.DataTable();
            dt1.Columns.Add(dcol0, typeof(string));
            dt1.Columns.Add(dcol1, typeof(string));
            dt1.Columns.Add(dcol2, typeof(double));
            dt1.Columns.Add(dcol3, typeof(double));
            dt1.Columns.Add(dcol4, typeof(double));
            dt1.Columns.Add(dcol5, typeof(double));
            dt1.Columns.Add(dcol6, typeof(double));
            dt1.Columns.Add(dcol7, typeof(double));
            dt1.Columns.Add(dcol8, typeof(double));
            dt1.Columns.Add(dcol9, typeof(double));
            dt1.Columns.Add(dcol10, typeof(string));


            return dt1;
        }

        public static System.Data.DataTable Creaza_dt_mat_point_structure()
        {
            string dcol0 = "MMID";
            string dcol1 = "ItemNo";
            string dcol2 = "2DSta";
            string dcol3 = "3DSta";
            string dcol4 = "EqSta";
            string dcol5 = "Symbol";
            string dcol6 = "AltDesc";


            System.Data.DataTable dt1 = new System.Data.DataTable();
            dt1.Columns.Add(dcol0, typeof(string));
            dt1.Columns.Add(dcol1, typeof(string));
            dt1.Columns.Add(dcol2, typeof(double));
            dt1.Columns.Add(dcol3, typeof(double));
            dt1.Columns.Add(dcol4, typeof(double));
            dt1.Columns.Add(dcol5, typeof(double));
            dt1.Columns.Add(dcol6, typeof(double));



            return dt1;
        }

        public static System.Data.DataTable Creaza_sheet_index_datatable_structure()
        {

            string Col_MMid = "MMID";
            string Col_handle = "AcadHandle";
            string Col_dwg_name = "DwgNo";
            string Col_M1 = "StaBeg";
            string Col_M2 = "StaEnd";
            string Col_dispM1 = "Disp_StaBeg";
            string Col_dispM2 = "Disp_StaEnd";
            string Col_length = "Length";
            string Col_X = "X";
            string Col_Y = "Y";
            string Col_rot = "Rotation";
            string Col_Width = "Width";
            string Col_Height = "Height";
            string Col_X1 = "X_Beg";
            string Col_Y1 = "Y_Beg";
            string Col_X2 = "X_End";
            string Col_Y2 = "Y_End";

            System.Type type_MMid = typeof(string);
            System.Type type_handle = typeof(string);
            System.Type type_dwg_name = typeof(string);
            System.Type type_M1 = typeof(double);
            System.Type type_M2 = typeof(double);
            System.Type type_dispM1 = typeof(double);
            System.Type type_dispM2 = typeof(double);
            System.Type type_length = typeof(double);
            System.Type type_X = typeof(double);
            System.Type type_Y = typeof(double);
            System.Type type_rot = typeof(double);
            System.Type type_Width = typeof(double);
            System.Type type_double = typeof(double);


            List<string> Lista1 = new List<string>();
            List<System.Type> Lista2 = new List<System.Type>();

            Lista1.Add(Col_MMid);
            Lista1.Add(Col_handle);
            Lista1.Add(Col_dwg_name);
            Lista1.Add(Col_M1);
            Lista1.Add(Col_M2);
            Lista1.Add(Col_dispM1);
            Lista1.Add(Col_dispM2);
            Lista1.Add(Col_length);
            Lista1.Add(Col_X);
            Lista1.Add(Col_Y);
            Lista1.Add(Col_rot);
            Lista1.Add(Col_Width);
            Lista1.Add(Col_Height);
            Lista1.Add(Col_X1);
            Lista1.Add(Col_Y1);
            Lista1.Add(Col_X2);
            Lista1.Add(Col_Y2);

            Lista2.Add(type_MMid);
            Lista2.Add(type_handle);
            Lista2.Add(type_dwg_name);
            Lista2.Add(type_M1);
            Lista2.Add(type_M2);
            Lista2.Add(type_dispM1);
            Lista2.Add(type_dispM2);
            Lista2.Add(type_length);
            Lista2.Add(type_X);
            Lista2.Add(type_Y);
            Lista2.Add(type_rot);
            Lista2.Add(type_Width);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);

            System.Data.DataTable dt1 = new System.Data.DataTable();

            for (int i = 0; i < Lista1.Count; ++i)
            {
                dt1.Columns.Add(Lista1[i], Lista2[i]);
            }
            return dt1;
        }

        public static System.Data.DataTable Creaza_display_datatable_structure()
        {

            string Col_dwg_name = "DwgNo";
            string Col_M1 = "StaBeg";
            string Col_M2 = "StaEnd";



            System.Type type_dwg_name = typeof(string);
            System.Type type_M1 = typeof(double);
            System.Type type_M2 = typeof(double);



            List<string> Lista1 = new List<string>();
            List<System.Type> Lista2 = new List<System.Type>();


            Lista1.Add(Col_dwg_name);
            Lista1.Add(Col_M1);
            Lista1.Add(Col_M2);


            Lista2.Add(type_dwg_name);
            Lista2.Add(type_M1);
            Lista2.Add(type_M2);


            System.Data.DataTable dt1 = new System.Data.DataTable();

            for (int i = 0; i < Lista1.Count; ++i)
            {
                dt1.Columns.Add(Lista1[i], Lista2[i]);
            }
            return dt1;
        }

        public static System.Data.DataTable Build_Data_table_sheet_index_from_excel(Microsoft.Office.Interop.Excel.Worksheet W1, int Start_row)
        {


            System.Data.DataTable Data_table_sheet_index = Creaza_sheet_index_datatable_structure();
            string Col1 = "C";

            Range range2 = W1.Range[Col1 + Start_row.ToString() + ":" + Col1 + "30000"];
            object[,] values2 = new object[30000, 1];
            values2 = range2.Value2;


            bool is_data = false;
            for (int i = 1; i <= values2.Length; ++i)
            {
                object Valoare2 = values2[i, 1];
                if (Valoare2 != null)
                {
                    Data_table_sheet_index.Rows.Add();
                    is_data = true;
                }
                else
                {
                    i = values2.Length + 1;
                }
            }


            if (is_data == false)
            {
                MessageBox.Show("no data found in the SHEET INDEX FILE file");
                return Data_table_sheet_index;
            }

            int NrR = Data_table_sheet_index.Rows.Count;
            int NrC = Data_table_sheet_index.Columns.Count;


            if (is_data == true)
            {

                Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[Start_row, 1], W1.Cells[NrR + Start_row - 1, NrC]];

                object[,] values = new object[NrR - 1, NrC - 1];

                values = range1.Value2;

                for (int i = 0; i < Data_table_sheet_index.Rows.Count; ++i)
                {
                    for (int j = 0; j < Data_table_sheet_index.Columns.Count; ++j)
                    {
                        object Valoare = values[i + 1, j + 1];
                        if (Valoare == null) Valoare = DBNull.Value;

                        Data_table_sheet_index.Rows[i][j] = Valoare;
                    }
                }
            }



            return Data_table_sheet_index;


        }



        public static System.Data.DataTable Sort_data_table(System.Data.DataTable Datatable1, string Column1)
        {
            System.Data.DataTable Data_table_temp = new System.Data.DataTable();
            if (Datatable1 != null)
            {
                if (Datatable1.Rows.Count > 0)
                {
                    if (Datatable1.Columns.Contains(Column1) == true)
                    {
                        System.Data.DataView DataView1 = new System.Data.DataView(Datatable1);
                        DataView1.Sort = Column1 + " ASC";
                        Data_table_temp = Datatable1.Clone();
                        Data_table_temp.Rows.Clear();
                        for (int i = 0; i < DataView1.Count; ++i)
                        {
                            System.Data.DataRow Data_row1 = DataView1[i].Row;
                            Data_table_temp.Rows.Add();
                            for (int j = 0; j < Datatable1.Columns.Count; ++j)
                            {
                                Data_table_temp.Rows[Data_table_temp.Rows.Count - 1][j] = Data_row1[j];
                            }
                        }
                    }
                }
            }
            return Data_table_temp;

        }

        public static System.Data.DataTable Elimina_duplicates_from_data_table(System.Data.DataTable Datatable1)
        {


            List<int> lista1 = new List<int>();

            if (Datatable1 != null)
            {
                if (Datatable1.Rows.Count > 0)
                {
                    for (int i = 0; i < Datatable1.Rows.Count; ++i)
                    {
                        System.Data.DataRow row1 = Datatable1.Rows[i];
                        if (i < Datatable1.Rows.Count - 1)
                        {
                            for (int j = i + 1; j < Datatable1.Rows.Count; ++j)
                            {
                                int nr_eq = Datatable1.Columns.Count;
                                System.Data.DataRow row2 = Datatable1.Rows[j];
                                for (int k = 0; k < Datatable1.Columns.Count; ++k)
                                {
                                    string value1 = Convert.ToString(row1[k]);
                                    string value2 = Convert.ToString(row2[k]);

                                    if (value1 == value2)
                                    {
                                        nr_eq = nr_eq - 1;
                                    }
                                }
                                if (nr_eq == 0)
                                {
                                    if (lista1.Contains(j) == false) lista1.Add(j);
                                }
                            }
                        }
                    }

                    if (lista1.Count > 0)
                    {
                        for (int i = lista1.Count - 1; i >= 0; --i)
                        {
                            Datatable1.Rows.RemoveAt(i);
                        }
                    }

                }
            }

            return Datatable1;
        }

        public static System.Data.DataTable Sort_data_table_Nu_e_gata(System.Data.DataTable Datatable1, string Col1, string Col2)
        {
            System.Data.DataTable Data_table_sorted = new System.Data.DataTable();
            if (Datatable1 != null)
            {
                if (Datatable1.Rows.Count > 0)
                {
                    if (Datatable1.Columns.Contains(Col1) == true && Datatable1.Columns.Contains(Col2) == true)
                    {


                        Data_table_sorted = Datatable1.Clone();
                        Data_table_sorted.Rows.Clear();




                        int i = 0;
                        do
                        {
                            string Val1 = "";
                            string Val2 = "";

                            if (Datatable1.Rows[i][Col1] != DBNull.Value)
                            {
                                Val1 = Datatable1.Rows[i][Col1].ToString();
                            }

                            if (Datatable1.Rows[i][Col2] != DBNull.Value)
                            {
                                Val2 = Datatable1.Rows[i][Col2].ToString();
                            }


                            Data_table_sorted.Rows.Add();

                            for (int k = 0; k < Datatable1.Columns.Count; ++k)
                            {
                                Data_table_sorted.Rows[Data_table_sorted.Rows.Count - 1][k] = Datatable1.Rows[i][k];

                            }
                            Datatable1.Rows[i].Delete();

                            for (int j = i; j < Datatable1.Rows.Count; ++j)
                            {
                                if (Datatable1.Rows[j][Col2] != DBNull.Value)
                                {

                                    string v2 = Datatable1.Rows[j][Col2].ToString();

                                    if (Val2 == v2)
                                    {
                                        Data_table_sorted.Rows.Add();

                                        for (int k = 0; k < Datatable1.Columns.Count; ++k)
                                        {
                                            Data_table_sorted.Rows[Data_table_sorted.Rows.Count - 1][k] = Datatable1.Rows[j][k];
                                        }

                                    }

                                    Datatable1.Rows[j].Delete();

                                }
                            }

                        }
                        while (i < Datatable1.Rows.Count);

                    }
                }
            }
            return Data_table_sorted;

        }

        static public void Incarca_existing_Blocks_to_combobox(System.Windows.Forms.ComboBox Combo_blockname)
        {

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                    Combo_blockname.Items.Clear();
                    Combo_blockname.Items.Add("");
                    foreach (ObjectId Block_id in BlockTable_data1)
                    {
                        BlockTableRecord Block1 = (BlockTableRecord)Trans1.GetObject(Block_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);


                        if (Block1.Name.Contains("*") == false && Block1.Name.Contains("|") == false &&
                            Block1.Name.Contains("$") == false && Block1.IsFromExternalReference == false &&
                            Block1.IsFromOverlayReference == false &&
                            Block1.IsLayout == false)
                        {

                            Combo_blockname.Items.Add(Block1.Name);
                        }
                    }
                    Trans1.Dispose();
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        static public void Incarca_existing_Blocks_with_attributes_to_combobox(System.Windows.Forms.ComboBox Combo_blockname)
        {

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                Combo_blockname.Items.Clear();
                Combo_blockname.Items.Add("");
                Combo_blockname.Text = "";
                foreach (ObjectId Block_id in BlockTable_data1)
                {
                    BlockTableRecord Block1 = (BlockTableRecord)Trans1.GetObject(Block_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

                    if (Block1.HasAttributeDefinitions == true)
                    {
                        if (Block1.Name.Contains("*") == false && Block1.Name.Contains("|") == false && Block1.Name.Contains("$") == false)
                        {
                            Combo_blockname.Items.Add(Block1.Name);
                        }
                    }
                }
                Trans1.Dispose();
            }
        }

        static public void Incarca_existing_Atributes_to_combobox(string BlockName, ComboBox Combo_atributes)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                Autodesk.AutoCAD.DatabaseServices.BlockTable Block_table = Trans1.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.BlockTable;
                Combo_atributes.Items.Clear();
                Combo_atributes.Items.Add("");
                if (BlockName != "" && Block_table != null)
                {
                    if (Block_table.Has(BlockName) == true)
                    {
                        BlockTableRecord BTrecordBlock = Trans1.GetObject(Block_table[BlockName], OpenMode.ForRead) as BlockTableRecord;
                        if (BTrecordBlock != null)
                        {
                            foreach (ObjectId Id1 in BTrecordBlock)
                            {
                                Entity ent = Trans1.GetObject(Id1, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Entity;
                                if (ent != null)
                                {
                                    AttributeDefinition attDefinition1 = ent as AttributeDefinition;
                                    if (attDefinition1 != null)
                                    {
                                        Combo_atributes.Items.Add(attDefinition1.Tag);
                                    }
                                }
                            }
                        }
                    }
                }
                Trans1.Dispose();
            }
        }

        static public List<string> Incarca_existing_Atributes_to_list(string BlockName)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                Autodesk.AutoCAD.DatabaseServices.BlockTable Block_table = Trans1.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.BlockTable;

                List<string> Lista1 = new List<string>();

                if (BlockName != "" && Block_table != null)
                {
                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecordBlock = Trans1.GetObject(Block_table[BlockName], Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.BlockTableRecord;
                    if (BTrecordBlock != null)
                    {
                        foreach (ObjectId Id1 in BTrecordBlock)
                        {
                            Entity ent = Trans1.GetObject(Id1, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Entity;
                            if (ent != null)
                            {
                                AttributeDefinition attDefinition1 = ent as AttributeDefinition;
                                if (attDefinition1 != null)
                                {
                                    Lista1.Add(attDefinition1.Tag);
                                }
                            }
                        }
                    }
                }
                Trans1.Dispose();
                return Lista1;
            }
        }

        static public void Incarca_existing_layers_to_combobox(System.Windows.Forms.ComboBox Combo_layer)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                Autodesk.AutoCAD.DatabaseServices.LayerTable layer_table = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                Combo_layer.Items.Clear();

                System.Data.DataTable dt1 = new System.Data.DataTable();
                dt1.Columns.Add("ln", typeof(string));


                foreach (ObjectId Layer_id in layer_table)
                {
                    LayerTableRecord Layer1 = (LayerTableRecord)Trans1.GetObject(Layer_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                    string Name_of_layer = Layer1.Name;
                    if (Name_of_layer.Contains("|") == false & Name_of_layer.Contains("$") == false)
                    {
                        dt1.Rows.Add();
                        dt1.Rows[dt1.Rows.Count - 1][0] = Name_of_layer;


                    }
                }

                System.Data.DataTable dt2 = Sort_data_table(dt1, "ln");
                for (int i = 0; i < dt2.Rows.Count; ++i)
                {
                    Combo_layer.Items.Add(dt2.Rows[i][0].ToString());
                }
                Combo_layer.SelectedIndex = 0;
                Trans1.Dispose();
            }
        }

        static public void Incarca_existing_textstyles_to_combobox(System.Windows.Forms.ComboBox Combo_text_style)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                Autodesk.AutoCAD.DatabaseServices.TextStyleTable Text_style_table = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as TextStyleTable;
                Combo_text_style.Items.Clear();
                foreach (ObjectId TextStyle_id in Text_style_table)
                {
                    TextStyleTableRecord TextStyle1 = Trans1.GetObject(TextStyle_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as TextStyleTableRecord;
                    Combo_text_style.Items.Add(TextStyle1.Name);
                }
                Combo_text_style.SelectedIndex = 0;
                Trans1.Dispose();
            }
        }

        static public ObjectId Get_textstyle_id_from_combobox(System.Windows.Forms.ComboBox Combo_text_style)
        {

            ObjectId Id = ObjectId.Null;
            if (Combo_text_style.Text != "")
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.TextStyleTable Text_style_table = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as TextStyleTable;
                    Id = Text_style_table["Standard"];

                    foreach (ObjectId TextStyle_id in Text_style_table)
                    {
                        TextStyleTableRecord TextStyle1 = Trans1.GetObject(TextStyle_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as TextStyleTableRecord;
                        if (TextStyle1.Name == Combo_text_style.Text)
                        {
                            Id = TextStyle_id;
                        }

                    }
                    Trans1.Dispose();
                }
            }
            return Id;
        }

        static public ObjectId Get_textstyle_id(string text_style)
        {

            ObjectId Id = ObjectId.Null;

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                Autodesk.AutoCAD.DatabaseServices.TextStyleTable Text_style_table = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as TextStyleTable;
                Id = Text_style_table["Standard"];

                foreach (ObjectId TextStyle_id in Text_style_table)
                {
                    TextStyleTableRecord TextStyle1 = Trans1.GetObject(TextStyle_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as TextStyleTableRecord;
                    if (TextStyle1.Name == text_style)
                    {
                        Id = TextStyle_id;
                    }

                }
                Trans1.Dispose();
            }

            return Id;
        }

        static public double Get_text_height_from_textstyle(string text_style)
        {

            double Texth = 0;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {

                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as BlockTableRecord;

                    ObjectId Tid = Get_textstyle_id(text_style);
                    if (Tid != null)
                    {
                        TextStyleTableRecord TextStyle1 = Trans1.GetObject(Tid, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as TextStyleTableRecord;
                        if (TextStyle1 != null)
                        {
                            return TextStyle1.TextSize;

                        }
                    }
                }
            }
            return Texth;

        }

        static public double Get_text_width_factor_from_textstyle(string text_style)
        {

            double width1 = 1;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {

                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as BlockTableRecord;

                    ObjectId Tid = Get_textstyle_id(text_style);
                    if (Tid != null)
                    {
                        TextStyleTableRecord TextStyle1 = Trans1.GetObject(Tid, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as TextStyleTableRecord;
                        if (TextStyle1 != null)
                        {
                            return TextStyle1.XScale;

                        }
                    }
                }
            }
            return width1;

        }
        static public String verify_layername_from_combobox(System.Windows.Forms.ComboBox Combo_layername)
        {

            string Layer_name = "0";
            if (Combo_layername.Text != "")
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.LayerTable Layer_table = Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as LayerTable;

                    if (Layer_table.Has(Combo_layername.Text) == true)
                    {
                        Layer_name = Combo_layername.Text;
                    }
                    Trans1.Dispose();
                }
            }
            return Layer_name;
        }

        static public string verify_layername(string layer1)
        {

            string Layer_name = "0";

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                Autodesk.AutoCAD.DatabaseServices.LayerTable Layer_table = Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as LayerTable;

                if (Layer_table.Has(layer1) == true)
                {
                    Layer_name = layer1;
                }
                Trans1.Dispose();

            }
            return Layer_name;
        }

        static public String get_block_name(BlockReference Block1)
        {
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                    BlockTableRecord Btr = null;
                    if (Block1.IsDynamicBlock == true)
                    {

                        Btr = (BlockTableRecord)Trans1.GetObject(Block1.DynamicBlockTableRecord, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        return Btr.Name;
                    }
                    else
                    {
                        Btr = (BlockTableRecord)Trans1.GetObject(Block1.BlockTableRecord, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        return Btr.Name;
                    }
                }
            }
            catch (System.Exception ex)
            {
                return "";
            }
        }

        static public String get_block_name_another_database(BlockReference Block1, Database database2)
        {
            try
            {

                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = database2.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)database2.BlockTableId.GetObject(OpenMode.ForRead);
                    BlockTableRecord Btr = null;
                    if (Block1.IsDynamicBlock == true)
                    {

                        Btr = (BlockTableRecord)Trans1.GetObject(Block1.DynamicBlockTableRecord, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        return Btr.Name;
                    }
                    else
                    {
                        Btr = (BlockTableRecord)Trans1.GetObject(Block1.BlockTableRecord, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        return Btr.Name;
                    }
                }
            }
            catch (System.Exception ex)
            {
                return "";
            }
        }

        public static string Extract_text_from_name(string OldName)
        {
            ;
            string String1 = Extract_number_from_name(OldName);
            return OldName.Substring(0, OldName.Length - String1.Length);



        }

        public static string Extract_number_from_name(string OldName)
        {

            string String1 = "";
            for (int i = OldName.Length - 1; i > 0; --i)
            {
                if (IsNumeric(OldName.Substring(i, 1)) == true)
                {
                    String1 = OldName.Substring(i, 1) + String1;

                }
                else
                {
                    return String1;
                }
            }


            return String1;

        }

        public static Autodesk.AutoCAD.DatabaseServices.BlockTableRecord get_modelspace(Autodesk.AutoCAD.DatabaseServices.Transaction Trans1, Database Database1)
        {
            HostApplicationServices.WorkingDatabase = Database1;
            Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (Autodesk.AutoCAD.DatabaseServices.BlockTable)Database1.BlockTableId.GetObject(OpenMode.ForRead);
            return (Autodesk.AutoCAD.DatabaseServices.BlockTableRecord)Trans1.GetObject(BlockTable_data1[BlockTableRecord.ModelSpace], OpenMode.ForRead);
        }

        public static Autodesk.AutoCAD.DatabaseServices.BlockTableRecord get_first_layout_as_paperspace(Autodesk.AutoCAD.DatabaseServices.Transaction Trans1, Database Database1)
        {

            HostApplicationServices.WorkingDatabase = Database1;

            DBDictionary Layoutdict = (DBDictionary)Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead);

            LayoutManager LayoutManager1 = (LayoutManager)Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;

            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecordPS = null;
            foreach (DBDictionaryEntry entry in Layoutdict)
            {
                Layout Layout0 = (Layout)Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForRead);
                if (Layout0.TabOrder == 1)
                {
                    return (BlockTableRecord)Trans1.GetObject(Layout0.BlockTableRecordId, OpenMode.ForRead);
                }

            }
            return BTrecordPS;


        }

        public static Layout get_first_layout(Autodesk.AutoCAD.DatabaseServices.Transaction Trans1, Database Database1)
        {
            HostApplicationServices.WorkingDatabase = Database1;

            DBDictionary Layoutdict = (DBDictionary)Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead);

            LayoutManager LayoutManager1 = (LayoutManager)Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;

            Layout Layout0 = null;
            foreach (DBDictionaryEntry entry in Layoutdict)
            {
                Layout0 = (Layout)Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForRead);
                if (Layout0.TabOrder == 1)
                {
                    return Layout0;
                }

            }
            return Layout0;
        }

        public static void make_first_layout_active(Autodesk.AutoCAD.DatabaseServices.Transaction Trans1, Database Database1)
        {
            HostApplicationServices.WorkingDatabase = Database1;

            DBDictionary Layoutdict = (DBDictionary)Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead);

            LayoutManager LayoutManager1 = (LayoutManager)Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;

            Layout Layout0 = null;
            foreach (DBDictionaryEntry entry in Layoutdict)
            {
                Layout0 = (Layout)Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForRead);
                if (Layout0.TabOrder == 1)
                {
                    LayoutManager1.CurrentLayout = Layout0.LayoutName;
                    return;
                }

            }

        }

        public static Viewport Create_viewport(Point3d MSpoint, Point3d PSpoint, double Width, double Height, double Scale, double Twist_rad)
        {
            Viewport Viewport = new Viewport();

            Viewport.SetDatabaseDefaults();
            Viewport.CenterPoint = PSpoint;
            Viewport.Height = Height;
            Viewport.Width = Width;
            Viewport.ViewDirection = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis;
            Viewport.ViewTarget = MSpoint;
            Viewport.ViewCenter = Autodesk.AutoCAD.Geometry.Point2d.Origin;
            Viewport.TwistAngle = Twist_rad;
            Viewport.CustomScale = Scale;
            Viewport.Locked = true;

            return Viewport;
        }

        static public BlockReference InsertBlock_with_multiple_atributes_with_database(Database Database1, Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord,
            string Nume_fisier, string NumeBlock, Point3d Insertion_point, double Scale_xyz, double Rotation1, string Layer1,
             System.Collections.Specialized.StringCollection Colectie_nume_atribute, System.Collections.Specialized.StringCollection Colectie_valori_atribute)
        {

            BlockReference Block1 = null;


            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = Database1.TransactionManager.StartTransaction())
            {

                BlockTable BlockTable1 = (BlockTable)Trans1.GetObject(Database1.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                if (BlockTable1.Has(NumeBlock) == false)
                {
                    if (System.IO.File.Exists(Nume_fisier) == true)
                    {
                        using (Database Database2 = new Database(false, false))
                        {
                            Database2.ReadDwgFile(Nume_fisier, System.IO.FileShare.Read, true, null);
                            Database1.Insert(NumeBlock, Database2, false);
                        }
                    }


                }

                Trans1.Commit();
            }

            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = Database1.TransactionManager.StartTransaction())
            {

                BlockTable BlockTable1 = (BlockTable)Trans1.GetObject(Database1.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

                if (BlockTable1.Has(NumeBlock) == true)
                {


                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTR = (BlockTableRecord)Trans1.GetObject(BlockTable1[NumeBlock], OpenMode.ForRead);

                    Block1 = new BlockReference(Insertion_point, BTR.ObjectId);
                    Block1.Layer = Layer1;
                    Block1.ScaleFactors = new Autodesk.AutoCAD.Geometry.Scale3d(Scale_xyz, Scale_xyz, Scale_xyz);
                    Block1.Rotation = Rotation1;
                    BTrecord.AppendEntity(Block1);
                    Trans1.AddNewlyCreatedDBObject(Block1, true);
                    Autodesk.AutoCAD.DatabaseServices.AttributeCollection attColl = Block1.AttributeCollection;
                    BlockTableRecordEnumerator BTR_enum = BTR.GetEnumerator();
                    while (BTR_enum.MoveNext())
                    {
                        Entity Ent1 = (Entity)Trans1.GetObject(BTR_enum.Current, OpenMode.ForWrite);
                        if (Ent1 is AttributeDefinition)
                        {
                            AttributeDefinition Attdef = (AttributeDefinition)Ent1;
                            AttributeReference Attref = new AttributeReference();
                            Attref.SetAttributeFromBlock(Attdef, Block1.BlockTransform);

                            for (int i = 0; i < Colectie_nume_atribute.Count; ++i)
                            {
                                string Tag1 = Colectie_nume_atribute[i];
                                string Valoare = Colectie_valori_atribute[i];
                                if (Attref.Tag.ToLower() == Tag1.ToLower())
                                {
                                    Attref.TextString = Valoare;
                                    i = Colectie_nume_atribute.Count;
                                }
                            }
                            if (Attref != null)
                            {
                                attColl.AppendAttribute(Attref);
                                Trans1.AddNewlyCreatedDBObject(Attref, true);
                            }
                        }

                    }

                }

                Trans1.Commit();
            }

            return Block1;
        }

        static public BlockReference InsertBlock_with_2scales(Database Database1, Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord, string NumeBlock, Point3d Insertion_point, double Scale_x, double Scale_y, double Rotation1, string Layer1)
        {

            BlockReference Block1 = null;

            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = Database1.TransactionManager.StartTransaction())
            {

                BlockTable BlockTable1 = (BlockTable)Trans1.GetObject(Database1.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

                if (BlockTable1.Has(NumeBlock) == true)
                {


                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTR = (BlockTableRecord)Trans1.GetObject(BlockTable1[NumeBlock], OpenMode.ForRead);

                    Block1 = new BlockReference(Insertion_point, BTR.ObjectId);
                    Block1.Layer = Layer1;
                    Block1.ScaleFactors = new Autodesk.AutoCAD.Geometry.Scale3d(Scale_x, Scale_y, 1);
                    Block1.Rotation = Rotation1;
                    BTrecord.AppendEntity(Block1);
                    Trans1.AddNewlyCreatedDBObject(Block1, true);

                }

                Trans1.Commit();
            }

            return Block1;
        }

        static public MLeader Create_mleader_on_profile_with_database(Database Database1, Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord, Point3d Insertion_point, string Layer1, string descriptie, double Text_height, ObjectId textid)
        {

            MLeader mleader1 = null;

            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = Database1.TransactionManager.StartTransaction())
            {
                MText mtext1 = new MText();
                mtext1.Contents = descriptie;

                mtext1.Attachment = AttachmentPoint.MiddleLeft;
                mtext1.BackgroundFill = true;
                mtext1.UseBackgroundColor = true;
                mtext1.BackgroundScaleFactor = 1.2;
                if (textid != ObjectId.Null) mtext1.TextStyleId = textid;

                var leaderline_index = new int();
                var leaderpoint = new int();

                mleader1 = new MLeader();
                mleader1.Layer = Layer1;
                mleader1.SetDatabaseDefaults();
                leaderline_index = mleader1.AddLeader();
                leaderpoint = mleader1.AddLeaderLine(leaderline_index);
                mleader1.AddFirstVertex(leaderpoint, Insertion_point);
                mleader1.AddLastVertex(leaderpoint, new Point3d(Insertion_point.X + Text_height, Insertion_point.Y, 0));
                // calling AddLastVertex again will add more vertices to the leader line. The last vertex you add is where the text will appear.
                mleader1.AddLastVertex(leaderline_index, new Point3d(Insertion_point.X + 1.5 * Text_height, Insertion_point.Y, 0));

                mleader1.LeaderLineType = LeaderType.StraightLeader;
                mleader1.ContentType = ContentType.MTextContent;
                mleader1.MText = mtext1;
                mleader1.TextHeight = Text_height;
                mleader1.LandingGap = Text_height;
                mleader1.ArrowSize = 0;
                mleader1.DoglegLength = Text_height / 2;
                BTrecord.AppendEntity(mleader1);
                Trans1.AddNewlyCreatedDBObject(mleader1, true);

                // rotate leader to make it vertical

                mleader1.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, Insertion_point));

                // MoveMLeader will reset the MLeader and fix a bug with the text showing up in a wrong location

                mleader1.MoveMLeader(new Vector3d(), MoveType.MoveAllPoints);

                //you must call set text attachment for right and left leader after the entitiy is appended before you commit.

                mleader1.SetTextAttachmentType(TextAttachmentType.AttachmentBottomLine, LeaderDirectionType.RightLeader);
                mleader1.SetTextAttachmentType(TextAttachmentType.AttachmentBottomLine, LeaderDirectionType.LeftLeader);

                Trans1.Commit();
            }

            return mleader1;
        }

        static public int Round_Up(double numToRound, int multiple)
        {
            if (multiple == 0)
            {
                return Convert.ToInt32(numToRound);
            }

            int remainder = Convert.ToInt32(numToRound) % multiple;
            if (remainder == 0)
            {
                return Convert.ToInt32(numToRound);
            }
            return Convert.ToInt32(numToRound) + multiple - remainder;
        }

        static public int Round_Down(double numToRound, int multiple)
        {
            if (multiple == 0)
            {
                return Convert.ToInt32(numToRound);
            }

            int remainder = Convert.ToInt32(numToRound) % multiple;
            if (remainder == 0)
            {
                return Convert.ToInt32(numToRound);
            }

            return Convert.ToInt32(numToRound) - remainder;
        }

        static public int Round_Closest(double numToRound, int multiple)
        {
            Int32 Numar = Convert.ToInt32(numToRound);
            Int32 Up = Round_Up(numToRound, multiple);
            Int32 Down = Round_Down(numToRound, multiple);
            if (Math.Abs(Numar - Up) < Math.Abs(Numar - Down))
            {
                return Up;
            }
            else
            {
                return Down;
            }
        }

        public static Polyline Build_2dpoly_from_3d(Polyline3d Poly3D)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Polyline Poly2D = new Polyline();
                    int Index1 = 0;
                    if (Poly3D.Length > 0)
                    {

                        double last_param = Poly3D.EndParam;

                        for (int i = 0; i <= last_param; ++i)
                        {
                            try
                            {
                                Poly2D.AddVertexAt(Index1, new Point2d(Poly3D.GetPointAtParameter(i).X, Poly3D.GetPointAtParameter(i).Y), 0, 0, 0);
                                Index1 = Index1 + 1;

                            }
                            catch (System.Exception ex)
                            {

                            }
                        }
                    }
                    return Poly2D;
                }
            }
        }

        public static void Transfer_datatable_to_new_excel_spreadsheet(System.Data.DataTable dt1)
        {
            if (dt1 != null)
            {
                if (dt1.Rows.Count > 0)
                {
                    Microsoft.Office.Interop.Excel.Worksheet W1 = Get_NEW_worksheet_from_Excel();
                    W1.Cells.NumberFormat = "@";
                    int maxRows = dt1.Rows.Count;
                    int maxCols = dt1.Columns.Count;
                    Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[2, 1], W1.Cells[maxRows + 1, maxCols]];
                    object[,] values1 = new object[maxRows, maxCols];

                    for (int i = 0; i < maxRows; ++i)
                    {
                        for (int j = 0; j < maxCols; ++j)
                        {
                            if (dt1.Rows[i][j] != DBNull.Value)
                            {
                                values1[i, j] = Convert.ToString(dt1.Rows[i][j]);
                            }
                        }
                    }

                    for (int i = 0; i < dt1.Columns.Count; ++i)
                    {
                        W1.Cells[1, i + 1].value2 = dt1.Columns[i].ColumnName;
                    }
                    range1.Value2 = values1;
                }
            }
        }

        public static void Transfer_datatable_to_new_excel_spreadsheet_formated_general(System.Data.DataTable dt1)
        {
            if (dt1 != null)
            {
                if (dt1.Rows.Count > 0)
                {
                    Microsoft.Office.Interop.Excel.Worksheet W1 = Get_NEW_worksheet_from_Excel();
                    W1.Cells.NumberFormat = "General";
                    int maxRows = dt1.Rows.Count;
                    int maxCols = dt1.Columns.Count;
                    Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[2, 1], W1.Cells[maxRows + 1, maxCols]];
                    object[,] values1 = new object[maxRows, maxCols];

                    for (int i = 0; i < maxRows; ++i)
                    {
                        for (int j = 0; j < maxCols; ++j)
                        {
                            if (dt1.Rows[i][j] != DBNull.Value)
                            {
                                values1[i, j] = Convert.ToString(dt1.Rows[i][j]);
                            }
                        }
                    }

                    for (int i = 0; i < dt1.Columns.Count; ++i)
                    {
                        W1.Cells[1, i + 1].value2 = dt1.Columns[i].ColumnName;
                    }
                    range1.Value2 = values1;
                }
            }
        }

        public static void Transfer_datatable_to_existing_excel_spreadsheet(System.Data.DataTable dt1, int start1)
        {
            if (dt1 != null)
            {
                if (dt1.Rows.Count > 0)
                {
                    Microsoft.Office.Interop.Excel.Worksheet W1 = Get_active_worksheet_from_Excel();
                    W1.Cells.NumberFormat = "@";
                    int maxRows = dt1.Rows.Count;
                    int maxCols = dt1.Columns.Count;
                    Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[1 + start1, 1], W1.Cells[maxRows + start1, maxCols]];
                    object[,] values1 = new object[maxRows, maxCols];

                    for (int i = 0; i < maxRows; ++i)
                    {
                        for (int j = 0; j < maxCols; ++j)
                        {
                            if (dt1.Rows[i][j] != DBNull.Value)
                            {
                                values1[i, j] = Convert.ToString(dt1.Rows[i][j]);
                            }
                        }
                    }

                    for (int i = 0; i < dt1.Columns.Count; ++i)
                    {
                        W1.Cells[start1, i + 1].value2 = dt1.Columns[i].ColumnName;
                    }
                    range1.Value2 = values1;
                }
            }
        }

        static public double Round_Up_as_double(double numToRound, double multiple)
        {
            if (multiple == 0)
            {
                return numToRound;
            }

            return Math.Ceiling(numToRound / multiple) * multiple;

        }

        static public double Round_Down_as_double(double numToRound, double multiple)
        {
            if (multiple == 0)
            {
                return numToRound;
            }


            return Math.Floor(numToRound / multiple) * multiple;

        }

        static public string Get_chainage_from_double(double Numar, string units, int Nr_dec)
        {

            String String2, String3;
            String3 = "";
            String String_minus = "";

            if (Numar < 0)
            {
                String_minus = "-";
                Numar = -Numar;
            }




            String2 = Get_String_Rounded(Numar, Nr_dec);




            int Punct;
            if (String2.Contains(".") == false)
            {
                Punct = 0;
            }
            else
            {
                Punct = 1;
            }


            if (String2.Length - Nr_dec - Punct >= 4)
            {
                if (units == "f") String3 = String2.Substring(0, String2.Length - 2 - Nr_dec - Punct) + "+" + String2.Substring(String2.Length - (2 + Nr_dec + Punct));
                if (units == "m") String3 = String2.Substring(0, String2.Length - 3 - Nr_dec - Punct) + "+" + String2.Substring(String2.Length - (3 + Nr_dec + Punct));
            }
            else
            {
                if (units == "f")
                {
                    if (String2.Length - Nr_dec - Punct == 1) String3 = "0+0" + String2;
                    if (String2.Length - Nr_dec - Punct == 2) String3 = "0+" + String2;
                    if (String2.Length - Nr_dec - Punct == 3) String3 = String2.Substring(0, 1) + "+" + String2.Substring(String2.Length - (2 + Nr_dec + Punct));
                }
                if (units == "m")
                {
                    if (String2.Length - Nr_dec - Punct == 1) String3 = "0+00" + String2;
                    if (String2.Length - Nr_dec - Punct == 2) String3 = "0+0" + String2;
                    if (String2.Length - Nr_dec - Punct == 3) String3 = "0+" + String2;
                }
            }


            return String_minus + String3;

        }

        static public string Get_String_Rounded(double Numar, int Nr_dec)
        {

            String String1, String2, Zero, zero1;
            Zero = "";
            zero1 = "";

            String String_punct = "";

            if (Nr_dec > 0)
            {
                String_punct = ".";
                for (int i = 1; i <= Nr_dec; i = i + 1)
                {
                    Zero = Zero + "0";
                }
            }

            String String_minus = "";

            if (Numar < 0)
            {
                String_minus = "-";
                Numar = -Numar;
            }

            String1 = Math.Round(Numar, Nr_dec, MidpointRounding.AwayFromZero).ToString();

            String2 = String1;

            if (String1.Contains(".") == false)
            {
                String2 = String1 + String_punct + Zero;
                goto end;
            }

            if (String1.Length - String1.IndexOf(".") - 1 - Nr_dec != 0)
            {
                for (int i = 1; i <= String1.IndexOf(".") + 1 + Nr_dec - String1.Length; i = i + 1)
                {
                    zero1 = zero1 + "0";
                }

                String2 = String1 + zero1;
            }

        end:
            return String_minus + String2;

        }
















        static public void add_OD_fieds_to_combobox(ComboBox Combobox_table_name, ComboBox Combobox1)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Combobox1.Items.Clear();
            Combobox1.Items.Add("");
            using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    if (Tables1.IsTableDefined(Combobox_table_name.Text) == true)
                    {
                        Autodesk.Gis.Map.ObjectData.Table tabla1 = Tables1[Combobox_table_name.Text];
                        Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = tabla1.FieldDefinitions;
                        for (int i = 0; i < Field_defs1.Count; ++i)
                        {
                            Autodesk.Gis.Map.ObjectData.FieldDefinition fielddef1 = Field_defs1[i];
                            Combobox1.Items.Add(fielddef1.Name);
                        }
                    }
                    else
                    {
                        Combobox1.Items.Clear();
                    }
                    Trans1.Commit();
                }
            }
        }

        static public void append_OD_fieds_to_combobox(ComboBox Combobox_table_name, ComboBox Combobox1)
        {

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            if (Combobox1.Items.Count == 0)
            {
                Combobox1.Items.Add("");
            }

            if (Combobox1.Items[0].ToString() != "")
            {
                Combobox1.Items.Insert(0, "");
            }

            using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    if (Tables1.IsTableDefined(Combobox_table_name.Text) == true)
                    {
                        Autodesk.Gis.Map.ObjectData.Table tabla1 = Tables1[Combobox_table_name.Text];
                        Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = tabla1.FieldDefinitions;
                        for (int i = 0; i < Field_defs1.Count; ++i)
                        {
                            Autodesk.Gis.Map.ObjectData.FieldDefinition fielddef1 = Field_defs1[i];
                            if (Combobox1.Items.Contains(fielddef1.Name) == false)
                            {
                                Combobox1.Items.Add(fielddef1.Name);
                            }


                        }
                    }

                    Trans1.Commit();
                }
            }
        }

        static public double Get_distance1_block(BlockReference BR)
        {
            using (DynamicBlockReferencePropertyCollection pc = BR.DynamicBlockReferencePropertyCollection)
            {
                foreach (DynamicBlockReferenceProperty prop in pc)
                {
                    if (prop.PropertyName == "Distance1" && prop.UnitsType == DynamicBlockReferencePropertyUnitsType.Distance)
                    {
                        return Convert.ToDouble(prop.Value);

                    }
                }
            }
            return 0;
        }

        static public void Stretch_block(BlockReference BR, String Prop_name, double Prop_value)
        {
            using (DynamicBlockReferencePropertyCollection pc = BR.DynamicBlockReferencePropertyCollection)
            {
                foreach (DynamicBlockReferenceProperty prop in pc)
                {
                    if (prop.PropertyName == Prop_name && prop.UnitsType == DynamicBlockReferencePropertyUnitsType.Distance)
                    {
                        prop.Value = Prop_value;
                        return;
                    }
                }
            }
        }

        static public void set_block_visibility(BlockReference BR, String visibility_name)
        {
            using (DynamicBlockReferencePropertyCollection pc = BR.DynamicBlockReferencePropertyCollection)
            {
                foreach (DynamicBlockReferenceProperty prop in pc)
                {


                    if (prop.PropertyName == "Visibility1" && !prop.ReadOnly)
                    {
                        object[] existing = prop.GetAllowedValues();
                        bool found = false;

                        foreach (object ob in existing)
                        {
                            if (ob.ToString() == visibility_name)
                            {
                                found = true;
                            }
                        }

                        if (found == true)
                        {
                            if (prop.Value.ToString() != visibility_name)
                            {
                                prop.Value = visibility_name;
                            }
                        }
                    }



                }
                return;
            }
        }

        static public string get_block_visibility_value(BlockReference BR, string visibility_name)
        {
            using (DynamicBlockReferencePropertyCollection pc = BR.DynamicBlockReferencePropertyCollection)
            {
                foreach (DynamicBlockReferenceProperty prop in pc)
                {


                    if (prop.PropertyName == visibility_name)
                    {
                        return prop.Value.ToString();

                    }
                }
                return "";
            }
        }












        static public void textbox_input_only_pozitive_and_negative_doubles_at_keypress(object sender, KeyPressEventArgs e)
        {

            if (char.IsControl(e.KeyChar) == false && char.IsDigit(e.KeyChar) == false && (e.KeyChar != '.') && (e.KeyChar != '-'))
            {
                e.Handled = true;
            }



        }


        static public void textbox_input_only_pozitive_doubles_at_keypress(object sender, KeyPressEventArgs e)
        {

            if (char.IsControl(e.KeyChar) == false && char.IsDigit(e.KeyChar) == false && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            if ((e.KeyChar == '-') && ((sender as System.Windows.Forms.TextBox).Text.Contains(".") == true))
            {
                e.Handled = true;
            }
        }
        static public void textbox_input_doubles_at_keypress(object sender, KeyPressEventArgs e)
        {

            string Ex_txt = (sender as System.Windows.Forms.TextBox).Text;

            if (e.KeyChar == '-')
            {
                if (Ex_txt.Contains("-") == true)
                {

                    e.Handled = true;
                    return;
                }

                else
                {
                    return;
                }
            }

            else if (e.KeyChar == '.')
            {
                if (Ex_txt.Contains(".") == true)
                {
                    e.Handled = true;
                    return;
                }
                else
                {
                    return;
                }
            }
            else if (char.IsControl(e.KeyChar) == false && char.IsDigit(e.KeyChar) == false)
            {
                e.Handled = true;
            }


        }
        static public void textbox_input_only_integer_pozitive_at_keypress(object sender, KeyPressEventArgs e)
        {

            if ((char.IsControl(e.KeyChar) == false && char.IsDigit(e.KeyChar) == false) || (e.KeyChar == '.') || e.KeyChar == '-')
            {
                e.Handled = true;
            }

        }








        public static MLeader creaza_mleader(Point3d pt_ins, string continut, double texth, double delta_x, double delta_y, double lgap, double dogl, double arrow)
        {



            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            MLeader mleader1 = new MLeader();


            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {

                BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                MText mtext1 = new MText();

                mtext1.Contents = continut;
                mtext1.TextHeight = texth;
                mtext1.BackgroundFill = true;
                mtext1.UseBackgroundColor = true;
                mtext1.BackgroundScaleFactor = 1.2;
                mtext1.ColorIndex = 0;

                mleader1.SetDatabaseDefaults();
                int index1 = mleader1.AddLeader();
                int index2 = mleader1.AddLeaderLine(index1);
                mleader1.AddFirstVertex(index2, pt_ins);
                mleader1.AddLastVertex(index2, new Point3d(pt_ins.X + delta_x, pt_ins.Y + delta_y, pt_ins.Z));
                mleader1.LeaderLineType = LeaderType.StraightLeader;
                mleader1.ContentType = ContentType.MTextContent;
                mleader1.MText = mtext1;
                mleader1.TextHeight = texth;
                mleader1.LandingGap = lgap;
                mleader1.ArrowSize = arrow;
                mleader1.DoglegLength = dogl;
                mleader1.Annotative = AnnotativeStates.False;
                mleader1.ColorIndex = 256;

                BTrecord.AppendEntity(mleader1);
                Trans1.AddNewlyCreatedDBObject(mleader1, true);
                Trans1.Commit();
            }




            return mleader1;







        }

        public static MLeader creaza_mleader_with_style(Point3d pt_ins, string continut, double Texth, string ml_style_name, string text_style_name, double DeltaX, double DeltaY, string layer_name)
        {



            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            MLeader mleader1 = new MLeader();


            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {

                BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                DBDictionary Leader_style_table = Trans1.GetObject(ThisDrawing.Database.MLeaderStyleDictionaryId, OpenMode.ForRead) as DBDictionary;
                TextStyleTable Text_style_table = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;

                if (Leader_style_table.Contains(ml_style_name) == true)
                {
                    ObjectId ID1 = Leader_style_table.GetAt(ml_style_name);
                    foreach (ObjectId TextStyle_id in Text_style_table)
                    {
                        TextStyleTableRecord TextStyle1 = Trans1.GetObject(TextStyle_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as TextStyleTableRecord;
                        if (TextStyle1.Name.ToLower() == text_style_name.ToLower())
                        {
                            MText mtext1 = new MText();
                            mtext1.Contents = continut;
                            mtext1.TextHeight = Texth;
                            mtext1.BackgroundFill = true;
                            mtext1.UseBackgroundColor = true;
                            mtext1.BackgroundScaleFactor = 1.2;
                            mtext1.ColorIndex = 0;
                            mtext1.TextStyleId = TextStyle1.ObjectId;


                            mleader1.SetDatabaseDefaults();
                            int index1 = mleader1.AddLeader();
                            int index2 = mleader1.AddLeaderLine(index1);
                            mleader1.AddFirstVertex(index2, pt_ins);
                            mleader1.AddLastVertex(index2, new Point3d(pt_ins.X + DeltaX, pt_ins.Y + DeltaY, pt_ins.Z));
                            mleader1.LeaderLineType = LeaderType.StraightLeader;
                            mleader1.ContentType = ContentType.MTextContent;
                            mleader1.MText = mtext1;
                            mleader1.MLeaderStyle = ID1;
                            mleader1.ColorIndex = 256;
                            mleader1.Layer = layer_name;
                            BTrecord.AppendEntity(mleader1);
                            Trans1.AddNewlyCreatedDBObject(mleader1, true);

                        }
                    }



                }



                Trans1.Commit();
            }




            return mleader1;







        }



        public static MLeader creaza_mleader_with_style_IN_BTR(Point3d pt_ins, string continut, double Texth, string ml_style_name, string text_style_name, double DeltaX, double DeltaY, string layer_name, BlockTableRecord BTR, Autodesk.AutoCAD.DatabaseServices.Transaction Trans1, BlockTableRecord BTrecord)
        {

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            MLeader mleader1 = new MLeader();

            DBDictionary Leader_style_table = Trans1.GetObject(ThisDrawing.Database.MLeaderStyleDictionaryId, OpenMode.ForRead) as DBDictionary;
            TextStyleTable Text_style_table = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;

            if (Leader_style_table.Contains(ml_style_name) == true)
            {
                ObjectId ID1 = Leader_style_table.GetAt(ml_style_name);
                foreach (ObjectId TextStyle_id in Text_style_table)
                {
                    TextStyleTableRecord TextStyle1 = Trans1.GetObject(TextStyle_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as TextStyleTableRecord;
                    if (TextStyle1.Name.ToLower() == text_style_name.ToLower())
                    {
                        MText mtext1 = new MText();
                        mtext1.Contents = continut;
                        mtext1.TextHeight = Texth;
                        mtext1.BackgroundFill = true;
                        mtext1.UseBackgroundColor = true;
                        mtext1.BackgroundScaleFactor = 1.2;
                        mtext1.ColorIndex = 0;
                        mtext1.TextStyleId = TextStyle1.ObjectId;


                        mleader1.SetDatabaseDefaults();
                        int index1 = mleader1.AddLeader();
                        int index2 = mleader1.AddLeaderLine(index1);
                        mleader1.AddFirstVertex(index2, pt_ins);
                        mleader1.AddLastVertex(index2, new Point3d(pt_ins.X + DeltaX, pt_ins.Y + DeltaY, pt_ins.Z));
                        mleader1.LeaderLineType = LeaderType.StraightLeader;
                        mleader1.ContentType = ContentType.MTextContent;
                        mleader1.MText = mtext1;
                        mleader1.MLeaderStyle = ID1;
                        mleader1.ColorIndex = 256;
                        mleader1.Layer = layer_name;
                        BTR.AppendEntity(mleader1);


                    }
                }



            }









            return mleader1;







        }


        public static MLeader creaza_mleader_as_label_profile(Point3d pt_ins, string continut)
        {
            double texth = 16;
            double delta_x1 = 32;// 0.16;
            double delta_x2 = 0.1253;
            double delta_y = 30;// 0.3214;
            double lgap = 16;
            double dogl = 16;
            double arrow = 0.00001;


            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            MLeader mleader1 = new MLeader();


            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {

                BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                MText mtext1 = new MText();

                mtext1.Contents = continut;
                mtext1.TextHeight = texth;
                mtext1.BackgroundFill = true;
                mtext1.UseBackgroundColor = true;
                mtext1.BackgroundScaleFactor = 1.2;
                mtext1.ColorIndex = 0;

                mleader1.SetDatabaseDefaults();

                int index1 = mleader1.AddLeader();
                int index2 = mleader1.AddLeaderLine(index1);
                mleader1.AddFirstVertex(index2, pt_ins);
                mleader1.AddLastVertex(index2, new Point3d(pt_ins.X + delta_x1, pt_ins.Y + delta_y, pt_ins.Z));

                mleader1.LeaderLineType = LeaderType.StraightLeader;
                mleader1.ContentType = ContentType.MTextContent;
                mleader1.MText = mtext1;
                mleader1.TextHeight = texth;
                mleader1.LandingGap = lgap;
                mleader1.ArrowSize = arrow;
                mleader1.DoglegLength = dogl;
                mleader1.Annotative = AnnotativeStates.False;
                mleader1.ColorIndex = 256;



                BTrecord.AppendEntity(mleader1);
                Trans1.AddNewlyCreatedDBObject(mleader1, true);
                Trans1.Commit();
            }




            return mleader1;







        }

        public static MText creaza_mtext_label(Point3d pt_ins, string continut, double texth)
        {


            MText mtext1 = new MText();
            mtext1.Attachment = AttachmentPoint.MiddleCenter;
            mtext1.Contents = continut;
            mtext1.TextHeight = texth;
            mtext1.BackgroundFill = true;
            mtext1.UseBackgroundColor = true;
            mtext1.BackgroundScaleFactor = 1.2;
            mtext1.Location = pt_ins;



            return mtext1;


        }

        public static System.Data.DataTable creeaza_custom_band_data_table_structure()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("band_name", typeof(string));
            dt.Columns.Add("Custom_scale", typeof(double));
            dt.Columns.Add("OD_table_name", typeof(string));
            dt.Columns.Add("OD_field1", typeof(string));
            dt.Columns.Add("OD_field2", typeof(string));
            dt.Columns.Add("block_name", typeof(string));
            dt.Columns.Add("block_sta_atr1", typeof(string));
            dt.Columns.Add("block_sta_atr2", typeof(string));
            dt.Columns.Add("block_len_atr", typeof(string));
            dt.Columns.Add("block_field1", typeof(string));
            dt.Columns.Add("block_field2", typeof(string));
            dt.Columns.Add("band_separation", typeof(double));
            dt.Columns.Add("viewport_width", typeof(double));
            dt.Columns.Add("viewport_height", typeof(double));
            dt.Columns.Add("viewport_ps_x", typeof(double));
            dt.Columns.Add("viewport_ps_y", typeof(double));
            dt.Columns.Add("viewport_ms_x", typeof(double));
            dt.Columns.Add("viewport_ms_y", typeof(double));
            dt.Columns.Add("viewport_twist", typeof(double));
            return dt;
        }

        public static System.Data.DataTable creeaza_regular_band_data_table_structure()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("band_name", typeof(string));
            dt.Columns.Add("Custom_scale", typeof(double));
            dt.Columns.Add("OD_table_name", typeof(string));
            dt.Columns.Add("OD_field1", typeof(string));
            dt.Columns.Add("OD_field2", typeof(string));
            dt.Columns.Add("block_name", typeof(string));
            dt.Columns.Add("block_sta_atr1", typeof(string));
            dt.Columns.Add("block_sta_atr2", typeof(string));
            dt.Columns.Add("block_len_atr", typeof(string));
            dt.Columns.Add("block_field1", typeof(string));
            dt.Columns.Add("block_field2", typeof(string));
            dt.Columns.Add("band_separation", typeof(double));
            dt.Columns.Add("viewport_width", typeof(double));
            dt.Columns.Add("viewport_height", typeof(double));
            dt.Columns.Add("viewport_ps_x", typeof(double));
            dt.Columns.Add("viewport_ps_y", typeof(double));
            dt.Columns.Add("viewport_ms_x", typeof(double));
            dt.Columns.Add("viewport_ms_y", typeof(double));
            dt.Columns.Add("viewport_twist", typeof(double));
            return dt;
        }

        public static Point3d Convert_coordinate_to_new_CS(Point3d Point1, string to_coord_system)
        {
            Autodesk.Gis.Map.Platform.AcMapMap Acmap = Autodesk.Gis.Map.Platform.AcMapMap.GetCurrentMap();
            Point3d Point2 = new Point3d();
            string Curent_system = Acmap.GetMapSRS();
            if (string.IsNullOrEmpty(Curent_system) == true)
            {
                MessageBox.Show("Please set your coordinate system");
                return Point2;
            }

            OSGeo.MapGuide.MgCoordinateSystemFactory Coord_factory1 = new OSGeo.MapGuide.MgCoordinateSystemFactory();
            OSGeo.MapGuide.MgCoordinateSystemCatalog Catalog1 = Coord_factory1.GetCatalog();
            OSGeo.MapGuide.MgCoordinateSystemDictionary Dictionary1 = Catalog1.GetCoordinateSystemDictionary();
            OSGeo.MapGuide.MgCoordinateSystemEnum Enum1 = Dictionary1.GetEnum();

            OSGeo.MapGuide.MgCoordinateSystem CoordSys1 = Coord_factory1.Create(Curent_system);

            OSGeo.MapGuide.MgCoordinateSystem CoordSys2 = Dictionary1.GetCoordinateSystem(to_coord_system);

            OSGeo.MapGuide.MgCoordinateSystemTransform Transform1 = Coord_factory1.GetTransform(CoordSys1, CoordSys2);
            OSGeo.MapGuide.MgCoordinate Coord1 = Transform1.Transform(Point1.X, Point1.Y);

            Point2 = new Point3d(Coord1.X, Coord1.Y, 0);
            return Point2;
        }

        public static Point3d Convert_coordinate_from_CS_to_new_CS(Point3d Point1, string from_coord_system, string to_coord_system)
        {
            Autodesk.Gis.Map.Platform.AcMapMap Acmap = Autodesk.Gis.Map.Platform.AcMapMap.GetCurrentMap();
            Point3d Point2 = new Point3d();

            OSGeo.MapGuide.MgCoordinateSystemFactory Coord_factory1 = new OSGeo.MapGuide.MgCoordinateSystemFactory();
            OSGeo.MapGuide.MgCoordinateSystemCatalog Catalog1 = Coord_factory1.GetCatalog();
            OSGeo.MapGuide.MgCoordinateSystemDictionary Dictionary1 = Catalog1.GetCoordinateSystemDictionary();


            OSGeo.MapGuide.MgCoordinateSystem CoordSys1 = Dictionary1.GetCoordinateSystem(from_coord_system);

            OSGeo.MapGuide.MgCoordinateSystem CoordSys2 = Dictionary1.GetCoordinateSystem(to_coord_system);

            OSGeo.MapGuide.MgCoordinateSystemTransform Transform1 = Coord_factory1.GetTransform(CoordSys1, CoordSys2);
            OSGeo.MapGuide.MgCoordinate Coord1 = Transform1.Transform(Point1.X, Point1.Y);

            Point2 = new Point3d(Coord1.X, Coord1.Y, 0);
            return Point2;
        }


        static public System.Data.DataTable Read_block_attributes_and_values(BlockReference Block1)
        {
            System.Data.DataTable Table1 = new System.Data.DataTable();
            Table1.Columns.Add("ATTRIB", typeof(string));
            Table1.Columns.Add("VALUE", typeof(string));


            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);



                    if (Block1.AttributeCollection.Count > 0)
                    {


                        Autodesk.AutoCAD.DatabaseServices.AttributeCollection attColl = Block1.AttributeCollection;

                        foreach (ObjectId ID1 in attColl)
                        {
                            DBObject ent = Trans1.GetObject(ID1, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                            if (ent is AttributeReference)
                            {
                                AttributeReference attref = (AttributeReference)ent;
                                Table1.Rows.Add();
                                Table1.Rows[Table1.Rows.Count - 1]["ATTRIB"] = attref.Tag;
                                if (attref.IsMTextAttribute == false)
                                {
                                    Table1.Rows[Table1.Rows.Count - 1]["VALUE"] = attref.TextString;
                                }
                                if (attref.IsMTextAttribute == true)
                                {
                                    Table1.Rows[Table1.Rows.Count - 1]["VALUE"] = attref.MTextAttribute.Contents;
                                }
                            }

                        }

                    }
                    Trans1.Dispose();
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            return Table1;


        }
        static public void Update_Attrib_block_values(BlockReference Block1, System.Collections.Specialized.StringCollection Col_name, System.Collections.Specialized.StringCollection Col_value)
        {

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);



                    if (Block1.AttributeCollection.Count > 0 & Col_name != null & Col_value != null)
                    {

                        if (Col_name.Count == Col_value.Count)
                        {
                            Autodesk.AutoCAD.DatabaseServices.AttributeCollection attColl = Block1.AttributeCollection;

                            foreach (ObjectId ID1 in attColl)
                            {
                                DBObject ent = Trans1.GetObject(ID1, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                                if (ent is AttributeReference)
                                {
                                    AttributeReference attref = (AttributeReference)ent;
                                    attref.UpgradeOpen();

                                    if (Col_name.Contains(attref.Tag) == true)
                                    {
                                        int index1 = Col_name.IndexOf(attref.Tag);
                                        attref.TextString = Col_value[index1];
                                    }
                                }

                            }

                        }
                    }
                    Trans1.Commit();
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }



        }


        static public double calculate_vp_ms_y(double band_separation, double Y_ps)
        {
            double Y_ms = -10000;
            return Y_ms - band_separation + Y_ps;
        }

        static public void zoom_to_object(ObjectId ObjId)
        {

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;


            try
            {
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);


                        try
                        {
                            Entity Ent1 = Trans1.GetObject(ObjId, OpenMode.ForRead) as Entity;
                            if (Ent1 != null)
                            {

                                Point3d minx = Ent1.GeometricExtents.MinPoint;
                                Point3d maxx = Ent1.GeometricExtents.MaxPoint;

                                using (Autodesk.AutoCAD.GraphicsSystem.Manager GraphicsManager = ThisDrawing.GraphicsManager)
                                {

                                    int Cvport = Convert.ToInt32(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CVPORT"));

                                    //from here 2015 dlls:
                                    Autodesk.AutoCAD.GraphicsSystem.KernelDescriptor kd = new Autodesk.AutoCAD.GraphicsSystem.KernelDescriptor();
                                    kd.addRequirement(Autodesk.AutoCAD.UniqueString.Intern("3D Drawing"));
                                    Autodesk.AutoCAD.GraphicsSystem.View view = GraphicsManager.ObtainAcGsView(Cvport, kd);
                                    // to here 2015 dlls

                                    //from here 2013 dlls:

                                    //Autodesk.AutoCAD.GraphicsSystem.View view = GraphicsManager.GetGsView(Cvport, true);

                                    // to here 2013 dlls

                                    if (view != null)
                                    {
                                        using (view)
                                        {

                                            view.ZoomExtents(Ent1.GeometricExtents.MaxPoint, Ent1.GeometricExtents.MinPoint);

                                            view.Zoom(0.95);//<--optional 
                                            GraphicsManager.SetViewportFromView(Cvport, view, true, true, false);

                                        }
                                    }
                                    Trans1.Commit();
                                }

                            }
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }

                    }
                }
            }







            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        static public void zoom_to_Point(Point3d pt, double factor1)
        {

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;


            try
            {
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);


                        try
                        {



                            Point3d minx = new Point3d(pt.X - factor1, pt.Y - factor1, 0);
                            Point3d maxx = new Point3d(pt.X + factor1, pt.Y + factor1, 0);

                            using (Autodesk.AutoCAD.GraphicsSystem.Manager GraphicsManager = ThisDrawing.GraphicsManager)
                            {

                                int Cvport = Convert.ToInt32(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CVPORT"));

                                //from here 2015 dlls:
                                Autodesk.AutoCAD.GraphicsSystem.KernelDescriptor kd = new Autodesk.AutoCAD.GraphicsSystem.KernelDescriptor();
                                kd.addRequirement(Autodesk.AutoCAD.UniqueString.Intern("3D Drawing"));
                                Autodesk.AutoCAD.GraphicsSystem.View view = GraphicsManager.ObtainAcGsView(Cvport, kd);
                                // to here 2015 dlls

                                //from here 2013 dlls:

                                //Autodesk.AutoCAD.GraphicsSystem.View view = GraphicsManager.GetGsView(Cvport, true);

                                // to here 2013 dlls

                                if (view != null)
                                {
                                    using (view)
                                    {

                                        view.ZoomExtents(minx, maxx);

                                        view.Zoom(0.95);//<--optional 
                                        GraphicsManager.SetViewportFromView(Cvport, view, true, true, false);

                                    }
                                }
                                Trans1.Commit();
                            }


                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }

                    }
                }
            }







            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        public static void create_backup(string fisier1)
        {

            if (System.IO.File.Exists(fisier1) == true)
            {
                string Director1 = System.IO.Path.GetDirectoryName(fisier1);
                if (Director1.Substring(Director1.Length - 1, 1) != "\\")
                {
                    Director1 = Director1 + "\\";
                }

                string name1 = System.IO.Path.GetFileNameWithoutExtension(fisier1);
                string backup1 = Director1 + "~Archive";
                if (System.IO.Directory.Exists(backup1) == false)
                {
                    System.IO.Directory.CreateDirectory(backup1);
                }
                string backup2 = "C:\\Users\\Public\\" + "~Archive";
                if (System.IO.Directory.Exists(backup2) == false && Environment.UserName.ToUpper() == "POP70694")
                {
                    System.IO.Directory.CreateDirectory(backup2);
                }

                string new_name = name1 + "-[" + System.DateTime.Now.Year.ToString() + "_" + System.DateTime.Now.Month.ToString() + "_" + System.DateTime.Now.Day.ToString() +
                    "_" + System.DateTime.Now.Hour.ToString() + "_" + System.DateTime.Now.Minute.ToString() + "_" + System.DateTime.Now.Second.ToString() +
                  "]-" + Environment.UserName.ToUpper() + ".xlsx";
                backup1 = backup1 + "\\" + new_name;
                backup2 = backup2 + "\\" + new_name;

                System.IO.File.Copy(fisier1, backup1);

                if (Environment.UserName.ToUpper() == "POP70694")
                {
                    System.IO.File.Copy(fisier1, backup2);
                }
            }
        }

        public static string extrage_STATION_din_text_de_la_inceputul_textului(string string1)
        {
            string station = "";

            for (int i = 0; i < string1.Length; ++i)
            {
                string Litera = string1.Substring(i, 1);

                switch (Litera)
                {
                    case "0":
                        station = station + Litera;
                        break;
                    case "1":
                        station = station + Litera;
                        break;
                    case "2":
                        station = station + Litera;
                        break;
                    case "3":
                        station = station + Litera;
                        break;
                    case "4":
                        station = station + Litera;
                        break;
                    case "5":
                        station = station + Litera;
                        break;
                    case "6":
                        station = station + Litera;
                        break;
                    case "7":
                        station = station + Litera;
                        break;
                    case "8":
                        station = station + Litera;
                        break;
                    case "9":
                        station = station + Litera;
                        break;
                    case "+":
                        station = station + Litera;
                        break;
                    default:
                        i = string1.Length;
                        break;
                }
            }

            return station;


        }




        public static System.Data.DataTable Creaza_profile_band_datatable_structure()
        {

            string Col_MMid = "MMID";

            string Col_dwg_name = "DwgNo";
            string Col_M1 = "StaBeg";
            string Col_M2 = "StaEnd";
            string Col_zero = "Zero_position";


            System.Type type_string = typeof(string);
            System.Type type_double = typeof(double);



            List<string> Lista1 = new List<string>();
            List<System.Type> Lista2 = new List<System.Type>();

            Lista1.Add(Col_MMid);
            Lista1.Add(Col_dwg_name);
            Lista1.Add(Col_M1);
            Lista1.Add(Col_M2);
            Lista1.Add(Col_zero);
            Lista1.Add("x0");
            Lista1.Add("y0");
            Lista1.Add("height");
            Lista1.Add("length");
            Lista1.Add("Sta_Y");
            Lista1.Add("textH");

            Lista2.Add(type_string);
            Lista2.Add(type_string);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);

            System.Data.DataTable dt1 = new System.Data.DataTable();

            for (int i = 0; i < Lista1.Count; ++i)
            {
                dt1.Columns.Add(Lista1[i], Lista2[i]);
            }
            return dt1;
        }

        public static System.Data.DataTable Build_Data_table_profile_band_from_excel(Microsoft.Office.Interop.Excel.Worksheet W1, int Start_row)
        {


            System.Data.DataTable Data_profile_band = Creaza_profile_band_datatable_structure();
            string Col1 = "B";

            Range range2 = W1.Range[Col1 + Start_row.ToString() + ":" + Col1 + "30000"];
            object[,] values2 = new object[30000, 1];
            values2 = range2.Value2;


            bool is_data = false;
            for (int i = 1; i <= values2.Length; ++i)
            {
                object Valoare2 = values2[i, 1];
                if (Valoare2 != null)
                {
                    Data_profile_band.Rows.Add();
                    is_data = true;
                }
                else
                {
                    i = values2.Length + 1;
                }
            }


            if (is_data == false)
            {
                MessageBox.Show("no data found in the profile band file");
                return Data_profile_band;
            }

            int NrR = Data_profile_band.Rows.Count;
            int NrC = Data_profile_band.Columns.Count;


            if (is_data == true)
            {

                Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[Start_row, 1], W1.Cells[NrR + Start_row - 1, NrC]];

                object[,] values = new object[NrR - 1, NrC - 1];

                values = range1.Value2;

                for (int i = 0; i < Data_profile_band.Rows.Count; ++i)
                {
                    for (int j = 0; j < Data_profile_band.Columns.Count; ++j)
                    {
                        object Valoare = values[i + 1, j + 1];
                        if (Valoare == null) Valoare = DBNull.Value;

                        Data_profile_band.Rows[i][j] = Valoare;
                    }
                }
            }



            return Data_profile_band;


        }









        public static void Create_stationing_od_table()
        {

            {
                try
                {
                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {

                            List<string> List1 = new List<string>();
                            List<string> List2 = new List<string>();
                            List<Autodesk.Gis.Map.Constants.DataType> List3 = new List<Autodesk.Gis.Map.Constants.DataType>();



                            List1.Add("SegmentName");
                            List2.Add("Segment");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add("Note1");
                            List2.Add("Notes");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);



                            Get_object_data_table("Agen_stationing", "Generated by AGEN", List1, List2, List3);


                            Trans1.Commit();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

        }

        public static void Create_kpmp_od_table()
        {

            {
                try
                {
                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {

                            List<string> List1 = new List<string>();
                            List<string> List2 = new List<string>();
                            List<Autodesk.Gis.Map.Constants.DataType> List3 = new List<Autodesk.Gis.Map.Constants.DataType>();



                            List1.Add("SegmentName");
                            List2.Add("Segment");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add("Note1");
                            List2.Add("Notes");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);



                            Functions.Get_object_data_table("Agen_mp_block", "Generated by AGEN", List1, List2, List3);


                            Trans1.Commit();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

        }

        public static void Create_eq_od_table()
        {
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        List<string> List1 = new List<string>();
                        List<string> List2 = new List<string>();
                        List<Autodesk.Gis.Map.Constants.DataType> List3 = new List<Autodesk.Gis.Map.Constants.DataType>();

                        List1.Add("SegmentName");
                        List2.Add("Segment");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add("Note1");
                        List2.Add("Notes");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        Functions.Get_object_data_table("Agen_eq", "Generated by AGEN", List1, List2, List3);

                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public static void Create_northarrow_od_table()
        {

            {
                try
                {
                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {

                            List<string> List1 = new List<string>();
                            List<string> List2 = new List<string>();
                            List<Autodesk.Gis.Map.Constants.DataType> List3 = new List<Autodesk.Gis.Map.Constants.DataType>();



                            List1.Add("SegmentName");
                            List2.Add("Segment");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add("Note1");
                            List2.Add("Notes");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);



                            Functions.Get_object_data_table("Agen_Northarrow", "Generated by AGEN", List1, List2, List3);


                            Trans1.Commit();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

        }

        public static void Create_matchline_block_od_table()
        {

            {
                try
                {
                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {

                            List<string> List1 = new List<string>();
                            List<string> List2 = new List<string>();
                            List<Autodesk.Gis.Map.Constants.DataType> List3 = new List<Autodesk.Gis.Map.Constants.DataType>();



                            List1.Add("SegmentName");
                            List2.Add("Segment");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add("Note1");
                            List2.Add("Notes");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);



                            Functions.Get_object_data_table("Agen_mlblocks", "Generated by AGEN", List1, List2, List3);


                            Trans1.Commit();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

        }

        public static void Create_pi_od_table()
        {
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        List<string> List1 = new List<string>();
                        List<string> List2 = new List<string>();
                        List<Autodesk.Gis.Map.Constants.DataType> List3 = new List<Autodesk.Gis.Map.Constants.DataType>();

                        List1.Add("SegmentName");
                        List2.Add("Segment");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add("Note1");
                        List2.Add("Notes");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        Functions.Get_object_data_table("Agen_pi", "Generated by AGEN", List1, List2, List3);
                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public static Autodesk.Gis.Map.ObjectData.Table Get_object_data_table(string Nume_table, string Description_table, List<string> List_Names, List<string> List_descriptions, List<Autodesk.Gis.Map.Constants.DataType> List_types)
        {
            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

            if (Tables1.IsTableDefined(Nume_table) == true)
            {
                return Tables1[Nume_table];
            }

            if (Tables1.IsTableDefined(Nume_table) == false)
            {
                using (Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_definitions = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.MapUtility.NewODFieldDefinitions())
                {
                    for (int i = 0; i < List_Names.Count; ++i)
                    {
                        Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_definitions.Add(List_Names[i], List_descriptions[i], List_types[i], i);
                    }

                    Tables1.Add(Nume_table, Field_definitions, Description_table, true);
                }
            }
            return Tables1[Nume_table];
        }

        public static void Populate_object_data_table_from_handle_string(Autodesk.Gis.Map.ObjectData.Tables Tables1, string ObjId, string Nume_table, List<object> List_value, List<Autodesk.Gis.Map.Constants.DataType> List_types)
        {
            if (Tables1.IsTableDefined(Nume_table) == true)
            {
                using (Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Nume_table])
                {
                    ObjectId oB1 = GetObjectId(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database, ObjId);
                    using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tabla1.GetObjectTableRecords(Convert.ToUInt32(0), oB1, Autodesk.Gis.Map.Constants.OpenMode.OpenForWrite, true))
                    {
                        if (Records1.Count > 0)
                        {
                            foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                            {
                                Autodesk.Gis.Map.Utilities.MapValue Valoare1;
                                for (int i = 0; i < List_value.Count; ++i)
                                {
                                    Valoare1 = Record1[i];
                                    if (List_types[i] == Autodesk.Gis.Map.Constants.DataType.Character)
                                    {
                                        Valoare1.Assign(List_value[i].ToString());
                                    }

                                    if (List_types[i] == Autodesk.Gis.Map.Constants.DataType.Real)
                                    {
                                        Valoare1.Assign(Convert.ToDouble(List_value[i]));
                                    }
                                    Records1.UpdateRecord(Record1);
                                }
                            }
                        }
                        else
                        {
                            using (Autodesk.Gis.Map.ObjectData.Record rec = Autodesk.Gis.Map.ObjectData.Record.Create())
                            {
                                Tabla1.InitRecord(rec);
                                for (int i = 0; i < List_value.Count; ++i)
                                {
                                    Autodesk.Gis.Map.Utilities.MapValue Val = rec[i];
                                    if (List_types[i] == Autodesk.Gis.Map.Constants.DataType.Character)
                                    {
                                        string Valoare = List_value[i].ToString();
                                        Val.Assign(Valoare);
                                    }
                                    if (List_types[i] == Autodesk.Gis.Map.Constants.DataType.Real)
                                    {
                                        double Valoare = Convert.ToDouble(List_value[i]);
                                        Val.Assign(Valoare);
                                    }
                                }
                                Tabla1.AddRecord(rec, oB1);
                            }
                        }
                    }
                }
            }
        }

        public static void Populate_object_data_table_from_objectid(Autodesk.Gis.Map.ObjectData.Tables Tables1, ObjectId id1, string Nume_table, List<object> List_value, List<Autodesk.Gis.Map.Constants.DataType> List_types)
        {
            if (Tables1.IsTableDefined(Nume_table) == true)
            {
                using (Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Nume_table])
                {

                    using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tabla1.GetObjectTableRecords(Convert.ToUInt32(0), id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForWrite, true))
                    {
                        if (Records1.Count > 0)
                        {
                            foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                            {
                                Autodesk.Gis.Map.Utilities.MapValue Valoare1;
                                for (int i = 0; i < List_value.Count; ++i)
                                {
                                    Valoare1 = Record1[i];
                                    if (List_types[i] == Autodesk.Gis.Map.Constants.DataType.Character)
                                    {
                                        Valoare1.Assign(List_value[i].ToString());
                                    }

                                    if (List_types[i] == Autodesk.Gis.Map.Constants.DataType.Real)
                                    {
                                        Valoare1.Assign(Convert.ToDouble(List_value[i]));
                                    }
                                    Records1.UpdateRecord(Record1);
                                }
                            }
                        }
                        else
                        {
                            using (Autodesk.Gis.Map.ObjectData.Record rec = Autodesk.Gis.Map.ObjectData.Record.Create())
                            {
                                Tabla1.InitRecord(rec);
                                for (int i = 0; i < List_value.Count; ++i)
                                {
                                    Autodesk.Gis.Map.Utilities.MapValue Val = rec[i];
                                    if (List_types[i] == Autodesk.Gis.Map.Constants.DataType.Character)
                                    {
                                        string Valoare = "";
                                        if (List_value[i] != null)
                                        {
                                            Valoare = List_value[i].ToString();
                                        }

                                        Val.Assign(Valoare);
                                    }
                                    if (List_types[i] == Autodesk.Gis.Map.Constants.DataType.Real)
                                    {
                                        double Valoare = 0;
                                        if (List_value[i] != null)
                                        {
                                            Valoare = Convert.ToDouble(List_value[i]);
                                        }

                                        Val.Assign(Valoare);
                                    }
                                }
                                Tabla1.AddRecord(rec, id1);
                            }
                        }
                    }
                }
            }
        }





        public static System.Data.DataTable creaza_crossing_block_table_record_structure()
        {
            System.Data.DataTable dt_rec = new System.Data.DataTable();
            dt_rec.Columns.Add("objectid", typeof(string));
            dt_rec.Columns.Add("layer", typeof(string));
            dt_rec.Columns.Add("stationprefix", typeof(string));
            dt_rec.Columns.Add("station", typeof(string));
            dt_rec.Columns.Add("descriptionprefix", typeof(string));
            dt_rec.Columns.Add("description", typeof(string));
            dt_rec.Columns.Add("x", typeof(double));
            dt_rec.Columns.Add("y", typeof(double));
            dt_rec.Columns.Add("textheight", typeof(double));
            dt_rec.Columns.Add("rotation", typeof(double));
            dt_rec.Columns.Add("underline", typeof(string));
            dt_rec.Columns.Add("widthfactor", typeof(double));
            dt_rec.Columns.Add("xins", typeof(double));
            dt_rec.Columns.Add("yins", typeof(double));
            dt_rec.Columns.Add("xm1", typeof(double));
            dt_rec.Columns.Add("ym1", typeof(double));
            dt_rec.Columns.Add("xm2", typeof(double));
            dt_rec.Columns.Add("ym2", typeof(double));
            dt_rec.Columns.Add("xsw", typeof(double));
            return dt_rec;
        }

        public static System.Data.DataTable build_crossing_block_record_from_excel(Microsoft.Office.Interop.Excel.Worksheet W1, int Start_row)
        {

            System.Data.DataTable dt_rec = creaza_crossing_block_table_record_structure();

            Range range2 = W1.Range["A" + Start_row.ToString() + ":A30000"];
            object[,] values2 = new object[30000, 1];
            values2 = range2.Value2;

            bool is_data = false;

            for (int i = 1; i <= values2.Length; ++i)
            {
                object Valoare2 = values2[i, 1];
                if (Valoare2 != null)
                {
                    dt_rec.Rows.Add();
                    is_data = true;
                }
                else
                {
                    i = values2.Length + 1;
                }
            }

            int NrR = dt_rec.Rows.Count;
            int NrC = dt_rec.Columns.Count;


            if (is_data == true)
            {
                Range range1 = W1.Range[W1.Cells[Start_row, 1], W1.Cells[NrR + Start_row - 1, NrC]];

                object[,] values = new object[NrR - 1, NrC - 1];

                values = range1.Value2;

                for (int i = 0; i < dt_rec.Rows.Count; ++i)
                {
                    for (int j = 0; j < dt_rec.Columns.Count; ++j)
                    {
                        object Valoare = values[i + 1, j + 1];
                        if (Valoare == null) Valoare = DBNull.Value;
                        dt_rec.Rows[i][j] = Valoare;
                    }
                }
            }

            return dt_rec;
        }

        public static string remove_space_from_start_and_end_of_a_string(string String1)
        {
            if (String1 != "" && String1.Length > 1)
            {
                if (String1.Substring(String1.Length - 1, 1) == " ")
                {
                    do
                    {
                        String1 = String1.Substring(0, String1.Length - 1);
                    }
                    while (String1.Substring(String1.Length - 1, 1) == " " && String1.Length > 1);
                }
                if (String1.Substring(0, 1) == " " && String1.Length > 1)
                {
                    do
                    {
                        String1 = String1.Substring(1, String1.Length - 1);
                    }
                    while (String1.Substring(0, 1) == " " && String1.Length > 1);
                }
            }

            return String1;
        }

        public static double get_width_factor_of_an_mtext(string Content1)
        {
            double wf = 1;

            if (Content1.Contains("\\W") == true && Content1.Contains(";") == true)
            {
                int poz1 = Content1.IndexOf("\\W");
                int poz2 = Content1.IndexOf(";");
                string width1 = Content1.Substring(poz1 + 2, poz2 - poz1 - 2);
                if (IsNumeric(width1) == true)
                {
                    wf = Convert.ToDouble(width1);
                }
            }

            return wf;
        }

        public static string extract_station_from_mtext(string Content1)
        {
            string station = "";

            if (Content1.Contains("+") == true)
            {
                int poz1 = Content1.IndexOf("+");
                int start1 = -1;
                int end1 = -1;
                for (int i = 0; i < poz1; ++i)
                {
                    string litera = Content1.Substring(i, 1);
                    if (start1 == -1 && litera == "-")
                    {
                        start1 = i;
                        i = poz1;
                    }
                    else if (start1 == -1 && (IsNumeric(litera) == true))
                    {
                        start1 = i;
                        i = poz1;
                    }
                }

                for (int i = poz1 + 1; i < Content1.Length; ++i)
                {
                    string litera = Content1.Substring(i, 1);
                    if (litera != "." && IsNumeric(litera) == false)
                    {
                        end1 = i - 1;
                        i = Content1.Length;
                    }
                }

                if (start1 > -1 && end1 > -1)
                {
                    station = Content1.Substring(start1, end1 - start1 + 1);
                }

            }

            return remove_space_from_start_and_end_of_a_string(station);
        }

        public static string extract_stationprefix_from_mtext(string Content1)
        {
            Content1 = remove_space_from_start_and_end_of_a_string(Content1);
            string station = extract_station_from_mtext(Content1);
            string station_prefix = "";
            int poz1 = 0;
            if (station != "")
            {
                poz1 = Content1.IndexOf(station);
            }

            if (poz1 > 0)
            {
                station_prefix = Content1.Substring(0, poz1);
            }
            station_prefix = remove_space_from_start_and_end_of_a_string(station_prefix);
            if (station_prefix.Replace(" ", "") == "") return "";
            return remove_space_from_start_and_end_of_a_string(station_prefix);
        }

        public static string extract_description_from_mtext(string Content1)
        {
            string descr = Content1;

            if (Content1.Contains("+") == true)
            {
                int poz1 = Content1.IndexOf("+");

                int end1 = -1;


                for (int i = poz1 + 1; i < Content1.Length; ++i)
                {
                    string litera = Content1.Substring(i, 1);
                    if (litera != "." && IsNumeric(litera) == false)
                    {
                        end1 = i;
                        i = Content1.Length;
                    }
                }

                if (end1 > -1)
                {
                    descr = Content1.Substring(end1, Content1.Length - end1);
                }

            }


            return remove_space_from_start_and_end_of_a_string(descr);
        }

        public static void Create_ownership_od_table()
        {
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        List<string> List1 = new List<string>();
                        List<string> List2 = new List<string>();
                        List<Autodesk.Gis.Map.Constants.DataType> List3 = new List<Autodesk.Gis.Map.Constants.DataType>();

                        List1.Add("SegmentName");
                        List2.Add("Segment");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add("Note1");
                        List2.Add("Notes");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add("xbeg");
                        List2.Add("X BEG");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                        List1.Add("ybeg");
                        List2.Add("Y BEG");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                        List1.Add("xend");
                        List2.Add("X END");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                        List1.Add("yend");
                        List2.Add("Y END");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                        Functions.Get_object_data_table("Agen_owner", "Generated by AGEN", List1, List2, List3);
                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public static void Create_crossing_od_table()
        {
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        List<string> List1 = new List<string>();
                        List<string> List2 = new List<string>();
                        List<Autodesk.Gis.Map.Constants.DataType> List3 = new List<Autodesk.Gis.Map.Constants.DataType>();

                        List1.Add("SegmentName");
                        List2.Add("Segment");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add("Note1");
                        List2.Add("Notes");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add("x");
                        List2.Add("X crossing");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                        List1.Add("y");
                        List2.Add("Y crossing");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);



                        Functions.Get_object_data_table("Agen_crossing", "Generated by AGEN", List1, List2, List3);
                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }




        public static Matrix3d WCS_align()
        {
            Matrix3d m3d = new Matrix3d();
            try
            {
                Point3d pt1 = new Point3d(0, 0, 0);
                Point3d pt2 = new Point3d(0, 0, 11);
                Point3d pt3 = new Point3d(0, 2, 0);
                Vector3d zaxis = pt1.GetVectorTo(pt2).GetNormal();
                Vector3d yaxis = pt1.GetVectorTo(pt3).GetNormal();
                Vector3d xaxis = yaxis.CrossProduct(zaxis).GetNormal();
                m3d = Matrix3d.AlignCoordinateSystem(Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis, pt1, xaxis, yaxis, zaxis);
            }
            catch (System.Exception ex)
            {

            }
            return m3d;
        }


        public static Matrix3d ModelToPaper(Viewport vp)
        {
            Vector3d vd = vp.ViewDirection;
            Point3d vc = new Point3d(vp.ViewCenter.X, vp.ViewCenter.Y, 0);
            Point3d vt = vp.ViewTarget;
            Point3d cp = vp.CenterPoint;
            double ta = -vp.TwistAngle;
            double vh = vp.ViewHeight;
            double height = vp.Height;
            double width = vp.Width;
            double scale = vh / height;
            double lensLength = vp.LensLength;
            Vector3d zaxis = vd.GetNormal();
            Vector3d xaxis = Vector3d.ZAxis.CrossProduct(vd);
            Vector3d yaxis;

            if (!xaxis.IsZeroLength())
            {
                xaxis = xaxis.GetNormal();
                yaxis = zaxis.CrossProduct(xaxis);
            }
            else if (zaxis.Z < 0)
            {
                xaxis = Vector3d.XAxis * -1;
                yaxis = Vector3d.YAxis;
                zaxis = Vector3d.ZAxis * -1;
            }
            else
            {
                xaxis = Vector3d.XAxis;
                yaxis = Vector3d.YAxis;
                zaxis = Vector3d.ZAxis;
            }
            Matrix3d pcsToDCS = Matrix3d.Displacement(Point3d.Origin - cp);
            pcsToDCS = pcsToDCS * Matrix3d.Scaling(scale, cp);
            Matrix3d dcsToWcs = Matrix3d.Displacement(vc - Point3d.Origin);
            Matrix3d mxCoords = Matrix3d.AlignCoordinateSystem(
                  Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis,
                 Vector3d.ZAxis, Point3d.Origin,
                xaxis, yaxis, zaxis);
            dcsToWcs = mxCoords * dcsToWcs;
            dcsToWcs = Matrix3d.Displacement(vt - Point3d.Origin) * dcsToWcs;
            dcsToWcs = Matrix3d.Rotation(ta, zaxis, vt) * dcsToWcs;

            Matrix3d perspectiveMx = Matrix3d.Identity;
            if (vp.PerspectiveOn)
            {
                double vSize = vh;
                double aspectRatio = width / height;
                double adjustFactor = 1.0 / 42.0;
                double adjstLenLgth = vSize * lensLength *
                   Math.Sqrt(1.0 + aspectRatio * aspectRatio) * adjustFactor;
                double iDist = vd.Length;
                double lensDist = iDist - adjstLenLgth;
                double[] dataAry = new double[]
                {
                     1,0,0,0,0,1,0,0,0,0,
                       (adjstLenLgth-lensDist)/adjstLenLgth,
                       lensDist*(iDist-adjstLenLgth)/adjstLenLgth,
                     0,0,-1.0/adjstLenLgth,iDist/adjstLenLgth
               };

                perspectiveMx = new Matrix3d(dataAry);
            }

            Matrix3d finalMx =
             pcsToDCS.Inverse() * perspectiveMx * dcsToWcs.Inverse();

            return finalMx;
        }
        public static Matrix3d PaperToModel(Viewport vp)
        {
            Matrix3d mx = ModelToPaper(vp);
            return mx.Inverse();
        }

        public static void Create_vp_grab_od_table(string tablename)
        {
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        List<string> List1 = new List<string>();
                        List<string> List2 = new List<string>();
                        List<Autodesk.Gis.Map.Constants.DataType> List3 = new List<Autodesk.Gis.Map.Constants.DataType>();

                        List1.Add("Drawing_Number");
                        List2.Add("dwg name");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add("Drawing_Type");
                        List2.Add("dwg type");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add("Note1");
                        List2.Add("Notes");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);



                        Functions.Get_object_data_table(tablename, "Generated by AGEN", List1, List2, List3);
                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public static System.Data.DataTable Creaza_weldmap_pipelist_datatable_structure()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("Pipe ID", typeof(string));
            dt.Columns.Add("Heat", typeof(string));
            dt.Columns.Add("Length", typeof(string));
            dt.Columns.Add("Wall Thickness", typeof(string));
            dt.Columns.Add("Diameter", typeof(string));
            dt.Columns.Add("Grade", typeof(string));
            dt.Columns.Add("Coating", typeof(string));
            dt.Columns.Add("Manufacture", typeof(string));
            dt.Columns.Add("DoubleJointNo", typeof(string));
            return dt;
        }


        public static System.Data.DataTable Creaza_weldmap_dbl_joint_datatable_structure()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("DoubleJoint#", typeof(string));
            dt.Columns.Add("Pipe ID 1", typeof(string));
            dt.Columns.Add("Heat 1", typeof(string));
            dt.Columns.Add("Pipe ID 2", typeof(string));
            dt.Columns.Add("Heat 2", typeof(string));
            dt.Columns.Add("Length", typeof(string));
            dt.Columns.Add("Wall Thickness", typeof(string));
            dt.Columns.Add("Diameter", typeof(string));
            dt.Columns.Add("Grade", typeof(string));
            dt.Columns.Add("Coating", typeof(string));
            dt.Columns.Add("Manufacture", typeof(string));
            return dt;
        }

        public static System.Data.DataTable Creaza_weldmap_pipe_tally_datatable_structure()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("MMID", typeof(string));
            dt.Columns.Add("Pipe", typeof(string));
            dt.Columns.Add("Heat", typeof(string));
            dt.Columns.Add("OriginalLength", typeof(string));
            dt.Columns.Add("NewLength", typeof(string));
            dt.Columns.Add("WallThickness", typeof(string));
            dt.Columns.Add("Diameter", typeof(string));
            dt.Columns.Add("Grade", typeof(string));
            dt.Columns.Add("Coating", typeof(string));
            dt.Columns.Add("Manufacture", typeof(string));
            dt.Columns.Add("DoubleJointNo", typeof(string));
            return dt;
        }

        static public System.Data.DataTable Populate_data_table_from_excel(System.Data.DataTable dt1, Worksheet W1, int start_row,
            string checkColumn1, string checkColumn2, string checkColumn3, string checkColumn4, string checkColumn5, string checkColumn6, string checkColumn7, string checkColumn8, string checkColumn9, string checkColumn10, string checkColumn11)
        {
            if (W1 == null) return dt1;


            if (checkColumn1 != "")
            {
                Range range1 = W1.Range[checkColumn1 + start_row.ToString() + ":" + checkColumn1 + "300000"];
                object[,] values2 = new object[300000, 1];
                values2 = range1.Value2;


                for (int i = 1; i <= values2.Length; ++i)
                {
                    object Valoare2 = values2[i, 1];
                    if (Valoare2 != null)
                    {
                        dt1.Rows.Add();
                    }
                    else
                    {
                        i = values2.Length + 1;
                    }
                }
            }

            if (checkColumn2 != "")
            {
                Range range2 = W1.Range[checkColumn2 + Convert.ToString(start_row + dt1.Rows.Count) + ":" + checkColumn2 + "300000"];
                object[,] values3 = new object[300000, 1];
                values3 = range2.Value2;

                for (int i = 1; i <= values3.Length; ++i)
                {
                    object Valoare3 = values3[i, 1];
                    if (Valoare3 != null)
                    {
                        dt1.Rows.Add();
                    }
                    else
                    {
                        i = values3.Length + 1;
                    }
                }
            }

            if (checkColumn3 != "")
            {
                Range range3 = W1.Range[checkColumn3 + Convert.ToString(start_row + dt1.Rows.Count) + ":" + checkColumn3 + "300000"];
                object[,] values3 = new object[300000, 1];
                values3 = range3.Value2;

                for (int i = 1; i <= values3.Length; ++i)
                {
                    object Valoare3 = values3[i, 1];
                    if (Valoare3 != null)
                    {
                        dt1.Rows.Add();
                    }
                    else
                    {
                        i = values3.Length + 1;
                    }
                }
            }

            if (checkColumn4 != "")
            {
                Range range4 = W1.Range[checkColumn4 + Convert.ToString(start_row + dt1.Rows.Count) + ":" + checkColumn4 + "300000"];
                object[,] values4 = new object[300000, 1];
                values4 = range4.Value2;

                for (int i = 1; i <= values4.Length; ++i)
                {
                    object Valoare4 = values4[i, 1];
                    if (Valoare4 != null)
                    {
                        dt1.Rows.Add();
                    }
                    else
                    {
                        i = values4.Length + 1;
                    }
                }
            }

            if (checkColumn5 != "")
            {
                Range range5 = W1.Range[checkColumn5 + Convert.ToString(start_row + dt1.Rows.Count) + ":" + checkColumn5 + "300000"];
                object[,] values5 = new object[300000, 1];
                values5 = range5.Value2;

                for (int i = 1; i <= values5.Length; ++i)
                {
                    object Valoare5 = values5[i, 1];
                    if (Valoare5 != null)
                    {
                        dt1.Rows.Add();
                    }
                    else
                    {
                        i = values5.Length + 1;
                    }
                }
            }

            if (checkColumn6 != "")
            {
                Range range6 = W1.Range[checkColumn6 + Convert.ToString(start_row + dt1.Rows.Count) + ":" + checkColumn6 + "300000"];
                object[,] values6 = new object[300000, 1];
                values6 = range6.Value2;

                for (int i = 1; i <= values6.Length; ++i)
                {
                    object Valoare6 = values6[i, 1];
                    if (Valoare6 != null)
                    {
                        dt1.Rows.Add();
                    }
                    else
                    {
                        i = values6.Length + 1;
                    }
                }
            }

            if (checkColumn7 != "")
            {
                Range range7 = W1.Range[checkColumn7 + Convert.ToString(start_row + dt1.Rows.Count) + ":" + checkColumn7 + "300000"];
                object[,] values7 = new object[300000, 1];
                values7 = range7.Value2;

                for (int i = 1; i <= values7.Length; ++i)
                {
                    object Valoare7 = values7[i, 1];
                    if (Valoare7 != null)
                    {
                        dt1.Rows.Add();
                    }
                    else
                    {
                        i = values7.Length + 1;
                    }
                }
            }

            if (checkColumn8 != "")
            {
                Range range8 = W1.Range[checkColumn8 + Convert.ToString(start_row + dt1.Rows.Count) + ":" + checkColumn8 + "300000"];
                object[,] values8 = new object[300000, 1];
                values8 = range8.Value2;

                for (int i = 1; i <= values8.Length; ++i)
                {
                    object Valoare8 = values8[i, 1];
                    if (Valoare8 != null)
                    {
                        dt1.Rows.Add();
                    }
                    else
                    {
                        i = values8.Length + 1;
                    }
                }
            }

            if (checkColumn9 != "")
            {
                Range range9 = W1.Range[checkColumn9 + Convert.ToString(start_row + dt1.Rows.Count) + ":" + checkColumn9 + "300000"];
                object[,] values9 = new object[300000, 1];
                values9 = range9.Value2;

                for (int i = 1; i <= values9.Length; ++i)
                {
                    object Valoare9 = values9[i, 1];
                    if (Valoare9 != null)
                    {
                        dt1.Rows.Add();
                    }
                    else
                    {
                        i = values9.Length + 1;
                    }
                }
            }

            if (checkColumn10 != "")
            {
                Range range10 = W1.Range[checkColumn10 + Convert.ToString(start_row + dt1.Rows.Count) + ":" + checkColumn10 + "300000"];
                object[,] values10 = new object[300000, 1];
                values10 = range10.Value2;

                for (int i = 1; i <= values10.Length; ++i)
                {
                    object Valoare10 = values10[i, 1];
                    if (Valoare10 != null)
                    {
                        dt1.Rows.Add();
                    }
                    else
                    {
                        i = values10.Length + 1;
                    }
                }
            }

            if (checkColumn11 != "")
            {
                Range range11 = W1.Range[checkColumn11 + Convert.ToString(start_row + dt1.Rows.Count) + ":" + checkColumn11 + "300000"];
                object[,] values11 = new object[300000, 1];
                values11 = range11.Value2;

                for (int i = 1; i <= values11.Length; ++i)
                {
                    object Valoare11 = values11[i, 1];
                    if (Valoare11 != null)
                    {
                        dt1.Rows.Add();
                    }
                    else
                    {
                        i = values11.Length + 1;
                    }
                }
            }

            int NrC = dt1.Columns.Count;
            int NrR = dt1.Rows.Count;

            if (dt1.Rows.Count == 0)
            {
                MessageBox.Show("no data found in the file");
                return dt1;
            }

            if (dt1.Rows.Count > 0)
            {
                Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[start_row, 1], W1.Cells[NrR + start_row - 1, NrC]];
                object[,] values = new object[NrR - 1, NrC - 1];
                values = range1.Value2;
                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    for (int j = 0; j < dt1.Columns.Count; ++j)
                    {
                        object Valoare = values[i + 1, j + 1];
                        if (Valoare == null) Valoare = DBNull.Value;
                        dt1.Rows[i][j] = Valoare;
                    }
                }
            }

            return dt1;
        }

        static public Worksheet Get_opened_worksheet_from_Excel_by_name(string filename, string SheetName)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application Excel1;
                Microsoft.Office.Interop.Excel.Workbook Workbook1;
                Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                if (Excel1 == null) return null;
                for (int j = 1; j <= Excel1.Workbooks.Count; ++j)
                {
                    Workbook1 = Excel1.Workbooks[j];
                    if (Workbook1.Name == filename)
                    {
                        for (int i = 1; i <= Workbook1.Worksheets.Count; ++i)
                        {
                            if (Workbook1.Worksheets[i].name == SheetName)
                            {
                                return Workbook1.Worksheets[i];
                            }
                        }
                    }
                }
                return null;
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                return null;
            }



        }

        static public Workbook Get_opened_workbook_from_Excel(string filename)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application Excel1;
                Microsoft.Office.Interop.Excel.Workbook Workbook1;
                Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                if (Excel1 == null) return null;
                for (int j = 1; j <= Excel1.Workbooks.Count; ++j)
                {
                    Workbook1 = Excel1.Workbooks[j];
                    if (Workbook1.Name == filename)
                    {
                        return Workbook1;
                    }
                }
                return null;
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                return null;
            }
        }

        static public void Load_opened_worksheets_to_combobox(ComboBox combo1)
        {
            combo1.Items.Clear();
            try
            {
                Microsoft.Office.Interop.Excel.Application Excel1;
                Microsoft.Office.Interop.Excel.Workbook Workbook1;
                Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                if (Excel1 == null) return;
                for (int j = 1; j <= Excel1.Workbooks.Count; ++j)
                {
                    Workbook1 = Excel1.Workbooks[j];
                    string wn = Workbook1.Name;
                    for (int i = 1; i <= Workbook1.Worksheets.Count; ++i)
                    {
                        combo1.Items.Add("[" + Workbook1.Worksheets[i].name + "] - " + wn);
                    }
                }
                if (combo1.Items.Count > 0) combo1.SelectedIndex = 0;

            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }



        }

        static public void Load_opened_workbooks_to_combobox(ComboBox combo1)
        {
            combo1.Items.Clear();
            try
            {
                Microsoft.Office.Interop.Excel.Application Excel1;
                Microsoft.Office.Interop.Excel.Workbook Workbook1;
                Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                if (Excel1 == null) return;
                for (int j = 1; j <= Excel1.Workbooks.Count; ++j)
                {
                    Workbook1 = Excel1.Workbooks[j];
                    string wn = Workbook1.Name;
                    combo1.Items.Add(wn);
                }
                if (combo1.Items.Count > 0) combo1.SelectedIndex = 0;

            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }



        }


        static public System.Data.DataTable build_data_table_from_excel_based_on_columns
      (
          System.Data.DataTable dt1,
          Microsoft.Office.Interop.Excel.Worksheet W1, int start_row,
          string col1, string colpt1, string col2, string colpt2, string col3, string colpt3, string col4, string colpt4,
          string col5, string colpt5, string col6, string colpt6, string col7, string colpt7, string col8, string colpt8,
          string col9, string colpt9, string col10, string colpt10, string col11, string colpt11,
          string col12, string colpt12, string col13, string colpt13, string col14, string colpt14,
          string col15, string colpt15, string col16, string colpt16, string col17, string colpt17, string col18, string colpt18,
          string col19, string colpt19, string col20, string colpt20, string col21, string colpt21,
          string col22, string colpt22, string col23, string colpt23, string col24, string colpt24,
          string col25, string colpt25, string col26, string colpt26, string col27, string colpt27
      )
        {
            if (W1 == null) return dt1;


            object[,] values1 = new object[300000, 1];
            object[,] values2 = new object[300000, 1];
            object[,] values3 = new object[300000, 1];
            object[,] values4 = new object[300000, 1];
            object[,] values5 = new object[300000, 1];
            object[,] values6 = new object[300000, 1];
            object[,] values7 = new object[300000, 1];
            object[,] values8 = new object[300000, 1];
            object[,] values9 = new object[300000, 1];
            object[,] values10 = new object[300000, 1];
            object[,] values11 = new object[300000, 1];
            object[,] values12 = new object[300000, 1];
            object[,] values13 = new object[300000, 1];
            object[,] values14 = new object[300000, 1];
            object[,] values15 = new object[300000, 1];
            object[,] values16 = new object[300000, 1];
            object[,] values17 = new object[300000, 1];
            object[,] values18 = new object[300000, 1];
            object[,] values19 = new object[300000, 1];
            object[,] values20 = new object[300000, 1];
            object[,] values21 = new object[300000, 1];
            object[,] values22 = new object[300000, 1];
            object[,] values23 = new object[300000, 1];
            object[,] values24 = new object[300000, 1];
            object[,] values25 = new object[300000, 1];
            object[,] values26 = new object[300000, 1];
            object[,] values27 = new object[300000, 1];

            #region 1
            if (colpt1 != "")
            {
                Microsoft.Office.Interop.Excel.Range range1 = W1.Range[colpt1 + start_row.ToString() + ":" + colpt1 + "300000"];
                values1 = range1.Value2;

                for (int i = 1; i <= values1.Length; ++i)
                {
                    object Valoare1 = values1[i, 1];
                    if (Valoare1 != null)
                    {
                        dt1.Rows.Add();
                    }
                    else
                    {
                        i = values1.Length + 1;
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values1[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col1] = Valoare;
                }
            }
            #endregion

            #region 2
            if (colpt2 != "")
            {
                Microsoft.Office.Interop.Excel.Range range2 = W1.Range[colpt2 + Convert.ToString(start_row) + ":" + colpt2 + "300000"];

                values2 = range2.Value2;

                for (int i = 1; i <= values2.Length; ++i)
                {
                    object Valoare2 = values2[i, 1];
                    if (Valoare2 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();
                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }

                        }

                    }
                    else
                    {
                        i = values2.Length + 1;
                    }

                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values2[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col2] = Valoare;
                }
            }
            #endregion

            #region 3
            if (colpt3 != "")
            {
                Microsoft.Office.Interop.Excel.Range range3 = W1.Range[colpt3 + Convert.ToString(start_row) + ":" + colpt3 + "300000"];

                values3 = range3.Value2;

                for (int i = 1; i <= values3.Length; ++i)
                {
                    object Valoare3 = values3[i, 1];
                    if (Valoare3 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();
                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }

                        }

                    }
                    else
                    {
                        i = values3.Length + 1;
                    }

                }



                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values3[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col3] = Valoare;
                }
            }
            #endregion

            #region 4
            if (colpt4 != "")
            {
                Microsoft.Office.Interop.Excel.Range range4 = W1.Range[colpt4 + Convert.ToString(start_row) + ":" + colpt4 + "300000"];

                values4 = range4.Value2;

                for (int i = 1; i <= values4.Length; ++i)
                {
                    object Valoare4 = values4[i, 1];
                    if (Valoare4 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                        }

                    }
                    else
                    {
                        i = values4.Length + 1;
                    }

                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values4[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col4] = Valoare;
                }
            }
            #endregion

            #region 5
            if (colpt5 != "")
            {
                Microsoft.Office.Interop.Excel.Range range5 = W1.Range[colpt5 + Convert.ToString(start_row) + ":" + colpt5 + "300000"];

                values5 = range5.Value2;

                for (int i = 1; i <= values5.Length; ++i)
                {
                    object Valoare5 = values5[i, 1];
                    if (Valoare5 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }
                        }

                    }
                    else
                    {
                        i = values5.Length + 1;
                    }

                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values5[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col5] = Valoare;
                }
            }
            #endregion

            #region 6
            if (colpt6 != "")
            {
                Microsoft.Office.Interop.Excel.Range range6 = W1.Range[colpt6 + Convert.ToString(start_row) + ":" + colpt6 + "300000"];

                values6 = range6.Value2;

                for (int i = 1; i <= values6.Length; ++i)
                {
                    object Valoare6 = values6[i, 1];
                    if (Valoare6 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();


                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                        }

                    }
                    else
                    {
                        i = values6.Length + 1;
                    }

                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values6[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col6] = Valoare;
                }
            }
            #endregion

            #region 7
            if (colpt7 != "")
            {
                Microsoft.Office.Interop.Excel.Range range7 = W1.Range[colpt7 + Convert.ToString(start_row) + ":" + colpt7 + "300000"];

                values7 = range7.Value2;

                for (int i = 1; i <= values7.Length; ++i)
                {
                    object Valoare7 = values7[i, 1];
                    if (Valoare7 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                        }

                    }
                    else
                    {
                        i = values7.Length + 1;
                    }

                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values7[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col7] = Valoare;
                }
            }
            #endregion

            #region 8
            if (colpt8 != "")
            {
                Microsoft.Office.Interop.Excel.Range range8 = W1.Range[colpt8 + Convert.ToString(start_row) + ":" + colpt8 + "300000"];

                values8 = range8.Value2;

                for (int i = 1; i <= values8.Length; ++i)
                {
                    object Valoare8 = values8[i, 1];
                    if (Valoare8 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }

                        }

                    }
                    else
                    {
                        i = values8.Length + 1;
                    }

                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values8[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col8] = Valoare;
                }
            }
            #endregion

            #region 9
            if (colpt9 != "")
            {
                Microsoft.Office.Interop.Excel.Range range9 = W1.Range[colpt9 + Convert.ToString(start_row) + ":" + colpt9 + "300000"];

                values9 = range9.Value2;

                for (int i = 1; i <= values9.Length; ++i)
                {
                    object Valoare9 = values9[i, 1];
                    if (Valoare9 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();
                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }
                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }
                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }
                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }
                        }
                    }
                    else
                    {
                        i = values9.Length + 1;
                    }

                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values9[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col9] = Valoare;
                }
            }
            #endregion

            #region 10
            if (colpt10 != "")
            {
                Microsoft.Office.Interop.Excel.Range range10 = W1.Range[colpt10 + Convert.ToString(start_row) + ":" + colpt10 + "300000"];

                values10 = range10.Value2;

                for (int i = 1; i <= values10.Length; ++i)
                {
                    object Valoare10 = values10[i, 1];
                    if (Valoare10 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                        }

                    }
                    else
                    {
                        i = values10.Length + 1;
                    }

                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values10[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col10] = Valoare;
                }
            }
            #endregion

            #region 11
            if (colpt11 != "")
            {
                Microsoft.Office.Interop.Excel.Range range11 = W1.Range[colpt11 + Convert.ToString(start_row) + ":" + colpt11 + "300000"];

                values11 = range11.Value2;

                for (int i = 1; i <= values11.Length; ++i)
                {
                    object Valoare11 = values11[i, 1];
                    if (Valoare11 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }
                        }

                    }
                    else
                    {
                        i = values11.Length + 1;
                    }



                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values11[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col11] = Valoare;
                }
            }
            #endregion

            #region 12

            if (colpt12 != "")
            {
                Microsoft.Office.Interop.Excel.Range range12 = W1.Range[colpt12 + Convert.ToString(start_row) + ":" + colpt12 + "300000"];

                values12 = range12.Value2;

                for (int i = 1; i <= values12.Length; ++i)
                {
                    object Valoare12 = values12[i, 1];
                    if (Valoare12 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }
                        }

                    }
                    else
                    {
                        i = values12.Length + 1;
                    }



                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values12[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col12] = Valoare;
                }
            }
            #endregion

            #region 13
            if (colpt13 != "")
            {
                Microsoft.Office.Interop.Excel.Range range13 = W1.Range[colpt13 + Convert.ToString(start_row) + ":" + colpt13 + "300000"];

                values13 = range13.Value2;

                for (int i = 1; i <= values13.Length; ++i)
                {
                    object Valoare13 = values13[i, 1];
                    if (Valoare13 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }
                        }

                    }
                    else
                    {
                        i = values13.Length + 1;
                    }



                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values13[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col13] = Valoare;
                }
            }
            #endregion

            #region 14
            if (colpt14 != "")
            {
                Microsoft.Office.Interop.Excel.Range range14 = W1.Range[colpt14 + Convert.ToString(start_row) + ":" + colpt14 + "300000"];

                values14 = range14.Value2;

                for (int i = 1; i <= values14.Length; ++i)
                {
                    object Valoare14 = values14[i, 1];
                    if (Valoare14 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }

                            if (colpt13 != "")
                            {
                                object Valoare13 = values13[i + 1, 1];
                                if (Valoare13 == null) Valoare13 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col13] = Valoare13;
                            }
                        }

                    }
                    else
                    {
                        i = values14.Length + 1;
                    }



                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values14[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col14] = Valoare;
                }
            }
            #endregion

            #region 15
            if (colpt15 != "")
            {
                Microsoft.Office.Interop.Excel.Range range15 = W1.Range[colpt15 + Convert.ToString(start_row) + ":" + colpt15 + "300000"];

                values15 = range15.Value2;

                for (int i = 1; i <= values15.Length; ++i)
                {
                    object Valoare15 = values15[i, 1];
                    if (Valoare15 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }

                            if (colpt13 != "")
                            {
                                object Valoare13 = values13[i + 1, 1];
                                if (Valoare13 == null) Valoare13 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col13] = Valoare13;
                            }

                            if (colpt14 != "")
                            {
                                object Valoare14 = values14[i + 1, 1];
                                if (Valoare14 == null) Valoare14 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col14] = Valoare14;
                            }
                        }

                    }
                    else
                    {
                        i = values15.Length + 1;
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values15[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col15] = Valoare;
                }
            }
            #endregion

            #region 16
            if (colpt16 != "")
            {
                Microsoft.Office.Interop.Excel.Range range16 = W1.Range[colpt16 + Convert.ToString(start_row) + ":" + colpt16 + "300000"];

                values16 = range16.Value2;

                for (int i = 1; i <= values16.Length; ++i)
                {
                    object Valoare16 = values16[i, 1];
                    if (Valoare16 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }

                            if (colpt13 != "")
                            {
                                object Valoare13 = values13[i + 1, 1];
                                if (Valoare13 == null) Valoare13 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col13] = Valoare13;
                            }

                            if (colpt14 != "")
                            {
                                object Valoare14 = values14[i + 1, 1];
                                if (Valoare14 == null) Valoare14 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col14] = Valoare14;
                            }

                            if (colpt15 != "")
                            {
                                object Valoare15 = values15[i + 1, 1];
                                if (Valoare15 == null) Valoare15 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col15] = Valoare15;
                            }
                        }

                    }
                    else
                    {
                        i = values16.Length + 1;
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values16[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col16] = Valoare;
                }
            }
            #endregion

            #region 17
            if (colpt17 != "")
            {
                Microsoft.Office.Interop.Excel.Range range17 = W1.Range[colpt17 + Convert.ToString(start_row) + ":" + colpt17 + "300000"];

                values17 = range17.Value2;

                for (int i = 1; i <= values17.Length; ++i)
                {
                    object Valoare17 = values17[i, 1];
                    if (Valoare17 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }

                            if (colpt13 != "")
                            {
                                object Valoare13 = values13[i + 1, 1];
                                if (Valoare13 == null) Valoare13 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col13] = Valoare13;
                            }

                            if (colpt14 != "")
                            {
                                object Valoare14 = values14[i + 1, 1];
                                if (Valoare14 == null) Valoare14 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col14] = Valoare14;
                            }

                            if (colpt15 != "")
                            {
                                object Valoare15 = values15[i + 1, 1];
                                if (Valoare15 == null) Valoare15 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col15] = Valoare15;
                            }

                            if (colpt16 != "")
                            {
                                object Valoare16 = values16[i + 1, 1];
                                if (Valoare16 == null) Valoare16 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col16] = Valoare16;
                            }
                        }

                    }
                    else
                    {
                        i = values17.Length + 1;
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values17[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col17] = Valoare;
                }
            }

            #endregion

            #region 18
            if (colpt18 != "")
            {
                Microsoft.Office.Interop.Excel.Range range18 = W1.Range[colpt18 + Convert.ToString(start_row) + ":" + colpt18 + "300000"];

                values18 = range18.Value2;

                for (int i = 1; i <= values18.Length; ++i)
                {
                    object Valoare18 = values18[i, 1];
                    if (Valoare18 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }

                            if (colpt13 != "")
                            {
                                object Valoare13 = values13[i + 1, 1];
                                if (Valoare13 == null) Valoare13 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col13] = Valoare13;
                            }

                            if (colpt14 != "")
                            {
                                object Valoare14 = values14[i + 1, 1];
                                if (Valoare14 == null) Valoare14 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col14] = Valoare14;
                            }

                            if (colpt15 != "")
                            {
                                object Valoare15 = values15[i + 1, 1];
                                if (Valoare15 == null) Valoare15 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col15] = Valoare15;
                            }

                            if (colpt16 != "")
                            {
                                object Valoare16 = values16[i + 1, 1];
                                if (Valoare16 == null) Valoare16 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col16] = Valoare16;
                            }

                            if (colpt17 != "")
                            {
                                object Valoare17 = values17[i + 1, 1];
                                if (Valoare17 == null) Valoare17 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col17] = Valoare17;
                            }

                        }

                    }
                    else
                    {
                        i = values18.Length + 1;
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values18[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col18] = Valoare;
                }
            }
            #endregion

            #region 19
            if (colpt19 != "")
            {
                Microsoft.Office.Interop.Excel.Range range19 = W1.Range[colpt19 + Convert.ToString(start_row) + ":" + colpt19 + "300000"];

                values19 = range19.Value2;

                for (int i = 1; i <= values19.Length; ++i)
                {
                    object Valoare19 = values19[i, 1];
                    if (Valoare19 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }

                            if (colpt13 != "")
                            {
                                object Valoare13 = values13[i + 1, 1];
                                if (Valoare13 == null) Valoare13 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col13] = Valoare13;
                            }

                            if (colpt14 != "")
                            {
                                object Valoare14 = values14[i + 1, 1];
                                if (Valoare14 == null) Valoare14 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col14] = Valoare14;
                            }

                            if (colpt15 != "")
                            {
                                object Valoare15 = values15[i + 1, 1];
                                if (Valoare15 == null) Valoare15 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col15] = Valoare15;
                            }

                            if (colpt16 != "")
                            {
                                object Valoare16 = values16[i + 1, 1];
                                if (Valoare16 == null) Valoare16 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col16] = Valoare16;
                            }

                            if (colpt17 != "")
                            {
                                object Valoare17 = values17[i + 1, 1];
                                if (Valoare17 == null) Valoare17 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col17] = Valoare17;
                            }

                            if (colpt18 != "")
                            {
                                object Valoare18 = values18[i + 1, 1];
                                if (Valoare18 == null) Valoare18 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col18] = Valoare18;
                            }

                        }

                    }
                    else
                    {
                        i = values19.Length + 1;
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values19[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col19] = Valoare;
                }
            }
            #endregion

            #region 20
            if (colpt20 != "")
            {
                Microsoft.Office.Interop.Excel.Range range20 = W1.Range[colpt20 + Convert.ToString(start_row) + ":" + colpt20 + "300000"];

                values20 = range20.Value2;

                for (int i = 1; i <= values20.Length; ++i)
                {
                    object Valoare20 = values20[i, 1];
                    if (Valoare20 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }

                            if (colpt13 != "")
                            {
                                object Valoare13 = values13[i + 1, 1];
                                if (Valoare13 == null) Valoare13 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col13] = Valoare13;
                            }

                            if (colpt14 != "")
                            {
                                object Valoare14 = values14[i + 1, 1];
                                if (Valoare14 == null) Valoare14 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col14] = Valoare14;
                            }

                            if (colpt15 != "")
                            {
                                object Valoare15 = values15[i + 1, 1];
                                if (Valoare15 == null) Valoare15 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col15] = Valoare15;
                            }

                            if (colpt16 != "")
                            {
                                object Valoare16 = values16[i + 1, 1];
                                if (Valoare16 == null) Valoare16 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col16] = Valoare16;
                            }

                            if (colpt17 != "")
                            {
                                object Valoare17 = values17[i + 1, 1];
                                if (Valoare17 == null) Valoare17 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col17] = Valoare17;
                            }

                            if (colpt18 != "")
                            {
                                object Valoare18 = values18[i + 1, 1];
                                if (Valoare18 == null) Valoare18 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col18] = Valoare18;
                            }

                            if (colpt19 != "")
                            {
                                object Valoare19 = values19[i + 1, 1];
                                if (Valoare19 == null) Valoare19 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col19] = Valoare19;
                            }

                        }

                    }
                    else
                    {
                        i = values20.Length + 1;
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values20[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col20] = Valoare;
                }
            }
            #endregion

            #region 21
            if (colpt21 != "")
            {
                Microsoft.Office.Interop.Excel.Range range21 = W1.Range[colpt21 + Convert.ToString(start_row) + ":" + colpt21 + "300000"];

                values21 = range21.Value2;

                for (int i = 1; i <= values21.Length; ++i)
                {
                    object Valoare21 = values21[i, 1];
                    if (Valoare21 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }

                            if (colpt13 != "")
                            {
                                object Valoare13 = values13[i + 1, 1];
                                if (Valoare13 == null) Valoare13 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col13] = Valoare13;
                            }

                            if (colpt14 != "")
                            {
                                object Valoare14 = values14[i + 1, 1];
                                if (Valoare14 == null) Valoare14 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col14] = Valoare14;
                            }

                            if (colpt15 != "")
                            {
                                object Valoare15 = values15[i + 1, 1];
                                if (Valoare15 == null) Valoare15 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col15] = Valoare15;
                            }

                            if (colpt16 != "")
                            {
                                object Valoare16 = values16[i + 1, 1];
                                if (Valoare16 == null) Valoare16 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col16] = Valoare16;
                            }

                            if (colpt17 != "")
                            {
                                object Valoare17 = values17[i + 1, 1];
                                if (Valoare17 == null) Valoare17 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col17] = Valoare17;
                            }

                            if (colpt18 != "")
                            {
                                object Valoare18 = values18[i + 1, 1];
                                if (Valoare18 == null) Valoare18 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col18] = Valoare18;
                            }

                            if (colpt19 != "")
                            {
                                object Valoare19 = values19[i + 1, 1];
                                if (Valoare19 == null) Valoare19 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col19] = Valoare19;
                            }

                            if (colpt20 != "")
                            {
                                object Valoare20 = values20[i + 1, 1];
                                if (Valoare20 == null) Valoare20 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col20] = Valoare20;
                            }

                        }

                    }
                    else
                    {
                        i = values21.Length + 1;
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values21[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col21] = Valoare;
                }
            }
            #endregion

            #region 22
            if (colpt22 != "")
            {
                Microsoft.Office.Interop.Excel.Range range22 = W1.Range[colpt22 + Convert.ToString(start_row) + ":" + colpt22 + "300000"];

                values22 = range22.Value2;

                for (int i = 1; i <= values22.Length; ++i)
                {
                    object Valoare22 = values22[i, 1];
                    if (Valoare22 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }

                            if (colpt13 != "")
                            {
                                object Valoare13 = values13[i + 1, 1];
                                if (Valoare13 == null) Valoare13 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col13] = Valoare13;
                            }

                            if (colpt14 != "")
                            {
                                object Valoare14 = values14[i + 1, 1];
                                if (Valoare14 == null) Valoare14 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col14] = Valoare14;
                            }

                            if (colpt15 != "")
                            {
                                object Valoare15 = values15[i + 1, 1];
                                if (Valoare15 == null) Valoare15 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col15] = Valoare15;
                            }

                            if (colpt16 != "")
                            {
                                object Valoare16 = values16[i + 1, 1];
                                if (Valoare16 == null) Valoare16 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col16] = Valoare16;
                            }

                            if (colpt17 != "")
                            {
                                object Valoare17 = values17[i + 1, 1];
                                if (Valoare17 == null) Valoare17 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col17] = Valoare17;
                            }

                            if (colpt18 != "")
                            {
                                object Valoare18 = values18[i + 1, 1];
                                if (Valoare18 == null) Valoare18 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col18] = Valoare18;
                            }

                            if (colpt19 != "")
                            {
                                object Valoare19 = values19[i + 1, 1];
                                if (Valoare19 == null) Valoare19 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col19] = Valoare19;
                            }

                            if (colpt20 != "")
                            {
                                object Valoare20 = values20[i + 1, 1];
                                if (Valoare20 == null) Valoare20 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col20] = Valoare20;
                            }

                            if (colpt21 != "")
                            {
                                object Valoare21 = values21[i + 1, 1];
                                if (Valoare21 == null) Valoare21 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col21] = Valoare21;
                            }

                        }

                    }
                    else
                    {
                        i = values22.Length + 1;
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values22[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col22] = Valoare;
                }
            }
            #endregion

            #region 23
            if (colpt23 != "")
            {
                Microsoft.Office.Interop.Excel.Range range23 = W1.Range[colpt23 + Convert.ToString(start_row) + ":" + colpt23 + "300000"];

                values23 = range23.Value2;

                for (int i = 1; i <= values23.Length; ++i)
                {
                    object Valoare23 = values23[i, 1];
                    if (Valoare23 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }

                            if (colpt13 != "")
                            {
                                object Valoare13 = values13[i + 1, 1];
                                if (Valoare13 == null) Valoare13 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col13] = Valoare13;
                            }

                            if (colpt14 != "")
                            {
                                object Valoare14 = values14[i + 1, 1];
                                if (Valoare14 == null) Valoare14 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col14] = Valoare14;
                            }

                            if (colpt15 != "")
                            {
                                object Valoare15 = values15[i + 1, 1];
                                if (Valoare15 == null) Valoare15 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col15] = Valoare15;
                            }

                            if (colpt16 != "")
                            {
                                object Valoare16 = values16[i + 1, 1];
                                if (Valoare16 == null) Valoare16 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col16] = Valoare16;
                            }

                            if (colpt17 != "")
                            {
                                object Valoare17 = values17[i + 1, 1];
                                if (Valoare17 == null) Valoare17 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col17] = Valoare17;
                            }

                            if (colpt18 != "")
                            {
                                object Valoare18 = values18[i + 1, 1];
                                if (Valoare18 == null) Valoare18 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col18] = Valoare18;
                            }

                            if (colpt19 != "")
                            {
                                object Valoare19 = values19[i + 1, 1];
                                if (Valoare19 == null) Valoare19 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col19] = Valoare19;
                            }

                            if (colpt20 != "")
                            {
                                object Valoare20 = values20[i + 1, 1];
                                if (Valoare20 == null) Valoare20 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col20] = Valoare20;
                            }

                            if (colpt21 != "")
                            {
                                object Valoare21 = values21[i + 1, 1];
                                if (Valoare21 == null) Valoare21 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col21] = Valoare21;
                            }


                            if (colpt22 != "")
                            {
                                object Valoare22 = values22[i + 1, 1];
                                if (Valoare22 == null) Valoare22 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col22] = Valoare22;
                            }

                        }

                    }
                    else
                    {
                        i = values23.Length + 1;
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values23[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col23] = Valoare;
                }
            }

            #endregion

            #region 24

            if (colpt24 != "")
            {
                Microsoft.Office.Interop.Excel.Range range24 = W1.Range[colpt24 + Convert.ToString(start_row) + ":" + colpt24 + "300000"];

                values24 = range24.Value2;

                for (int i = 1; i <= values24.Length; ++i)
                {
                    object Valoare24 = values24[i, 1];
                    if (Valoare24 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }

                            if (colpt13 != "")
                            {
                                object Valoare13 = values13[i + 1, 1];
                                if (Valoare13 == null) Valoare13 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col13] = Valoare13;
                            }

                            if (colpt14 != "")
                            {
                                object Valoare14 = values14[i + 1, 1];
                                if (Valoare14 == null) Valoare14 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col14] = Valoare14;
                            }

                            if (colpt15 != "")
                            {
                                object Valoare15 = values15[i + 1, 1];
                                if (Valoare15 == null) Valoare15 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col15] = Valoare15;
                            }

                            if (colpt16 != "")
                            {
                                object Valoare16 = values16[i + 1, 1];
                                if (Valoare16 == null) Valoare16 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col16] = Valoare16;
                            }

                            if (colpt17 != "")
                            {
                                object Valoare17 = values17[i + 1, 1];
                                if (Valoare17 == null) Valoare17 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col17] = Valoare17;
                            }

                            if (colpt18 != "")
                            {
                                object Valoare18 = values18[i + 1, 1];
                                if (Valoare18 == null) Valoare18 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col18] = Valoare18;
                            }

                            if (colpt19 != "")
                            {
                                object Valoare19 = values19[i + 1, 1];
                                if (Valoare19 == null) Valoare19 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col19] = Valoare19;
                            }

                            if (colpt20 != "")
                            {
                                object Valoare20 = values20[i + 1, 1];
                                if (Valoare20 == null) Valoare20 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col20] = Valoare20;
                            }

                            if (colpt21 != "")
                            {
                                object Valoare21 = values21[i + 1, 1];
                                if (Valoare21 == null) Valoare21 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col21] = Valoare21;
                            }


                            if (colpt22 != "")
                            {
                                object Valoare22 = values22[i + 1, 1];
                                if (Valoare22 == null) Valoare22 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col22] = Valoare22;
                            }

                            if (colpt23 != "")
                            {
                                object Valoare23 = values23[i + 1, 1];
                                if (Valoare23 == null) Valoare23 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col23] = Valoare23;
                            }

                        }

                    }
                    else
                    {
                        i = values24.Length + 1;
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values24[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col24] = Valoare;
                }
            }

            #endregion

            #region 25
            if (colpt25 != "")
            {
                Microsoft.Office.Interop.Excel.Range range25 = W1.Range[colpt25 + Convert.ToString(start_row) + ":" + colpt25 + "300000"];

                values25 = range25.Value2;

                for (int i = 1; i <= values25.Length; ++i)
                {
                    object Valoare25 = values25[i, 1];
                    if (Valoare25 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }

                            if (colpt13 != "")
                            {
                                object Valoare13 = values13[i + 1, 1];
                                if (Valoare13 == null) Valoare13 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col13] = Valoare13;
                            }

                            if (colpt14 != "")
                            {
                                object Valoare14 = values14[i + 1, 1];
                                if (Valoare14 == null) Valoare14 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col14] = Valoare14;
                            }

                            if (colpt15 != "")
                            {
                                object Valoare15 = values15[i + 1, 1];
                                if (Valoare15 == null) Valoare15 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col15] = Valoare15;
                            }

                            if (colpt16 != "")
                            {
                                object Valoare16 = values16[i + 1, 1];
                                if (Valoare16 == null) Valoare16 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col16] = Valoare16;
                            }

                            if (colpt17 != "")
                            {
                                object Valoare17 = values17[i + 1, 1];
                                if (Valoare17 == null) Valoare17 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col17] = Valoare17;
                            }

                            if (colpt18 != "")
                            {
                                object Valoare18 = values18[i + 1, 1];
                                if (Valoare18 == null) Valoare18 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col18] = Valoare18;
                            }

                            if (colpt19 != "")
                            {
                                object Valoare19 = values19[i + 1, 1];
                                if (Valoare19 == null) Valoare19 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col19] = Valoare19;
                            }

                            if (colpt20 != "")
                            {
                                object Valoare20 = values20[i + 1, 1];
                                if (Valoare20 == null) Valoare20 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col20] = Valoare20;
                            }

                            if (colpt21 != "")
                            {
                                object Valoare21 = values21[i + 1, 1];
                                if (Valoare21 == null) Valoare21 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col21] = Valoare21;
                            }


                            if (colpt22 != "")
                            {
                                object Valoare22 = values22[i + 1, 1];
                                if (Valoare22 == null) Valoare22 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col22] = Valoare22;
                            }

                            if (colpt23 != "")
                            {
                                object Valoare23 = values23[i + 1, 1];
                                if (Valoare23 == null) Valoare23 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col23] = Valoare23;
                            }

                            if (colpt24 != "")
                            {
                                object Valoare24 = values24[i + 1, 1];
                                if (Valoare24 == null) Valoare24 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col24] = Valoare24;
                            }

                        }

                    }
                    else
                    {
                        i = values25.Length + 1;
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values25[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col25] = Valoare;
                }
            }
            #endregion

            #region 26
            if (colpt26 != "")
            {
                Microsoft.Office.Interop.Excel.Range range26 = W1.Range[colpt26 + Convert.ToString(start_row) + ":" + colpt26 + "300000"];

                values26 = range26.Value2;

                for (int i = 1; i <= values26.Length; ++i)
                {
                    object Valoare26 = values26[i, 1];
                    if (Valoare26 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }

                            if (colpt13 != "")
                            {
                                object Valoare13 = values13[i + 1, 1];
                                if (Valoare13 == null) Valoare13 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col13] = Valoare13;
                            }

                            if (colpt14 != "")
                            {
                                object Valoare14 = values14[i + 1, 1];
                                if (Valoare14 == null) Valoare14 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col14] = Valoare14;
                            }

                            if (colpt15 != "")
                            {
                                object Valoare15 = values15[i + 1, 1];
                                if (Valoare15 == null) Valoare15 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col15] = Valoare15;
                            }

                            if (colpt16 != "")
                            {
                                object Valoare16 = values16[i + 1, 1];
                                if (Valoare16 == null) Valoare16 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col16] = Valoare16;
                            }

                            if (colpt17 != "")
                            {
                                object Valoare17 = values17[i + 1, 1];
                                if (Valoare17 == null) Valoare17 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col17] = Valoare17;
                            }

                            if (colpt18 != "")
                            {
                                object Valoare18 = values18[i + 1, 1];
                                if (Valoare18 == null) Valoare18 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col18] = Valoare18;
                            }

                            if (colpt19 != "")
                            {
                                object Valoare19 = values19[i + 1, 1];
                                if (Valoare19 == null) Valoare19 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col19] = Valoare19;
                            }

                            if (colpt20 != "")
                            {
                                object Valoare20 = values20[i + 1, 1];
                                if (Valoare20 == null) Valoare20 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col20] = Valoare20;
                            }

                            if (colpt21 != "")
                            {
                                object Valoare21 = values21[i + 1, 1];
                                if (Valoare21 == null) Valoare21 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col21] = Valoare21;
                            }


                            if (colpt22 != "")
                            {
                                object Valoare22 = values22[i + 1, 1];
                                if (Valoare22 == null) Valoare22 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col22] = Valoare22;
                            }

                            if (colpt23 != "")
                            {
                                object Valoare23 = values23[i + 1, 1];
                                if (Valoare23 == null) Valoare23 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col23] = Valoare23;
                            }

                            if (colpt24 != "")
                            {
                                object Valoare24 = values24[i + 1, 1];
                                if (Valoare24 == null) Valoare24 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col24] = Valoare24;
                            }


                            if (colpt25 != "")
                            {
                                object Valoare25 = values25[i + 1, 1];
                                if (Valoare25 == null) Valoare25 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col25] = Valoare25;
                            }

                        }

                    }
                    else
                    {
                        i = values26.Length + 1;
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values26[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col26] = Valoare;
                }
            }
            #endregion

            #region 27
            if (colpt27 != "")
            {
                Microsoft.Office.Interop.Excel.Range range27 = W1.Range[colpt27 + Convert.ToString(start_row) + ":" + colpt27 + "300000"];

                values27 = range27.Value2;

                for (int i = 1; i <= values27.Length; ++i)
                {
                    object Valoare27 = values27[i, 1];
                    if (Valoare27 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }

                            if (colpt13 != "")
                            {
                                object Valoare13 = values13[i + 1, 1];
                                if (Valoare13 == null) Valoare13 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col13] = Valoare13;
                            }

                            if (colpt14 != "")
                            {
                                object Valoare14 = values14[i + 1, 1];
                                if (Valoare14 == null) Valoare14 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col14] = Valoare14;
                            }

                            if (colpt15 != "")
                            {
                                object Valoare15 = values15[i + 1, 1];
                                if (Valoare15 == null) Valoare15 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col15] = Valoare15;
                            }

                            if (colpt16 != "")
                            {
                                object Valoare16 = values16[i + 1, 1];
                                if (Valoare16 == null) Valoare16 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col16] = Valoare16;
                            }

                            if (colpt17 != "")
                            {
                                object Valoare17 = values17[i + 1, 1];
                                if (Valoare17 == null) Valoare17 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col17] = Valoare17;
                            }

                            if (colpt18 != "")
                            {
                                object Valoare18 = values18[i + 1, 1];
                                if (Valoare18 == null) Valoare18 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col18] = Valoare18;
                            }

                            if (colpt19 != "")
                            {
                                object Valoare19 = values19[i + 1, 1];
                                if (Valoare19 == null) Valoare19 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col19] = Valoare19;
                            }

                            if (colpt20 != "")
                            {
                                object Valoare20 = values20[i + 1, 1];
                                if (Valoare20 == null) Valoare20 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col20] = Valoare20;
                            }

                            if (colpt21 != "")
                            {
                                object Valoare21 = values21[i + 1, 1];
                                if (Valoare21 == null) Valoare21 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col21] = Valoare21;
                            }


                            if (colpt22 != "")
                            {
                                object Valoare22 = values22[i + 1, 1];
                                if (Valoare22 == null) Valoare22 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col22] = Valoare22;
                            }

                            if (colpt23 != "")
                            {
                                object Valoare23 = values23[i + 1, 1];
                                if (Valoare23 == null) Valoare23 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col23] = Valoare23;
                            }

                            if (colpt24 != "")
                            {
                                object Valoare24 = values24[i + 1, 1];
                                if (Valoare24 == null) Valoare24 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col24] = Valoare24;
                            }


                            if (colpt25 != "")
                            {
                                object Valoare25 = values25[i + 1, 1];
                                if (Valoare25 == null) Valoare25 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col25] = Valoare25;
                            }

                            if (colpt26 != "")
                            {
                                object Valoare26 = values26[i + 1, 1];
                                if (Valoare26 == null) Valoare26 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col26] = Valoare26;
                            }

                        }

                    }
                    else
                    {
                        i = values27.Length + 1;
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values27[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col27] = Valoare;
                }
            }
            #endregion

            return dt1;
        }

        public static void Transfer_weldmap_datatable_to_new_excel_spreadsheet_formated_general_and_colored(System.Data.DataTable dt1, System.Data.DataTable dt2)
        {
            if (dt1 != null)
            {
                if (dt1.Rows.Count > 0)
                {
                    Microsoft.Office.Interop.Excel.Worksheet W1 = Get_NEW_worksheet_from_Excel();
                    W1.Cells.NumberFormat = "General";
                    int maxRows = dt1.Rows.Count;
                    int maxCols = dt1.Columns.Count;
                    Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[2, 1], W1.Cells[maxRows + 1, maxCols]];
                    object[,] values1 = new object[maxRows, maxCols];

                    for (int i = 0; i < maxRows; ++i)
                    {
                        for (int j = 0; j < maxCols; ++j)
                        {
                            if (dt1.Rows[i][j] != DBNull.Value)
                            {
                                values1[i, j] = Convert.ToString(dt1.Rows[i][j]);
                            }
                        }
                    }

                    for (int i = 0; i < dt1.Columns.Count; ++i)
                    {
                        W1.Cells[1, i + 1].value2 = dt1.Columns[i].ColumnName;
                    }
                    range1.Value2 = values1;

                    W1.Range["A:A"].Font.Bold = true;
                    W1.Range["H:H"].Font.Bold = true;
                    W1.Range["M:M"].Font.Bold = true;
                    if (dt2 != null && dt2.Rows.Count > 0)
                    {
                        for (int i = 0; i < maxRows; ++i)
                        {
                            if (dt2.Rows[i][0] != DBNull.Value)
                            {
                                int color1 = Convert.ToInt32(dt2.Rows[i][0]);
                                W1.Rows[2 + i].Interior.Color = color1;
                            }
                            if (dt2.Rows[i][1] != DBNull.Value)
                            {
                                int cidx = Convert.ToInt32(dt2.Rows[i][1]);
                                W1.Rows[2 + i].Interior.ColorIndex = cidx;
                            }
                            if (dt2.Rows[i][2] != DBNull.Value)
                            {
                                int thc = Convert.ToInt32(dt2.Rows[i][2]);
                                W1.Rows[2 + i].Interior.ThemeColor = thc;
                            }
                            if (dt2.Rows[i][3] != DBNull.Value)
                            {
                                double tint1 = Convert.ToDouble(dt2.Rows[i][3]);
                                W1.Rows[2 + i].Interior.PatternTintAndShade = tint1;
                            }
                        }
                    }
                    W1.Range["A:A"].ColumnWidth = 7.29;
                    W1.Range["B:D"].ColumnWidth = 0;
                    W1.Range["E:E"].ColumnWidth = 22.86;
                    W1.Range["F:F"].ColumnWidth = 24.57;
                    W1.Range["G:G"].ColumnWidth = 16.86;
                    W1.Name = "WELD_MAP";
                }
            }
        }

        static public System.Data.DataTable build_data_table_from_excel_based_on_columns_with_type_check
        (
            System.Data.DataTable dt1,
            Microsoft.Office.Interop.Excel.Worksheet W1, int start_row, int end_row,
            string col1, string colxl1, string col2, string colxl2, string col3, string colxl3, string col4, string colxl4, string col5, string colxl5, string col6, string colxl6
        )

        {
            if (W1 == null) return dt1;
            if (end_row - start_row < 0) return dt1;

            object[,] values1 = new object[end_row - start_row + 1, 1];
            object[,] values2 = new object[end_row - start_row + 1, 1];
            object[,] values3 = new object[end_row - start_row + 1, 1];
            object[,] values4 = new object[end_row - start_row + 1, 1];
            object[,] values5 = new object[end_row - start_row + 1, 1];
            object[,] values6 = new object[end_row - start_row + 1, 1];


            #region 1
            if (colxl1 != "")
            {
                Microsoft.Office.Interop.Excel.Range range1 = W1.Range[colxl1 + start_row.ToString() + ":" + colxl1 + end_row.ToString()];
                values1 = range1.Value2;

                for (int i = 1; i <= values1.Length; ++i)
                {
                    object Valoare1 = values1[i, 1];
                    dt1.Rows.Add();

                    if (Valoare1 == null) Valoare1 = DBNull.Value;

                    if (Valoare1 != null && dt1.Columns[col1].GetType() == typeof(double))
                    {
                        Valoare1 = Valoare1.ToString().Replace("+", "");
                        if (IsNumeric(Valoare1.ToString()) == true)
                        {
                            Valoare1 = Convert.ToDouble(Valoare1);
                        }
                        else
                        {
                            Valoare1 = DBNull.Value;
                        }
                    }

                    dt1.Rows[i - 1][col1] = Valoare1;
                }
            }
            #endregion

            #region 2
            if (colxl2 != "")
            {
                Microsoft.Office.Interop.Excel.Range range2 = W1.Range[colxl2 + start_row.ToString() + ":" + colxl2 + end_row.ToString()];
                values2 = range2.Value2;

                for (int i = 1; i <= values2.Length; ++i)
                {
                    object Valoare2 = values2[i, 1];
                    if (colxl1 == "") dt1.Rows.Add();

                    if (Valoare2 == null) Valoare2 = DBNull.Value;

                    if (Valoare2 != null && dt1.Columns[col2].GetType() == typeof(double))
                    {
                        Valoare2 = Valoare2.ToString().Replace("+", "");
                        if (IsNumeric(Valoare2.ToString()) == true)
                        {
                            Valoare2 = Convert.ToDouble(Valoare2);
                        }
                        else
                        {
                            Valoare2 = DBNull.Value;
                        }
                    }

                    dt1.Rows[i - 1][col2] = Valoare2;
                }
            }
            #endregion

            #region 3
            if (colxl3 != "")
            {
                Microsoft.Office.Interop.Excel.Range range3 = W1.Range[colxl3 + start_row.ToString() + ":" + colxl3 + end_row.ToString()];
                values3 = range3.Value2;

                for (int i = 1; i <= values3.Length; ++i)
                {
                    object Valoare3 = values3[i, 1];
                    if (colxl1 == "" && colxl2 == "") dt1.Rows.Add();

                    if (Valoare3 == null) Valoare3 = DBNull.Value;

                    if (Valoare3 != null && dt1.Columns[col3].GetType() == typeof(double))
                    {
                        Valoare3 = Valoare3.ToString().Replace("+", "");
                        if (IsNumeric(Valoare3.ToString()) == true)
                        {
                            Valoare3 = Convert.ToDouble(Valoare3);
                        }
                        else
                        {
                            Valoare3 = DBNull.Value;
                        }
                    }

                    dt1.Rows[i - 1][col3] = Valoare3;
                }
            }
            #endregion

            #region 4
            if (colxl4 != "")
            {
                Microsoft.Office.Interop.Excel.Range range4 = W1.Range[colxl4 + start_row.ToString() + ":" + colxl4 + end_row.ToString()];
                values4 = range4.Value2;

                for (int i = 1; i <= values4.Length; ++i)
                {
                    object Valoare4 = values4[i, 1];
                    if (colxl1 == "" && colxl2 == "" && colxl3 == "") dt1.Rows.Add();

                    if (Valoare4 == null) Valoare4 = DBNull.Value;

                    if (Valoare4 != null && dt1.Columns[col4].GetType() == typeof(double))
                    {
                        Valoare4 = Valoare4.ToString().Replace("+", "");
                        if (IsNumeric(Valoare4.ToString()) == true)
                        {
                            Valoare4 = Convert.ToDouble(Valoare4);
                        }
                        else
                        {
                            Valoare4 = DBNull.Value;
                        }
                    }

                    dt1.Rows[i - 1][col4] = Valoare4;
                }
            }
            #endregion

            #region 5
            if (colxl5 != "")
            {
                Microsoft.Office.Interop.Excel.Range range5 = W1.Range[colxl5 + start_row.ToString() + ":" + colxl5 + end_row.ToString()];
                values5 = range5.Value2;

                for (int i = 1; i <= values5.Length; ++i)
                {
                    object Valoare5 = values5[i, 1];
                    if (colxl1 == "" && colxl2 == "" && colxl3 == "") dt1.Rows.Add();

                    if (Valoare5 == null) Valoare5 = DBNull.Value;

                    if (Valoare5 != null && dt1.Columns[col5].GetType() == typeof(double))
                    {
                        Valoare5 = Valoare5.ToString().Replace("+", "");
                        if (IsNumeric(Valoare5.ToString()) == true)
                        {
                            Valoare5 = Convert.ToDouble(Valoare5);
                        }
                        else
                        {
                            Valoare5 = DBNull.Value;
                        }
                    }

                    dt1.Rows[i - 1][col5] = Valoare5;
                }
            }
            #endregion

            #region 6
            if (colxl6 != "")
            {
                Microsoft.Office.Interop.Excel.Range range6 = W1.Range[colxl6 + start_row.ToString() + ":" + colxl6 + end_row.ToString()];
                values6 = range6.Value2;

                for (int i = 1; i <= values6.Length; ++i)
                {
                    object Valoare6 = values6[i, 1];
                    if (colxl1 == "" && colxl2 == "" && colxl3 == "") dt1.Rows.Add();

                    if (Valoare6 == null) Valoare6 = DBNull.Value;

                    if (Valoare6 != null && dt1.Columns[col6].GetType() == typeof(double))
                    {
                        Valoare6 = Valoare6.ToString().Replace("+", "");
                        if (IsNumeric(Valoare6.ToString()) == true)
                        {
                            Valoare6 = Convert.ToDouble(Valoare6);
                        }
                        else
                        {
                            Valoare6 = DBNull.Value;
                        }
                    }

                    dt1.Rows[i - 1][col6] = Valoare6;
                }
            }
            #endregion

            return dt1;
        }
        public static double get_stationCSF_from_point(Polyline Poly2D, Point3d Pt1, double meas_2D_at_pt1, System.Data.DataTable Data_table_centerline)
        {
            double station = -1.23456;
            string Col_3DSta = "3DSta";
            string Col_back = "BackSta";
            string Col_ahead = "AheadSta";

            if (meas_2D_at_pt1 >= Poly2D.Length) meas_2D_at_pt1 = Poly2D.Length - 0.0001;

            double param1 = Poly2D.GetParameterAtDistance(meas_2D_at_pt1);
            int index1 = Convert.ToInt32(Math.Floor(param1));
            if (index1 == Poly2D.EndParam)
            {
                index1 = Data_table_centerline.Rows.Count - 2;
            }

            Point2d pt01 = Poly2D.GetPoint2dAt(index1);
            Point2d pt11 = Poly2D.GetPoint2dAt(index1 + 1);

            double d_l1 = pt01.GetDistanceTo(pt11);

            if (
                (
                    Data_table_centerline.Rows[index1][Col_3DSta] != DBNull.Value ||
                    Data_table_centerline.Rows[index1][Col_ahead] != DBNull.Value
                ) &&
                (
                    Functions.IsNumeric(Convert.ToString(Data_table_centerline.Rows[index1][Col_3DSta])) == true ||
                    Functions.IsNumeric(Convert.ToString(Data_table_centerline.Rows[index1][Col_ahead])) == true
                ) &&
                (
                    Data_table_centerline.Rows[index1 + 1][Col_3DSta] != DBNull.Value ||
                    Data_table_centerline.Rows[index1 + 1][Col_back] != DBNull.Value
                ) &&
                (
                    Functions.IsNumeric(Convert.ToString(Data_table_centerline.Rows[index1 + 1][Col_3DSta])) == true ||
                    Functions.IsNumeric(Convert.ToString(Data_table_centerline.Rows[index1 + 1][Col_back])) == true
                )
             )
            {

                double sta01 = -1.23456;

                if (Data_table_centerline.Rows[index1][Col_3DSta] != DBNull.Value &&
                   Functions.IsNumeric(Convert.ToString(Data_table_centerline.Rows[index1][Col_3DSta])) == true)
                {
                    sta01 = Convert.ToDouble(Data_table_centerline.Rows[index1][Col_3DSta]);
                }

                if (Data_table_centerline.Rows[index1][Col_ahead] != DBNull.Value &&
                    Functions.IsNumeric(Convert.ToString(Data_table_centerline.Rows[index1][Col_ahead])) == true)
                {
                    sta01 = Convert.ToDouble(Data_table_centerline.Rows[index1][Col_ahead]);
                }

                double sta11 = -1.23456;

                if (Data_table_centerline.Rows[index1 + 1][Col_3DSta] != DBNull.Value &&
                   Functions.IsNumeric(Convert.ToString(Data_table_centerline.Rows[index1 + 1][Col_3DSta])) == true)
                {
                    sta11 = Convert.ToDouble(Data_table_centerline.Rows[index1 + 1][Col_3DSta]);
                }

                if (Data_table_centerline.Rows[index1 + 1][Col_back] != DBNull.Value &&
                    Functions.IsNumeric(Convert.ToString(Data_table_centerline.Rows[index1 + 1][Col_back])) == true)
                {
                    sta11 = Convert.ToDouble(Data_table_centerline.Rows[index1 + 1][Col_back]);
                }

                double factor1 = (sta11 - sta01) / d_l1;

                Polyline poly01 = new Polyline();
                poly01.AddVertexAt(0, pt01, 0, 0, 0);
                poly01.AddVertexAt(1, pt11, 0, 0, 0);
                double d01 = poly01.GetDistAtPoint(poly01.GetClosestPointTo(new Point3d(Pt1.X, Pt1.Y, 0), Vector3d.ZAxis, false));

                station = sta01 + d01 * factor1;
            }

            return station;
        }

        private static void Create_profile_od_table()
        {
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        List<string> List1 = new List<string>();
                        List<string> List2 = new List<string>();
                        List<Autodesk.Gis.Map.Constants.DataType> List3 = new List<Autodesk.Gis.Map.Constants.DataType>();

                        List1.Add("Version");
                        List2.Add("Version");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add("BeginSta");
                        List2.Add("Profile start");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                        List1.Add("Segment");
                        List2.Add("Segment");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add("Note1");
                        List2.Add("Notes");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        Functions.Get_object_data_table("Agen_profile", "Generated by AGEN", List1, List2, List3);

                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public static ObjectId Get_Arrow_dimension_ID(string NUME_variabila, string NUME_ARROW)
        {

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            string OLD_VALUE = (string)Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable(NUME_variabila);
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable(NUME_variabila, NUME_ARROW);
            if (OLD_VALUE.Length != 0) Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable(NUME_variabila, OLD_VALUE);
            ObjectId ID1;
            using (Transaction Trans1 = ThisDrawing.Database.TransactionManager.StartTransaction())
            {
                BlockTable BlockTable1 = (BlockTable)Trans1.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                ID1 = BlockTable1[NUME_ARROW];
                Trans1.Commit();


            }

            return ID1;

        }

    }



}
