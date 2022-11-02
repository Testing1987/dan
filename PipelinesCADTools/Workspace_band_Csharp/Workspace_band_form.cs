using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;


using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;


namespace Workspace_band_Csharp
{
    public partial class Workspace_band_form : Form
    {
        public Workspace_band_form()
        {
            InitializeComponent();
        }


        System.Data.DataTable ROW_DATA_TABLE;
        System.Data.DataTable WS_DATA_TABLE;
        System.Data.DataTable ATWS_DATA_TABLE;
        System.Data.DataTable PPE_DATA_TABLE;

        System.Data.DataTable EPE_L_DATA_TABLE;
        System.Data.DataTable EPE_R_DATA_TABLE;
        System.Data.DataTable TWS_L_ON_EPE_TABLE;
        System.Data.DataTable TWS_R_ON_EPE_TABLE;

        System.Data.DataTable ATWS_L_ON_EPE_TABLE;
        System.Data.DataTable ATWS_R_ON_EPE_TABLE;
        System.Data.DataTable ATWS_L_ON_TWS_TABLE;
        System.Data.DataTable ATWS_R_ON_TWS_TABLE;
        System.Data.DataTable Data_table_easement_left;
        System.Data.DataTable Data_table_easement_right;

        Polyline EPE_L;
        Polyline EPE_R;
        Polyline TWS_L;
        Polyline TWS_R;

        Polyline PolyCL_MS;

        System.Data.DataTable Data_table_easement;
        System.Data.DataTable Compiled_DATA_TABLE;

        Boolean Freeze_operations = false;
        int Rounding_no = 0;
        int number_of_workspaces = 0;
        double Sta_max = 0;

        private void Button_connect_to_access_DB_Click(object sender, EventArgs e)
        {
            try
            {
                string Table1 = textBox_row_table_name.Text;
                string query = "SELECT * FROM " + Table1;
                string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + textBox_access_database_location.Text + ";";
                DataSet ROW_DATASET = new DataSet();
                OleDbConnection cnn;
                try
                {
                    cnn = new OleDbConnection(ConnectionString);
                    cnn.Open();
                    OleDbCommand cmd = new OleDbCommand(query, cnn);
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(ROW_DATASET, Table1);
                    cnn.Close();
                }
                catch (OleDbException ex)
                {
                    MessageBox.Show(ex.Message);
                }
                ROW_DATA_TABLE = new System.Data.DataTable();

                ROW_DATA_TABLE = ROW_DATASET.Tables[Table1];

                //MessageBox.Show(ROW_DATA_TABLE.Columns[4].ColumnName + " Row 2 = " + ROW_DATA_TABLE.Rows[1][4]);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        private void Button_connect_to_wspaceDB_Click(object sender, EventArgs e)
        {
            try
            {
                string Table1 = textBox_WORKSPACE_TABLE_NAME.Text;
                string query = "SELECT * FROM " + Table1;
                string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + textBox_access_database_location.Text + ";";
                DataSet WS_DATASET = new DataSet();
                OleDbConnection cnn;
                try
                {
                    cnn = new OleDbConnection(ConnectionString);
                    cnn.Open();
                    OleDbCommand cmd = new OleDbCommand(query, cnn);
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(WS_DATASET, Table1);
                    cnn.Close();
                }
                catch (OleDbException ex)
                {
                    MessageBox.Show(ex.Message);
                }
                WS_DATA_TABLE = new System.Data.DataTable();

                WS_DATA_TABLE = WS_DATASET.Tables[Table1];

                //MessageBox.Show(WS_DATA_TABLE.Columns[1].ColumnName + " Row 1 = " + WS_DATA_TABLE.Rows[0][1]);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Button_connect_to_atws_DB_Click(object sender, EventArgs e)
        {
            try
            {
                string Table1 = textBox_ATWS_table_name.Text;
                string query = "SELECT * FROM " + Table1;
                string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + textBox_access_database_location.Text + ";";
                DataSet ATWS_DATASET = new DataSet();
                OleDbConnection cnn;
                try
                {
                    cnn = new OleDbConnection(ConnectionString);
                    cnn.Open();
                    OleDbCommand cmd = new OleDbCommand(query, cnn);
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(ATWS_DATASET, Table1);
                    cnn.Close();
                }
                catch (OleDbException ex)
                {
                    MessageBox.Show(ex.Message);
                }
                ATWS_DATA_TABLE = new System.Data.DataTable();

                ATWS_DATA_TABLE = ATWS_DATASET.Tables[Table1];

                //MessageBox.Show(ROW_DATA_TABLE.Columns[4].ColumnName + " Row 2 = " + ROW_DATA_TABLE.Rows[1][4]);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }


        private void Button_draw_click(object sender, EventArgs e)
        {
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();



                if (PolyCL_MS == null)
                {

                    //ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                    //return;
                }


                string Column_row_config_in_WS_schemathic = "RID";
                string Column_row_config_in_ROW_CONFIG = "RID";
                string Column_WS_Begin = "BEG_CONFIG";
                string Column_WS_End = "END_CONFIG";
                string Column_PE_L = "PPE_L";
                string Column_PE_R = "PPE_R";
                string Column_TWS_L = "TWS_L";
                string Column_TWS_R = "TWS_R";
                string Column_ATWS_BEGIN = "BEGIN";
                string Column_ATWS_END = "END";
                string Column_ATWS_WIDTH = "WIDTH";
                string Column_ATWS_DIRECTION = "DIRECTION";
                string Symbol_Left = "L";






                Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                    if (WS_DATA_TABLE != null & ROW_DATA_TABLE != null)
                    {
                        if (WS_DATA_TABLE.Rows.Count > 0 & ROW_DATA_TABLE.Rows.Count > 0)
                        {

                            Double Match1 = Convert.ToDouble(textBox_Matchline_start.Text);
                            Double Match2 = Convert.ToDouble(textBox_Matchline_end.Text);
                            if (Match1 > Match2)
                            {
                                double T = Match1;
                                Match1 = Match2;
                                Match2 = T;
                            }

                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                            Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                            PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the starting point");
                            PP1.AllowNone = false;
                            Point_res1 = Editor1.GetPoint(PP1);

                            if (Point_res1.Status != PromptStatus.OK)
                            {

                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            double Len1 = Match2 - Match1;// PolyCL_MS.Length;

                            Polyline PolyCL_PS = new Polyline();
                            PolyCL_PS.AddVertexAt(0, new Point2d(Point_res1.Value.X, Point_res1.Value.Y), 0, 0, 0);
                            PolyCL_PS.AddVertexAt(1, new Point2d(Point_res1.Value.X + Len1, Point_res1.Value.Y), 0, 0, 0);
                            PolyCL_PS.ColorIndex = 1;
                            BTrecord.AppendEntity(PolyCL_PS);
                            Trans1.AddNewlyCreatedDBObject(PolyCL_PS, true);

                            Point3dCollection Colectie_puncte_PPE_L = new Point3dCollection();
                            Point3dCollection Colectie_puncte_PPE_R = new Point3dCollection();
                            Point3dCollection Colectie_stanga1 = new Point3dCollection();
                            Point3dCollection Colectie_dreapta1 = new Point3dCollection();

                            Point3dCollection[] Colectie_puncte_TWS_L;
                            int NrL = 0;
                            Colectie_puncte_TWS_L = new Point3dCollection[NrL + 1];
                            Colectie_puncte_TWS_L[NrL] = new Point3dCollection();

                            Point3dCollection[] Colectie_puncte_TWS_R;
                            int NrR = 0;
                            Colectie_puncte_TWS_R = new Point3dCollection[NrR + 1];
                            Colectie_puncte_TWS_R[NrR] = new Point3dCollection();

                            for (int i = 0; i < WS_DATA_TABLE.Rows.Count; ++i)
                            {
                                Double Start1 = -1;
                                Double End1 = -1;
                                string Config_No = "";

                                if ((WS_DATA_TABLE.Rows[i][Column_WS_Begin] != System.DBNull.Value)) Start1 = Convert.ToDouble(WS_DATA_TABLE.Rows[i][Column_WS_Begin]);
                                if ((WS_DATA_TABLE.Rows[i][Column_WS_End] != System.DBNull.Value)) End1 = Convert.ToDouble(WS_DATA_TABLE.Rows[i][Column_WS_End]);
                                if ((WS_DATA_TABLE.Rows[i][Column_row_config_in_WS_schemathic] != System.DBNull.Value)) Config_No = Convert.ToString(WS_DATA_TABLE.Rows[i][Column_row_config_in_WS_schemathic]);

                                Double TWS_L, PPE_L, PPE_R, TWS_R;
                                PPE_L = 0;
                                PPE_R = 0;
                                TWS_L = 0;
                                TWS_R = 0;

                                if (Start1 != -1 & End1 != -1 & Config_No != "")
                                {

                                    for (int j = 0; j < ROW_DATA_TABLE.Rows.Count; ++j)
                                    {
                                        if (Convert.ToString(ROW_DATA_TABLE.Rows[j][Column_row_config_in_ROW_CONFIG]) == Config_No)
                                        {
                                            if ((ROW_DATA_TABLE.Rows[j][Column_PE_L] != System.DBNull.Value))
                                            {
                                                PPE_L = Convert.ToDouble((ROW_DATA_TABLE.Rows[j][Column_PE_L]));

                                            }
                                            if ((ROW_DATA_TABLE.Rows[j][Column_PE_R] != System.DBNull.Value))
                                            {
                                                PPE_R = Convert.ToDouble((ROW_DATA_TABLE.Rows[j][Column_PE_R]));

                                            }
                                            if ((ROW_DATA_TABLE.Rows[j][Column_TWS_L] != System.DBNull.Value))
                                            {
                                                TWS_L = Convert.ToDouble((ROW_DATA_TABLE.Rows[j][Column_TWS_L]));

                                            }
                                            if ((ROW_DATA_TABLE.Rows[j][Column_TWS_R] != System.DBNull.Value))
                                            {
                                                TWS_R = Convert.ToDouble((ROW_DATA_TABLE.Rows[j][Column_TWS_R]));
                                            }


                                            j = ROW_DATA_TABLE.Rows.Count;
                                        }
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Non valid workspace config on line no. " + i);
                                    return;
                                }

                                if (Start1 <= Match1 & End1 >= Match2)
                                {
                                    Start1 = Match1;
                                    End1 = Match2;
                                }
                                if (Start1 <= Match1 & End1 <= Match2)
                                {
                                    Start1 = Match1;
                                }

                                if (Start1 >= Match1 & End1 >= Match2)
                                {
                                    End1 = Match2;
                                }


                                if (Start1 >= Match1 & Start1 <= Match2 & End1 >= Match1 & End1 <= Match2)
                                {

                                    Point3d Point_start1 = new Point3d();
                                    Point_start1 = PolyCL_PS.GetPointAtDist(Start1);
                                    Point3d Point_end1 = new Point3d();
                                    Point_end1 = PolyCL_PS.GetPointAtDist(End1);


                                    Polyline Poly_start_end = new Polyline();
                                    Poly_start_end.AddVertexAt(0, new Point2d(Point_start1.X, Point_start1.Y), 0, 0, 0);
                                    Poly_start_end.AddVertexAt(1, new Point2d(Point_end1.X, Point_end1.Y), 0, 0, 0);

                                    //MessageBox.Show("PPE_L=" + Convert.ToString(PPE_L) + "\r\n" + "PPE_R=" + Convert.ToString(PPE_R) + "\r\n" + "TWS_L=" + Convert.ToString(TWS_L) + "\r\n" + "TWS_R=" + Convert.ToString(TWS_R));

                                    if ((PPE_L != 0) & (PPE_R != 0))
                                    {

                                        Autodesk.AutoCAD.DatabaseServices.DBObjectCollection Colectie_offset_PPE_L = Poly_start_end.GetOffsetCurves(-PPE_L);

                                        //MessageBox.Show(Convert.ToString(Colectie_offset_PPE_L.Count));

                                        Polyline Poly_PPE_L = new Polyline();
                                        foreach (Polyline obj in Colectie_offset_PPE_L)
                                        {
                                            if (obj != null)
                                            {
                                                Poly_PPE_L = obj;

                                            }
                                        }
                                        if (Poly_PPE_L != null)
                                        {
                                            Colectie_puncte_PPE_L.Add(Poly_PPE_L.GetPoint3dAt(0));
                                            Colectie_puncte_PPE_L.Add(Poly_PPE_L.GetPoint3dAt(1));
                                            Colectie_stanga1.Add(Poly_PPE_L.GetPoint3dAt(0));
                                            Colectie_stanga1.Add(Poly_PPE_L.GetPoint3dAt(1));
                                        }

                                        Autodesk.AutoCAD.DatabaseServices.DBObjectCollection Colectie_offset_PPE_R = Poly_start_end.GetOffsetCurves(+PPE_R);

                                        Polyline Poly_PPE_R = new Polyline();
                                        foreach (Polyline obj in Colectie_offset_PPE_R)
                                        {
                                            if (obj != null)
                                            {
                                                Poly_PPE_R = obj;

                                            }
                                        }
                                        if (Poly_PPE_R != null)
                                        {

                                            Colectie_puncte_PPE_R.Add(Poly_PPE_R.GetPoint3dAt(0));
                                            Colectie_puncte_PPE_R.Add(Poly_PPE_R.GetPoint3dAt(1));
                                            Colectie_dreapta1.Add(Poly_PPE_R.GetPoint3dAt(0));
                                            Colectie_dreapta1.Add(Poly_PPE_R.GetPoint3dAt(1));
                                        }

                                    }

                                    if ((TWS_L != 0) & (PPE_L != 0) & (PPE_R != 0))
                                    {
                                        Autodesk.AutoCAD.DatabaseServices.DBObjectCollection Colectie_offset_TWS_L = Poly_start_end.GetOffsetCurves(-PPE_L - TWS_L);

                                        Polyline Poly_TWS_L = new Polyline();
                                        foreach (Polyline obj in Colectie_offset_TWS_L)
                                        {
                                            if (obj != null)
                                            {
                                                Poly_TWS_L = obj;
                                            }

                                            if (Poly_TWS_L != null)
                                            {
                                                Colectie_puncte_TWS_L[NrL].Add(Poly_TWS_L.GetPoint3dAt(0));
                                                Colectie_puncte_TWS_L[NrL].Add(Poly_TWS_L.GetPoint3dAt(1));
                                                Colectie_stanga1.RemoveAt(Colectie_stanga1.Count - 1);
                                                Colectie_stanga1.RemoveAt(Colectie_stanga1.Count - 1);
                                                Colectie_stanga1.Add(Poly_TWS_L.GetPoint3dAt(0));
                                                Colectie_stanga1.Add(Poly_TWS_L.GetPoint3dAt(1));
                                            }
                                        }
                                    }
                                    else
                                    {
                                        NrL = NrL + 1;
                                        Array.Resize(ref Colectie_puncte_TWS_L, NrL + 1);
                                        Colectie_puncte_TWS_L[NrL] = new Point3dCollection();
                                    }

                                    if ((TWS_R != 0) & (PPE_L != 0) & (PPE_R != 0))
                                    {
                                        Autodesk.AutoCAD.DatabaseServices.DBObjectCollection Colectie_offset_TWS_R = Poly_start_end.GetOffsetCurves(PPE_R + TWS_R);
                                        Polyline Poly_TWS_R = new Polyline();
                                        foreach (Polyline obj in Colectie_offset_TWS_R)
                                        {
                                            if (obj != null)
                                            {
                                                Poly_TWS_R = obj;
                                            }
                                            if (Poly_TWS_R != null)
                                            {
                                                Colectie_puncte_TWS_R[NrR].Add(Poly_TWS_R.GetPoint3dAt(0));
                                                Colectie_puncte_TWS_R[NrR].Add(Poly_TWS_R.GetPoint3dAt(1));
                                                Colectie_dreapta1.RemoveAt(Colectie_dreapta1.Count - 1);
                                                Colectie_dreapta1.RemoveAt(Colectie_dreapta1.Count - 1);
                                                Colectie_dreapta1.Add(Poly_TWS_R.GetPoint3dAt(0));
                                                Colectie_dreapta1.Add(Poly_TWS_R.GetPoint3dAt(1));
                                            }
                                        }
                                    }
                                    else
                                    {
                                        NrR = NrR + 1;
                                        Array.Resize(ref Colectie_puncte_TWS_R, NrR + 1);
                                        Colectie_puncte_TWS_R[NrR] = new Point3dCollection();
                                    }
                                }
                            }

                            if (Colectie_puncte_PPE_L.Count > 0)
                            {
                                Polyline Poly_PPE_L = new Polyline();
                                for (int i = 0; i < Colectie_puncte_PPE_L.Count; ++i)
                                {
                                    double X = Colectie_puncte_PPE_L[i].X;
                                    double Y = Colectie_puncte_PPE_L[i].Y;
                                    Point2d Pt_vertx = new Point2d(X, Y);
                                    Poly_PPE_L.AddVertexAt(i, Pt_vertx, 0, 0, 0);
                                }
                                Poly_PPE_L.ColorIndex = 3;
                                BTrecord.AppendEntity(Poly_PPE_L);
                                Trans1.AddNewlyCreatedDBObject(Poly_PPE_L, true);

                                if (Colectie_puncte_TWS_L.Length > 0)
                                {
                                    for (int j = 0; j < Colectie_puncte_TWS_L.Length; ++j)
                                    {
                                        if (Colectie_puncte_TWS_L[j].Count > 0)
                                        {
                                            Polyline Poly_TWS_L = new Polyline();

                                            for (int i = 0; i < Colectie_puncte_TWS_L[j].Count; ++i)
                                            {
                                                double X = Colectie_puncte_TWS_L[j][i].X;
                                                double Y = Colectie_puncte_TWS_L[j][i].Y;
                                                Point2d Pt_vertx = new Point2d(X, Y);
                                                Poly_TWS_L.AddVertexAt(i, Pt_vertx, 0, 0, 0);
                                            }

                                            Point3d Start_poly = new Point3d();
                                            Start_poly = Poly_TWS_L.GetPoint3dAt(0);

                                            Point3d End_poly = new Point3d();
                                            End_poly = Poly_TWS_L.GetPoint3dAt(Poly_TWS_L.NumberOfVertices - 1);

                                            Point3d Point_0 = new Point3d();
                                            Point_0 = Poly_PPE_L.GetClosestPointTo(Start_poly, Vector3d.ZAxis, false);

                                            Point3d Point_1 = new Point3d();
                                            Point_1 = Poly_PPE_L.GetClosestPointTo(End_poly, Vector3d.ZAxis, false);

                                            Poly_TWS_L.AddVertexAt(0, new Point2d(Point_0.X, Point_0.Y), 0, 0, 0);
                                            Poly_TWS_L.AddVertexAt(Poly_TWS_L.NumberOfVertices, new Point2d(Point_1.X, Point_1.Y), 0, 0, 0);
                                            Poly_TWS_L.Closed = true;

                                            Poly_TWS_L.ColorIndex = 2;
                                            BTrecord.AppendEntity(Poly_TWS_L);
                                            Trans1.AddNewlyCreatedDBObject(Poly_TWS_L, true);
                                        }
                                    }
                                }
                            }

                            if (Colectie_puncte_PPE_R.Count > 0)
                            {
                                Polyline Poly_PPE_R = new Polyline();
                                for (int i = 0; i < Colectie_puncte_PPE_R.Count; ++i)
                                {
                                    double X = Colectie_puncte_PPE_R[i].X;
                                    double Y = Colectie_puncte_PPE_R[i].Y;
                                    Point2d Pt_vertx = new Point2d(X, Y);
                                    Poly_PPE_R.AddVertexAt(i, Pt_vertx, 0, 0, 0);
                                }

                                Poly_PPE_R.ColorIndex = 3;
                                BTrecord.AppendEntity(Poly_PPE_R);
                                Trans1.AddNewlyCreatedDBObject(Poly_PPE_R, true);

                                if (Colectie_puncte_TWS_R.Length > 0)
                                {
                                    for (int j = 0; j < Colectie_puncte_TWS_R.Length; ++j)
                                    {
                                        if (Colectie_puncte_TWS_R[j].Count > 0)
                                        {
                                            Polyline Poly_TWS_R = new Polyline();
                                            for (int i = 0; i < Colectie_puncte_TWS_R[j].Count; ++i)
                                            {
                                                double X = Colectie_puncte_TWS_R[j][i].X;
                                                double Y = Colectie_puncte_TWS_R[j][i].Y;
                                                Point2d Pt_vertx = new Point2d(X, Y);
                                                Poly_TWS_R.AddVertexAt(i, Pt_vertx, 0, 0, 0);
                                            }

                                            Point3d Start_poly = new Point3d();
                                            Start_poly = Poly_TWS_R.GetPoint3dAt(0);
                                            Point3d End_poly = new Point3d();
                                            End_poly = Poly_TWS_R.GetPoint3dAt(Poly_TWS_R.NumberOfVertices - 1);
                                            Point3d Point_0 = new Point3d();
                                            Point_0 = Poly_PPE_R.GetClosestPointTo(Start_poly, Vector3d.ZAxis, false);
                                            Point3d Point_1 = new Point3d();
                                            Point_1 = Poly_PPE_R.GetClosestPointTo(End_poly, Vector3d.ZAxis, false);
                                            Poly_TWS_R.AddVertexAt(0, new Point2d(Point_0.X, Point_0.Y), 0, 0, 0);
                                            Poly_TWS_R.AddVertexAt(Poly_TWS_R.NumberOfVertices, new Point2d(Point_1.X, Point_1.Y), 0, 0, 0);
                                            Poly_TWS_R.Closed = true;
                                            Poly_TWS_R.ColorIndex = 2;
                                            BTrecord.AppendEntity(Poly_TWS_R);
                                            Trans1.AddNewlyCreatedDBObject(Poly_TWS_R, true);
                                        }
                                    }
                                }
                            }
                            Polyline Stanga1 = null;
                            Polyline Dreapta1 = null;
                            if (Colectie_stanga1.Count > 0)
                            {
                                Stanga1 = new Polyline();
                                for (int i = 0; i < Colectie_stanga1.Count; ++i)
                                {
                                    double X = Colectie_stanga1[i].X;
                                    double Y = Colectie_stanga1[i].Y;
                                    Point2d Pt_vertx = new Point2d(X, Y);
                                    Stanga1.AddVertexAt(i, Pt_vertx, 0, 0, 0);
                                }

                            }

                            if (Colectie_dreapta1.Count > 0)
                            {
                                Dreapta1 = new Polyline();
                                for (int i = 0; i < Colectie_dreapta1.Count; ++i)
                                {
                                    double X = Colectie_dreapta1[i].X;
                                    double Y = Colectie_dreapta1[i].Y;
                                    Point2d Pt_vertx = new Point2d(X, Y);
                                    Dreapta1.AddVertexAt(i, Pt_vertx, 0, 0, 0);
                                }

                            }
                            if (Stanga1 != null & Dreapta1 != null)
                            {
                                if (ATWS_DATA_TABLE.Rows.Count > 0)
                                {
                                    for (int i = 0; i < ATWS_DATA_TABLE.Rows.Count; ++i)
                                    {
                                        Double Start1 = -1;
                                        Double End1 = -1;
                                        Double Width1 = -1;
                                        string Direct1 = "";

                                        if ((ATWS_DATA_TABLE.Rows[i][Column_ATWS_BEGIN] != System.DBNull.Value)) Start1 = Convert.ToDouble(ATWS_DATA_TABLE.Rows[i][Column_ATWS_BEGIN]);
                                        if ((ATWS_DATA_TABLE.Rows[i][Column_ATWS_END] != System.DBNull.Value)) End1 = Convert.ToDouble(ATWS_DATA_TABLE.Rows[i][Column_ATWS_END]);
                                        if ((ATWS_DATA_TABLE.Rows[i][Column_ATWS_WIDTH] != System.DBNull.Value)) Width1 = Convert.ToDouble(ATWS_DATA_TABLE.Rows[i][Column_ATWS_WIDTH]);
                                        if ((ATWS_DATA_TABLE.Rows[i][Column_ATWS_DIRECTION] != System.DBNull.Value)) Direct1 = Convert.ToString(ATWS_DATA_TABLE.Rows[i][Column_ATWS_DIRECTION]);

                                        if (Start1 <= Match1 & End1 >= Match2)
                                        {
                                            Start1 = Match1;
                                            End1 = Match2;
                                        }
                                        if (Start1 <= Match1 & End1 <= Match2)
                                        {
                                            Start1 = Match1;
                                        }

                                        if (Start1 >= Match1 & End1 >= Match2)
                                        {
                                            End1 = Match2;
                                        }


                                        if (Start1 >= Match1 & Start1 <= Match2 & End1 >= Match1 & End1 <= Match2)
                                        {
                                            int Directie1 = 1;
                                            if (Direct1 == Symbol_Left) Directie1 = -1;
                                            Point3d Point_start1 = new Point3d();
                                            Point_start1 = PolyCL_PS.GetPointAtDist(Start1);
                                            Point3d Point_end1 = new Point3d();
                                            Point_end1 = PolyCL_PS.GetPointAtDist(End1);

                                            Double Param1 = PolyCL_PS.GetParameterAtDistance(Start1);
                                            Point3d Nod11 = new Point3d();
                                            Point3d Nod12 = new Point3d();

                                            Nod11 = PolyCL_PS.GetPointAtParameter(Math.Floor(Param1));
                                            Nod12 = PolyCL_PS.GetPointAtParameter(Math.Ceiling(Param1));

                                            if (Math.Floor(Param1) == Math.Ceiling(Param1))
                                            {
                                                if (Param1 >= 1)
                                                {
                                                    Nod11 = PolyCL_PS.GetPointAtParameter(Math.Floor(Param1) - 1);
                                                    Nod12 = PolyCL_PS.GetPointAtParameter(Math.Floor(Param1) + 1);
                                                }
                                                else
                                                {
                                                    Nod11 = PolyCL_PS.GetPointAtParameter(0);
                                                    Nod12 = PolyCL_PS.GetPointAtParameter(1);
                                                }

                                            }
                                            Line Linie1 = new Line(Nod11, Nod12);
                                            Linie1.TransformBy(Matrix3d.Displacement(Nod11.GetVectorTo(Point_start1)));
                                            Linie1.TransformBy(Matrix3d.Rotation(-Directie1 * Math.PI / 2, Vector3d.ZAxis, Point_start1));



                                            Point3dCollection Col_int11 = new Point3dCollection();
                                            if (Directie1 == -1 & Stanga1 != null)
                                            {
                                                Linie1.IntersectWith(Stanga1, Autodesk.AutoCAD.DatabaseServices.Intersect.ExtendThis, Col_int11, IntPtr.Zero, IntPtr.Zero);
                                            }

                                            if (Directie1 == 1 & Dreapta1 != null)
                                            {
                                                Linie1.IntersectWith(Dreapta1, Autodesk.AutoCAD.DatabaseServices.Intersect.ExtendThis, Col_int11, IntPtr.Zero, IntPtr.Zero);
                                            }




                                            Point3d Point_start2_stanga = new Point3d();
                                            Point3d Point_start2_dreapta = new Point3d();

                                            if (Col_int11.Count > 0)
                                            {
                                                if (Directie1 == -1)
                                                {
                                                    Point_start2_stanga = Col_int11[0];
                                                }
                                                if (Directie1 == 1)
                                                {
                                                    Point_start2_dreapta = Col_int11[0];
                                                }

                                            }

                                            Double Param2 = PolyCL_PS.GetParameterAtDistance(End1);
                                            Point3d Nod21 = new Point3d();
                                            Point3d Nod22 = new Point3d();

                                            Nod21 = PolyCL_PS.GetPointAtParameter(Math.Floor(Param2));
                                            Nod22 = PolyCL_PS.GetPointAtParameter(Math.Ceiling(Param2));
                                            if (Math.Floor(Param2) == Math.Ceiling(Param2))
                                            {
                                                if (Param2 >= 1)
                                                {
                                                    Nod21 = PolyCL_PS.GetPointAtParameter(Math.Floor(Param2) - 1);
                                                    Nod22 = PolyCL_PS.GetPointAtParameter(Math.Floor(Param2) + 1);
                                                }
                                                else
                                                {
                                                    Nod21 = PolyCL_PS.GetPointAtParameter(0);
                                                    Nod22 = PolyCL_PS.GetPointAtParameter(1);
                                                }

                                            }

                                            Line Linie2 = new Line(Nod21, Nod22);
                                            Linie2.TransformBy(Matrix3d.Displacement(Nod21.GetVectorTo(Point_end1)));
                                            Linie2.TransformBy(Matrix3d.Rotation(-Directie1 * Math.PI / 2, Vector3d.ZAxis, Point_end1));



                                            Point3dCollection Col_int22 = new Point3dCollection();
                                            if (Directie1 == -1 & Stanga1 != null)
                                            {
                                                Linie2.IntersectWith(Stanga1, Autodesk.AutoCAD.DatabaseServices.Intersect.ExtendThis, Col_int22, IntPtr.Zero, IntPtr.Zero);
                                            }

                                            if (Directie1 == 1 & Dreapta1 != null)
                                            {
                                                Linie2.IntersectWith(Dreapta1, Autodesk.AutoCAD.DatabaseServices.Intersect.ExtendThis, Col_int22, IntPtr.Zero, IntPtr.Zero);
                                            }



                                            Point3d Point_end2_stanga = new Point3d();
                                            Point3d Point_end2_dreapta = new Point3d();

                                            if (Col_int22.Count > 0)
                                            {
                                                if (Directie1 == -1)
                                                {
                                                    Point_end2_stanga = Col_int22[0];
                                                }
                                                if (Directie1 == 1)
                                                {
                                                    Point_end2_dreapta = Col_int22[0];
                                                }

                                            }


                                            Polyline Poly_start_end = new Polyline();

                                            if (Directie1 == -1)
                                            {
                                                Poly_start_end.AddVertexAt(0, new Point2d(Point_start2_stanga.X, Point_start2_stanga.Y), 0, 0, 0);

                                                int Index_ATWS = 1;

                                                double Param_atws1 = Stanga1.GetParameterAtPoint(Point_start2_stanga);
                                                double Param_atws2 = Stanga1.GetParameterAtPoint(Point_end2_stanga);
                                                if (Math.Floor(Param_atws1) == Math.Floor(Param_atws2))
                                                {
                                                    Poly_start_end.AddVertexAt(1, new Point2d(Point_end2_stanga.X, Point_end2_stanga.Y), 0, 0, 0);
                                                }
                                                else
                                                {
                                                    for (int j = Convert.ToInt32(Math.Ceiling(Param_atws1)); j <= Math.Floor(Param_atws2); ++j)
                                                    {
                                                        Poly_start_end.AddVertexAt(Index_ATWS, Stanga1.GetPoint2dAt(j), 0, 0, 0);
                                                        Index_ATWS = Index_ATWS + 1;
                                                    }
                                                    Poly_start_end.AddVertexAt(Index_ATWS, new Point2d(Point_end2_stanga.X, Point_end2_stanga.Y), 0, 0, 0);

                                                }



                                                Autodesk.AutoCAD.DatabaseServices.DBObjectCollection Colectie_offset_ATWS_L = Poly_start_end.GetOffsetCurves(Directie1 * Width1);

                                                Polyline Poly_ATWS_L = new Polyline();
                                                foreach (Polyline obj in Colectie_offset_ATWS_L)
                                                {
                                                    if (obj != null)
                                                    {
                                                        Poly_ATWS_L = obj;


                                                    }
                                                    if (Poly_ATWS_L != null)
                                                    {
                                                        int Nr_vertic_L = Poly_start_end.NumberOfVertices - 1;
                                                        Poly_ATWS_L.AddVertexAt(0, Poly_start_end.GetPoint2dAt(0), 0, 0, 0);

                                                        if (Poly_start_end.NumberOfVertices > 2)
                                                        {
                                                            for (int j = Nr_vertic_L; j > 0; --j)
                                                            {
                                                                Poly_ATWS_L.AddVertexAt(Poly_ATWS_L.NumberOfVertices, Poly_start_end.GetPoint2dAt(j), 0, 0, 0);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            Poly_ATWS_L.AddVertexAt(Poly_ATWS_L.NumberOfVertices, Poly_start_end.GetPoint2dAt(Nr_vertic_L), 0, 0, 0);
                                                        }

                                                        Poly_ATWS_L.Closed = true;
                                                        Poly_ATWS_L.ColorIndex = 1;
                                                        BTrecord.AppendEntity(Poly_ATWS_L);
                                                        Trans1.AddNewlyCreatedDBObject(Poly_ATWS_L, true);
                                                    }
                                                }
                                            }

                                            if (Directie1 == 1)
                                            {
                                                Poly_start_end.AddVertexAt(0, new Point2d(Point_start2_dreapta.X, Point_start2_dreapta.Y), 0, 0, 0);

                                                int Index_ATWS = 1;

                                                double Param_atws1 = Dreapta1.GetParameterAtPoint(Point_start2_dreapta);
                                                double Param_atws2 = Dreapta1.GetParameterAtPoint(Point_end2_dreapta);
                                                if (Math.Floor(Param_atws1) == Math.Floor(Param_atws2))
                                                {
                                                    Poly_start_end.AddVertexAt(1, new Point2d(Point_end2_dreapta.X, Point_end2_dreapta.Y), 0, 0, 0);
                                                }
                                                else
                                                {
                                                    for (int j = Convert.ToInt32(Math.Ceiling(Param_atws1)); j <= Math.Floor(Param_atws2); ++j)
                                                    {
                                                        Poly_start_end.AddVertexAt(Index_ATWS, Dreapta1.GetPoint2dAt(j), 0, 0, 0);
                                                        Index_ATWS = Index_ATWS + 1;
                                                    }
                                                    Poly_start_end.AddVertexAt(Index_ATWS, new Point2d(Point_end2_dreapta.X, Point_end2_dreapta.Y), 0, 0, 0);
                                                }

                                                Autodesk.AutoCAD.DatabaseServices.DBObjectCollection Colectie_offset_ATWS_R = Poly_start_end.GetOffsetCurves(Directie1 * Width1);

                                                Polyline Poly_ATWS_R = new Polyline();
                                                foreach (Polyline obj in Colectie_offset_ATWS_R)
                                                {
                                                    if (obj != null)
                                                    {
                                                        Poly_ATWS_R = obj;
                                                    }

                                                    if (Poly_ATWS_R != null)
                                                    {
                                                        int Nr_vertic_R = Poly_start_end.NumberOfVertices - 1;
                                                        Poly_ATWS_R.AddVertexAt(0, Poly_start_end.GetPoint2dAt(0), 0, 0, 0);

                                                        if (Poly_start_end.NumberOfVertices > 2)
                                                        {
                                                            for (int j = Nr_vertic_R; j > 0; --j)
                                                            {
                                                                Poly_ATWS_R.AddVertexAt(Poly_ATWS_R.NumberOfVertices, Poly_start_end.GetPoint2dAt(j), 0, 0, 0);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            Poly_ATWS_R.AddVertexAt(Poly_ATWS_R.NumberOfVertices, Poly_start_end.GetPoint2dAt(Nr_vertic_R), 0, 0, 0);
                                                        }

                                                        Poly_ATWS_R.Closed = true;
                                                        Poly_ATWS_R.ColorIndex = 1;
                                                        BTrecord.AppendEntity(Poly_ATWS_R);
                                                        Trans1.AddNewlyCreatedDBObject(Poly_ATWS_R, true);

                                                    }

                                                }
                                            }

                                        }


                                    }
                                }

                            }


                        }
                    }



                    Trans1.Commit();
                }

                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
            }

            catch (System.Exception ex)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\n" + "Command:");
                MessageBox.Show(ex.Message);
            }



        }

        private void Button_draw_from_compiled(object sender, EventArgs e)
        {
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();



                if (PolyCL_MS == null)
                {

                    //ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                    //return;
                }

                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the starting point");
                PP1.AllowNone = false;
                Point_res1 = Editor1.GetPoint(PP1);

                if (Point_res1.Status != PromptStatus.OK)
                {

                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                    return;
                }

                Connect_to_compiled_EPE_L();
                draw_easement_left(Point_res1.Value);
                Connect_to_compiled_EPE_R();

                draw_easement_right(Point_res1.Value);

                TWS_L_ON_EPE_TABLE = new System.Data.DataTable();
                TWS_L_ON_EPE_TABLE = Connect_to_compiled_TWS("TWS", "L", "EPE");
                draw_TWS_L_ON_EPE(Point_res1.Value);

                TWS_R_ON_EPE_TABLE = new System.Data.DataTable();
                TWS_R_ON_EPE_TABLE = Connect_to_compiled_TWS("TWS", "R", "EPE");
                draw_TWS_R_ON_EPE(Point_res1.Value);

                ATWS_L_ON_EPE_TABLE = new System.Data.DataTable();
                ATWS_L_ON_EPE_TABLE = Connect_to_compiled_TWS("ATWS", "L", "EPE");
                draw_ATWS_L_ON_EPE(Point_res1.Value);

                ATWS_R_ON_EPE_TABLE = new System.Data.DataTable();
                ATWS_R_ON_EPE_TABLE = Connect_to_compiled_TWS("ATWS", "R", "EPE");
                draw_ATWS_R_ON_EPE(Point_res1.Value);

                ATWS_L_ON_TWS_TABLE = new System.Data.DataTable();
                ATWS_L_ON_TWS_TABLE = Connect_to_compiled_TWS("ATWS", "L", "TWS");
                draw_ATWS_L_ON_TWS(Point_res1.Value);

                ATWS_R_ON_TWS_TABLE = new System.Data.DataTable();
                ATWS_R_ON_TWS_TABLE = Connect_to_compiled_TWS("ATWS", "R", "TWS");
                draw_ATWS_R_ON_TWS(Point_res1.Value);



            }

            catch (System.Exception ex)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\n" + "Command:");
                MessageBox.Show(ex.Message);
            }



        }

        private void Connect_to_compiled_EPE_L()
        {
            try
            {
                string Table1 = textBox_COMPILED.Text;
                string query = "SELECT * FROM " + Table1 + " WHERE TYPE='EPE' AND SIDE='L'";
                string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + textBox_access_database_location.Text + ";";
                DataSet EPE_L_DATASET = new DataSet();
                OleDbConnection cnn;
                try
                {
                    cnn = new OleDbConnection(ConnectionString);
                    cnn.Open();
                    OleDbCommand cmd = new OleDbCommand(query, cnn);
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(EPE_L_DATASET, Table1);
                    cnn.Close();
                }
                catch (OleDbException ex)
                {
                    MessageBox.Show(ex.Message);
                }
                EPE_L_DATA_TABLE = new System.Data.DataTable();
                EPE_L_DATASET.Tables[Table1].DefaultView.Sort = "START ASC";

                EPE_L_DATA_TABLE = EPE_L_DATASET.Tables[Table1];

                //MessageBox.Show(ROW_DATA_TABLE.Columns[4].ColumnName + " Row 2 = " + ROW_DATA_TABLE.Rows[1][4]);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }
        private void Connect_to_compiled_EPE_R()
        {
            try
            {
                string Table1 = textBox_COMPILED.Text;
                string query = "SELECT * FROM " + Table1 + " WHERE TYPE='EPE' AND SIDE='R'";
                string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + textBox_access_database_location.Text + ";";
                DataSet EPE_R_DATASET = new DataSet();
                OleDbConnection cnn;
                try
                {
                    cnn = new OleDbConnection(ConnectionString);
                    cnn.Open();
                    OleDbCommand cmd = new OleDbCommand(query, cnn);
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(EPE_R_DATASET, Table1);
                    cnn.Close();
                }
                catch (OleDbException ex)
                {
                    MessageBox.Show(ex.Message);
                }
                EPE_R_DATA_TABLE = new System.Data.DataTable();
                EPE_R_DATASET.Tables[Table1].DefaultView.Sort = "START ASC";

                EPE_R_DATA_TABLE = EPE_R_DATASET.Tables[Table1];

                //MessageBox.Show(ROW_DATA_TABLE.Columns[4].ColumnName + " Row 2 = " + ROW_DATA_TABLE.Rows[1][4]);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        public System.Data.DataTable Connect_to_compiled_TWS(string Type, string Side, string Bdy)
        {
            try
            {
                string Table1 = textBox_COMPILED.Text;
                string query = "SELECT * FROM " + Table1 + " WHERE TYPE='" + Type + "' AND SIDE='" + Side + "' AND BDY='" + Bdy + "'";
                string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + textBox_access_database_location.Text + ";";
                DataSet TWS_DATASET = new DataSet();
                OleDbConnection cnn;
                try
                {
                    cnn = new OleDbConnection(ConnectionString);
                    cnn.Open();
                    OleDbCommand cmd = new OleDbCommand(query, cnn);
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(TWS_DATASET, Table1);
                    cnn.Close();
                }
                catch (OleDbException ex)
                {
                    MessageBox.Show(ex.Message);
                }

                TWS_DATASET.Tables[Table1].DefaultView.Sort = "START ASC";

                return TWS_DATASET.Tables[Table1];

                //MessageBox.Show(MemoryTable.Columns[4].ColumnName + " Row 2 = " + MemoryTable.Rows[1][4]);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }


        }



        private void draw_easement_left(Point3d Point1)
        {
            try
            {
                string Column_EPE_L_Begin = "START";
                string Column_EPE_L_End = "END";

                string Column_EPE_L_Width_Begin = "START_WIDTH";
                string Column_EPE_L_Width_End = "END_WIDTH";


                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        if (EPE_L_DATA_TABLE != null)
                        {
                            if (EPE_L_DATA_TABLE.Rows.Count > 0)
                            {
                                Double Match1 = Convert.ToDouble(textBox_Matchline_start.Text);
                                Double Match2 = Convert.ToDouble(textBox_Matchline_end.Text);
                                if (Match1 > Match2)
                                {
                                    double T = Match1;
                                    Match1 = Match2;
                                    Match2 = T;
                                }



                                double Len1 = Match2 - Match1;// PolyCL_MS.Length;

                                Polyline PolyCL_PS = new Polyline();
                                PolyCL_PS.AddVertexAt(0, new Point2d(Point1.X, Point1.Y), 0, 0, 0);
                                PolyCL_PS.AddVertexAt(1, new Point2d(Point1.X + Len1, Point1.Y), 0, 0, 0);
                                PolyCL_PS.ColorIndex = 1;
                                BTrecord.AppendEntity(PolyCL_PS);
                                Trans1.AddNewlyCreatedDBObject(PolyCL_PS, true);
                                Point3dCollection Colectie_puncte_EPE_L = new Point3dCollection();

                                for (int i = 0; i < EPE_L_DATA_TABLE.Rows.Count; ++i)
                                {
                                    Double Start1 = -1;
                                    Double End1 = -1;
                                    Double Start_w1 = -1;
                                    Double End_w1 = -1;
                                    if ((EPE_L_DATA_TABLE.Rows[i][Column_EPE_L_Begin] != System.DBNull.Value)) Start1 = Convert.ToDouble(EPE_L_DATA_TABLE.Rows[i][Column_EPE_L_Begin]);
                                    if ((EPE_L_DATA_TABLE.Rows[i][Column_EPE_L_End] != System.DBNull.Value)) End1 = Convert.ToDouble(EPE_L_DATA_TABLE.Rows[i][Column_EPE_L_End]);
                                    if ((EPE_L_DATA_TABLE.Rows[i][Column_EPE_L_Width_Begin] != System.DBNull.Value)) Start_w1 = Convert.ToDouble(EPE_L_DATA_TABLE.Rows[i][Column_EPE_L_Width_Begin]);
                                    if ((EPE_L_DATA_TABLE.Rows[i][Column_EPE_L_Width_End] != System.DBNull.Value)) End_w1 = Convert.ToDouble(EPE_L_DATA_TABLE.Rows[i][Column_EPE_L_Width_End]);
                                    if (Start1 < Match1 & End1 <= Match2)
                                    {
                                        Start1 = Match1;
                                    }

                                    if (Start1 >= Match1 & Start1 <= Match2 & End1 >= Match2)
                                    {
                                        End1 = Match2;

                                    }

                                    if (Start1 <= Match1 & End1 >= Match2)
                                    {
                                        Start1 = Match1;
                                        End1 = Match2;

                                    }
                                    if (Start1 >= Match1 & End1 <= Match2)
                                    {
                                        if (Start1 != -1 & End1 != -1 & Start_w1 != -1 & End_w1 != -1)
                                        {
                                            Point3d PointCL_PS_start1 = new Point3d();
                                            PointCL_PS_start1 = PolyCL_PS.GetPointAtDist(Start1);
                                            Point3d PointCL_PS_end1 = new Point3d();
                                            PointCL_PS_end1 = PolyCL_PS.GetPointAtDist(End1);
                                            Colectie_puncte_EPE_L.Add(new Point3d(PointCL_PS_start1.X, PointCL_PS_start1.Y + Start_w1, 0));
                                            Colectie_puncte_EPE_L.Add(new Point3d(PointCL_PS_end1.X, PointCL_PS_end1.Y + End_w1, 0));

                                        }
                                    }
                                }

                                if (Colectie_puncte_EPE_L.Count > 1)
                                {
                                    Polyline Poly_EPE_L = new Polyline();
                                    for (int i = 0; i < Colectie_puncte_EPE_L.Count; ++i)
                                    {
                                        Poly_EPE_L.AddVertexAt(i, new Point2d(Colectie_puncte_EPE_L[i].X, Colectie_puncte_EPE_L[i].Y), 0, 0, 0);

                                    }


                                    Poly_EPE_L.ColorIndex = 6;
                                    BTrecord.AppendEntity(Poly_EPE_L);
                                    Trans1.AddNewlyCreatedDBObject(Poly_EPE_L, true);
                                    EPE_L = Poly_EPE_L;
                                }


                            }
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
        private void draw_easement_right(Point3d Point1)
        {
            try
            {
                string Column_EPE_R_Begin = "START";
                string Column_EPE_R_End = "END";

                string Column_EPE_R_Width_Begin = "START_WIDTH";
                string Column_EPE_R_Width_End = "END_WIDTH";


                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        if (EPE_R_DATA_TABLE != null)
                        {
                            if (EPE_R_DATA_TABLE.Rows.Count > 0)
                            {
                                Double Match1 = Convert.ToDouble(textBox_Matchline_start.Text);
                                Double Match2 = Convert.ToDouble(textBox_Matchline_end.Text);
                                if (Match1 > Match2)
                                {
                                    double T = Match1;
                                    Match1 = Match2;
                                    Match2 = T;
                                }


                                double Len1 = Match2 - Match1;// PolyCL_MS.Length;

                                Polyline PolyCL_PS = new Polyline();
                                PolyCL_PS.AddVertexAt(0, new Point2d(Point1.X, Point1.Y), 0, 0, 0);
                                PolyCL_PS.AddVertexAt(1, new Point2d(Point1.X + Len1, Point1.Y), 0, 0, 0);
                                PolyCL_PS.ColorIndex = 1;
                                BTrecord.AppendEntity(PolyCL_PS);
                                Trans1.AddNewlyCreatedDBObject(PolyCL_PS, true);
                                Point3dCollection Colectie_puncte_EPE_R = new Point3dCollection();

                                for (int i = 0; i < EPE_R_DATA_TABLE.Rows.Count; ++i)
                                {
                                    Double Start1 = -1;
                                    Double End1 = -1;
                                    Double Start_w1 = -1;
                                    Double End_w1 = -1;
                                    if ((EPE_R_DATA_TABLE.Rows[i][Column_EPE_R_Begin] != System.DBNull.Value)) Start1 = Convert.ToDouble(EPE_R_DATA_TABLE.Rows[i][Column_EPE_R_Begin]);
                                    if ((EPE_R_DATA_TABLE.Rows[i][Column_EPE_R_End] != System.DBNull.Value)) End1 = Convert.ToDouble(EPE_R_DATA_TABLE.Rows[i][Column_EPE_R_End]);
                                    if ((EPE_R_DATA_TABLE.Rows[i][Column_EPE_R_Width_Begin] != System.DBNull.Value)) Start_w1 = Convert.ToDouble(EPE_R_DATA_TABLE.Rows[i][Column_EPE_R_Width_Begin]);
                                    if ((EPE_R_DATA_TABLE.Rows[i][Column_EPE_R_Width_End] != System.DBNull.Value)) End_w1 = Convert.ToDouble(EPE_R_DATA_TABLE.Rows[i][Column_EPE_R_Width_End]);

                                    if (Start1 < Match1 & End1 <= Match2)
                                    {
                                        Start1 = Match1;
                                    }

                                    if (Start1 >= Match1 & Start1 <= Match2 & End1 >= Match2)
                                    {
                                        End1 = Match2;

                                    }

                                    if (Start1 <= Match1 & End1 >= Match2)
                                    {
                                        Start1 = Match1;
                                        End1 = Match2;

                                    }
                                    if (Start1 >= Match1 & End1 <= Match2)
                                    {
                                        if (Start1 != -1 & End1 != -1 & Start_w1 != -1 & End_w1 != -1)
                                        {
                                            Point3d PointCL_PS_start1 = new Point3d();
                                            PointCL_PS_start1 = PolyCL_PS.GetPointAtDist(Start1);
                                            Point3d PointCL_PS_end1 = new Point3d();
                                            PointCL_PS_end1 = PolyCL_PS.GetPointAtDist(End1);
                                            Colectie_puncte_EPE_R.Add(new Point3d(PointCL_PS_start1.X, PointCL_PS_start1.Y - Start_w1, 0));
                                            Colectie_puncte_EPE_R.Add(new Point3d(PointCL_PS_end1.X, PointCL_PS_end1.Y - End_w1, 0));

                                        }
                                    }



                                }
                                if (Colectie_puncte_EPE_R.Count > 1)
                                {
                                    Polyline Poly_EPE_R = new Polyline();
                                    for (int i = 0; i < Colectie_puncte_EPE_R.Count; ++i)
                                    {
                                        Poly_EPE_R.AddVertexAt(i, new Point2d(Colectie_puncte_EPE_R[i].X, Colectie_puncte_EPE_R[i].Y), 0, 0, 0);

                                    }


                                    Poly_EPE_R.ColorIndex = 6;
                                    BTrecord.AppendEntity(Poly_EPE_R);
                                    Trans1.AddNewlyCreatedDBObject(Poly_EPE_R, true);
                                    EPE_R = Poly_EPE_R;
                                }


                            }
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



        private void draw_TWS_L_ON_EPE(Point3d Point1)
        {
            string eRR_Stat = "";
            try
            {
                string Column_TWS_L_ON_EPE_Begin = "START";
                string Column_TWS_L_ON_EPE_End = "END";

                string Column_TWS_L_ON_EPE_Width_Begin = "START_WIDTH";
                string Column_TWS_L_ON_EPE_Width_End = "END_WIDTH";


                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        if (TWS_L_ON_EPE_TABLE != null & EPE_L != null)
                        {
                            if (TWS_L_ON_EPE_TABLE.Rows.Count > 0)
                            {
                                Double Match1 = Convert.ToDouble(textBox_Matchline_start.Text);
                                Double Match2 = Convert.ToDouble(textBox_Matchline_end.Text);
                                if (Match1 > Match2)
                                {
                                    double T = Match1;
                                    Match1 = Match2;
                                    Match2 = T;
                                }

                                double Len1 = Match2 - Match1;// PolyCL_MS.Length;

                                Polyline PolyCL_PS = new Polyline();
                                PolyCL_PS.AddVertexAt(0, new Point2d(Point1.X, Point1.Y), 0, 0, 0);
                                PolyCL_PS.AddVertexAt(1, new Point2d(Point1.X + Len1, Point1.Y), 0, 0, 0);
                                PolyCL_PS.ColorIndex = 1;
                                PolyCL_PS.Elevation = 0;
                                BTrecord.AppendEntity(PolyCL_PS);
                                Trans1.AddNewlyCreatedDBObject(PolyCL_PS, true);

                                Point3dCollection[] Colectie_puncte_TWS_L_ON_EPE;
                                int Index_TWS_L_ON_EPE = 0;
                                Colectie_puncte_TWS_L_ON_EPE = new Point3dCollection[Index_TWS_L_ON_EPE + 1];
                                Colectie_puncte_TWS_L_ON_EPE[Index_TWS_L_ON_EPE] = new Point3dCollection();

                                Point3dCollection[] Colectie_puncte_sus_TWS_L_ON_EPE;
                                int Index_sus_TWS_L_ON_EPE = 0;
                                Colectie_puncte_sus_TWS_L_ON_EPE = new Point3dCollection[Index_sus_TWS_L_ON_EPE + 1];
                                Colectie_puncte_sus_TWS_L_ON_EPE[Index_sus_TWS_L_ON_EPE] = new Point3dCollection();




                                for (int i = 0; i < TWS_L_ON_EPE_TABLE.Rows.Count; ++i)
                                {
                                    Double Start1 = -1;
                                    Double End1 = -1;
                                    Double Start_w1 = -1;
                                    Double End_w1 = -1;
                                    Double Start2 = -1;


                                    if ((TWS_L_ON_EPE_TABLE.Rows[i][Column_TWS_L_ON_EPE_Begin] != System.DBNull.Value)) Start1 = Convert.ToDouble(TWS_L_ON_EPE_TABLE.Rows[i][Column_TWS_L_ON_EPE_Begin]);
                                    if ((TWS_L_ON_EPE_TABLE.Rows[i][Column_TWS_L_ON_EPE_End] != System.DBNull.Value)) End1 = Convert.ToDouble(TWS_L_ON_EPE_TABLE.Rows[i][Column_TWS_L_ON_EPE_End]);
                                    if ((TWS_L_ON_EPE_TABLE.Rows[i][Column_TWS_L_ON_EPE_Width_Begin] != System.DBNull.Value)) Start_w1 = Convert.ToDouble(TWS_L_ON_EPE_TABLE.Rows[i][Column_TWS_L_ON_EPE_Width_Begin]);
                                    if ((TWS_L_ON_EPE_TABLE.Rows[i][Column_TWS_L_ON_EPE_Width_End] != System.DBNull.Value)) End_w1 = Convert.ToDouble(TWS_L_ON_EPE_TABLE.Rows[i][Column_TWS_L_ON_EPE_Width_End]);

                                    if (i + 1 < TWS_L_ON_EPE_TABLE.Rows.Count)
                                    {
                                        if ((TWS_L_ON_EPE_TABLE.Rows[i + 1][Column_TWS_L_ON_EPE_Begin] != System.DBNull.Value))
                                        {
                                            Start2 = Convert.ToDouble(TWS_L_ON_EPE_TABLE.Rows[i + 1][Column_TWS_L_ON_EPE_Begin]);
                                        }
                                    }


                                    if (Start1 < Match1 & End1 <= Match2)
                                    {
                                        Start1 = Match1;
                                    }

                                    if (Start1 >= Match1 & Start1 <= Match2 & End1 >= Match2)
                                    {
                                        End1 = Match2;

                                    }

                                    if (Start1 <= Match1 & End1 >= Match2)
                                    {
                                        Start1 = Match1;
                                        End1 = Match2;

                                    }
                                    if (Start1 >= Match1 & End1 <= Match2)
                                    {
                                        if (Start1 != -1 & End1 != -1 & Start_w1 != -1 & End_w1 != -1)
                                        {
                                            Point3d PointCL_PS_start1 = new Point3d();
                                            PointCL_PS_start1 = PolyCL_PS.GetPointAtDist(Start1);
                                            Point3d PointCL_PS_end1 = new Point3d();
                                            PointCL_PS_end1 = PolyCL_PS.GetPointAtDist(End1);
                                            Line Linie_int_start = new Line(PointCL_PS_start1, new Point3d(PointCL_PS_start1.X, PointCL_PS_start1.Y + 1000, 0));
                                            Line Linie_int_end = new Line(PointCL_PS_end1, new Point3d(PointCL_PS_end1.X, PointCL_PS_end1.Y + 1000, 0));
                                            //BTrecord.AppendEntity(Linie_int_start);
                                            //BTrecord.AppendEntity(Linie_int_end);
                                            //Trans1.AddNewlyCreatedDBObject(Linie_int_end, true);
                                            //Trans1.AddNewlyCreatedDBObject(Linie_int_start, true);

                                            Point3dCollection Col_int_start = new Point3dCollection();
                                            Point3dCollection Col_int_end = new Point3dCollection();
                                            Point3dCollection Colectie_puncte_poly = new Point3dCollection();

                                            EPE_L.IntersectWith(Linie_int_start, Autodesk.AutoCAD.DatabaseServices.Intersect.ExtendThis, Col_int_start, IntPtr.Zero, IntPtr.Zero);
                                            EPE_L.IntersectWith(Linie_int_end, Autodesk.AutoCAD.DatabaseServices.Intersect.ExtendThis, Col_int_end, IntPtr.Zero, IntPtr.Zero);

                                            Colectie_puncte_TWS_L_ON_EPE[Index_TWS_L_ON_EPE].Add(Col_int_start[0]);
                                            Colectie_puncte_poly.Add(Col_int_start[0]);
                                            double Param_start = EPE_L.GetParameterAtPoint(Col_int_start[0]);
                                            double Param_end = EPE_L.GetParameterAtPoint(Col_int_end[0]);

                                            //Line TEST_int_start = new Line(PointCL_PS_start1, Col_int_start[0]);
                                            //Line TEST_int_end = new Line(PointCL_PS_end1, Col_int_end[0]);
                                            //BTrecord.AppendEntity(TEST_int_start);
                                            //BTrecord.AppendEntity(TEST_int_end);
                                            //Trans1.AddNewlyCreatedDBObject(TEST_int_end, true);
                                            //Trans1.AddNewlyCreatedDBObject(TEST_int_start, true);

                                            if (Math.Floor(Param_end) - Math.Floor(Param_start) >= 1)
                                            {

                                                int P1 = (int)Math.Ceiling(Param_start);
                                                int P2 = (int)Math.Floor(Param_end);

                                                for (int j = P1; j <= P2; ++j)
                                                {
                                                    Colectie_puncte_TWS_L_ON_EPE[Index_TWS_L_ON_EPE].Add(EPE_L.GetPointAtParameter(j));
                                                    Colectie_puncte_poly.Add(EPE_L.GetPointAtParameter(j));
                                                }
                                            }

                                            Colectie_puncte_TWS_L_ON_EPE[Index_TWS_L_ON_EPE].Add(Col_int_end[0]);
                                            Colectie_puncte_poly.Add(Col_int_end[0]);

                                            if (Start_w1 == End_w1)
                                            {

                                                Polyline Temp_poly = new Polyline();

                                                for (int j = 0; j < Colectie_puncte_poly.Count; ++j)
                                                {
                                                    Temp_poly.AddVertexAt(j, new Point2d(Colectie_puncte_poly[j].X, Colectie_puncte_poly[j].Y), 0, 0, 0);
                                                }

                                                //BTrecord.AppendEntity(Temp_poly);
                                                //Trans1.AddNewlyCreatedDBObject(Temp_poly, true);

                                                Autodesk.AutoCAD.DatabaseServices.DBObjectCollection Colectie_offset_TWS_L_ON_EPE = Temp_poly.GetOffsetCurves(-Start_w1);

                                                Polyline Temp_poly_sus = new Polyline();

                                                foreach (Polyline obj in Colectie_offset_TWS_L_ON_EPE)
                                                {
                                                    if (obj != null)
                                                    {
                                                        Temp_poly_sus = obj;
                                                    }

                                                    if (Temp_poly_sus != null)
                                                    {
                                                        for (int j = 0; j < Temp_poly_sus.NumberOfVertices; ++j)
                                                        {
                                                            Colectie_puncte_sus_TWS_L_ON_EPE[Index_sus_TWS_L_ON_EPE].Add(Temp_poly_sus.GetPointAtParameter(j));
                                                        }
                                                    }
                                                }
                                                //BTrecord.AppendEntity(Temp_poly_sus);
                                                //Trans1.AddNewlyCreatedDBObject(Temp_poly_sus, true);
                                            }
                                            else
                                            {
                                                Colectie_puncte_sus_TWS_L_ON_EPE[Index_sus_TWS_L_ON_EPE].Add(new Point3d(Col_int_start[0].X, Col_int_start[0].Y + Start_w1, 0));
                                                Colectie_puncte_sus_TWS_L_ON_EPE[Index_sus_TWS_L_ON_EPE].Add(new Point3d(Col_int_end[0].X, Col_int_end[0].Y + End_w1, 0));
                                            }

                                            if (Start2 != End1)
                                            {
                                                Index_TWS_L_ON_EPE = Index_TWS_L_ON_EPE + 1;
                                                Array.Resize(ref Colectie_puncte_TWS_L_ON_EPE, Index_TWS_L_ON_EPE + 1);
                                                Colectie_puncte_TWS_L_ON_EPE[Index_TWS_L_ON_EPE] = new Point3dCollection();

                                                Index_sus_TWS_L_ON_EPE = Index_sus_TWS_L_ON_EPE + 1;
                                                Array.Resize(ref Colectie_puncte_sus_TWS_L_ON_EPE, Index_sus_TWS_L_ON_EPE + 1);
                                                Colectie_puncte_sus_TWS_L_ON_EPE[Index_sus_TWS_L_ON_EPE] = new Point3dCollection();

                                            }


                                        }
                                    }
                                }

                                TWS_L = new Polyline();
                                int IdxX = 0;

                                for (int i = 0; i <= Index_TWS_L_ON_EPE; ++i)
                                {
                                    Polyline TWS_L_ON_EPE = new Polyline();
                                    int Idx1 = 0;
                                    for (int j = 0; j < Colectie_puncte_sus_TWS_L_ON_EPE[i].Count; ++j)
                                    {
                                        TWS_L_ON_EPE.AddVertexAt(Idx1, new Point2d(Colectie_puncte_sus_TWS_L_ON_EPE[i][j].X, Colectie_puncte_sus_TWS_L_ON_EPE[i][j].Y), 0, 0, 0);
                                        Idx1 = Idx1 + 1;

                                        TWS_L.AddVertexAt(IdxX, new Point2d(Colectie_puncte_sus_TWS_L_ON_EPE[i][j].X, Colectie_puncte_sus_TWS_L_ON_EPE[i][j].Y), 0, 0, 0);
                                        IdxX = IdxX + 1;

                                    }
                                    for (int j = Colectie_puncte_TWS_L_ON_EPE[i].Count - 1; j >= 0; --j)
                                    {
                                        TWS_L_ON_EPE.AddVertexAt(Idx1, new Point2d(Colectie_puncte_TWS_L_ON_EPE[i][j].X, Colectie_puncte_TWS_L_ON_EPE[i][j].Y), 0, 0, 0);
                                        Idx1 = Idx1 + 1;
                                    }

                                    TWS_L_ON_EPE.ColorIndex = 3;
                                    TWS_L_ON_EPE.Closed = true;
                                    BTrecord.AppendEntity(TWS_L_ON_EPE);
                                    Trans1.AddNewlyCreatedDBObject(TWS_L_ON_EPE, true);
                                }


                            }


                            Trans1.Commit();
                        }
                    }


                }

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message + "\nProblem around station " + eRR_Stat + " on TWS Left on EPE");
            }
        }


        private void draw_ATWS_L_ON_EPE(Point3d Point1)
        {
            string eRR_Stat = "";
            try
            {
                string Column_ATWS_L_ON_EPE_Begin = "START";
                string Column_ATWS_L_ON_EPE_End = "END";

                string Column_ATWS_L_ON_EPE_Width_Begin = "START_WIDTH";
                string Column_ATWS_L_ON_EPE_Width_End = "END_WIDTH";


                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        if (ATWS_L_ON_EPE_TABLE != null & EPE_L != null)
                        {
                            if (ATWS_L_ON_EPE_TABLE.Rows.Count > 0)
                            {
                                Double Match1 = Convert.ToDouble(textBox_Matchline_start.Text);
                                Double Match2 = Convert.ToDouble(textBox_Matchline_end.Text);
                                if (Match1 > Match2)
                                {
                                    double T = Match1;
                                    Match1 = Match2;
                                    Match2 = T;
                                }

                                double Len1 = Match2 - Match1;// PolyCL_MS.Length;

                                Polyline PolyCL_PS = new Polyline();
                                PolyCL_PS.AddVertexAt(0, new Point2d(Point1.X, Point1.Y), 0, 0, 0);
                                PolyCL_PS.AddVertexAt(1, new Point2d(Point1.X + Len1, Point1.Y), 0, 0, 0);
                                PolyCL_PS.ColorIndex = 1;
                                PolyCL_PS.Elevation = 0;
                                BTrecord.AppendEntity(PolyCL_PS);
                                Trans1.AddNewlyCreatedDBObject(PolyCL_PS, true);

                                Point3dCollection[] Colectie_puncte_ATWS_L_ON_EPE;
                                int Index_ATWS_L_ON_EPE = 0;
                                Colectie_puncte_ATWS_L_ON_EPE = new Point3dCollection[Index_ATWS_L_ON_EPE + 1];
                                Colectie_puncte_ATWS_L_ON_EPE[Index_ATWS_L_ON_EPE] = new Point3dCollection();

                                Point3dCollection[] Colectie_puncte_sus_ATWS_L_ON_EPE;
                                int Index_sus_ATWS_L_ON_EPE = 0;
                                Colectie_puncte_sus_ATWS_L_ON_EPE = new Point3dCollection[Index_sus_ATWS_L_ON_EPE + 1];
                                Colectie_puncte_sus_ATWS_L_ON_EPE[Index_sus_ATWS_L_ON_EPE] = new Point3dCollection();




                                for (int i = 0; i < ATWS_L_ON_EPE_TABLE.Rows.Count; ++i)
                                {
                                    Double Start1 = -1;
                                    Double End1 = -1;
                                    Double Start_w1 = -1;
                                    Double End_w1 = -1;
                                    Double Start2 = -1;


                                    if ((ATWS_L_ON_EPE_TABLE.Rows[i][Column_ATWS_L_ON_EPE_Begin] != System.DBNull.Value)) Start1 = Convert.ToDouble(ATWS_L_ON_EPE_TABLE.Rows[i][Column_ATWS_L_ON_EPE_Begin]);
                                    if ((ATWS_L_ON_EPE_TABLE.Rows[i][Column_ATWS_L_ON_EPE_End] != System.DBNull.Value)) End1 = Convert.ToDouble(ATWS_L_ON_EPE_TABLE.Rows[i][Column_ATWS_L_ON_EPE_End]);
                                    if ((ATWS_L_ON_EPE_TABLE.Rows[i][Column_ATWS_L_ON_EPE_Width_Begin] != System.DBNull.Value)) Start_w1 = Convert.ToDouble(ATWS_L_ON_EPE_TABLE.Rows[i][Column_ATWS_L_ON_EPE_Width_Begin]);
                                    if ((ATWS_L_ON_EPE_TABLE.Rows[i][Column_ATWS_L_ON_EPE_Width_End] != System.DBNull.Value)) End_w1 = Convert.ToDouble(ATWS_L_ON_EPE_TABLE.Rows[i][Column_ATWS_L_ON_EPE_Width_End]);

                                    if (i + 1 < ATWS_L_ON_EPE_TABLE.Rows.Count)
                                    {
                                        if ((ATWS_L_ON_EPE_TABLE.Rows[i + 1][Column_ATWS_L_ON_EPE_Begin] != System.DBNull.Value))
                                        {
                                            Start2 = Convert.ToDouble(ATWS_L_ON_EPE_TABLE.Rows[i + 1][Column_ATWS_L_ON_EPE_Begin]);
                                        }
                                    }


                                    if (Start1 < Match1 & End1 <= Match2)
                                    {
                                        Start1 = Match1;
                                    }

                                    if (Start1 >= Match1 & Start1 <= Match2 & End1 >= Match2)
                                    {
                                        End1 = Match2;

                                    }

                                    if (Start1 <= Match1 & End1 >= Match2)
                                    {
                                        Start1 = Match1;
                                        End1 = Match2;

                                    }
                                    if (Start1 >= Match1 & End1 <= Match2)
                                    {
                                        if (Start1 != -1 & End1 != -1 & Start_w1 != -1 & End_w1 != -1)
                                        {
                                            Point3d PointCL_PS_start1 = new Point3d();
                                            PointCL_PS_start1 = PolyCL_PS.GetPointAtDist(Start1);
                                            Point3d PointCL_PS_end1 = new Point3d();
                                            PointCL_PS_end1 = PolyCL_PS.GetPointAtDist(End1);
                                            Line Linie_int_start = new Line(PointCL_PS_start1, new Point3d(PointCL_PS_start1.X, PointCL_PS_start1.Y + 1000, 0));
                                            Line Linie_int_end = new Line(PointCL_PS_end1, new Point3d(PointCL_PS_end1.X, PointCL_PS_end1.Y + 1000, 0));
                                            //BTrecord.AppendEntity(Linie_int_start);
                                            //BTrecord.AppendEntity(Linie_int_end);
                                            //Trans1.AddNewlyCreatedDBObject(Linie_int_end, true);
                                            //Trans1.AddNewlyCreatedDBObject(Linie_int_start, true);

                                            Point3dCollection Col_int_start = new Point3dCollection();
                                            Point3dCollection Col_int_end = new Point3dCollection();
                                            Point3dCollection Colectie_puncte_poly = new Point3dCollection();


                                            EPE_L.IntersectWith(Linie_int_start, Autodesk.AutoCAD.DatabaseServices.Intersect.ExtendThis, Col_int_start, IntPtr.Zero, IntPtr.Zero);
                                            EPE_L.IntersectWith(Linie_int_end, Autodesk.AutoCAD.DatabaseServices.Intersect.ExtendThis, Col_int_end, IntPtr.Zero, IntPtr.Zero);

                                            Colectie_puncte_ATWS_L_ON_EPE[Index_ATWS_L_ON_EPE].Add(Col_int_start[0]);
                                            Colectie_puncte_poly.Add(Col_int_start[0]);
                                            double Param_start = EPE_L.GetParameterAtPoint(Col_int_start[0]);
                                            double Param_end = EPE_L.GetParameterAtPoint(Col_int_end[0]);

                                            //Line TEST_int_start = new Line(PointCL_PS_start1, Col_int_start[0]);
                                            //Line TEST_int_end = new Line(PointCL_PS_end1, Col_int_end[0]);
                                            //BTrecord.AppendEntity(TEST_int_start);
                                            //BTrecord.AppendEntity(TEST_int_end);
                                            //Trans1.AddNewlyCreatedDBObject(TEST_int_end, true);
                                            //Trans1.AddNewlyCreatedDBObject(TEST_int_start, true);

                                            if (Math.Floor(Param_end) - Math.Floor(Param_start) >= 1)
                                            {

                                                int P1 = (int)Math.Ceiling(Param_start);
                                                int P2 = (int)Math.Floor(Param_end);

                                                for (int j = P1; j <= P2; ++j)
                                                {
                                                    Colectie_puncte_ATWS_L_ON_EPE[Index_ATWS_L_ON_EPE].Add(EPE_L.GetPointAtParameter(j));
                                                    Colectie_puncte_poly.Add(EPE_L.GetPointAtParameter(j));
                                                }
                                            }

                                            Colectie_puncte_ATWS_L_ON_EPE[Index_ATWS_L_ON_EPE].Add(Col_int_end[0]);
                                            Colectie_puncte_poly.Add(Col_int_end[0]);

                                            if (Start_w1 == End_w1)
                                            {

                                                Polyline Temp_poly = new Polyline();

                                                for (int j = 0; j < Colectie_puncte_poly.Count; ++j)
                                                {
                                                    Temp_poly.AddVertexAt(j, new Point2d(Colectie_puncte_poly[j].X, Colectie_puncte_poly[j].Y), 0, 0, 0);
                                                }

                                                //BTrecord.AppendEntity(Temp_poly);
                                                //Trans1.AddNewlyCreatedDBObject(Temp_poly, true);

                                                Autodesk.AutoCAD.DatabaseServices.DBObjectCollection Colectie_offset_ATWS_L_ON_EPE = Temp_poly.GetOffsetCurves(-Start_w1);

                                                Polyline Temp_poly_sus = new Polyline();

                                                foreach (Polyline obj in Colectie_offset_ATWS_L_ON_EPE)
                                                {
                                                    if (obj != null)
                                                    {
                                                        Temp_poly_sus = obj;
                                                    }

                                                    if (Temp_poly_sus != null)
                                                    {
                                                        for (int j = 0; j < Temp_poly_sus.NumberOfVertices; ++j)
                                                        {
                                                            Colectie_puncte_sus_ATWS_L_ON_EPE[Index_sus_ATWS_L_ON_EPE].Add(Temp_poly_sus.GetPointAtParameter(j));
                                                        }
                                                    }
                                                }
                                                //Temp_poly_sus.ColorIndex = 2;
                                                //BTrecord.AppendEntity(Temp_poly_sus);
                                                //Trans1.AddNewlyCreatedDBObject(Temp_poly_sus, true);
                                            }
                                            else
                                            {
                                                Colectie_puncte_sus_ATWS_L_ON_EPE[Index_sus_ATWS_L_ON_EPE].Add(new Point3d(Col_int_start[0].X, Col_int_start[0].Y + Start_w1, 0));
                                                Colectie_puncte_sus_ATWS_L_ON_EPE[Index_sus_ATWS_L_ON_EPE].Add(new Point3d(Col_int_end[0].X, Col_int_end[0].Y + End_w1, 0));
                                            }

                                            if (Start2 != End1)
                                            {
                                                Index_ATWS_L_ON_EPE = Index_ATWS_L_ON_EPE + 1;
                                                Array.Resize(ref Colectie_puncte_ATWS_L_ON_EPE, Index_ATWS_L_ON_EPE + 1);
                                                Colectie_puncte_ATWS_L_ON_EPE[Index_ATWS_L_ON_EPE] = new Point3dCollection();

                                                Index_sus_ATWS_L_ON_EPE = Index_sus_ATWS_L_ON_EPE + 1;
                                                Array.Resize(ref Colectie_puncte_sus_ATWS_L_ON_EPE, Index_sus_ATWS_L_ON_EPE + 1);
                                                Colectie_puncte_sus_ATWS_L_ON_EPE[Index_sus_ATWS_L_ON_EPE] = new Point3dCollection();

                                            }


                                        }
                                    }
                                }

                                for (int i = 0; i <= Index_ATWS_L_ON_EPE; ++i)
                                {
                                    Polyline ATWS_L_ON_EPE = new Polyline();
                                    int Idx1 = 0;
                                    for (int j = 0; j < Colectie_puncte_sus_ATWS_L_ON_EPE[i].Count; ++j)
                                    {
                                        ATWS_L_ON_EPE.AddVertexAt(Idx1, new Point2d(Colectie_puncte_sus_ATWS_L_ON_EPE[i][j].X, Colectie_puncte_sus_ATWS_L_ON_EPE[i][j].Y), 0, 0, 0);
                                        Idx1 = Idx1 + 1;

                                    }
                                    for (int j = Colectie_puncte_ATWS_L_ON_EPE[i].Count - 1; j >= 0; --j)
                                    {
                                        ATWS_L_ON_EPE.AddVertexAt(Idx1, new Point2d(Colectie_puncte_ATWS_L_ON_EPE[i][j].X, Colectie_puncte_ATWS_L_ON_EPE[i][j].Y), 0, 0, 0);
                                        Idx1 = Idx1 + 1;
                                    }

                                    ATWS_L_ON_EPE.ColorIndex = 253;
                                    ATWS_L_ON_EPE.Closed = true;
                                    BTrecord.AppendEntity(ATWS_L_ON_EPE);
                                    Trans1.AddNewlyCreatedDBObject(ATWS_L_ON_EPE, true);
                                }


                            }


                            Trans1.Commit();
                        }
                    }


                }

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message + "\nProblem around station " + eRR_Stat + " on ATWS Left on EPE");
            }
        }


        private void draw_TWS_R_ON_EPE(Point3d Point1)
        {
            string eRR_Stat = "";
            try
            {
                string Column_TWS_R_ON_EPE_Begin = "START";
                string Column_TWS_R_ON_EPE_End = "END";

                string Column_TWS_R_ON_EPE_Width_Begin = "START_WIDTH";
                string Column_TWS_R_ON_EPE_Width_End = "END_WIDTH";


                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        if (TWS_R_ON_EPE_TABLE != null & EPE_R != null)
                        {
                            if (TWS_R_ON_EPE_TABLE.Rows.Count > 0)
                            {
                                Double Match1 = Convert.ToDouble(textBox_Matchline_start.Text);
                                Double Match2 = Convert.ToDouble(textBox_Matchline_end.Text);
                                if (Match1 > Match2)
                                {
                                    double T = Match1;
                                    Match1 = Match2;
                                    Match2 = T;
                                }

                                double Len1 = Match2 - Match1;// PolyCL_MS.Length;

                                Polyline PolyCL_PS = new Polyline();
                                PolyCL_PS.AddVertexAt(0, new Point2d(Point1.X, Point1.Y), 0, 0, 0);
                                PolyCL_PS.AddVertexAt(1, new Point2d(Point1.X + Len1, Point1.Y), 0, 0, 0);
                                PolyCL_PS.ColorIndex = 1;
                                PolyCL_PS.Elevation = 0;
                                BTrecord.AppendEntity(PolyCL_PS);
                                Trans1.AddNewlyCreatedDBObject(PolyCL_PS, true);

                                Point3dCollection[] Colectie_puncte_TWS_R_ON_EPE;
                                int Index_TWS_R_ON_EPE = 0;
                                Colectie_puncte_TWS_R_ON_EPE = new Point3dCollection[Index_TWS_R_ON_EPE + 1];
                                Colectie_puncte_TWS_R_ON_EPE[Index_TWS_R_ON_EPE] = new Point3dCollection();

                                Point3dCollection[] Colectie_puncte_sus_TWS_R_ON_EPE;
                                int Index_sus_TWS_R_ON_EPE = 0;
                                Colectie_puncte_sus_TWS_R_ON_EPE = new Point3dCollection[Index_sus_TWS_R_ON_EPE + 1];
                                Colectie_puncte_sus_TWS_R_ON_EPE[Index_sus_TWS_R_ON_EPE] = new Point3dCollection();




                                for (int i = 0; i < TWS_R_ON_EPE_TABLE.Rows.Count; ++i)
                                {
                                    Double Start1 = -1;
                                    Double End1 = -1;
                                    Double Start_w1 = -1;
                                    Double End_w1 = -1;
                                    Double Start2 = -1;


                                    if ((TWS_R_ON_EPE_TABLE.Rows[i][Column_TWS_R_ON_EPE_Begin] != System.DBNull.Value)) Start1 = Convert.ToDouble(TWS_R_ON_EPE_TABLE.Rows[i][Column_TWS_R_ON_EPE_Begin]);
                                    if ((TWS_R_ON_EPE_TABLE.Rows[i][Column_TWS_R_ON_EPE_End] != System.DBNull.Value)) End1 = Convert.ToDouble(TWS_R_ON_EPE_TABLE.Rows[i][Column_TWS_R_ON_EPE_End]);
                                    if ((TWS_R_ON_EPE_TABLE.Rows[i][Column_TWS_R_ON_EPE_Width_Begin] != System.DBNull.Value)) Start_w1 = Convert.ToDouble(TWS_R_ON_EPE_TABLE.Rows[i][Column_TWS_R_ON_EPE_Width_Begin]);
                                    if ((TWS_R_ON_EPE_TABLE.Rows[i][Column_TWS_R_ON_EPE_Width_End] != System.DBNull.Value)) End_w1 = Convert.ToDouble(TWS_R_ON_EPE_TABLE.Rows[i][Column_TWS_R_ON_EPE_Width_End]);

                                    if (i + 1 < TWS_R_ON_EPE_TABLE.Rows.Count)
                                    {
                                        if ((TWS_R_ON_EPE_TABLE.Rows[i + 1][Column_TWS_R_ON_EPE_Begin] != System.DBNull.Value))
                                        {
                                            Start2 = Convert.ToDouble(TWS_R_ON_EPE_TABLE.Rows[i + 1][Column_TWS_R_ON_EPE_Begin]);
                                        }
                                    }


                                    if (Start1 < Match1 & End1 <= Match2)
                                    {
                                        Start1 = Match1;
                                    }

                                    if (Start1 >= Match1 & Start1 <= Match2 & End1 >= Match2)
                                    {
                                        End1 = Match2;

                                    }

                                    if (Start1 <= Match1 & End1 >= Match2)
                                    {
                                        Start1 = Match1;
                                        End1 = Match2;

                                    }
                                    if (Start1 >= Match1 & End1 <= Match2)
                                    {
                                        if (Start1 != -1 & End1 != -1 & Start_w1 != -1 & End_w1 != -1)
                                        {
                                            Point3d PointCL_PS_start1 = new Point3d();
                                            PointCL_PS_start1 = PolyCL_PS.GetPointAtDist(Start1);
                                            Point3d PointCL_PS_end1 = new Point3d();
                                            PointCL_PS_end1 = PolyCL_PS.GetPointAtDist(End1);
                                            Line Linie_int_start = new Line(PointCL_PS_start1, new Point3d(PointCL_PS_start1.X, PointCL_PS_start1.Y - 1000, 0));
                                            Line Linie_int_end = new Line(PointCL_PS_end1, new Point3d(PointCL_PS_end1.X, PointCL_PS_end1.Y - 1000, 0));
                                            //BTrecord.AppendEntity(Linie_int_start);
                                            //BTrecord.AppendEntity(Linie_int_end);
                                            //Trans1.AddNewlyCreatedDBObject(Linie_int_end, true);
                                            //Trans1.AddNewlyCreatedDBObject(Linie_int_start, true);

                                            Point3dCollection Col_int_start = new Point3dCollection();
                                            Point3dCollection Col_int_end = new Point3dCollection();
                                            Point3dCollection Colectie_puncte_poly = new Point3dCollection();

                                            EPE_R.IntersectWith(Linie_int_start, Autodesk.AutoCAD.DatabaseServices.Intersect.ExtendThis, Col_int_start, IntPtr.Zero, IntPtr.Zero);
                                            EPE_R.IntersectWith(Linie_int_end, Autodesk.AutoCAD.DatabaseServices.Intersect.ExtendThis, Col_int_end, IntPtr.Zero, IntPtr.Zero);

                                            Colectie_puncte_TWS_R_ON_EPE[Index_TWS_R_ON_EPE].Add(Col_int_start[0]);
                                            Colectie_puncte_poly.Add(Col_int_start[0]);

                                            double Param_start = EPE_R.GetParameterAtPoint(Col_int_start[0]);
                                            double Param_end = EPE_R.GetParameterAtPoint(Col_int_end[0]);

                                            //Line TEST_int_start = new Line(PointCL_PS_start1, Col_int_start[0]);
                                            //Line TEST_int_end = new Line(PointCL_PS_end1, Col_int_end[0]);
                                            //BTrecord.AppendEntity(TEST_int_start);
                                            //BTrecord.AppendEntity(TEST_int_end);
                                            //Trans1.AddNewlyCreatedDBObject(TEST_int_end, true);
                                            //Trans1.AddNewlyCreatedDBObject(TEST_int_start, true);

                                            if (Math.Floor(Param_end) - Math.Floor(Param_start) >= 1)
                                            {

                                                int P1 = (int)Math.Ceiling(Param_start);
                                                int P2 = (int)Math.Floor(Param_end);

                                                for (int j = P1; j <= P2; ++j)
                                                {
                                                    Colectie_puncte_TWS_R_ON_EPE[Index_TWS_R_ON_EPE].Add(EPE_R.GetPointAtParameter(j));
                                                    Colectie_puncte_poly.Add(EPE_R.GetPointAtParameter(j));
                                                }
                                            }

                                            Colectie_puncte_TWS_R_ON_EPE[Index_TWS_R_ON_EPE].Add(Col_int_end[0]);
                                            Colectie_puncte_poly.Add(Col_int_end[0]);

                                            if (Start_w1 == End_w1)
                                            {

                                                Polyline Temp_poly = new Polyline();

                                                for (int j = 0; j < Colectie_puncte_poly.Count; ++j)
                                                {
                                                    Temp_poly.AddVertexAt(j, new Point2d(Colectie_puncte_poly[j].X, Colectie_puncte_poly[j].Y), 0, 0, 0);
                                                }

                                                //BTrecord.AppendEntity(Temp_poly);
                                                //Trans1.AddNewlyCreatedDBObject(Temp_poly, true);

                                                Autodesk.AutoCAD.DatabaseServices.DBObjectCollection Colectie_offset_TWS_R_ON_EPE = Temp_poly.GetOffsetCurves(Start_w1);

                                                Polyline Temp_poly_sus = new Polyline();

                                                foreach (Polyline obj in Colectie_offset_TWS_R_ON_EPE)
                                                {
                                                    if (obj != null)
                                                    {
                                                        Temp_poly_sus = obj;
                                                    }

                                                    if (Temp_poly_sus != null)
                                                    {
                                                        for (int j = 0; j < Temp_poly_sus.NumberOfVertices; ++j)
                                                        {
                                                            Colectie_puncte_sus_TWS_R_ON_EPE[Index_sus_TWS_R_ON_EPE].Add(Temp_poly_sus.GetPointAtParameter(j));
                                                        }
                                                    }
                                                }
                                                //BTrecord.AppendEntity(Temp_poly_sus);
                                                //Trans1.AddNewlyCreatedDBObject(Temp_poly_sus, true);
                                            }
                                            else
                                            {
                                                Colectie_puncte_sus_TWS_R_ON_EPE[Index_sus_TWS_R_ON_EPE].Add(new Point3d(Col_int_start[0].X, Col_int_start[0].Y - Start_w1, 0));
                                                Colectie_puncte_sus_TWS_R_ON_EPE[Index_sus_TWS_R_ON_EPE].Add(new Point3d(Col_int_end[0].X, Col_int_end[0].Y - End_w1, 0));
                                            }

                                            if (Start2 != End1)
                                            {
                                                Index_TWS_R_ON_EPE = Index_TWS_R_ON_EPE + 1;
                                                Array.Resize(ref Colectie_puncte_TWS_R_ON_EPE, Index_TWS_R_ON_EPE + 1);
                                                Colectie_puncte_TWS_R_ON_EPE[Index_TWS_R_ON_EPE] = new Point3dCollection();

                                                Index_sus_TWS_R_ON_EPE = Index_sus_TWS_R_ON_EPE + 1;
                                                Array.Resize(ref Colectie_puncte_sus_TWS_R_ON_EPE, Index_sus_TWS_R_ON_EPE + 1);
                                                Colectie_puncte_sus_TWS_R_ON_EPE[Index_sus_TWS_R_ON_EPE] = new Point3dCollection();

                                            }


                                        }
                                    }
                                }

                                TWS_R = new Polyline();
                                int IdxX = 0;

                                for (int i = 0; i <= Index_TWS_R_ON_EPE; ++i)
                                {
                                    Polyline TWS_R_ON_EPE = new Polyline();
                                    int Idx1 = 0;
                                    for (int j = 0; j < Colectie_puncte_sus_TWS_R_ON_EPE[i].Count; ++j)
                                    {
                                        TWS_R_ON_EPE.AddVertexAt(Idx1, new Point2d(Colectie_puncte_sus_TWS_R_ON_EPE[i][j].X, Colectie_puncte_sus_TWS_R_ON_EPE[i][j].Y), 0, 0, 0);
                                        Idx1 = Idx1 + 1;
                                        TWS_R.AddVertexAt(IdxX, new Point2d(Colectie_puncte_sus_TWS_R_ON_EPE[i][j].X, Colectie_puncte_sus_TWS_R_ON_EPE[i][j].Y), 0, 0, 0);
                                        IdxX = IdxX + 1;

                                    }
                                    for (int j = Colectie_puncte_TWS_R_ON_EPE[i].Count - 1; j >= 0; --j)
                                    {
                                        TWS_R_ON_EPE.AddVertexAt(Idx1, new Point2d(Colectie_puncte_TWS_R_ON_EPE[i][j].X, Colectie_puncte_TWS_R_ON_EPE[i][j].Y), 0, 0, 0);
                                        Idx1 = Idx1 + 1;


                                    }

                                    TWS_R_ON_EPE.ColorIndex = 3;
                                    TWS_R_ON_EPE.Closed = true;
                                    BTrecord.AppendEntity(TWS_R_ON_EPE);
                                    Trans1.AddNewlyCreatedDBObject(TWS_R_ON_EPE, true);
                                }


                            }


                            Trans1.Commit();
                        }
                    }


                }

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message + "\nProblem around station " + eRR_Stat + " on TWS Right on EPE");
            }
        }

        private void draw_ATWS_R_ON_EPE(Point3d Point1)
        {
            string eRR_Stat = "";
            try
            {
                string Column_ATWS_R_ON_EPE_Begin = "START";
                string Column_ATWS_R_ON_EPE_End = "END";

                string Column_ATWS_R_ON_EPE_Width_Begin = "START_WIDTH";
                string Column_ATWS_R_ON_EPE_Width_End = "END_WIDTH";


                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        if (ATWS_R_ON_EPE_TABLE != null & EPE_R != null)
                        {
                            if (ATWS_R_ON_EPE_TABLE.Rows.Count > 0)
                            {
                                Double Match1 = Convert.ToDouble(textBox_Matchline_start.Text);
                                Double Match2 = Convert.ToDouble(textBox_Matchline_end.Text);
                                if (Match1 > Match2)
                                {
                                    double T = Match1;
                                    Match1 = Match2;
                                    Match2 = T;
                                }

                                double Len1 = Match2 - Match1;// PolyCL_MS.Length;

                                Polyline PolyCL_PS = new Polyline();
                                PolyCL_PS.AddVertexAt(0, new Point2d(Point1.X, Point1.Y), 0, 0, 0);
                                PolyCL_PS.AddVertexAt(1, new Point2d(Point1.X + Len1, Point1.Y), 0, 0, 0);
                                PolyCL_PS.ColorIndex = 1;
                                PolyCL_PS.Elevation = 0;
                                BTrecord.AppendEntity(PolyCL_PS);
                                Trans1.AddNewlyCreatedDBObject(PolyCL_PS, true);

                                Point3dCollection[] Colectie_puncte_ATWS_R_ON_EPE;
                                int Index_ATWS_R_ON_EPE = 0;
                                Colectie_puncte_ATWS_R_ON_EPE = new Point3dCollection[Index_ATWS_R_ON_EPE + 1];
                                Colectie_puncte_ATWS_R_ON_EPE[Index_ATWS_R_ON_EPE] = new Point3dCollection();

                                Point3dCollection[] Colectie_puncte_sus_ATWS_R_ON_EPE;
                                int Index_sus_ATWS_R_ON_EPE = 0;
                                Colectie_puncte_sus_ATWS_R_ON_EPE = new Point3dCollection[Index_sus_ATWS_R_ON_EPE + 1];
                                Colectie_puncte_sus_ATWS_R_ON_EPE[Index_sus_ATWS_R_ON_EPE] = new Point3dCollection();




                                for (int i = 0; i < ATWS_R_ON_EPE_TABLE.Rows.Count; ++i)
                                {
                                    Double Start1 = -1;
                                    Double End1 = -1;
                                    Double Start_w1 = -1;
                                    Double End_w1 = -1;
                                    Double Start2 = -1;


                                    if ((ATWS_R_ON_EPE_TABLE.Rows[i][Column_ATWS_R_ON_EPE_Begin] != System.DBNull.Value)) Start1 = Convert.ToDouble(ATWS_R_ON_EPE_TABLE.Rows[i][Column_ATWS_R_ON_EPE_Begin]);
                                    if ((ATWS_R_ON_EPE_TABLE.Rows[i][Column_ATWS_R_ON_EPE_End] != System.DBNull.Value)) End1 = Convert.ToDouble(ATWS_R_ON_EPE_TABLE.Rows[i][Column_ATWS_R_ON_EPE_End]);
                                    if ((ATWS_R_ON_EPE_TABLE.Rows[i][Column_ATWS_R_ON_EPE_Width_Begin] != System.DBNull.Value)) Start_w1 = Convert.ToDouble(ATWS_R_ON_EPE_TABLE.Rows[i][Column_ATWS_R_ON_EPE_Width_Begin]);
                                    if ((ATWS_R_ON_EPE_TABLE.Rows[i][Column_ATWS_R_ON_EPE_Width_End] != System.DBNull.Value)) End_w1 = Convert.ToDouble(ATWS_R_ON_EPE_TABLE.Rows[i][Column_ATWS_R_ON_EPE_Width_End]);

                                    if (i + 1 < ATWS_R_ON_EPE_TABLE.Rows.Count)
                                    {
                                        if ((ATWS_R_ON_EPE_TABLE.Rows[i + 1][Column_ATWS_R_ON_EPE_Begin] != System.DBNull.Value))
                                        {
                                            Start2 = Convert.ToDouble(ATWS_R_ON_EPE_TABLE.Rows[i + 1][Column_ATWS_R_ON_EPE_Begin]);
                                        }
                                    }


                                    if (Start1 < Match1 & End1 <= Match2)
                                    {
                                        Start1 = Match1;
                                    }

                                    if (Start1 >= Match1 & Start1 <= Match2 & End1 >= Match2)
                                    {
                                        End1 = Match2;

                                    }

                                    if (Start1 <= Match1 & End1 >= Match2)
                                    {
                                        Start1 = Match1;
                                        End1 = Match2;

                                    }
                                    if (Start1 >= Match1 & End1 <= Match2)
                                    {
                                        if (Start1 != -1 & End1 != -1 & Start_w1 != -1 & End_w1 != -1)
                                        {
                                            Point3d PointCL_PS_start1 = new Point3d();
                                            PointCL_PS_start1 = PolyCL_PS.GetPointAtDist(Start1);
                                            Point3d PointCL_PS_end1 = new Point3d();
                                            PointCL_PS_end1 = PolyCL_PS.GetPointAtDist(End1);
                                            Line Linie_int_start = new Line(PointCL_PS_start1, new Point3d(PointCL_PS_start1.X, PointCL_PS_start1.Y - 1000, 0));
                                            Line Linie_int_end = new Line(PointCL_PS_end1, new Point3d(PointCL_PS_end1.X, PointCL_PS_end1.Y - 1000, 0));
                                            //BTrecord.AppendEntity(Linie_int_start);
                                            //BTrecord.AppendEntity(Linie_int_end);
                                            //Trans1.AddNewlyCreatedDBObject(Linie_int_end, true);
                                            //Trans1.AddNewlyCreatedDBObject(Linie_int_start, true);

                                            Point3dCollection Col_int_start = new Point3dCollection();
                                            Point3dCollection Col_int_end = new Point3dCollection();
                                            Point3dCollection Colectie_puncte_poly = new Point3dCollection();

                                            EPE_R.IntersectWith(Linie_int_start, Autodesk.AutoCAD.DatabaseServices.Intersect.ExtendThis, Col_int_start, IntPtr.Zero, IntPtr.Zero);
                                            EPE_R.IntersectWith(Linie_int_end, Autodesk.AutoCAD.DatabaseServices.Intersect.ExtendThis, Col_int_end, IntPtr.Zero, IntPtr.Zero);

                                            Colectie_puncte_ATWS_R_ON_EPE[Index_ATWS_R_ON_EPE].Add(Col_int_start[0]);
                                            Colectie_puncte_poly.Add(Col_int_start[0]);

                                            double Param_start = EPE_R.GetParameterAtPoint(Col_int_start[0]);
                                            double Param_end = EPE_R.GetParameterAtPoint(Col_int_end[0]);

                                            //Line TEST_int_start = new Line(PointCL_PS_start1, Col_int_start[0]);
                                            //Line TEST_int_end = new Line(PointCL_PS_end1, Col_int_end[0]);
                                            //BTrecord.AppendEntity(TEST_int_start);
                                            //BTrecord.AppendEntity(TEST_int_end);
                                            //Trans1.AddNewlyCreatedDBObject(TEST_int_end, true);
                                            //Trans1.AddNewlyCreatedDBObject(TEST_int_start, true);

                                            if (Math.Floor(Param_end) - Math.Floor(Param_start) >= 1)
                                            {

                                                int P1 = (int)Math.Ceiling(Param_start);
                                                int P2 = (int)Math.Floor(Param_end);

                                                for (int j = P1; j <= P2; ++j)
                                                {
                                                    Colectie_puncte_ATWS_R_ON_EPE[Index_ATWS_R_ON_EPE].Add(EPE_R.GetPointAtParameter(j));
                                                    Colectie_puncte_poly.Add(EPE_R.GetPointAtParameter(j));
                                                }
                                            }

                                            Colectie_puncte_ATWS_R_ON_EPE[Index_ATWS_R_ON_EPE].Add(Col_int_end[0]);
                                            Colectie_puncte_poly.Add(Col_int_end[0]);

                                            if (Start_w1 == End_w1)
                                            {

                                                Polyline Temp_poly = new Polyline();

                                                for (int j = 0; j < Colectie_puncte_poly.Count; ++j)
                                                {
                                                    Temp_poly.AddVertexAt(j, new Point2d(Colectie_puncte_poly[j].X, Colectie_puncte_poly[j].Y), 0, 0, 0);
                                                }

                                                //BTrecord.AppendEntity(Temp_poly);
                                                //Trans1.AddNewlyCreatedDBObject(Temp_poly, true);

                                                Autodesk.AutoCAD.DatabaseServices.DBObjectCollection Colectie_offset_ATWS_R_ON_EPE = Temp_poly.GetOffsetCurves(Start_w1);

                                                Polyline Temp_poly_sus = new Polyline();

                                                foreach (Polyline obj in Colectie_offset_ATWS_R_ON_EPE)
                                                {
                                                    if (obj != null)
                                                    {
                                                        Temp_poly_sus = obj;
                                                    }

                                                    if (Temp_poly_sus != null)
                                                    {
                                                        for (int j = 0; j < Temp_poly_sus.NumberOfVertices; ++j)
                                                        {
                                                            Colectie_puncte_sus_ATWS_R_ON_EPE[Index_sus_ATWS_R_ON_EPE].Add(Temp_poly_sus.GetPointAtParameter(j));
                                                        }
                                                    }
                                                }
                                                //BTrecord.AppendEntity(Temp_poly_sus);
                                                //Trans1.AddNewlyCreatedDBObject(Temp_poly_sus, true);
                                            }
                                            else
                                            {
                                                Colectie_puncte_sus_ATWS_R_ON_EPE[Index_sus_ATWS_R_ON_EPE].Add(new Point3d(Col_int_start[0].X, Col_int_start[0].Y - Start_w1, 0));
                                                Colectie_puncte_sus_ATWS_R_ON_EPE[Index_sus_ATWS_R_ON_EPE].Add(new Point3d(Col_int_end[0].X, Col_int_end[0].Y - End_w1, 0));
                                            }

                                            if (Start2 != End1)
                                            {
                                                Index_ATWS_R_ON_EPE = Index_ATWS_R_ON_EPE + 1;
                                                Array.Resize(ref Colectie_puncte_ATWS_R_ON_EPE, Index_ATWS_R_ON_EPE + 1);
                                                Colectie_puncte_ATWS_R_ON_EPE[Index_ATWS_R_ON_EPE] = new Point3dCollection();

                                                Index_sus_ATWS_R_ON_EPE = Index_sus_ATWS_R_ON_EPE + 1;
                                                Array.Resize(ref Colectie_puncte_sus_ATWS_R_ON_EPE, Index_sus_ATWS_R_ON_EPE + 1);
                                                Colectie_puncte_sus_ATWS_R_ON_EPE[Index_sus_ATWS_R_ON_EPE] = new Point3dCollection();

                                            }


                                        }
                                    }
                                }

                                for (int i = 0; i <= Index_ATWS_R_ON_EPE; ++i)
                                {
                                    Polyline ATWS_R_ON_EPE = new Polyline();
                                    int Idx1 = 0;
                                    for (int j = 0; j < Colectie_puncte_sus_ATWS_R_ON_EPE[i].Count; ++j)
                                    {
                                        ATWS_R_ON_EPE.AddVertexAt(Idx1, new Point2d(Colectie_puncte_sus_ATWS_R_ON_EPE[i][j].X, Colectie_puncte_sus_ATWS_R_ON_EPE[i][j].Y), 0, 0, 0);
                                        Idx1 = Idx1 + 1;

                                    }
                                    for (int j = Colectie_puncte_ATWS_R_ON_EPE[i].Count - 1; j >= 0; --j)
                                    {
                                        ATWS_R_ON_EPE.AddVertexAt(Idx1, new Point2d(Colectie_puncte_ATWS_R_ON_EPE[i][j].X, Colectie_puncte_ATWS_R_ON_EPE[i][j].Y), 0, 0, 0);
                                        Idx1 = Idx1 + 1;
                                    }

                                    ATWS_R_ON_EPE.ColorIndex = 253;
                                    ATWS_R_ON_EPE.Closed = true;
                                    BTrecord.AppendEntity(ATWS_R_ON_EPE);
                                    Trans1.AddNewlyCreatedDBObject(ATWS_R_ON_EPE, true);
                                }


                            }


                            Trans1.Commit();
                        }
                    }


                }

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message + "\nProblem around station " + eRR_Stat + " on ATWS Right on EPE");
            }
        }


        private void draw_ATWS_L_ON_TWS(Point3d Point1)
        {
            string eRR_Stat = "";
            try
            {
                string Column_ATWS_L_ON_TWS_Begin = "START";
                string Column_ATWS_L_ON_TWS_End = "END";

                string Column_ATWS_L_ON_TWS_Width_Begin = "START_WIDTH";
                string Column_ATWS_L_ON_TWS_Width_End = "END_WIDTH";


                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        if (ATWS_L_ON_TWS_TABLE != null & TWS_L != null)
                        {
                            if (ATWS_L_ON_TWS_TABLE.Rows.Count > 0)
                            {
                                Double Match1 = Convert.ToDouble(textBox_Matchline_start.Text);
                                Double Match2 = Convert.ToDouble(textBox_Matchline_end.Text);
                                if (Match1 > Match2)
                                {
                                    double T = Match1;
                                    Match1 = Match2;
                                    Match2 = T;
                                }

                                double Len1 = Match2 - Match1;// PolyCL_MS.Length;

                                Polyline PolyCL_PS = new Polyline();
                                PolyCL_PS.AddVertexAt(0, new Point2d(Point1.X, Point1.Y), 0, 0, 0);
                                PolyCL_PS.AddVertexAt(1, new Point2d(Point1.X + Len1, Point1.Y), 0, 0, 0);
                                PolyCL_PS.ColorIndex = 1;
                                PolyCL_PS.Elevation = 0;
                                BTrecord.AppendEntity(PolyCL_PS);
                                Trans1.AddNewlyCreatedDBObject(PolyCL_PS, true);

                                Point3dCollection[] Colectie_puncte_ATWS_L_ON_TWS;
                                int Index_ATWS_L_ON_TWS = 0;
                                Colectie_puncte_ATWS_L_ON_TWS = new Point3dCollection[Index_ATWS_L_ON_TWS + 1];
                                Colectie_puncte_ATWS_L_ON_TWS[Index_ATWS_L_ON_TWS] = new Point3dCollection();

                                Point3dCollection[] Colectie_puncte_sus_ATWS_L_ON_TWS;
                                int Index_sus_ATWS_L_ON_TWS = 0;
                                Colectie_puncte_sus_ATWS_L_ON_TWS = new Point3dCollection[Index_sus_ATWS_L_ON_TWS + 1];
                                Colectie_puncte_sus_ATWS_L_ON_TWS[Index_sus_ATWS_L_ON_TWS] = new Point3dCollection();




                                for (int i = 0; i < ATWS_L_ON_TWS_TABLE.Rows.Count; ++i)
                                {
                                    Double Start1 = -1;
                                    Double End1 = -1;
                                    Double Start_w1 = -1;
                                    Double End_w1 = -1;
                                    Double Start2 = -1;


                                    if ((ATWS_L_ON_TWS_TABLE.Rows[i][Column_ATWS_L_ON_TWS_Begin] != System.DBNull.Value)) Start1 = Convert.ToDouble(ATWS_L_ON_TWS_TABLE.Rows[i][Column_ATWS_L_ON_TWS_Begin]);
                                    if ((ATWS_L_ON_TWS_TABLE.Rows[i][Column_ATWS_L_ON_TWS_End] != System.DBNull.Value)) End1 = Convert.ToDouble(ATWS_L_ON_TWS_TABLE.Rows[i][Column_ATWS_L_ON_TWS_End]);
                                    if ((ATWS_L_ON_TWS_TABLE.Rows[i][Column_ATWS_L_ON_TWS_Width_Begin] != System.DBNull.Value)) Start_w1 = Convert.ToDouble(ATWS_L_ON_TWS_TABLE.Rows[i][Column_ATWS_L_ON_TWS_Width_Begin]);
                                    if ((ATWS_L_ON_TWS_TABLE.Rows[i][Column_ATWS_L_ON_TWS_Width_End] != System.DBNull.Value)) End_w1 = Convert.ToDouble(ATWS_L_ON_TWS_TABLE.Rows[i][Column_ATWS_L_ON_TWS_Width_End]);

                                    if (i + 1 < ATWS_L_ON_TWS_TABLE.Rows.Count)
                                    {
                                        if ((ATWS_L_ON_TWS_TABLE.Rows[i + 1][Column_ATWS_L_ON_TWS_Begin] != System.DBNull.Value))
                                        {
                                            Start2 = Convert.ToDouble(ATWS_L_ON_TWS_TABLE.Rows[i + 1][Column_ATWS_L_ON_TWS_Begin]);
                                        }
                                    }


                                    if (Start1 < Match1 & End1 <= Match2)
                                    {
                                        Start1 = Match1;
                                    }

                                    if (Start1 >= Match1 & Start1 <= Match2 & End1 >= Match2)
                                    {
                                        End1 = Match2;

                                    }

                                    if (Start1 <= Match1 & End1 >= Match2)
                                    {
                                        Start1 = Match1;
                                        End1 = Match2;

                                    }
                                    if (Start1 >= Match1 & End1 <= Match2)
                                    {
                                        if (Start1 != -1 & End1 != -1 & Start_w1 != -1 & End_w1 != -1)
                                        {
                                            Point3d PointCL_PS_start1 = new Point3d();
                                            PointCL_PS_start1 = PolyCL_PS.GetPointAtDist(Start1);
                                            Point3d PointCL_PS_end1 = new Point3d();
                                            PointCL_PS_end1 = PolyCL_PS.GetPointAtDist(End1);
                                            Line Linie_int_start = new Line(PointCL_PS_start1, new Point3d(PointCL_PS_start1.X, PointCL_PS_start1.Y + 1000, 0));
                                            Line Linie_int_end = new Line(PointCL_PS_end1, new Point3d(PointCL_PS_end1.X, PointCL_PS_end1.Y + 1000, 0));
                                            //BTrecord.AppendEntity(Linie_int_start);
                                            //BTrecord.AppendEntity(Linie_int_end);
                                            //Trans1.AddNewlyCreatedDBObject(Linie_int_end, true);
                                            //Trans1.AddNewlyCreatedDBObject(Linie_int_start, true);

                                            Point3dCollection Col_int_start = new Point3dCollection();
                                            Point3dCollection Col_int_end = new Point3dCollection();
                                            Point3dCollection Colectie_puncte_poly = new Point3dCollection();


                                            TWS_L.IntersectWith(Linie_int_start, Autodesk.AutoCAD.DatabaseServices.Intersect.ExtendThis, Col_int_start, IntPtr.Zero, IntPtr.Zero);
                                            TWS_L.IntersectWith(Linie_int_end, Autodesk.AutoCAD.DatabaseServices.Intersect.ExtendThis, Col_int_end, IntPtr.Zero, IntPtr.Zero);

                                            Colectie_puncte_ATWS_L_ON_TWS[Index_ATWS_L_ON_TWS].Add(Col_int_start[0]);
                                            Colectie_puncte_poly.Add(Col_int_start[0]);
                                            double Param_start = TWS_L.GetParameterAtPoint(Col_int_start[0]);
                                            double Param_end = TWS_L.GetParameterAtPoint(Col_int_end[0]);

                                            //Line TEST_int_start = new Line(PointCL_PS_start1, Col_int_start[0]);
                                            //Line TEST_int_end = new Line(PointCL_PS_end1, Col_int_end[0]);
                                            //BTrecord.AppendEntity(TEST_int_start);
                                            //BTrecord.AppendEntity(TEST_int_end);
                                            //Trans1.AddNewlyCreatedDBObject(TEST_int_end, true);
                                            //Trans1.AddNewlyCreatedDBObject(TEST_int_start, true);

                                            if (Math.Floor(Param_end) - Math.Floor(Param_start) >= 1)
                                            {

                                                int P1 = (int)Math.Ceiling(Param_start);
                                                int P2 = (int)Math.Floor(Param_end);

                                                for (int j = P1; j <= P2; ++j)
                                                {
                                                    Colectie_puncte_ATWS_L_ON_TWS[Index_ATWS_L_ON_TWS].Add(TWS_L.GetPointAtParameter(j));
                                                    Colectie_puncte_poly.Add(TWS_L.GetPointAtParameter(j));
                                                }
                                            }

                                            Colectie_puncte_ATWS_L_ON_TWS[Index_ATWS_L_ON_TWS].Add(Col_int_end[0]);
                                            Colectie_puncte_poly.Add(Col_int_end[0]);

                                            if (Start_w1 == End_w1)
                                            {

                                                Polyline Temp_poly = new Polyline();

                                                for (int j = 0; j < Colectie_puncte_poly.Count; ++j)
                                                {
                                                    Temp_poly.AddVertexAt(j, new Point2d(Colectie_puncte_poly[j].X, Colectie_puncte_poly[j].Y), 0, 0, 0);
                                                }

                                                //BTrecord.AppendEntity(Temp_poly);
                                                //Trans1.AddNewlyCreatedDBObject(Temp_poly, true);

                                                Autodesk.AutoCAD.DatabaseServices.DBObjectCollection Colectie_offset_ATWS_L_ON_TWS = Temp_poly.GetOffsetCurves(-Start_w1);

                                                Polyline Temp_poly_sus = new Polyline();

                                                foreach (Polyline obj in Colectie_offset_ATWS_L_ON_TWS)
                                                {
                                                    if (obj != null)
                                                    {
                                                        Temp_poly_sus = obj;
                                                    }

                                                    if (Temp_poly_sus != null)
                                                    {
                                                        for (int j = 0; j < Temp_poly_sus.NumberOfVertices; ++j)
                                                        {
                                                            Colectie_puncte_sus_ATWS_L_ON_TWS[Index_sus_ATWS_L_ON_TWS].Add(Temp_poly_sus.GetPointAtParameter(j));
                                                        }
                                                    }
                                                }
                                                //Temp_poly_sus.ColorIndex = 2;
                                                //BTrecord.AppendEntity(Temp_poly_sus);
                                                //Trans1.AddNewlyCreatedDBObject(Temp_poly_sus, true);
                                            }
                                            else
                                            {
                                                Colectie_puncte_sus_ATWS_L_ON_TWS[Index_sus_ATWS_L_ON_TWS].Add(new Point3d(Col_int_start[0].X, Col_int_start[0].Y + Start_w1, 0));
                                                Colectie_puncte_sus_ATWS_L_ON_TWS[Index_sus_ATWS_L_ON_TWS].Add(new Point3d(Col_int_end[0].X, Col_int_end[0].Y + End_w1, 0));
                                            }

                                            if (Start2 != End1)
                                            {
                                                Index_ATWS_L_ON_TWS = Index_ATWS_L_ON_TWS + 1;
                                                Array.Resize(ref Colectie_puncte_ATWS_L_ON_TWS, Index_ATWS_L_ON_TWS + 1);
                                                Colectie_puncte_ATWS_L_ON_TWS[Index_ATWS_L_ON_TWS] = new Point3dCollection();

                                                Index_sus_ATWS_L_ON_TWS = Index_sus_ATWS_L_ON_TWS + 1;
                                                Array.Resize(ref Colectie_puncte_sus_ATWS_L_ON_TWS, Index_sus_ATWS_L_ON_TWS + 1);
                                                Colectie_puncte_sus_ATWS_L_ON_TWS[Index_sus_ATWS_L_ON_TWS] = new Point3dCollection();

                                            }


                                        }
                                    }
                                }

                                for (int i = 0; i <= Index_ATWS_L_ON_TWS; ++i)
                                {
                                    Polyline ATWS_L_ON_TWS = new Polyline();
                                    int Idx1 = 0;
                                    for (int j = 0; j < Colectie_puncte_sus_ATWS_L_ON_TWS[i].Count; ++j)
                                    {
                                        ATWS_L_ON_TWS.AddVertexAt(Idx1, new Point2d(Colectie_puncte_sus_ATWS_L_ON_TWS[i][j].X, Colectie_puncte_sus_ATWS_L_ON_TWS[i][j].Y), 0, 0, 0);
                                        Idx1 = Idx1 + 1;

                                    }
                                    for (int j = Colectie_puncte_ATWS_L_ON_TWS[i].Count - 1; j >= 0; --j)
                                    {
                                        ATWS_L_ON_TWS.AddVertexAt(Idx1, new Point2d(Colectie_puncte_ATWS_L_ON_TWS[i][j].X, Colectie_puncte_ATWS_L_ON_TWS[i][j].Y), 0, 0, 0);
                                        Idx1 = Idx1 + 1;
                                    }

                                    ATWS_L_ON_TWS.ColorIndex = 253;
                                    ATWS_L_ON_TWS.Closed = true;
                                    BTrecord.AppendEntity(ATWS_L_ON_TWS);
                                    Trans1.AddNewlyCreatedDBObject(ATWS_L_ON_TWS, true);
                                }


                            }


                            Trans1.Commit();
                        }
                    }


                }

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message + "\nProblem around station " + eRR_Stat + " on ATWS Left on TWS");
            }
        }
        private void draw_ATWS_R_ON_TWS(Point3d Point1)
        {

            string eRR_Stat = "";
            try
            {
                string Column_ATWS_R_ON_TWS_Begin = "START";
                string Column_ATWS_R_ON_TWS_End = "END";

                string Column_ATWS_R_ON_TWS_Width_Begin = "START_WIDTH";
                string Column_ATWS_R_ON_TWS_Width_End = "END_WIDTH";


                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        if (ATWS_R_ON_TWS_TABLE != null & TWS_R != null)
                        {
                            if (ATWS_R_ON_TWS_TABLE.Rows.Count > 0)
                            {
                                Double Match1 = Convert.ToDouble(textBox_Matchline_start.Text);
                                Double Match2 = Convert.ToDouble(textBox_Matchline_end.Text);
                                if (Match1 > Match2)
                                {
                                    double T = Match1;
                                    Match1 = Match2;
                                    Match2 = T;
                                }

                                double Len1 = Match2 - Match1;// PolyCL_MS.Length;

                                Polyline PolyCL_PS = new Polyline();
                                PolyCL_PS.AddVertexAt(0, new Point2d(Point1.X, Point1.Y), 0, 0, 0);
                                PolyCL_PS.AddVertexAt(1, new Point2d(Point1.X + Len1, Point1.Y), 0, 0, 0);
                                PolyCL_PS.ColorIndex = 1;
                                PolyCL_PS.Elevation = 0;
                                BTrecord.AppendEntity(PolyCL_PS);
                                Trans1.AddNewlyCreatedDBObject(PolyCL_PS, true);

                                Point3dCollection[] Colectie_puncte_ATWS_R_ON_TWS;
                                int Index_ATWS_R_ON_TWS = 0;
                                Colectie_puncte_ATWS_R_ON_TWS = new Point3dCollection[Index_ATWS_R_ON_TWS + 1];
                                Colectie_puncte_ATWS_R_ON_TWS[Index_ATWS_R_ON_TWS] = new Point3dCollection();

                                Point3dCollection[] Colectie_puncte_sus_ATWS_R_ON_TWS;
                                int Index_sus_ATWS_R_ON_TWS = 0;
                                Colectie_puncte_sus_ATWS_R_ON_TWS = new Point3dCollection[Index_sus_ATWS_R_ON_TWS + 1];
                                Colectie_puncte_sus_ATWS_R_ON_TWS[Index_sus_ATWS_R_ON_TWS] = new Point3dCollection();




                                for (int i = 0; i < ATWS_R_ON_TWS_TABLE.Rows.Count; ++i)
                                {
                                    Double Start1 = -1;
                                    Double End1 = -1;
                                    Double Start_w1 = -1;
                                    Double End_w1 = -1;
                                    Double Start2 = -1;


                                    if ((ATWS_R_ON_TWS_TABLE.Rows[i][Column_ATWS_R_ON_TWS_Begin] != System.DBNull.Value)) Start1 = Convert.ToDouble(ATWS_R_ON_TWS_TABLE.Rows[i][Column_ATWS_R_ON_TWS_Begin]);
                                    if ((ATWS_R_ON_TWS_TABLE.Rows[i][Column_ATWS_R_ON_TWS_End] != System.DBNull.Value)) End1 = Convert.ToDouble(ATWS_R_ON_TWS_TABLE.Rows[i][Column_ATWS_R_ON_TWS_End]);
                                    if ((ATWS_R_ON_TWS_TABLE.Rows[i][Column_ATWS_R_ON_TWS_Width_Begin] != System.DBNull.Value)) Start_w1 = Convert.ToDouble(ATWS_R_ON_TWS_TABLE.Rows[i][Column_ATWS_R_ON_TWS_Width_Begin]);
                                    if ((ATWS_R_ON_TWS_TABLE.Rows[i][Column_ATWS_R_ON_TWS_Width_End] != System.DBNull.Value)) End_w1 = Convert.ToDouble(ATWS_R_ON_TWS_TABLE.Rows[i][Column_ATWS_R_ON_TWS_Width_End]);
                                    eRR_Stat = Start1.ToString();
                                    if (i + 1 < ATWS_R_ON_TWS_TABLE.Rows.Count)
                                    {
                                        if ((ATWS_R_ON_TWS_TABLE.Rows[i + 1][Column_ATWS_R_ON_TWS_Begin] != System.DBNull.Value))
                                        {
                                            Start2 = Convert.ToDouble(ATWS_R_ON_TWS_TABLE.Rows[i + 1][Column_ATWS_R_ON_TWS_Begin]);
                                        }
                                    }


                                    if (Start1 < Match1 & End1 <= Match2)
                                    {
                                        Start1 = Match1;
                                    }

                                    if (Start1 >= Match1 & Start1 <= Match2 & End1 >= Match2)
                                    {
                                        End1 = Match2;

                                    }

                                    if (Start1 <= Match1 & End1 >= Match2)
                                    {
                                        Start1 = Match1;
                                        End1 = Match2;

                                    }
                                    if (Start1 >= Match1 & End1 <= Match2)
                                    {
                                        if (Start1 != -1 & End1 != -1 & Start_w1 != -1 & End_w1 != -1)
                                        {
                                            Point3d PointCL_PS_start1 = new Point3d();
                                            PointCL_PS_start1 = PolyCL_PS.GetPointAtDist(Start1);
                                            Point3d PointCL_PS_end1 = new Point3d();
                                            PointCL_PS_end1 = PolyCL_PS.GetPointAtDist(End1);
                                            Line Linie_int_start = new Line(PointCL_PS_start1, new Point3d(PointCL_PS_start1.X, PointCL_PS_start1.Y - 1000, 0));
                                            Line Linie_int_end = new Line(PointCL_PS_end1, new Point3d(PointCL_PS_end1.X, PointCL_PS_end1.Y - 1000, 0));
                                            //BTrecord.AppendEntity(Linie_int_start);
                                            //BTrecord.AppendEntity(Linie_int_end);
                                            //Trans1.AddNewlyCreatedDBObject(Linie_int_end, true);
                                            //Trans1.AddNewlyCreatedDBObject(Linie_int_start, true);

                                            Point3dCollection Col_int_start = new Point3dCollection();
                                            Point3dCollection Col_int_end = new Point3dCollection();
                                            Point3dCollection Colectie_puncte_poly = new Point3dCollection();

                                            TWS_R.IntersectWith(Linie_int_start, Autodesk.AutoCAD.DatabaseServices.Intersect.ExtendThis, Col_int_start, IntPtr.Zero, IntPtr.Zero);
                                            TWS_R.IntersectWith(Linie_int_end, Autodesk.AutoCAD.DatabaseServices.Intersect.ExtendThis, Col_int_end, IntPtr.Zero, IntPtr.Zero);

                                            Colectie_puncte_ATWS_R_ON_TWS[Index_ATWS_R_ON_TWS].Add(Col_int_start[0]);
                                            Colectie_puncte_poly.Add(Col_int_start[0]);

                                            double Param_start = TWS_R.GetParameterAtPoint(Col_int_start[0]);
                                            double Param_end = TWS_R.GetParameterAtPoint(Col_int_end[0]);

                                            //Line TEST_int_start = new Line(PointCL_PS_start1, Col_int_start[0]);
                                            //Line TEST_int_end = new Line(PointCL_PS_end1, Col_int_end[0]);
                                            //BTrecord.AppendEntity(TEST_int_start);
                                            //BTrecord.AppendEntity(TEST_int_end);
                                            //Trans1.AddNewlyCreatedDBObject(TEST_int_end, true);
                                            //Trans1.AddNewlyCreatedDBObject(TEST_int_start, true);

                                            if (Math.Floor(Param_end) - Math.Floor(Param_start) >= 1)
                                            {

                                                int P1 = (int)Math.Ceiling(Param_start);
                                                int P2 = (int)Math.Floor(Param_end);

                                                for (int j = P1; j <= P2; ++j)
                                                {
                                                    Colectie_puncte_ATWS_R_ON_TWS[Index_ATWS_R_ON_TWS].Add(TWS_R.GetPointAtParameter(j));
                                                    Colectie_puncte_poly.Add(TWS_R.GetPointAtParameter(j));
                                                }
                                            }

                                            Colectie_puncte_ATWS_R_ON_TWS[Index_ATWS_R_ON_TWS].Add(Col_int_end[0]);
                                            Colectie_puncte_poly.Add(Col_int_end[0]);

                                            if (Start_w1 == End_w1)
                                            {

                                                Polyline Temp_poly = new Polyline();

                                                for (int j = 0; j < Colectie_puncte_poly.Count; ++j)
                                                {
                                                    Temp_poly.AddVertexAt(j, new Point2d(Colectie_puncte_poly[j].X, Colectie_puncte_poly[j].Y), 0, 0, 0);
                                                }

                                                //BTrecord.AppendEntity(Temp_poly);
                                                //Trans1.AddNewlyCreatedDBObject(Temp_poly, true);

                                                Autodesk.AutoCAD.DatabaseServices.DBObjectCollection Colectie_offset_ATWS_R_ON_TWS = Temp_poly.GetOffsetCurves(Start_w1);

                                                Polyline Temp_poly_sus = new Polyline();

                                                foreach (Polyline obj in Colectie_offset_ATWS_R_ON_TWS)
                                                {
                                                    if (obj != null)
                                                    {
                                                        Temp_poly_sus = obj;
                                                    }

                                                    if (Temp_poly_sus != null)
                                                    {
                                                        for (int j = 0; j < Temp_poly_sus.NumberOfVertices; ++j)
                                                        {
                                                            Colectie_puncte_sus_ATWS_R_ON_TWS[Index_sus_ATWS_R_ON_TWS].Add(Temp_poly_sus.GetPointAtParameter(j));
                                                        }
                                                    }
                                                }
                                                //BTrecord.AppendEntity(Temp_poly_sus);
                                                //Trans1.AddNewlyCreatedDBObject(Temp_poly_sus, true);
                                            }
                                            else
                                            {
                                                Colectie_puncte_sus_ATWS_R_ON_TWS[Index_sus_ATWS_R_ON_TWS].Add(new Point3d(Col_int_start[0].X, Col_int_start[0].Y - Start_w1, 0));
                                                Colectie_puncte_sus_ATWS_R_ON_TWS[Index_sus_ATWS_R_ON_TWS].Add(new Point3d(Col_int_end[0].X, Col_int_end[0].Y - End_w1, 0));
                                            }

                                            if (Start2 != End1)
                                            {
                                                Index_ATWS_R_ON_TWS = Index_ATWS_R_ON_TWS + 1;
                                                Array.Resize(ref Colectie_puncte_ATWS_R_ON_TWS, Index_ATWS_R_ON_TWS + 1);
                                                Colectie_puncte_ATWS_R_ON_TWS[Index_ATWS_R_ON_TWS] = new Point3dCollection();

                                                Index_sus_ATWS_R_ON_TWS = Index_sus_ATWS_R_ON_TWS + 1;
                                                Array.Resize(ref Colectie_puncte_sus_ATWS_R_ON_TWS, Index_sus_ATWS_R_ON_TWS + 1);
                                                Colectie_puncte_sus_ATWS_R_ON_TWS[Index_sus_ATWS_R_ON_TWS] = new Point3dCollection();

                                            }


                                        }
                                    }
                                }

                                for (int i = 0; i <= Index_ATWS_R_ON_TWS; ++i)
                                {
                                    Polyline ATWS_R_ON_TWS = new Polyline();
                                    int Idx1 = 0;
                                    for (int j = 0; j < Colectie_puncte_sus_ATWS_R_ON_TWS[i].Count; ++j)
                                    {
                                        ATWS_R_ON_TWS.AddVertexAt(Idx1, new Point2d(Colectie_puncte_sus_ATWS_R_ON_TWS[i][j].X, Colectie_puncte_sus_ATWS_R_ON_TWS[i][j].Y), 0, 0, 0);
                                        Idx1 = Idx1 + 1;

                                    }
                                    for (int j = Colectie_puncte_ATWS_R_ON_TWS[i].Count - 1; j >= 0; --j)
                                    {
                                        ATWS_R_ON_TWS.AddVertexAt(Idx1, new Point2d(Colectie_puncte_ATWS_R_ON_TWS[i][j].X, Colectie_puncte_ATWS_R_ON_TWS[i][j].Y), 0, 0, 0);
                                        Idx1 = Idx1 + 1;
                                    }

                                    ATWS_R_ON_TWS.ColorIndex = 253;
                                    ATWS_R_ON_TWS.Closed = true;
                                    BTrecord.AppendEntity(ATWS_R_ON_TWS);
                                    Trans1.AddNewlyCreatedDBObject(ATWS_R_ON_TWS, true);
                                }


                            }


                            Trans1.Commit();
                        }
                    }


                }

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message + "\nProblem around station " + eRR_Stat + " on ATWS Right on TWS");
            }
        }




        private void OLD_Button_pick_existing_easement_left_click(object sender, EventArgs e)
        {
            if (Freeze_operations == false)
            {
                Freeze_operations = true;
                try
                {

                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                    if (PolyCL_MS != null)
                    {

                        Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                        using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                        {
                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                            {
                                Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                                Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                                Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_easement_L;
                                Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_easement_L;
                                Prompt_easement_L = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the existing left easement:");
                                Prompt_easement_L.SetRejectMessage("\nSelect a polyline!");
                                Prompt_easement_L.AllowNone = true;
                                Prompt_easement_L.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                                Rezultat_easement_L = ThisDrawing.Editor.GetEntity(Prompt_easement_L);

                                if (Rezultat_easement_L.Status != PromptStatus.OK)
                                {
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    Freeze_operations = false;
                                    return;
                                }

                                Plane Planul_curent = new Plane(new Point3d(0, 0, 0), Vector3d.ZAxis);

                                Data_table_easement_left = new System.Data.DataTable();

                                Data_table_easement_left.Columns.Add("STATION", typeof(Double));
                                Data_table_easement_left.Columns.Add("OFFSET", typeof(Double));


                                Polyline Poly_easement_L = (Polyline)Trans1.GetObject(Rezultat_easement_L.ObjectId, OpenMode.ForRead);

                                if (Poly_easement_L.StartPoint.GetVectorTo(PolyCL_MS.StartPoint).Length > Poly_easement_L.EndPoint.GetVectorTo(PolyCL_MS.StartPoint).Length)
                                {
                                    try
                                    {
                                        var Obj1 = Poly_easement_L as Polyline;
                                        Poly_easement_L.UpgradeOpen();
                                        Obj1.ReverseCurve();
                                        Poly_easement_L = Obj1;
                                        Poly_easement_L.DowngradeOpen();
                                    }
                                    catch
                                    {

                                    }
                                }


                                int Idx = 0;


                                for (int i = 0; i < Poly_easement_L.NumberOfVertices; ++i)
                                {

                                    Point3d Nod_poly_easement = new Point3d();
                                    Nod_poly_easement = Poly_easement_L.GetPoint3dAt(i);
                                    Point3d Point_on_CL = new Point3d();
                                    Point_on_CL = PolyCL_MS.GetClosestPointTo(Nod_poly_easement, Vector3d.ZAxis, false);
                                    double Station1 = Math.Round(PolyCL_MS.GetDistAtPoint(Point_on_CL), Rounding_no);
                                    double Dist1 = Math.Round(Nod_poly_easement.GetVectorTo(Point_on_CL).Length, Rounding_no);
                                    double Param1_CL = PolyCL_MS.GetParameterAtPoint(Point_on_CL);
                                    int Param_C = Convert.ToInt32(Math.Round(Param1_CL, 0));

                                    Line Line1 = null;
                                    Line Line2 = null;
                                    Line LineC1 = null;
                                    Line LineC2 = null;

                                    double Len1 = 0;

                                    if (i - 1 >= 0)
                                    {
                                        int k = 0;
                                        do
                                        {
                                            if (i - 1 - k >= 0)
                                            {
                                                Line1 = new Line(Poly_easement_L.GetPoint3dAt(i), Poly_easement_L.GetPoint3dAt(i - 1 - k));
                                                k = k + 1;
                                                Len1 = Line1.Length;
                                            }
                                            else
                                            {
                                                Line1 = null;
                                                goto calcs;
                                            }

                                        } while (Len1 <= 0.01);


                                    }

                                    Len1 = 0;

                                    if (i + 1 < Poly_easement_L.NumberOfVertices)
                                    {
                                        int k = 0;
                                        do
                                        {
                                            if (i + 1 + k < Poly_easement_L.NumberOfVertices)
                                            {
                                                Line2 = new Line(Poly_easement_L.GetPoint3dAt(i), Poly_easement_L.GetPoint3dAt(i + 1 + k));
                                                k = k + 1;
                                                Len1 = Line2.Length;
                                            }
                                            else
                                            {
                                                Line2 = null;
                                                goto calcs;
                                            }

                                        } while (Len1 <= 0.01);



                                    }

                                    Len1 = 0;

                                    if (Param_C - 1 >= 0)
                                    {

                                        int k = 0;
                                        do
                                        {
                                            if (Param_C - 1 - k >= 0)
                                            {
                                                LineC1 = new Line(PolyCL_MS.GetPoint3dAt(Param_C), PolyCL_MS.GetPoint3dAt(Param_C - 1 - k));
                                                k = k + 1;
                                                Len1 = LineC1.Length;
                                            }
                                            else
                                            {
                                                LineC1 = null;
                                                goto calcs;
                                            }

                                        } while (Len1 <= 0.01);


                                    }

                                    Len1 = 0;

                                    if (Param_C + 1 < PolyCL_MS.NumberOfVertices)
                                    {

                                        int k = 0;
                                        do
                                        {
                                            if (Param_C + 1 + k < PolyCL_MS.NumberOfVertices)
                                            {
                                                LineC2 = new Line(PolyCL_MS.GetPoint3dAt(Param_C), PolyCL_MS.GetPoint3dAt(Param_C + 1 + k));
                                                k = k + 1;
                                                Len1 = LineC2.Length;
                                            }
                                            else
                                            {
                                                LineC2 = null;
                                                goto calcs;
                                            }

                                        } while (Len1 <= 0.01);
                                    }

                                calcs:

                                    if (Line1 != null & Line2 != null & LineC1 != null & LineC2 != null)
                                    {
                                        double Bearing1 = Math.Round(Workspace_Band.Functions.GET_Bearing_rad(Line1.StartPoint.X, Line1.StartPoint.Y, Line1.EndPoint.X, Line1.EndPoint.Y), 2);
                                        double BearingC1 = Math.Round(Workspace_Band.Functions.GET_Bearing_rad(LineC1.StartPoint.X, LineC1.StartPoint.Y, LineC1.EndPoint.X, LineC1.EndPoint.Y), 2);

                                        double Bearing2 = Math.Round(Workspace_Band.Functions.GET_Bearing_rad(Line2.StartPoint.X, Line2.StartPoint.Y, Line2.EndPoint.X, Line2.EndPoint.Y), 2);
                                        double BearingC2 = Math.Round(Workspace_Band.Functions.GET_Bearing_rad(LineC2.StartPoint.X, LineC2.StartPoint.Y, LineC2.EndPoint.X, LineC2.EndPoint.Y), 2);


                                        if (Bearing1 == BearingC1 & Bearing2 == BearingC2)
                                        {
                                            Vector3d vector1 = LineC1.EndPoint.GetVectorTo(LineC1.StartPoint);
                                            Vector3d vector2 = LineC2.StartPoint.GetVectorTo(LineC2.EndPoint);
                                            double Defl = Math.Round(vector2.GetAngleTo(vector1), 2);
                                            if (Defl != 0)
                                            {
                                                Point3d Pt1 = new Point3d();
                                                Pt1 = Line1.GetClosestPointTo(LineC1.StartPoint, Vector3d.ZAxis, true);
                                                Point3d Pt2 = new Point3d();
                                                Pt2 = Line2.GetClosestPointTo(LineC2.StartPoint, Vector3d.ZAxis, true);
                                                double L1 = Math.Round(LineC1.StartPoint.GetVectorTo(Pt1).Length, Rounding_no);
                                                double L2 = Math.Round(LineC2.StartPoint.GetVectorTo(Pt2).Length, Rounding_no);

                                                if (L1 == L2 & Dist1 >= L1)
                                                {

                                                    Data_table_easement_left.Rows.Add();
                                                    Data_table_easement_left.Rows[Idx]["OFFSET"] = L1;
                                                    Data_table_easement_left.Rows[Idx]["STATION"] = Station1;
                                                    Idx = Idx + 1;
                                                }
                                                else if (L1 != L2 & (Dist1 >= L1 | Dist1 >= L2))
                                                {
                                                    Station1 = Math.Round(PolyCL_MS.GetDistanceAtParameter(Convert.ToDouble(Param_C)), Rounding_no);

                                                    Data_table_easement_left.Rows.Add();
                                                    Data_table_easement_left.Rows[Idx]["OFFSET"] = L1;
                                                    Data_table_easement_left.Rows[Idx]["STATION"] = Station1;
                                                    Idx = Idx + 1;
                                                    Data_table_easement_left.Rows.Add();
                                                    Data_table_easement_left.Rows[Idx]["OFFSET"] = L2;
                                                    Data_table_easement_left.Rows[Idx]["STATION"] = Station1;
                                                    Idx = Idx + 1;
                                                }
                                                else
                                                {
                                                    Data_table_easement_left.Rows.Add();
                                                    Data_table_easement_left.Rows[Idx]["OFFSET"] = Dist1;
                                                    Data_table_easement_left.Rows[Idx]["STATION"] = Station1;
                                                    Idx = Idx + 1;
                                                }
                                            }
                                            else
                                            {
                                                Data_table_easement_left.Rows.Add();
                                                Data_table_easement_left.Rows[Idx]["OFFSET"] = Dist1;
                                                Data_table_easement_left.Rows[Idx]["STATION"] = Station1;
                                                Idx = Idx + 1;
                                            }
                                        }
                                        else
                                        {
                                            Data_table_easement_left.Rows.Add();
                                            Data_table_easement_left.Rows[Idx]["OFFSET"] = Dist1;
                                            Data_table_easement_left.Rows[Idx]["STATION"] = Station1;
                                            Idx = Idx + 1;
                                        }
                                    }
                                    else
                                    {
                                        Data_table_easement_left.Rows.Add();
                                        Data_table_easement_left.Rows[Idx]["OFFSET"] = Dist1;
                                        Data_table_easement_left.Rows[Idx]["STATION"] = Station1;
                                        Idx = Idx + 1;
                                    }

                                }

                                Trans1.Commit();
                            }
                        }

                        if (Data_table_easement_left != null)
                        {
                            string Table1 = textBox_COMPILED.Text;

                            string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + textBox_access_database_location.Text + ";";

                            OleDbConnection cnn;
                            try
                            {
                                cnn = new OleDbConnection(ConnectionString);
                                cnn.Open();

                                OleDbCommand cmd = new OleDbCommand();
                                cmd.CommandType = CommandType.Text;
                                cmd.Connection = cnn;

                                Double Station2_SET = -111.111;

                                for (int i = 0; i < Data_table_easement_left.Rows.Count - 1; ++i)
                                {
                                    Double Offset0 = -111.111;
                                    if (i > 0)
                                    {
                                        Offset0 = (double)Data_table_easement_left.Rows[i - 1]["OFFSET"];
                                    }
                                    Double Offset1 = (double)Data_table_easement_left.Rows[i]["OFFSET"];
                                    Double Station1 = (double)Data_table_easement_left.Rows[i]["STATION"];

                                    Double Offset2 = (double)Data_table_easement_left.Rows[i + 1]["OFFSET"];
                                    Double Station2 = (double)Data_table_easement_left.Rows[i + 1]["STATION"];

                                    if (Offset0 != Offset1 | Offset0 != Offset2)
                                    {
                                        cmd.CommandText = "INSERT INTO " + Table1 + "(TYPE,SIDE,BDY,START_STA,END_STA,START_WIDTH,END_WIDTH,WS_NO) VALUES " +
                                            "('EPE','L','CL'," + Station1 + "," + Station2 + "," + Offset1 + "," + Offset2 + ",0)";
                                        cmd.ExecuteNonQuery();

                                        if (Station2_SET != -111.111)
                                        {
                                            cmd.CommandText = "UPDATE " + Table1 + " SET END_STA = " + Station1 + " WHERE END_STA = " + Station2_SET;
                                            cmd.ExecuteNonQuery();

                                        }

                                        Station2_SET = Station2;
                                    }

                                    else if (i == Data_table_easement_left.Rows.Count - 2)
                                    {
                                        cmd.CommandText = "UPDATE " + Table1 + " SET END_STA = " + Station2 + " WHERE END_STA = " + Station2_SET;
                                        cmd.ExecuteNonQuery();
                                    }



                                }



                                cnn.Close();
                            }
                            catch (OleDbException ex)
                            {
                                MessageBox.Show(ex.Message);
                                Freeze_operations = false;
                            }
                        }

                    }
                    else
                    {
                        MessageBox.Show("Centerline has not been loaded");
                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                        Freeze_operations = false;
                    }


                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    Freeze_operations = false;
                }

                Freeze_operations = false;
            }

        }

        private void old_Button_pick_easement_right_click(object sender, EventArgs e)
        {
            if (Freeze_operations == false)
            {
                Freeze_operations = true;
                try
                {

                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                    if (PolyCL_MS != null)
                    {

                        Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                        using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                        {
                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                            {
                                Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                                Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                                Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_easement_R;
                                Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_easement_R;
                                Prompt_easement_R = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the existing right easement:");
                                Prompt_easement_R.SetRejectMessage("\nSelect a polyline!");
                                Prompt_easement_R.AllowNone = true;
                                Prompt_easement_R.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                                Rezultat_easement_R = ThisDrawing.Editor.GetEntity(Prompt_easement_R);

                                if (Rezultat_easement_R.Status != PromptStatus.OK)
                                {
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    Freeze_operations = false;
                                    return;
                                }

                                Plane Planul_curent = new Plane(new Point3d(0, 0, 0), Vector3d.ZAxis);

                                Data_table_easement_right = new System.Data.DataTable();

                                Data_table_easement_right.Columns.Add("STATION", typeof(Double));
                                Data_table_easement_right.Columns.Add("OFFSET", typeof(Double));


                                Polyline Poly_easement_R = (Polyline)Trans1.GetObject(Rezultat_easement_R.ObjectId, OpenMode.ForRead);

                                if (Poly_easement_R.StartPoint.GetVectorTo(PolyCL_MS.StartPoint).Length > Poly_easement_R.EndPoint.GetVectorTo(PolyCL_MS.StartPoint).Length)
                                {
                                    try
                                    {
                                        var Obj1 = Poly_easement_R as Polyline;
                                        Poly_easement_R.UpgradeOpen();
                                        Obj1.ReverseCurve();
                                        Poly_easement_R = Obj1;
                                        Poly_easement_R.DowngradeOpen();
                                    }
                                    catch
                                    {

                                    }
                                }


                                int Idx = 0;


                                for (int i = 0; i < Poly_easement_R.NumberOfVertices; ++i)
                                {

                                    Point3d Nod_poly_easement = new Point3d();
                                    Nod_poly_easement = Poly_easement_R.GetPoint3dAt(i);
                                    Point3d Point_on_CL = new Point3d();
                                    Point_on_CL = PolyCL_MS.GetClosestPointTo(Nod_poly_easement, Vector3d.ZAxis, false);
                                    double Station1 = Math.Round(PolyCL_MS.GetDistAtPoint(Point_on_CL), Rounding_no);
                                    double Dist1 = Math.Round(Nod_poly_easement.GetVectorTo(Point_on_CL).Length, Rounding_no);
                                    double Param1_CL = PolyCL_MS.GetParameterAtPoint(Point_on_CL);
                                    int Param_C = Convert.ToInt32(Math.Round(Param1_CL, 0));

                                    Line Line1 = null;
                                    Line Line2 = null;
                                    Line LineC1 = null;
                                    Line LineC2 = null;

                                    double Len1 = 0;

                                    if (i - 1 >= 0)
                                    {
                                        int k = 0;
                                        do
                                        {
                                            if (i - 1 - k >= 0)
                                            {
                                                Line1 = new Line(Poly_easement_R.GetPoint3dAt(i), Poly_easement_R.GetPoint3dAt(i - 1 - k));
                                                k = k + 1;
                                                Len1 = Line1.Length;
                                            }
                                            else
                                            {
                                                Line1 = null;
                                                goto calcs;
                                            }

                                        } while (Len1 <= 0.01);


                                    }

                                    Len1 = 0;

                                    if (i + 1 < Poly_easement_R.NumberOfVertices)
                                    {
                                        int k = 0;
                                        do
                                        {
                                            if (i + 1 + k < Poly_easement_R.NumberOfVertices)
                                            {
                                                Line2 = new Line(Poly_easement_R.GetPoint3dAt(i), Poly_easement_R.GetPoint3dAt(i + 1 + k));
                                                k = k + 1;
                                                Len1 = Line2.Length;
                                            }
                                            else
                                            {
                                                Line2 = null;
                                                goto calcs;
                                            }

                                        } while (Len1 <= 0.01);



                                    }

                                    Len1 = 0;

                                    if (Param_C - 1 >= 0)
                                    {

                                        int k = 0;
                                        do
                                        {
                                            if (Param_C - 1 - k >= 0)
                                            {
                                                LineC1 = new Line(PolyCL_MS.GetPoint3dAt(Param_C), PolyCL_MS.GetPoint3dAt(Param_C - 1 - k));
                                                k = k + 1;
                                                Len1 = LineC1.Length;
                                            }
                                            else
                                            {
                                                LineC1 = null;
                                                goto calcs;
                                            }

                                        } while (Len1 <= 0.01);


                                    }

                                    Len1 = 0;

                                    if (Param_C + 1 < PolyCL_MS.NumberOfVertices)
                                    {

                                        int k = 0;
                                        do
                                        {
                                            if (Param_C + 1 + k < PolyCL_MS.NumberOfVertices)
                                            {
                                                LineC2 = new Line(PolyCL_MS.GetPoint3dAt(Param_C), PolyCL_MS.GetPoint3dAt(Param_C + 1 + k));
                                                k = k + 1;
                                                Len1 = LineC2.Length;
                                            }
                                            else
                                            {
                                                LineC2 = null;
                                                goto calcs;
                                            }

                                        } while (Len1 <= 0.01);
                                    }

                                calcs:

                                    if (Line1 != null & Line2 != null & LineC1 != null & LineC2 != null)
                                    {
                                        double Bearing1 = Math.Round(Workspace_Band.Functions.GET_Bearing_rad(Line1.StartPoint.X, Line1.StartPoint.Y, Line1.EndPoint.X, Line1.EndPoint.Y), 2);
                                        double BearingC1 = Math.Round(Workspace_Band.Functions.GET_Bearing_rad(LineC1.StartPoint.X, LineC1.StartPoint.Y, LineC1.EndPoint.X, LineC1.EndPoint.Y), 2);

                                        double Bearing2 = Math.Round(Workspace_Band.Functions.GET_Bearing_rad(Line2.StartPoint.X, Line2.StartPoint.Y, Line2.EndPoint.X, Line2.EndPoint.Y), 2);
                                        double BearingC2 = Math.Round(Workspace_Band.Functions.GET_Bearing_rad(LineC2.StartPoint.X, LineC2.StartPoint.Y, LineC2.EndPoint.X, LineC2.EndPoint.Y), 2);


                                        if (Bearing1 == BearingC1 & Bearing2 == BearingC2)
                                        {
                                            Vector3d vector1 = LineC1.EndPoint.GetVectorTo(LineC1.StartPoint);
                                            Vector3d vector2 = LineC2.StartPoint.GetVectorTo(LineC2.EndPoint);
                                            double Defl = Math.Round(vector2.GetAngleTo(vector1), 2);
                                            if (Defl != 0)
                                            {
                                                Point3d Pt1 = new Point3d();
                                                Pt1 = Line1.GetClosestPointTo(LineC1.StartPoint, Vector3d.ZAxis, true);
                                                Point3d Pt2 = new Point3d();
                                                Pt2 = Line2.GetClosestPointTo(LineC2.StartPoint, Vector3d.ZAxis, true);
                                                double L1 = Math.Round(LineC1.StartPoint.GetVectorTo(Pt1).Length, Rounding_no);
                                                double L2 = Math.Round(LineC2.StartPoint.GetVectorTo(Pt2).Length, Rounding_no);

                                                if (L1 == L2 & Dist1 >= L1)
                                                {

                                                    Data_table_easement_right.Rows.Add();
                                                    Data_table_easement_right.Rows[Idx]["OFFSET"] = L1;
                                                    Data_table_easement_right.Rows[Idx]["STATION"] = Station1;
                                                    Idx = Idx + 1;
                                                }
                                                else if (L1 != L2 & (Dist1 >= L1 | Dist1 >= L2))
                                                {
                                                    Station1 = Math.Round(PolyCL_MS.GetDistanceAtParameter(Convert.ToDouble(Param_C)), Rounding_no);

                                                    Data_table_easement_right.Rows.Add();
                                                    Data_table_easement_right.Rows[Idx]["OFFSET"] = L1;
                                                    Data_table_easement_right.Rows[Idx]["STATION"] = Station1;
                                                    Idx = Idx + 1;
                                                    Data_table_easement_right.Rows.Add();
                                                    Data_table_easement_right.Rows[Idx]["OFFSET"] = L2;
                                                    Data_table_easement_right.Rows[Idx]["STATION"] = Station1;
                                                    Idx = Idx + 1;
                                                }
                                                else
                                                {
                                                    Data_table_easement_right.Rows.Add();
                                                    Data_table_easement_right.Rows[Idx]["OFFSET"] = Dist1;
                                                    Data_table_easement_right.Rows[Idx]["STATION"] = Station1;
                                                    Idx = Idx + 1;
                                                }
                                            }
                                            else
                                            {
                                                Data_table_easement_right.Rows.Add();
                                                Data_table_easement_right.Rows[Idx]["OFFSET"] = Dist1;
                                                Data_table_easement_right.Rows[Idx]["STATION"] = Station1;
                                                Idx = Idx + 1;
                                            }
                                        }
                                        else
                                        {
                                            Data_table_easement_right.Rows.Add();
                                            Data_table_easement_right.Rows[Idx]["OFFSET"] = Dist1;
                                            Data_table_easement_right.Rows[Idx]["STATION"] = Station1;
                                            Idx = Idx + 1;
                                        }
                                    }
                                    else
                                    {
                                        Data_table_easement_right.Rows.Add();
                                        Data_table_easement_right.Rows[Idx]["OFFSET"] = Dist1;
                                        Data_table_easement_right.Rows[Idx]["STATION"] = Station1;
                                        Idx = Idx + 1;
                                    }

                                }

                                Trans1.Commit();
                            }
                        }

                        if (Data_table_easement_right != null)
                        {
                            string Table1 = textBox_COMPILED.Text;

                            string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + textBox_access_database_location.Text + ";";

                            OleDbConnection cnn;
                            try
                            {
                                cnn = new OleDbConnection(ConnectionString);
                                cnn.Open();

                                OleDbCommand cmd = new OleDbCommand();
                                cmd.CommandType = CommandType.Text;
                                cmd.Connection = cnn;

                                Double Station2_SET = -111.111;

                                for (int i = 0; i < Data_table_easement_right.Rows.Count - 1; ++i)
                                {
                                    Double Offset0 = -111.111;
                                    if (i > 0)
                                    {
                                        Offset0 = (double)Data_table_easement_right.Rows[i - 1]["OFFSET"];
                                    }
                                    Double Offset1 = (double)Data_table_easement_right.Rows[i]["OFFSET"];
                                    Double Station1 = (double)Data_table_easement_right.Rows[i]["STATION"];

                                    Double Offset2 = (double)Data_table_easement_right.Rows[i + 1]["OFFSET"];
                                    Double Station2 = (double)Data_table_easement_right.Rows[i + 1]["STATION"];

                                    if (Offset0 != Offset1 | Offset0 != Offset2)
                                    {
                                        cmd.CommandText = "INSERT INTO " + Table1 + "(TYPE,SIDE,BDY,START_STA,END_STA,START_WIDTH,END_WIDTH,WS_NO) VALUES " +
                                            "('EPE','R','CL'," + Station1 + "," + Station2 + "," + Offset1 + "," + Offset2 + ",0)";
                                        cmd.ExecuteNonQuery();

                                        if (Station2_SET != -111.111)
                                        {
                                            cmd.CommandText = "UPDATE " + Table1 + " SET END_STA = " + Station1 + " WHERE END_STA = " + Station2_SET;
                                            cmd.ExecuteNonQuery();

                                        }

                                        Station2_SET = Station2;
                                    }

                                    else if (i == Data_table_easement_right.Rows.Count - 2)
                                    {
                                        cmd.CommandText = "UPDATE " + Table1 + " SET END_STA = " + Station2 + " WHERE END_STA = " + Station2_SET;
                                        cmd.ExecuteNonQuery();
                                    }



                                }



                                cnn.Close();
                            }
                            catch (OleDbException ex)
                            {
                                MessageBox.Show(ex.Message);
                                Freeze_operations = false;
                            }
                        }

                    }
                    else
                    {
                        MessageBox.Show("Centerline has not been loaded");
                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                        Freeze_operations = false;
                    }


                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    Freeze_operations = false;
                }

                Freeze_operations = false;
            }

        }

        private void OLD_Populate_access_database_with_easement_LEFT(Polyline Poly_easement_L)
        {

            Data_table_easement_left = new System.Data.DataTable();

            Data_table_easement_left.Columns.Add("STATION", typeof(Double));
            Data_table_easement_left.Columns.Add("OFFSET", typeof(Double));

            int Idx = 0;


            for (int i = 0; i < Poly_easement_L.NumberOfVertices; ++i)
            {

                Point3d Nod_poly_easement = new Point3d();
                Nod_poly_easement = Poly_easement_L.GetPoint3dAt(i);
                Point3d Point_on_CL = new Point3d();
                Point_on_CL = PolyCL_MS.GetClosestPointTo(Nod_poly_easement, Vector3d.ZAxis, false);
                double Station1 = Math.Round(PolyCL_MS.GetDistAtPoint(Point_on_CL), Rounding_no);
                double Dist1 = Math.Round(Nod_poly_easement.GetVectorTo(Point_on_CL).Length, Rounding_no);
                double Param1_CL = PolyCL_MS.GetParameterAtPoint(Point_on_CL);
                int Param_C = Convert.ToInt32(Math.Round(Param1_CL, 0));

                Line Line_Easm1 = null;
                Line Line_Easm2 = null;
                Line Line_CL1 = null;
                Line Line_CL2 = null;

                double Line_Length = 0;

                if (i - 1 >= 0)
                {
                    int k = 0;
                    do
                    {
                        if (i - 1 - k >= 0)
                        {
                            Line_Easm1 = new Line(Poly_easement_L.GetPoint3dAt(i), Poly_easement_L.GetPoint3dAt(i - 1 - k));
                            k = k + 1;
                            Line_Length = Line_Easm1.Length;
                        }
                        else
                        {
                            Line_Easm1 = null;
                            goto calcs;
                        }

                    } while (Line_Length <= 0.01);


                }

                Line_Length = 0;

                if (i + 1 < Poly_easement_L.NumberOfVertices)
                {
                    int k = 0;
                    do
                    {
                        if (i + 1 + k < Poly_easement_L.NumberOfVertices)
                        {
                            Line_Easm2 = new Line(Poly_easement_L.GetPoint3dAt(i), Poly_easement_L.GetPoint3dAt(i + 1 + k));
                            k = k + 1;
                            Line_Length = Line_Easm2.Length;
                        }
                        else
                        {
                            Line_Easm2 = null;
                            goto calcs;
                        }

                    } while (Line_Length <= 0.01);



                }

                Line_Length = 0;

                if (Param_C - 1 >= 0)
                {

                    int k = 0;
                    do
                    {
                        if (Param_C - 1 - k >= 0)
                        {
                            Line_CL1 = new Line(PolyCL_MS.GetPoint3dAt(Param_C), PolyCL_MS.GetPoint3dAt(Param_C - 1 - k));
                            k = k + 1;
                            Line_Length = Line_CL1.Length;
                        }
                        else
                        {
                            Line_CL1 = null;
                            goto calcs;
                        }

                    } while (Line_Length <= 0.01);


                }

                Line_Length = 0;

                if (Param_C + 1 < PolyCL_MS.NumberOfVertices)
                {

                    int k = 0;
                    do
                    {
                        if (Param_C + 1 + k < PolyCL_MS.NumberOfVertices)
                        {
                            Line_CL2 = new Line(PolyCL_MS.GetPoint3dAt(Param_C), PolyCL_MS.GetPoint3dAt(Param_C + 1 + k));
                            k = k + 1;
                            Line_Length = Line_CL2.Length;
                        }
                        else
                        {
                            Line_CL2 = null;
                            goto calcs;
                        }

                    } while (Line_Length <= 0.01);
                }

            calcs:

                if (Line_Easm1 != null & Line_Easm2 != null)
                {
                    Vector3d vector1 = Line_Easm1.EndPoint.GetVectorTo(Line_Easm1.StartPoint);
                    Vector3d vector2 = Line_Easm2.StartPoint.GetVectorTo(Line_Easm2.EndPoint);
                    double Defl = Math.Round(vector2.GetAngleTo(vector1), 2);
                    if (Defl >= 1 * Math.PI / 180)
                    {

                        if (Line_CL1 != null & Line_CL2 != null)
                        {

                            double Bearing1 = Math.Round(Workspace_Band.Functions.GET_Bearing_rad(Line_Easm1.StartPoint.X, Line_Easm1.StartPoint.Y, Line_Easm1.EndPoint.X, Line_Easm1.EndPoint.Y), 2);
                            double BearingC1 = Math.Round(Workspace_Band.Functions.GET_Bearing_rad(Line_CL1.StartPoint.X, Line_CL1.StartPoint.Y, Line_CL1.EndPoint.X, Line_CL1.EndPoint.Y), 2);

                            double Bearing2 = Math.Round(Workspace_Band.Functions.GET_Bearing_rad(Line_Easm2.StartPoint.X, Line_Easm2.StartPoint.Y, Line_Easm2.EndPoint.X, Line_Easm2.EndPoint.Y), 2);
                            double BearingC2 = Math.Round(Workspace_Band.Functions.GET_Bearing_rad(Line_CL2.StartPoint.X, Line_CL2.StartPoint.Y, Line_CL2.EndPoint.X, Line_CL2.EndPoint.Y), 2);


                            if (Bearing1 == BearingC1 & Bearing2 == BearingC2)
                            {
                                Vector3d vectorC1 = Line_CL1.EndPoint.GetVectorTo(Line_CL1.StartPoint);
                                Vector3d vectorC2 = Line_CL2.StartPoint.GetVectorTo(Line_CL2.EndPoint);
                                double DeflC = Math.Round(vectorC2.GetAngleTo(vectorC1), 2);
                                if (DeflC >= 1 * Math.PI / 180)
                                {
                                    Point3d Pt1 = new Point3d();
                                    Pt1 = Line_Easm1.GetClosestPointTo(Line_CL1.StartPoint, Vector3d.ZAxis, true);
                                    Point3d Pt2 = new Point3d();
                                    Pt2 = Line_Easm2.GetClosestPointTo(Line_CL2.StartPoint, Vector3d.ZAxis, true);
                                    double L1 = Math.Round(Line_CL1.StartPoint.GetVectorTo(Pt1).Length, Rounding_no);
                                    double L2 = Math.Round(Line_CL2.StartPoint.GetVectorTo(Pt2).Length, Rounding_no);

                                    if (L1 == L2 & Dist1 >= L1)
                                    {

                                        Data_table_easement_left.Rows.Add();
                                        Data_table_easement_left.Rows[Idx]["OFFSET"] = L1;
                                        Data_table_easement_left.Rows[Idx]["STATION"] = Station1;
                                        Idx = Idx + 1;
                                    }
                                    else if (L1 != L2 & (Dist1 >= L1 | Dist1 >= L2))
                                    {
                                        Station1 = Math.Round(PolyCL_MS.GetDistanceAtParameter(Convert.ToDouble(Param_C)), Rounding_no);

                                        Data_table_easement_left.Rows.Add();
                                        Data_table_easement_left.Rows[Idx]["OFFSET"] = L1;
                                        Data_table_easement_left.Rows[Idx]["STATION"] = Station1;
                                        Idx = Idx + 1;
                                        Data_table_easement_left.Rows.Add();
                                        Data_table_easement_left.Rows[Idx]["OFFSET"] = L2;
                                        Data_table_easement_left.Rows[Idx]["STATION"] = Station1;
                                        Idx = Idx + 1;
                                    }
                                    else
                                    {
                                        Data_table_easement_left.Rows.Add();
                                        Data_table_easement_left.Rows[Idx]["OFFSET"] = Dist1;
                                        Data_table_easement_left.Rows[Idx]["STATION"] = Station1;
                                        Idx = Idx + 1;
                                    }
                                }
                                else
                                {
                                    Data_table_easement_left.Rows.Add();
                                    Data_table_easement_left.Rows[Idx]["OFFSET"] = Dist1;
                                    Data_table_easement_left.Rows[Idx]["STATION"] = Station1;
                                    Idx = Idx + 1;
                                }
                            }
                            else
                            {
                                Data_table_easement_left.Rows.Add();
                                Data_table_easement_left.Rows[Idx]["OFFSET"] = Dist1;
                                Data_table_easement_left.Rows[Idx]["STATION"] = Station1;
                                Idx = Idx + 1;
                            }
                        }
                    }
                }
                else
                {
                    Data_table_easement_left.Rows.Add();
                    Data_table_easement_left.Rows[Idx]["OFFSET"] = Dist1;
                    Data_table_easement_left.Rows[Idx]["STATION"] = Station1;
                    Idx = Idx + 1;
                }

            }

            if (Data_table_easement_left != null)
            {
                string Data_easement = "";
                for (int i = 0; i < Data_table_easement_left.Rows.Count - 1; ++i)
                {
                    Double Offset1 = (double)Data_table_easement_left.Rows[i]["OFFSET"];
                    Double Station1 = (double)Data_table_easement_left.Rows[i]["STATION"];
                    Data_easement = Data_easement + "\n" + Convert.ToString(Offset1) + (char)9 + Convert.ToString(Station1);


                }
                System.Windows.Forms.Clipboard.SetText(Data_easement);


            }


            if (Data_table_easement_left != null)
            {
                string Table1 = textBox_COMPILED.Text;

                string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + textBox_access_database_location.Text + ";";

                OleDbConnection cnn;
                try
                {
                    cnn = new OleDbConnection(ConnectionString);
                    cnn.Open();

                    OleDbCommand cmd = new OleDbCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = cnn;

                    Double Station2_SET = -111.111;

                    for (int i = 0; i < Data_table_easement_left.Rows.Count - 1; ++i)
                    {
                        Double Offset0 = -111.111;
                        if (i > 0)
                        {
                            Offset0 = (double)Data_table_easement_left.Rows[i - 1]["OFFSET"];
                        }
                        Double Offset1 = (double)Data_table_easement_left.Rows[i]["OFFSET"];
                        Double Station1 = (double)Data_table_easement_left.Rows[i]["STATION"];

                        Double Offset2 = (double)Data_table_easement_left.Rows[i + 1]["OFFSET"];
                        Double Station2 = (double)Data_table_easement_left.Rows[i + 1]["STATION"];

                        if (Offset0 != Offset1 | Offset0 != Offset2)
                        {
                            cmd.CommandText = "INSERT INTO " + Table1 + "(TYPE,SIDE,BDY,START_STA,END_STA,START_WIDTH,END_WIDTH,WS_NO) VALUES " +
                                "('EPE','L','CL'," + Station1 + "," + Station2 + "," + Offset1 + "," + Offset2 + ",0)";
                            cmd.ExecuteNonQuery();

                            if (Station2_SET != -111.111)
                            {
                                cmd.CommandText = "UPDATE " + Table1 + " SET END_STA = " + Station1 + " WHERE END_STA = " + Station2_SET;
                                cmd.ExecuteNonQuery();

                            }

                            Station2_SET = Station2;
                        }

                        else if (i == Data_table_easement_left.Rows.Count - 2)
                        {
                            cmd.CommandText = "UPDATE " + Table1 + " SET END_STA = " + Station2 + " WHERE END_STA = " + Station2_SET;
                            cmd.ExecuteNonQuery();
                        }



                    }



                    cnn.Close();
                }
                catch (OleDbException ex)
                {
                    MessageBox.Show(ex.Message);
                    Freeze_operations = false;
                }
            }
        }

        private void OLD_Populate_access_database_with_easement_RIGHT(Polyline Poly_easement_R)
        {

            Data_table_easement_right = new System.Data.DataTable();

            Data_table_easement_right.Columns.Add("STATION", typeof(Double));
            Data_table_easement_right.Columns.Add("OFFSET", typeof(Double));

            int Idx = 0;





            for (int i = 0; i < Poly_easement_R.NumberOfVertices; ++i)
            {

                Point3d Nod_poly_easement = new Point3d();
                Nod_poly_easement = Poly_easement_R.GetPoint3dAt(i);
                Point3d Point_on_CL = new Point3d();
                Point_on_CL = PolyCL_MS.GetClosestPointTo(Nod_poly_easement, Vector3d.ZAxis, false);
                double Station1 = Math.Round(PolyCL_MS.GetDistAtPoint(Point_on_CL), Rounding_no);
                double Dist1 = Math.Round(Nod_poly_easement.GetVectorTo(Point_on_CL).Length, Rounding_no);
                double Param1_CL = PolyCL_MS.GetParameterAtPoint(Point_on_CL);
                int Param_C = Convert.ToInt32(Math.Round(Param1_CL, 0));

                Line Line1 = null;
                Line Line2 = null;
                Line LineC1 = null;
                Line LineC2 = null;

                double Len1 = 0;

                if (i - 1 >= 0)
                {
                    int k = 0;
                    do
                    {
                        if (i - 1 - k >= 0)
                        {
                            Line1 = new Line(Poly_easement_R.GetPoint3dAt(i), Poly_easement_R.GetPoint3dAt(i - 1 - k));
                            k = k + 1;
                            Len1 = Line1.Length;
                        }
                        else
                        {
                            Line1 = null;
                            goto calcs;
                        }

                    } while (Len1 <= 0.01);


                }

                Len1 = 0;

                if (i + 1 < Poly_easement_R.NumberOfVertices)
                {
                    int k = 0;
                    do
                    {
                        if (i + 1 + k < Poly_easement_R.NumberOfVertices)
                        {
                            Line2 = new Line(Poly_easement_R.GetPoint3dAt(i), Poly_easement_R.GetPoint3dAt(i + 1 + k));
                            k = k + 1;
                            Len1 = Line2.Length;
                        }
                        else
                        {
                            Line2 = null;
                            goto calcs;
                        }

                    } while (Len1 <= 0.01);



                }

                Len1 = 0;

                if (Param_C - 1 >= 0)
                {

                    int k = 0;
                    do
                    {
                        if (Param_C - 1 - k >= 0)
                        {
                            LineC1 = new Line(PolyCL_MS.GetPoint3dAt(Param_C), PolyCL_MS.GetPoint3dAt(Param_C - 1 - k));
                            k = k + 1;
                            Len1 = LineC1.Length;
                        }
                        else
                        {
                            LineC1 = null;
                            goto calcs;
                        }

                    } while (Len1 <= 0.01);


                }

                Len1 = 0;

                if (Param_C + 1 < PolyCL_MS.NumberOfVertices)
                {

                    int k = 0;
                    do
                    {
                        if (Param_C + 1 + k < PolyCL_MS.NumberOfVertices)
                        {
                            LineC2 = new Line(PolyCL_MS.GetPoint3dAt(Param_C), PolyCL_MS.GetPoint3dAt(Param_C + 1 + k));
                            k = k + 1;
                            Len1 = LineC2.Length;
                        }
                        else
                        {
                            LineC2 = null;
                            goto calcs;
                        }

                    } while (Len1 <= 0.01);
                }

            calcs:

                if (Line1 != null & Line2 != null & LineC1 != null & LineC2 != null)
                {
                    double Bearing1 = Math.Round(Workspace_Band.Functions.GET_Bearing_rad(Line1.StartPoint.X, Line1.StartPoint.Y, Line1.EndPoint.X, Line1.EndPoint.Y), 2);
                    double BearingC1 = Math.Round(Workspace_Band.Functions.GET_Bearing_rad(LineC1.StartPoint.X, LineC1.StartPoint.Y, LineC1.EndPoint.X, LineC1.EndPoint.Y), 2);

                    double Bearing2 = Math.Round(Workspace_Band.Functions.GET_Bearing_rad(Line2.StartPoint.X, Line2.StartPoint.Y, Line2.EndPoint.X, Line2.EndPoint.Y), 2);
                    double BearingC2 = Math.Round(Workspace_Band.Functions.GET_Bearing_rad(LineC2.StartPoint.X, LineC2.StartPoint.Y, LineC2.EndPoint.X, LineC2.EndPoint.Y), 2);


                    if (Bearing1 == BearingC1 & Bearing2 == BearingC2)
                    {
                        Vector3d vector1 = LineC1.EndPoint.GetVectorTo(LineC1.StartPoint);
                        Vector3d vector2 = LineC2.StartPoint.GetVectorTo(LineC2.EndPoint);
                        double Defl = Math.Round(vector2.GetAngleTo(vector1), 2);
                        if (Defl != 0)
                        {
                            Point3d Pt1 = new Point3d();
                            Pt1 = Line1.GetClosestPointTo(LineC1.StartPoint, Vector3d.ZAxis, true);
                            Point3d Pt2 = new Point3d();
                            Pt2 = Line2.GetClosestPointTo(LineC2.StartPoint, Vector3d.ZAxis, true);
                            double L1 = Math.Round(LineC1.StartPoint.GetVectorTo(Pt1).Length, Rounding_no);
                            double L2 = Math.Round(LineC2.StartPoint.GetVectorTo(Pt2).Length, Rounding_no);

                            if (L1 == L2 & Dist1 >= L1)
                            {

                                Data_table_easement_right.Rows.Add();
                                Data_table_easement_right.Rows[Idx]["OFFSET"] = L1;
                                Data_table_easement_right.Rows[Idx]["STATION"] = Station1;
                                Idx = Idx + 1;
                            }
                            else if (L1 != L2 & (Dist1 >= L1 | Dist1 >= L2))
                            {
                                Station1 = Math.Round(PolyCL_MS.GetDistanceAtParameter(Convert.ToDouble(Param_C)), Rounding_no);

                                Data_table_easement_right.Rows.Add();
                                Data_table_easement_right.Rows[Idx]["OFFSET"] = L1;
                                Data_table_easement_right.Rows[Idx]["STATION"] = Station1;
                                Idx = Idx + 1;
                                Data_table_easement_right.Rows.Add();
                                Data_table_easement_right.Rows[Idx]["OFFSET"] = L2;
                                Data_table_easement_right.Rows[Idx]["STATION"] = Station1;
                                Idx = Idx + 1;
                            }
                            else
                            {
                                Data_table_easement_right.Rows.Add();
                                Data_table_easement_right.Rows[Idx]["OFFSET"] = Dist1;
                                Data_table_easement_right.Rows[Idx]["STATION"] = Station1;
                                Idx = Idx + 1;
                            }
                        }
                        else
                        {
                            Data_table_easement_right.Rows.Add();
                            Data_table_easement_right.Rows[Idx]["OFFSET"] = Dist1;
                            Data_table_easement_right.Rows[Idx]["STATION"] = Station1;
                            Idx = Idx + 1;
                        }
                    }
                    else
                    {
                        Data_table_easement_right.Rows.Add();
                        Data_table_easement_right.Rows[Idx]["OFFSET"] = Dist1;
                        Data_table_easement_right.Rows[Idx]["STATION"] = Station1;
                        Idx = Idx + 1;
                    }
                }
                else
                {
                    Data_table_easement_right.Rows.Add();
                    Data_table_easement_right.Rows[Idx]["OFFSET"] = Dist1;
                    Data_table_easement_right.Rows[Idx]["STATION"] = Station1;
                    Idx = Idx + 1;
                }
            }


            if (Data_table_easement_right != null)
            {
                string Table1 = textBox_COMPILED.Text;

                string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + textBox_access_database_location.Text + ";";

                OleDbConnection cnn;
                try
                {
                    cnn = new OleDbConnection(ConnectionString);
                    cnn.Open();

                    OleDbCommand cmd = new OleDbCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = cnn;

                    Double Station2_SET = -111.111;

                    for (int i = 0; i < Data_table_easement_right.Rows.Count - 1; ++i)
                    {
                        Double Offset0 = -111.111;
                        if (i > 0)
                        {
                            Offset0 = (double)Data_table_easement_right.Rows[i - 1]["OFFSET"];
                        }
                        Double Offset1 = (double)Data_table_easement_right.Rows[i]["OFFSET"];
                        Double Station1 = (double)Data_table_easement_right.Rows[i]["STATION"];

                        Double Offset2 = (double)Data_table_easement_right.Rows[i + 1]["OFFSET"];
                        Double Station2 = (double)Data_table_easement_right.Rows[i + 1]["STATION"];

                        if (Offset0 != Offset1 | Offset0 != Offset2)
                        {
                            cmd.CommandText = "INSERT INTO " + Table1 + "(TYPE,SIDE,BDY,START_STA,END_STA,START_WIDTH,END_WIDTH,WS_NO) VALUES " +
                                "('EPE','R','CL'," + Station1 + "," + Station2 + "," + Offset1 + "," + Offset2 + ",0)";
                            cmd.ExecuteNonQuery();

                            if (Station2_SET != -111.111)
                            {
                                cmd.CommandText = "UPDATE " + Table1 + " SET END_STA = " + Station1 + " WHERE END_STA = " + Station2_SET;
                                cmd.ExecuteNonQuery();

                            }

                            Station2_SET = Station2;
                        }

                        else if (i == Data_table_easement_right.Rows.Count - 2)
                        {
                            cmd.CommandText = "UPDATE " + Table1 + " SET END_STA = " + Station2 + " WHERE END_STA = " + Station2_SET;
                            cmd.ExecuteNonQuery();
                        }



                    }



                    cnn.Close();
                }
                catch (OleDbException ex)
                {
                    MessageBox.Show(ex.Message);
                    Freeze_operations = false;
                }
            }
        }



        private void Populate_Compiled_DATA_TABLE_with_easement_info(string Type1, string Side1, int Ws_no)
        {
            try
            {
                string Table1 = textBox_COMPILED.Text;
                string query = "SELECT * FROM " + Table1 + " WHERE TYPE='" + Type1 + "' AND SIDE='" + Side1 + "' AND WS_NO = " + Ws_no;
                string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + textBox_access_database_location.Text + ";";
                DataSet COMPILED_DATASET = new DataSet();
                OleDbConnection cnn;
                try
                {
                    cnn = new OleDbConnection(ConnectionString);
                    cnn.Open();
                    OleDbCommand cmd = new OleDbCommand(query, cnn);
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(COMPILED_DATASET, Table1);
                    cnn.Close();
                }
                catch (OleDbException ex)
                {
                    MessageBox.Show(ex.Message);
                }

                Compiled_DATA_TABLE = new System.Data.DataTable();


                Compiled_DATA_TABLE = COMPILED_DATASET.Tables[Table1];

                //MessageBox.Show(ROW_DATA_TABLE.Columns[4].ColumnName + " Row 2 = " + ROW_DATA_TABLE.Rows[1][4]);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }


        private void DRAW_easement_from_compiled(Polyline Poly_CL_Straight, Polyline Poly_easment, System.Data.DataTable Compiled_DATA_TABLE, int Up_down)
        {
            if (Compiled_DATA_TABLE != null)
            {
                if (Compiled_DATA_TABLE.Rows.Count > 0)
                {
                    try
                    {

                        int Index_poly = 0;
                        for (int i = 0; i < Compiled_DATA_TABLE.Rows.Count; i = i + 1)
                        {
                            double Sta1 = -1;
                            double Sta2 = -1;
                            if (Compiled_DATA_TABLE.Rows[i]["START_STA"] != DBNull.Value)
                            {
                                Sta1 = (double)Compiled_DATA_TABLE.Rows[i]["START_STA"];
                            }

                            if (Compiled_DATA_TABLE.Rows[i]["END_STA"] != DBNull.Value)
                            {
                                Sta2 = (double)Compiled_DATA_TABLE.Rows[i]["END_STA"];
                            }

                            double Wdth1 = -1;
                            double Wdth2 = -1;
                            if (Compiled_DATA_TABLE.Rows[i]["START_WIDTH"] != DBNull.Value)
                            {
                                Wdth1 = (double)Compiled_DATA_TABLE.Rows[i]["START_WIDTH"];
                            }

                            if (Compiled_DATA_TABLE.Rows[i]["END_WIDTH"] != DBNull.Value)
                            {
                                Wdth2 = (double)Compiled_DATA_TABLE.Rows[i]["END_WIDTH"];
                            }

                            if (Sta1 != -1 & Sta2 != -1 & Wdth1 != -1 & Wdth2 != -1)
                            {
                                Point3d Point1 = new Point3d();
                                Point1 = Poly_CL_Straight.GetPointAtDist(Sta1);
                                Point3d Point2 = new Point3d();
                                Point2 = Poly_CL_Straight.GetPointAtDist(Sta2);

                                Poly_easment.AddVertexAt(Index_poly, new Point2d(Point1.X, Point1.Y + Up_down * Wdth1), 0, 0, 0);
                                Index_poly = Index_poly + 1;
                                Poly_easment.AddVertexAt(Index_poly, new Point2d(Point2.X, Point2.Y + Up_down * Wdth2), 0, 0, 0);
                                Index_poly = Index_poly + 1;
                            }

                        }
                    }

                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
        }

        private void DRAW_workspace_from_compiled(Polyline Poly_CL_Straight, Polyline Poly_easmentR, Polyline Poly_easmentL, Polyline Poly_workspace, System.Data.DataTable Compiled_DATA_TABLE)
        {
            if (Compiled_DATA_TABLE != null)
            {
                if (Compiled_DATA_TABLE.Rows.Count > 0)
                {
                    try
                    {

                        int Index_poly = 0;
                        int UP_DOWN = -1;
                        
                        if (radioButton_left_UP.Checked == true)
                        {
                            UP_DOWN = 1;
                        }

                        Polyline Poly_easment = Poly_easmentR;
                        if (Compiled_DATA_TABLE.Rows[0]["SIDE"] != DBNull.Value)
                        {
                            string Side1 = Convert.ToString((Compiled_DATA_TABLE.Rows[0]["SIDE"]));

                            if (Side1 == "L")
                            {

                                Poly_easment = Poly_easmentL;
                            }


                            if (Side1 == "R")
                            {

                                UP_DOWN = -UP_DOWN;
                            }

                        }

                        for (int i = 0; i < Compiled_DATA_TABLE.Rows.Count; i = i + 1)
                        {
                            double Sta1 = -1;
                            double Sta2 = -1;
                            if (Compiled_DATA_TABLE.Rows[i]["START_STA"] != DBNull.Value)
                            {
                                Sta1 = (double)Compiled_DATA_TABLE.Rows[i]["START_STA"];
                            }

                            if (Compiled_DATA_TABLE.Rows[i]["END_STA"] != DBNull.Value)
                            {
                                Sta2 = (double)Compiled_DATA_TABLE.Rows[i]["END_STA"];
                            }

                            double Wdth1 = -1;
                            double Wdth2 = -1;
                            if (Compiled_DATA_TABLE.Rows[i]["START_WIDTH"] != DBNull.Value)
                            {
                                Wdth1 = (double)Compiled_DATA_TABLE.Rows[i]["START_WIDTH"];
                            }

                            if (Compiled_DATA_TABLE.Rows[i]["END_WIDTH"] != DBNull.Value)
                            {
                                Wdth2 = (double)Compiled_DATA_TABLE.Rows[i]["END_WIDTH"];
                            }



                            if (Sta1 != -1 & Sta2 != -1 & Wdth1 != -1 & Wdth2 != -1)
                            {
                                Point3d Point1 = new Point3d();
                                Point1 = Poly_CL_Straight.GetPointAtDist(Sta1);
                                Point3d Point2 = new Point3d();
                                Point2 = Poly_CL_Straight.GetPointAtDist(Sta2);

                                Line Line1 = new Line(Point1, new Point3d(Point1.X, Point1.Y + UP_DOWN * 100000, Poly_easment.Elevation));
                                Point3dCollection Collection1 = new Point3dCollection();
                                Line1.IntersectWith(Poly_easment, Intersect.OnBothOperands, Collection1, IntPtr.Zero, IntPtr.Zero);

                                Line Line2 = new Line(Point2, new Point3d(Point2.X, Point2.Y + UP_DOWN * 100000, Poly_easment.Elevation));
                                Point3dCollection Collection2 = new Point3dCollection();
                                Line2.IntersectWith(Poly_easment, Intersect.OnBothOperands, Collection2, IntPtr.Zero, IntPtr.Zero);

                                if (Collection1.Count > 0 & Collection2.Count > 0)
                                {
                                    Point3d PointE1 = Collection1[0];
                                    Point3d PointE2 = Collection2[0];
                                    Poly_workspace.AddVertexAt(Index_poly, new Point2d(PointE1.X, PointE1.Y + UP_DOWN * Wdth1), 0, 0, 0);
                                    Index_poly = Index_poly + 1;
                                    Poly_workspace.AddVertexAt(Index_poly, new Point2d(PointE2.X, PointE2.Y + UP_DOWN * Wdth2), 0, 0, 0);
                                    Index_poly = Index_poly + 1;
                                }




                            }

                        }
                    }

                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
        }

        private void Populate_Compiled_DATA_TABLE_with_workspace_info(string Type1, int Ws_no)
        {
            try
            {
                string Table1 = textBox_COMPILED.Text;
                string query = "SELECT * FROM " + Table1 + " WHERE TYPE='" + Type1 + "' AND WS_NO = " + Ws_no;
                string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + textBox_access_database_location.Text + ";";
                DataSet COMPILED_DATASET = new DataSet();
                OleDbConnection cnn;
                try
                {
                    cnn = new OleDbConnection(ConnectionString);
                    cnn.Open();
                    OleDbCommand cmd = new OleDbCommand(query, cnn);
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(COMPILED_DATASET, Table1);
                    cnn.Close();
                }
                catch (OleDbException ex)
                {
                    MessageBox.Show(ex.Message);
                }

                Compiled_DATA_TABLE = new System.Data.DataTable();


                Compiled_DATA_TABLE = COMPILED_DATASET.Tables[Table1];

                //MessageBox.Show(ROW_DATA_TABLE.Columns[4].ColumnName + " Row 2 = " + ROW_DATA_TABLE.Rows[1][4]);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }




        private void button_Pick_information_for_CL_EASEMENT_TWS(object sender, EventArgs e)
        {
            if (Freeze_operations == false)
            {
                Freeze_operations = true;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                try
                {
                    Plane Planul_curent = new Plane(new Point3d(0, 0, 0), Vector3d.ZAxis);
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                            Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                            Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the centerline:");
                            Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                            Prompt_centerline.AllowNone = true;
                            Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                            Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);

                            if (Rezultat_centerline.Status != PromptStatus.OK)
                            {
                                Freeze_operations = false;
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            PolyCL_MS = (Polyline)Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead);

                            Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_easement_L;
                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_easement_L;
                            Prompt_easement_L = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the existing left easement:");
                            Prompt_easement_L.SetRejectMessage("\nSelect a polyline!");
                            Prompt_easement_L.AllowNone = true;
                            Prompt_easement_L.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                            Rezultat_easement_L = ThisDrawing.Editor.GetEntity(Prompt_easement_L);


                            Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_easement_R;
                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_easement_R;
                            Prompt_easement_R = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the existing right easement:");
                            Prompt_easement_R.SetRejectMessage("\nSelect a polyline!");
                            Prompt_easement_R.AllowNone = true;
                            Prompt_easement_R.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                            Rezultat_easement_R = ThisDrawing.Editor.GetEntity(Prompt_easement_R);


                            Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_workspace;
                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_workspace;
                            Prompt_workspace = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect one object from the Workspace layer:");
                            Prompt_workspace.SetRejectMessage("\nSelect a polyline!");
                            Prompt_workspace.AllowNone = true;
                            Prompt_workspace.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                            Rezultat_workspace = ThisDrawing.Editor.GetEntity(Prompt_workspace);


                            Polyline Poly_easement_L = null;

                            if (Rezultat_easement_L.Status == PromptStatus.OK)
                            {
                                Poly_easement_L = (Polyline)Trans1.GetObject(Rezultat_easement_L.ObjectId, OpenMode.ForWrite);
                                Poly_easement_L.Elevation = PolyCL_MS.Elevation;


                                if (Poly_easement_L.StartPoint.GetVectorTo(PolyCL_MS.StartPoint).Length > Poly_easement_L.EndPoint.GetVectorTo(PolyCL_MS.StartPoint).Length)
                                {
                                    try
                                    {
                                        var Obj1 = Poly_easement_L as Polyline;

                                        Obj1.ReverseCurve();
                                        Poly_easement_L = Obj1;

                                    }
                                    catch
                                    {

                                    }
                                }

                                Populate_access_database_with_easement(Poly_easement_L, "L");
                            }


                            Polyline Poly_easement_R = null;

                            if (Rezultat_easement_R.Status == PromptStatus.OK)
                            {
                                Poly_easement_R = (Polyline)Trans1.GetObject(Rezultat_easement_R.ObjectId, OpenMode.ForWrite);
                                Poly_easement_R.Elevation = PolyCL_MS.Elevation;


                                if (Poly_easement_R.StartPoint.GetVectorTo(PolyCL_MS.StartPoint).Length > Poly_easement_R.EndPoint.GetVectorTo(PolyCL_MS.StartPoint).Length)
                                {
                                    try
                                    {
                                        var Obj1 = Poly_easement_R as Polyline;

                                        Obj1.ReverseCurve();
                                        Poly_easement_R = Obj1;

                                    }
                                    catch
                                    {

                                    }
                                }

                                Populate_access_database_with_easement(Poly_easement_R, "R");
                            }





                            if (Rezultat_workspace.Status == PromptStatus.OK)
                            {
                                Entity Entity_workspace = (Entity)Trans1.GetObject(Rezultat_workspace.ObjectId, OpenMode.ForRead);

                                if (Entity_workspace != null)
                                {
                                    Populate_Access_Database_with_Workspace_information(Trans1, BTrecord, Entity_workspace.Layer, Poly_easement_R, Poly_easement_L);
                                }
                            }



                            Trans1.Commit();
                        }
                    }








                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    Freeze_operations = false;
                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                }
                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                MessageBox.Show("Read complete");
                Freeze_operations = false;
            }

        }



        private void Populate_access_database_with_easement(Polyline Poly_easement, string Side1)
        {

            Data_table_easement = new System.Data.DataTable();

            Data_table_easement.Columns.Add("STATION", typeof(Double));
            Data_table_easement.Columns.Add("OFFSET", typeof(Double));

            int Idx = 0;


            for (int i = 0; i < Poly_easement.NumberOfVertices; ++i)
            {

                Point3d Nod_poly_easement = new Point3d();
                Nod_poly_easement = Poly_easement.GetPoint3dAt(i);
                Point3d Point_on_CL = new Point3d();
                Point_on_CL = PolyCL_MS.GetClosestPointTo(Nod_poly_easement, Vector3d.ZAxis, false);
                double Station1 = Math.Round(PolyCL_MS.GetDistAtPoint(Point_on_CL), Rounding_no);
                double Dist1 = Math.Round(Nod_poly_easement.GetVectorTo(Point_on_CL).Length, Rounding_no);
                double Param1_CL = PolyCL_MS.GetParameterAtPoint(Point_on_CL);
                int Param_C = Convert.ToInt32(Math.Round(Param1_CL, 0));

                Line Line_Easm1 = null;
                Line Line_Easm2 = null;
                Line Line_CL1 = null;
                Line Line_CL2 = null;

                double Line_Length = 0;

                if (i - 1 >= 0)
                {
                    int k = 0;
                    do
                    {
                        if (i - 1 - k >= 0)
                        {
                            Line_Easm1 = new Line(Poly_easement.GetPoint3dAt(i), Poly_easement.GetPoint3dAt(i - 1 - k));
                            k = k + 1;
                            Line_Length = Line_Easm1.Length;
                        }
                        else
                        {
                            Line_Easm1 = null;
                            goto calcs;
                        }

                    } while (Line_Length <= 0.01);


                }

                Line_Length = 0;

                if (i + 1 < Poly_easement.NumberOfVertices)
                {
                    int k = 0;
                    do
                    {
                        if (i + 1 + k < Poly_easement.NumberOfVertices)
                        {
                            Line_Easm2 = new Line(Poly_easement.GetPoint3dAt(i), Poly_easement.GetPoint3dAt(i + 1 + k));
                            k = k + 1;
                            Line_Length = Line_Easm2.Length;
                        }
                        else
                        {
                            Line_Easm2 = null;
                            goto calcs;
                        }

                    } while (Line_Length <= 0.01);



                }

                Line_Length = 0;

                if (Param_C - 1 >= 0)
                {

                    int k = 0;
                    do
                    {
                        if (Param_C - 1 - k >= 0)
                        {
                            Line_CL1 = new Line(PolyCL_MS.GetPoint3dAt(Param_C), PolyCL_MS.GetPoint3dAt(Param_C - 1 - k));
                            k = k + 1;
                            Line_Length = Line_CL1.Length;
                        }
                        else
                        {
                            Line_CL1 = null;
                            goto calcs;
                        }

                    } while (Line_Length <= 0.01);


                }

                Line_Length = 0;

                if (Param_C + 1 < PolyCL_MS.NumberOfVertices)
                {

                    int k = 0;
                    do
                    {
                        if (Param_C + 1 + k < PolyCL_MS.NumberOfVertices)
                        {
                            Line_CL2 = new Line(PolyCL_MS.GetPoint3dAt(Param_C), PolyCL_MS.GetPoint3dAt(Param_C + 1 + k));
                            k = k + 1;
                            Line_Length = Line_CL2.Length;
                        }
                        else
                        {
                            Line_CL2 = null;
                            goto calcs;
                        }

                    } while (Line_Length <= 0.01);
                }

            calcs:

                if (Line_Easm1 != null & Line_Easm2 != null)
                {
                    Vector3d vector1 = Line_Easm1.EndPoint.GetVectorTo(Line_Easm1.StartPoint);
                    Vector3d vector2 = Line_Easm2.StartPoint.GetVectorTo(Line_Easm2.EndPoint);
                    double Defl = Math.Round(vector2.GetAngleTo(vector1), 2);
                    if (Defl >= 1 * Math.PI / 180)
                    {

                        if (Line_CL1 != null & Line_CL2 != null)
                        {

                            double Bearing1 = Math.Round(Workspace_Band.Functions.GET_Bearing_rad(Line_Easm1.StartPoint.X, Line_Easm1.StartPoint.Y, Line_Easm1.EndPoint.X, Line_Easm1.EndPoint.Y), 2);
                            double BearingC1 = Math.Round(Workspace_Band.Functions.GET_Bearing_rad(Line_CL1.StartPoint.X, Line_CL1.StartPoint.Y, Line_CL1.EndPoint.X, Line_CL1.EndPoint.Y), 2);

                            double Bearing2 = Math.Round(Workspace_Band.Functions.GET_Bearing_rad(Line_Easm2.StartPoint.X, Line_Easm2.StartPoint.Y, Line_Easm2.EndPoint.X, Line_Easm2.EndPoint.Y), 2);
                            double BearingC2 = Math.Round(Workspace_Band.Functions.GET_Bearing_rad(Line_CL2.StartPoint.X, Line_CL2.StartPoint.Y, Line_CL2.EndPoint.X, Line_CL2.EndPoint.Y), 2);

                            if (Bearing1 >= 2 * 3.14) Bearing1 = Bearing1 - 2 * 3.14;
                            if (BearingC1 >= 2 * 3.14) BearingC1 = BearingC1 - 2 * 3.14;
                            if (Bearing2 >= 2 * 3.14) Bearing2 = Bearing2 - 2 * 3.14;
                            if (BearingC2 >= 2 * 3.14) BearingC2 = BearingC2 - 2 * 3.14;



                            if (Bearing1 == BearingC1 & Bearing2 == BearingC2)
                            {
                                Vector3d vectorC1 = Line_CL1.EndPoint.GetVectorTo(Line_CL1.StartPoint);
                                Vector3d vectorC2 = Line_CL2.StartPoint.GetVectorTo(Line_CL2.EndPoint);
                                double DeflC = Math.Round(vectorC2.GetAngleTo(vectorC1), 2);
                                if (DeflC >= 1 * Math.PI / 180)
                                {
                                    Point3d Pt1 = new Point3d();
                                    Pt1 = Line_Easm1.GetClosestPointTo(Line_CL1.StartPoint, Vector3d.ZAxis, true);
                                    Point3d Pt2 = new Point3d();
                                    Pt2 = Line_Easm2.GetClosestPointTo(Line_CL2.StartPoint, Vector3d.ZAxis, true);
                                    double L1 = Math.Round(Line_CL1.StartPoint.GetVectorTo(Pt1).Length, Rounding_no);
                                    double L2 = Math.Round(Line_CL2.StartPoint.GetVectorTo(Pt2).Length, Rounding_no);

                                    if (L1 == L2 & Dist1 >= L1)
                                    {

                                        Data_table_easement.Rows.Add();
                                        Data_table_easement.Rows[Idx]["OFFSET"] = L1;
                                        Data_table_easement.Rows[Idx]["STATION"] = Station1;
                                        Idx = Idx + 1;
                                    }
                                    else if (L1 != L2 & (Dist1 >= L1 | Dist1 >= L2))
                                    {
                                        Station1 = Math.Round(PolyCL_MS.GetDistanceAtParameter(Convert.ToDouble(Param_C)), Rounding_no);

                                        Data_table_easement.Rows.Add();
                                        Data_table_easement.Rows[Idx]["OFFSET"] = L1;
                                        Data_table_easement.Rows[Idx]["STATION"] = Station1;
                                        Idx = Idx + 1;
                                        Data_table_easement.Rows.Add();
                                        Data_table_easement.Rows[Idx]["OFFSET"] = L2;
                                        Data_table_easement.Rows[Idx]["STATION"] = Station1;
                                        Idx = Idx + 1;
                                    }
                                    else
                                    {
                                        Data_table_easement.Rows.Add();
                                        Data_table_easement.Rows[Idx]["OFFSET"] = Dist1;
                                        Data_table_easement.Rows[Idx]["STATION"] = Station1;
                                        Idx = Idx + 1;
                                    }
                                }
                                else
                                {
                                    Data_table_easement.Rows.Add();
                                    Data_table_easement.Rows[Idx]["OFFSET"] = Dist1;
                                    Data_table_easement.Rows[Idx]["STATION"] = Station1;
                                    Idx = Idx + 1;
                                }
                            }
                            else
                            {
                                Data_table_easement.Rows.Add();
                                Data_table_easement.Rows[Idx]["OFFSET"] = Dist1;
                                Data_table_easement.Rows[Idx]["STATION"] = Station1;
                                Idx = Idx + 1;
                            }
                        }
                    }
                }
                else
                {
                    Data_table_easement.Rows.Add();
                    Data_table_easement.Rows[Idx]["OFFSET"] = Dist1;
                    Data_table_easement.Rows[Idx]["STATION"] = Station1;
                    Idx = Idx + 1;
                }

            }

            if (Data_table_easement != null)
            {
                string Data_easement = "";
                for (int i = 0; i < Data_table_easement.Rows.Count - 1; ++i)
                {
                    Double Offset1 = (double)Data_table_easement.Rows[i]["OFFSET"];
                    Double Station1 = (double)Data_table_easement.Rows[i]["STATION"];
                    Data_easement = Data_easement + "\n" + Convert.ToString(Offset1) + (char)9 + Convert.ToString(Station1);


                }
                System.Windows.Forms.Clipboard.SetText(Data_easement);


            }


            if (Data_table_easement != null)
            {
                string Table1 = textBox_COMPILED.Text;

                string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + textBox_access_database_location.Text + ";";

                OleDbConnection cnn;
                try
                {
                    cnn = new OleDbConnection(ConnectionString);
                    cnn.Open();

                    OleDbCommand cmd = new OleDbCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = cnn;

                    Double Station2_SET = -111.111;

                    for (int i = 0; i < Data_table_easement.Rows.Count - 1; ++i)
                    {
                        Double Offset0 = -111.111;
                        if (i > 0)
                        {
                            Offset0 = (double)Data_table_easement.Rows[i - 1]["OFFSET"];
                        }
                        Double Offset1 = (double)Data_table_easement.Rows[i]["OFFSET"];
                        Double Station1 = (double)Data_table_easement.Rows[i]["STATION"];

                        Double Offset2 = (double)Data_table_easement.Rows[i + 1]["OFFSET"];
                        Double Station2 = (double)Data_table_easement.Rows[i + 1]["STATION"];

                        if (Offset0 != Offset1 | Offset0 != Offset2)
                        {
                            cmd.CommandText = "INSERT INTO " + Table1 + "(TYPE,SIDE,BDY,START_STA,END_STA,START_WIDTH,END_WIDTH,WS_NO) VALUES " +
                                "('EPE','" + Side1 + "','CL'," + Station1 + "," + Station2 + "," + Offset1 + "," + Offset2 + ",0)";
                            cmd.ExecuteNonQuery();

                            if (Station2_SET != -111.111)
                            {
                                cmd.CommandText = "UPDATE " + Table1 + " SET END_STA = " + Station1 + " WHERE END_STA = " + Station2_SET;
                                cmd.ExecuteNonQuery();

                            }

                            Station2_SET = Station2;
                        }

                        else if (i == Data_table_easement.Rows.Count - 2)
                        {
                            cmd.CommandText = "UPDATE " + Table1 + " SET END_STA = " + Station2 + " WHERE END_STA = " + Station2_SET;
                            cmd.ExecuteNonQuery();
                        }



                    }



                    cnn.Close();
                }
                catch (OleDbException ex)
                {
                    MessageBox.Show(ex.Message);
                    Freeze_operations = false;
                }
            }
        }

        private void Populate_Access_Database_with_Workspace_information(Autodesk.AutoCAD.DatabaseServices.Transaction Trans1, Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord,
                                                                            String Layer_workspace, Polyline Poly_easement_R, Polyline Poly_easement_L)
        {

            int WS_no = 1;
            foreach (ObjectId Obj_ID1 in BTrecord)
            {
                Entity Ent1 = (Entity)Trans1.GetObject(Obj_ID1, OpenMode.ForRead);
                if (Ent1 != null)
                {
                    if (Ent1.Layer == Layer_workspace & Ent1 is Polyline)
                    {
                        Polyline WS_poly = (Polyline)Ent1;
                        if (WS_poly.Closed == true)
                        {

                            int Extra_index = -1;
                            double Extra_offset = -1;
                            double Extra_station = -1;

                            WS_poly.UpgradeOpen();
                            WS_poly.Elevation = 0;

                            string Side_of_WS = "";
                            string BDY_WS = "EPE";
                            System.Data.DataTable Data_Table_WS = new System.Data.DataTable();
                            Data_Table_WS.Columns.Add("X", typeof(double));
                            Data_Table_WS.Columns.Add("Y", typeof(double));
                            Data_Table_WS.Columns.Add("STA", typeof(double));
                            Data_Table_WS.Columns.Add("OFFSET", typeof(double));
                            for (int i = 0; i < WS_poly.NumberOfVertices; ++i)
                            {
                                Point3d Point_WS = WS_poly.GetPointAtParameter(i);
                                Point3d Point_CL = PolyCL_MS.GetClosestPointTo(Point_WS, Vector3d.ZAxis, false);

                                Polyline Tie_WS_CL_poly = new Polyline();
                                Tie_WS_CL_poly.AddVertexAt(0, new Point2d(Point_WS.X, Point_WS.Y), 0, 0, 0);
                                Tie_WS_CL_poly.AddVertexAt(1, new Point2d(Point_CL.X, Point_CL.Y), 0, 0, 0);
                                Tie_WS_CL_poly.Elevation = PolyCL_MS.Elevation;

                                Point3dCollection Col_int_L;
                                Col_int_L = new Point3dCollection();
                                Poly_easement_L.IntersectWith(Tie_WS_CL_poly, Intersect.OnBothOperands, Col_int_L, IntPtr.Zero, IntPtr.Zero);

                                Point3dCollection Col_int_R = new Point3dCollection();
                                Poly_easement_R.IntersectWith(Tie_WS_CL_poly, Intersect.OnBothOperands, Col_int_R, IntPtr.Zero, IntPtr.Zero);
                                if (Side_of_WS == "")
                                {
                                    if (Col_int_L.Count > 0)
                                    {
                                        Side_of_WS = "L";
                                        i = WS_poly.NumberOfVertices;
                                    }
                                    if (Col_int_R.Count > 0)
                                    {
                                        Side_of_WS = "R";
                                        i = WS_poly.NumberOfVertices;

                                    }
                                }
                            }
                            Point3dCollection Col_pt = new Point3dCollection();
                            int j = 0;

                            Polyline WS_clean_poly = Workspace_Band.Functions.Clean_poly_of_duplicate_points(WS_poly, 1);
                            WS_clean_poly = Workspace_Band.Functions.Clean_poly_of_deflection_points(WS_clean_poly, 1 * Math.PI / 180);

                            Polyline PolyCL_cleaned = Workspace_Band.Functions.Clean_poly_of_duplicate_points(PolyCL_MS, 1);
                            PolyCL_cleaned = Workspace_Band.Functions.Clean_poly_of_deflection_points(PolyCL_cleaned, 1 * Math.PI / 180);


                            double Param_min, Param_max;

                            Param_min = 0;
                            Param_max = 0;

                            double Param_min_L, Param_max_L;
                            Param_min_L = 0;
                            Param_max_L = 0;

                            double Param_min_R, Param_max_R;
                            Param_min_R = 0;
                            Param_max_R = 0;

                            for (int i = 0; i < WS_clean_poly.NumberOfVertices; ++i)
                            {
                                Point3d Point_WS = WS_clean_poly.GetPointAtParameter(i);

                                Point3d Point_CL = PolyCL_cleaned.GetClosestPointTo(Point_WS, Vector3d.ZAxis, false);
                                double Param_CL = PolyCL_cleaned.GetParameterAtPoint(Point_CL);

                                Point3d Point_EL = Poly_easement_L.GetClosestPointTo(Point_WS, Vector3d.ZAxis, false);
                                double Param_EL = Poly_easement_L.GetParameterAtPoint(Point_EL);

                                Point3d Point_ER = Poly_easement_R.GetClosestPointTo(Point_WS, Vector3d.ZAxis, false);
                                double Param_ER = Poly_easement_R.GetParameterAtPoint(Point_ER);

                                if (i == 0)
                                {
                                    Param_max = Param_CL;
                                    Param_min = Param_CL;
                                    Param_max_L = Param_EL;
                                    Param_min_L = Param_EL;
                                    Param_max_R = Param_ER;
                                    Param_min_R = Param_ER;
                                }

                                if (Param_CL > Param_max)
                                {
                                    Param_max = Param_CL;

                                }

                                if (Param_CL < Param_min)
                                {

                                    Param_min = Param_CL;
                                }

                                if (Param_EL > Param_max_L)
                                {
                                    Param_max_L = Param_EL;

                                }

                                if (Param_EL < Param_min_L)
                                {

                                    Param_min_L = Param_EL;
                                }


                                if (Param_ER > Param_max_R)
                                {
                                    Param_max_R = Param_ER;

                                }

                                if (Param_ER < Param_min_R)
                                {

                                    Param_min_R = Param_ER;
                                }
                            }

                            for (int i = 0; i < WS_clean_poly.NumberOfVertices; ++i)
                            {
                                Point3d Point_WS = WS_clean_poly.GetPointAtParameter(i);
                                Point3d Point_CL = PolyCL_cleaned.GetClosestPointTo(Point_WS, Vector3d.ZAxis, false);
                                double Param_CL = PolyCL_cleaned.GetParameterAtPoint(Point_CL);

                                double Param_CL_rounded = Math.Round(Param_CL, 0);


                                do
                                {
                                    Param_CL_rounded = Param_CL_rounded + 1;
                                } while (Param_CL_rounded < Param_min);

                                do
                                {
                                    Param_CL_rounded = Param_CL_rounded - 1;
                                } while (Param_CL_rounded > Param_max);

                                Point3d Point_CL_rounded = PolyCL_cleaned.GetPointAtParameter(Param_CL_rounded);





                                Point3d Point_WS_minus1 = new Point3d();
                                Point3d Point_WS_plus1 = new Point3d();

                                Point3d Point_CL_minus1 = new Point3d();
                                Point3d Point_CL_plus1 = new Point3d();

                                if (i == WS_clean_poly.NumberOfVertices - 1)
                                {
                                    Point_WS_minus1 = WS_clean_poly.GetPointAtParameter(i - 1);
                                    Point_WS_plus1 = WS_clean_poly.GetPointAtParameter(0);

                                }
                                else if (i == 0)
                                {
                                    Point_WS_minus1 = WS_clean_poly.GetPointAtParameter(WS_clean_poly.NumberOfVertices - 1);
                                    Point_WS_plus1 = WS_clean_poly.GetPointAtParameter(i + 1);
                                }
                                else
                                {
                                    Point_WS_minus1 = WS_clean_poly.GetPointAtParameter(i - 1);
                                    Point_WS_plus1 = WS_clean_poly.GetPointAtParameter(i + 1);
                                }

                                Double ParamT_minus1 = PolyCL_cleaned.GetParameterAtPoint(PolyCL_cleaned.GetClosestPointTo(Point_WS_minus1, Vector3d.ZAxis, false));

                                if (Param_CL_rounded + 1 <= PolyCL_cleaned.NumberOfVertices - 1)
                                {

                                    if (ParamT_minus1 < Param_CL_rounded)
                                    {
                                        Point_CL_minus1 = PolyCL_cleaned.GetPointAtParameter(Param_CL_rounded - 1);
                                        Point_CL_plus1 = PolyCL_cleaned.GetPointAtParameter(Param_CL_rounded + 1);
                                    }
                                    else
                                    {
                                        Point_CL_minus1 = PolyCL_cleaned.GetPointAtParameter(Param_CL_rounded + 1);
                                        Point_CL_plus1 = PolyCL_cleaned.GetPointAtParameter(Param_CL_rounded - 1);
                                    }


                                }
                                else
                                {
                                    Point_CL_minus1 = PolyCL_cleaned.GetPointAtParameter(Param_CL_rounded);
                                    Point_CL_plus1 = PolyCL_cleaned.GetPointAtParameter(Param_CL_rounded - 0.5);
                                }


                                double Bearing1 = -100;
                                double BearingC1 = -100;
                                double Bearing2 = -100;
                                double BearingC2 = -100;


                                Bearing1 = Math.Round(Workspace_Band.Functions.GET_Bearing_rad(Point_WS.X, Point_WS.Y, Point_WS_plus1.X, Point_WS_plus1.Y), 2);
                                BearingC1 = Math.Round(Workspace_Band.Functions.GET_Bearing_rad(Point_CL_rounded.X, Point_CL_rounded.Y, Point_CL_plus1.X, Point_CL_plus1.Y), 2);
                                Bearing2 = Math.Round(Workspace_Band.Functions.GET_Bearing_rad(Point_WS.X, Point_WS.Y, Point_WS_minus1.X, Point_WS_minus1.Y), 2);
                                BearingC2 = Math.Round(Workspace_Band.Functions.GET_Bearing_rad(Point_CL_rounded.X, Point_CL_rounded.Y, Point_CL_minus1.X, Point_CL_minus1.Y), 2);

                                if (Bearing1 >= 3.14) Bearing1 = Bearing1 - 3.14;
                                if (BearingC1 >= 3.14) BearingC1 = BearingC1 - 3.14;
                                if (Bearing2 >= 3.14) Bearing2 = Bearing2 - 3.14;
                                if (BearingC2 >= 3.14) BearingC2 = BearingC2 - 3.14;



                                if (Bearing1 == BearingC1 & Bearing2 == BearingC2)
                                {

                                    Line CL_plus1 = new Line(Point_CL, Point_CL_plus1);
                                    Line CL_minus1 = new Line(Point_CL, Point_CL_minus1);
                                    double Dist_plus1 = CL_plus1.GetClosestPointTo(Point_WS, Vector3d.ZAxis, true).GetVectorTo(Point_WS).Length;
                                    double Dist_minus1 = CL_minus1.GetClosestPointTo(Point_WS, Vector3d.ZAxis, true).GetVectorTo(Point_WS).Length;

                                    Line EAS_plus1 = new Line();
                                    Line EAS_minus1 = new Line();

                                    if (Side_of_WS == "L")
                                    {

                                        Point3d Pt_E = Poly_easement_L.GetClosestPointTo(Point_WS, Vector3d.ZAxis, false);
                                        double Param_E = Poly_easement_L.GetParameterAtPoint(Pt_E);

                                        double Param_E_rounded = Math.Round(Param_E, 0);

                                        do
                                        {
                                            Param_E_rounded = Param_E_rounded + 1;
                                        } while (Param_E_rounded < Param_min_L);



                                        do
                                        {
                                            Param_E_rounded = Param_E_rounded - 1;
                                        } while (Param_E_rounded > Param_max_L);

                                        Point3d Point_E_rounded = Poly_easement_L.GetPointAtParameter(Param_E_rounded);

                                        Point3d Pt_E_minus1 = Poly_easement_L.GetPointAtParameter(Param_E_rounded - 1);
                                        Point3d Pt_E_plus1 = Poly_easement_L.GetPointAtParameter(Param_E_rounded + 1);

                                        EAS_minus1 = new Line(Point_E_rounded, Pt_E_minus1);
                                        EAS_plus1 = new Line(Point_E_rounded, Pt_E_plus1);

                                        Dist_plus1 = EAS_plus1.GetClosestPointTo(Point_WS, Vector3d.ZAxis, true).GetVectorTo(Point_WS).Length;
                                        Dist_minus1 = EAS_minus1.GetClosestPointTo(Point_WS, Vector3d.ZAxis, true).GetVectorTo(Point_WS).Length;
                                    }

                                    if (Side_of_WS == "R")
                                    {

                                        Point3d Pt_E = Poly_easement_R.GetClosestPointTo(Point_WS, Vector3d.ZAxis, false);
                                        double Param_E = Poly_easement_R.GetParameterAtPoint(Pt_E);

                                        double Param_E_rounded = Math.Round(Param_E, 0);

                                        do
                                        {
                                            Param_E_rounded = Param_E_rounded + 1;
                                        } while (Param_E_rounded < Param_min_R);



                                        do
                                        {
                                            Param_E_rounded = Param_E_rounded - 1;
                                        } while (Param_E_rounded > Param_max_R);

                                        Point3d Point_E_rounded = Poly_easement_R.GetPointAtParameter(Param_E_rounded);

                                        Point3d Pt_E_minus1 = Poly_easement_R.GetPointAtParameter(Param_E_rounded - 1);
                                        Point3d Pt_E_plus1 = Poly_easement_R.GetPointAtParameter(Param_E_rounded + 1);

                                        EAS_minus1 = new Line(Point_E_rounded, Pt_E_minus1);
                                        EAS_plus1 = new Line(Point_E_rounded, Pt_E_plus1);

                                        Dist_plus1 = EAS_plus1.GetClosestPointTo(Point_WS, Vector3d.ZAxis, true).GetVectorTo(Point_WS).Length;
                                        Dist_minus1 = EAS_minus1.GetClosestPointTo(Point_WS, Vector3d.ZAxis, true).GetVectorTo(Point_WS).Length;
                                    }

                                    if (Side_of_WS == "")
                                    {

                                        BDY_WS = "CL";
                                    }

                                    Data_Table_WS.Rows.Add();
                                    Data_Table_WS.Rows[j]["X"] = Math.Round(Point_WS.X, Rounding_no);
                                    Data_Table_WS.Rows[j]["Y"] = Math.Round(Point_WS.Y, Rounding_no);

                                    Point3d PT_CL_MS = PolyCL_MS.GetClosestPointTo(PolyCL_cleaned.GetPointAtParameter(Math.Round(Param_CL, 0)), Vector3d.ZAxis, false);
                                    Data_Table_WS.Rows[j]["STA"] = Math.Round(PolyCL_MS.GetDistAtPoint(PT_CL_MS), Rounding_no);

                                    Data_Table_WS.Rows[j]["OFFSET"] = Math.Round(Dist_minus1, Rounding_no);
                                    j = j + 1;

                                    if (Math.Round(Dist_minus1, Rounding_no) != Math.Round(Dist_plus1, Rounding_no))
                                    {

                                        Data_Table_WS.Rows.Add();
                                        Data_Table_WS.Rows[j]["X"] = Math.Round(Point_WS.X, Rounding_no);
                                        Data_Table_WS.Rows[j]["Y"] = Math.Round(Point_WS.Y, Rounding_no);

                                        Data_Table_WS.Rows[j]["STA"] = Math.Round(PolyCL_MS.GetDistAtPoint(PT_CL_MS), Rounding_no);

                                        Data_Table_WS.Rows[j]["OFFSET"] = Math.Round(Dist_plus1, Rounding_no);





                                        j = j + 1;


                                    }


                                }

                                else
                                {
                                    Point3d Point_WS_L = Poly_easement_L.GetClosestPointTo(Point_WS, Vector3d.ZAxis, false);
                                    Point3d Point_WS_R = Poly_easement_R.GetClosestPointTo(Point_WS, Vector3d.ZAxis, false);

                                    Line CL_WS = new Line(Point_WS, Point_CL);
                                    Line WS_L = new Line(Point_WS, Point_WS_L);
                                    Line WS_R = new Line(Point_WS, Point_WS_R);

                                    double WS_L_Length = WS_L.Length;
                                    double WS_R_Length = WS_R.Length;
                                    double CL_WS_Length = CL_WS.Length;

                                    Data_Table_WS.Rows.Add();
                                    Data_Table_WS.Rows[j]["X"] = Math.Round(Point_WS.X, Rounding_no);
                                    Data_Table_WS.Rows[j]["Y"] = Math.Round(Point_WS.Y, Rounding_no);
                                    Data_Table_WS.Rows[j]["STA"] = Math.Round(PolyCL_MS.GetDistAtPoint(Point_CL), Rounding_no);
                                    if (Side_of_WS == "L")
                                    {
                                        Data_Table_WS.Rows[j]["OFFSET"] = Math.Round(WS_L_Length, Rounding_no);
                                    }
                                    if (Side_of_WS == "R")
                                    {
                                        Data_Table_WS.Rows[j]["OFFSET"] = Math.Round(WS_R_Length, Rounding_no);
                                    }

                                    if (Side_of_WS == "")
                                    {
                                        Data_Table_WS.Rows[j]["OFFSET"] = Math.Round(CL_WS_Length, Rounding_no);
                                        BDY_WS = "CL";
                                    }
                                    j = j + 1;
                                }
                            }




                            for (int i = 0; i < Data_Table_WS.Rows.Count; ++i)
                            {
                                double Width1 = (double)Data_Table_WS.Rows[i]["OFFSET"];
                                double Station1 = (double)Data_Table_WS.Rows[i]["STA"];

                                if (i < WS_clean_poly.NumberOfVertices)
                                {


                                    Point3d Pt_WS_1 = WS_clean_poly.GetPointAtParameter(i);

                                    double Station2;
                                    double Width2;
                                    Point3d Pt_WS_2;


                                    if (i == Data_Table_WS.Rows.Count - 1)
                                    {
                                        Station2 = (double)Data_Table_WS.Rows[0]["STA"];
                                        Width2 = (double)Data_Table_WS.Rows[0]["OFFSET"];
                                        Pt_WS_2 = WS_clean_poly.GetPointAtParameter(0);

                                    }
                                    else
                                    {
                                        Station2 = (double)Data_Table_WS.Rows[i + 1]["STA"];
                                        Width2 = (double)Data_Table_WS.Rows[i + 1]["OFFSET"];
                                        if (i + 1 == WS_clean_poly.NumberOfVertices)
                                        {
                                            Pt_WS_2 = WS_clean_poly.GetPointAtParameter(0);
                                        }
                                        else
                                        {
                                            Pt_WS_2 = WS_clean_poly.GetPointAtParameter(i + 1);
                                        }

                                    }




                                    Point3d Pt_CL_1 = PolyCL_cleaned.GetClosestPointTo(Pt_WS_1, Vector3d.ZAxis, false);
                                    double ParamCL1 = PolyCL_cleaned.GetParameterAtPoint(Pt_CL_1);


                                    Point3d Pt_CL_2 = PolyCL_cleaned.GetClosestPointTo(Pt_WS_2, Vector3d.ZAxis, false);
                                    double ParamCL2 = PolyCL_cleaned.GetParameterAtPoint(Pt_CL_2);


                                    if (Width1 != Width2 & Width1 == 0)
                                    {
                                        if (Math.Abs(Math.Floor(ParamCL1) - Math.Floor(ParamCL2)) > 0)
                                        {


                                            Point3d Pt_CL_1_round = new Point3d();
                                            Point3d Pt_CL_2_round = new Point3d();
                                            double Bearing1;
                                            double BearingC1;



                                            if (ParamCL1 > ParamCL2)
                                            {
                                                Pt_CL_1_round = PolyCL_cleaned.GetPointAtParameter(Math.Floor(ParamCL1));
                                                Pt_CL_2_round = PolyCL_cleaned.GetPointAtParameter(Math.Floor(ParamCL1) - 1);

                                            }
                                            if (ParamCL1 < ParamCL2)
                                            {
                                                Pt_CL_1_round = PolyCL_cleaned.GetPointAtParameter(Math.Ceiling(ParamCL1));
                                                Pt_CL_2_round = PolyCL_cleaned.GetPointAtParameter(Math.Ceiling(ParamCL1) + 1);

                                            }

                                            Bearing1 = Math.Round(Workspace_Band.Functions.GET_Bearing_rad(Pt_WS_1.X, Pt_WS_1.Y, Pt_WS_2.X, Pt_WS_2.Y), 2);
                                            BearingC1 = Math.Round(Workspace_Band.Functions.GET_Bearing_rad(Pt_CL_1_round.X, Pt_CL_1_round.Y, Pt_CL_2_round.X, Pt_CL_2_round.Y), 2);

                                            if (Bearing1 >= 3.14) Bearing1 = Bearing1 - 3.14;
                                            if (BearingC1 >= 3.14) BearingC1 = BearingC1 - 3.14;

                                            if (Bearing1 == BearingC1)
                                            {
                                                Extra_offset = Width2;
                                                Extra_station = Math.Round(PolyCL_MS.GetDistAtPoint(PolyCL_MS.GetClosestPointTo(Pt_CL_1_round, Vector3d.ZAxis, false)), Rounding_no);
                                                Extra_index = i + 1;
                                            }

                                        }

                                    }

                                    if (Width1 != Width2 & Width2 == 0)
                                    {
                                        if (Math.Abs(Math.Floor(ParamCL1) - Math.Floor(ParamCL2)) > 0)
                                        {


                                            Point3d Pt_CL_1_round = new Point3d();
                                            Point3d Pt_CL_2_round = new Point3d();
                                            double Bearing1;
                                            double BearingC1;



                                            if (ParamCL1 > ParamCL2)
                                            {
                                                Pt_CL_1_round = PolyCL_cleaned.GetPointAtParameter(Math.Ceiling(ParamCL2));
                                                Pt_CL_2_round = PolyCL_cleaned.GetPointAtParameter(Math.Ceiling(ParamCL2) + 1);

                                            }
                                            if (ParamCL1 < ParamCL2)
                                            {
                                                Pt_CL_1_round = PolyCL_cleaned.GetPointAtParameter(Math.Floor(ParamCL2));
                                                Pt_CL_2_round = PolyCL_cleaned.GetPointAtParameter(Math.Floor(ParamCL2) - 1);

                                            }

                                            Bearing1 = Math.Round(Workspace_Band.Functions.GET_Bearing_rad(Pt_WS_1.X, Pt_WS_1.Y, Pt_WS_2.X, Pt_WS_2.Y), 2);
                                            BearingC1 = Math.Round(Workspace_Band.Functions.GET_Bearing_rad(Pt_CL_2_round.X, Pt_CL_2_round.Y, Pt_CL_1_round.X, Pt_CL_1_round.Y), 2);

                                            if (Bearing1 >= 3.14) Bearing1 = Bearing1 - 3.14;
                                            if (BearingC1 >= 3.14) BearingC1 = BearingC1 - 3.14;

                                            if (Bearing1 == BearingC1)
                                            {
                                                Extra_offset = Width1;
                                                Extra_station = Math.Round(PolyCL_MS.GetDistAtPoint(PolyCL_MS.GetClosestPointTo(Pt_CL_1_round, Vector3d.ZAxis, false)), Rounding_no);
                                                Extra_index = i + 1;
                                            }

                                        }

                                    }

                                }
                            }

                            if (Extra_index != -1)
                            {

                                System.Data.DataRow Row1;
                                Row1 = Data_Table_WS.NewRow();
                                Row1["OFFSET"] = Extra_offset;
                                Row1["STA"] = Extra_station;
                                Data_Table_WS.Rows.InsertAt(Row1, Extra_index);
                            }




                            string Table1 = textBox_COMPILED.Text;

                            string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + textBox_access_database_location.Text + ";";
                            OleDbConnection cnn;


                            try
                            {
                                cnn = new OleDbConnection(ConnectionString);
                                cnn.Open();

                                OleDbCommand cmd = new OleDbCommand();
                                cmd.CommandType = CommandType.Text;
                                cmd.Connection = cnn;



                                for (int i = 0; i < Data_Table_WS.Rows.Count; i = i + 1)
                                {
                                    if (i + 1 < Data_Table_WS.Rows.Count)
                                    {
                                        double STA1 = (double)Data_Table_WS.Rows[i]["STA"];
                                        double STA2 = (double)Data_Table_WS.Rows[i + 1]["STA"];
                                        double OFFSET1 = (double)Data_Table_WS.Rows[i]["OFFSET"];
                                        double OFFSET2 = (double)Data_Table_WS.Rows[i + 1]["OFFSET"];

                                        cmd.CommandText = "INSERT INTO " + Table1 + "(TYPE,SIDE,BDY,START_STA,END_STA,START_WIDTH,END_WIDTH,WS_NO) VALUES " +
                                                                                    "('TWS','" + Side_of_WS + "','" + BDY_WS + "'," + STA1 + "," + STA2 + "," + OFFSET1 + "," + OFFSET2 + "," + WS_no + ")";
                                        cmd.ExecuteNonQuery();

                                    }

                                    if (i == Data_Table_WS.Rows.Count - 1)
                                    {
                                        double STA1 = (double)Data_Table_WS.Rows[i]["STA"];
                                        double STA2 = (double)Data_Table_WS.Rows[0]["STA"];
                                        double OFFSET1 = (double)Data_Table_WS.Rows[i]["OFFSET"];
                                        double OFFSET2 = (double)Data_Table_WS.Rows[0]["OFFSET"];

                                        cmd.CommandText = "INSERT INTO " + Table1 + "(TYPE,SIDE,BDY,START_STA,END_STA,START_WIDTH,END_WIDTH,WS_NO) VALUES " +
                                                                                    "('TWS','" + Side_of_WS + "','" + BDY_WS + "'," + STA1 + "," + STA2 + "," + OFFSET1 + "," + OFFSET2 + "," + WS_no + ")";
                                        cmd.ExecuteNonQuery();

                                    }
                                }
                                WS_no = WS_no + 1;
                                cnn.Close();

                            }


                            catch (OleDbException ex)
                            {
                                MessageBox.Show(ex.Message);
                                Freeze_operations = false;
                            }

                        }
                    }
                }



            }


        }


        private void Button_draw_from_compiled_database(object sender, EventArgs e)
        {
            if (Freeze_operations == false)
            {
                Freeze_operations = true;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                try
                {


                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                    Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                    PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the insertion point");
                    PP1.AllowNone = false;
                    Point_res1 = ThisDrawing.Editor.GetPoint(PP1);

                    if (Point_res1.Status != PromptStatus.OK)
                    {

                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                        return;
                    }


                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                            Read_max_STA();
                            if (Sta_max != 0)
                            {
                                Point2d Start_pt = new Point2d(Point_res1.Value.X, Point_res1.Value.Y);
                                Point2d End_pt = new Point2d(Start_pt.X + Sta_max, Point_res1.Value.Y);
                                Polyline Poly_CL_Straight = new Polyline();
                                Poly_CL_Straight.AddVertexAt(0, Start_pt, 0, 0, 0);
                                if (radioButton_left_right.Checked == true)
                                {
                                    Poly_CL_Straight.AddVertexAt(1, End_pt, 0, 0, 0);
                                }
                                else
                                {
                                    Poly_CL_Straight.AddVertexAt(0, End_pt, 0, 0, 0);
                                }
                                BTrecord.AppendEntity(Poly_CL_Straight);
                                Trans1.AddNewlyCreatedDBObject(Poly_CL_Straight, true);
                                int UP_DOWN = -1;
                                if (radioButton_left_UP.Checked == true)
                                {
                                    UP_DOWN = 1;
                                }

                                Populate_Compiled_DATA_TABLE_with_easement_info("EPE", "L", 0);
                                Polyline Poly_easment_left = new Polyline();
                                DRAW_easement_from_compiled(Poly_CL_Straight, Poly_easment_left, Compiled_DATA_TABLE, UP_DOWN);
                                if (Poly_easment_left.NumberOfVertices > 0)
                                {
                                    BTrecord.AppendEntity(Poly_easment_left);
                                    Trans1.AddNewlyCreatedDBObject(Poly_easment_left, true);
                                }

                                Populate_Compiled_DATA_TABLE_with_easement_info("EPE", "R", 0);
                                Polyline Poly_easment_right = new Polyline();
                                DRAW_easement_from_compiled(Poly_CL_Straight, Poly_easment_right, Compiled_DATA_TABLE, -UP_DOWN);
                                if (Poly_easment_right.NumberOfVertices > 0)
                                {
                                    BTrecord.AppendEntity(Poly_easment_right);
                                    Trans1.AddNewlyCreatedDBObject(Poly_easment_right, true);
                                }

                                number_of_workspaces = 0;
                                Read_nr_max_workspace();
                                for (int i = 1; i <= number_of_workspaces; i = i + 1)
                                {
                                    Populate_Compiled_DATA_TABLE_with_workspace_info("TWS", i);
                                    Polyline Poly_WS1 = new Polyline();
                                    DRAW_workspace_from_compiled(Poly_CL_Straight, Poly_easment_right, Poly_easment_left, Poly_WS1, Compiled_DATA_TABLE);

                                    if (Poly_WS1.NumberOfVertices > 0)
                                    {
                                        Poly_WS1.Closed = true;
                                        BTrecord.AppendEntity(Poly_WS1);
                                        Trans1.AddNewlyCreatedDBObject(Poly_WS1, true);
                                    }
                                }
                                Trans1.Commit();
                            }
                        }
                    }

                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                }

                catch (System.Exception ex)
                {
                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\n" + "Command:");
                    MessageBox.Show(ex.Message);
                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                }

                Freeze_operations = false;
            }
        }

        private void Read_nr_max_workspace()
        {
            try
            {
                string Table1 = textBox_COMPILED.Text;
                string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + textBox_access_database_location.Text + ";";
                DataSet COMPILED_DATASET = new DataSet();
                OleDbConnection cnn;
                try
                {
                    cnn = new OleDbConnection(ConnectionString);
                    cnn.Open();
                    string query1 = "SELECT MAX(WS_NO) FROM " + Table1;
                    OleDbCommand cmd1 = new OleDbCommand(query1, cnn);
                    System.Data.OleDb.OleDbDataReader Reader_max = cmd1.ExecuteReader();

                    while (Reader_max.Read())
                    {
                        number_of_workspaces = Convert.ToInt32(Reader_max[0]);
                    }
                    cnn.Close();
                }
                catch (OleDbException ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void Read_max_STA()
        {
            try
            {
                string Table1 = textBox_COMPILED.Text;

                string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + textBox_access_database_location.Text + ";";
                DataSet COMPILED_DATASET = new DataSet();
                OleDbConnection cnn;
                try
                {
                    cnn = new OleDbConnection(ConnectionString);
                    cnn.Open();

                    double val1 = 0;
                    double val2 = 0;

                    string query1 = "SELECT MAX(END_STA) FROM " + Table1;
                    string query2 = "SELECT MAX(START_STA) FROM " + Table1;

                    OleDbCommand cmd1 = new OleDbCommand(query1, cnn);
                    System.Data.OleDb.OleDbDataReader Reader_max1 = cmd1.ExecuteReader();

                    while (Reader_max1.Read())
                    {
                        val1 = Convert.ToDouble(Reader_max1[0]);
                    }

                    OleDbCommand cmd2 = new OleDbCommand(query2, cnn);
                    System.Data.OleDb.OleDbDataReader Reader_max2 = cmd2.ExecuteReader();
                    while (Reader_max2.Read())
                    {
                        val2 = Convert.ToDouble(Reader_max2[0]);
                    }
                    cnn.Close();

                    if (val1 > val2)
                    {
                        Sta_max = val1;
                    }
                    else
                    {
                        Sta_max = val2;
                    }

                }
                catch (OleDbException ex)
                {
                    MessageBox.Show(ex.Message);
                }


            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

    }
}

