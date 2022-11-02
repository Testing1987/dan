using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;

namespace Alignment_mdi
{
    public partial class Lgen_label_page : Form
    {
        List<string> scales;

        object old_osnap;
        double txt_rot = 0;
        string label_layer = "_lgen";


        string current_labelstyle = "";
        string suff_layer = "";

        string primary_labelstyle = "";


        string table_bl_name = "";
        double scale1;
        int idx_alias;

        bool is_bdy = false;
        bool align_to_fea = false;

        string layer_prev = "";

        int i_desc = 1;
        int i_lay = 2;
        int i_bdy = 3;
        int i_align = 4;
        int i_prim = 5;
        int i_sec = 6;
        int i_ter = 7;
        int i_tst = 8;
        int i_tf = 9;
        int i_tw = 10;
        int i_to = 11;
        int i_th = 12;
        int i_tu = 13;
        int i_mst = 14;
        int i_ma = 15;
        int i_mg = 16;
        int i_md = 17;
        int i_nh = 18;
        int i_od = 19;
        int i_c = 20;
        int i_cl = 21;
        int i_clp = 22;
        int i_p1 = 23;
        int i_od1 = 24;
        int i_s1 = 25;
        int i_p2 = 26;
        int i_od2 = 27;
        int i_s2 = 28;
        int i_p3 = 29;
        int i_od3 = 30;
        int i_s3 = 31;
        int i_p4 = 32;
        int i_od4 = 33;
        int i_s4 = 34;
        int i_p5 = 35;
        int i_od5 = 36;
        int i_s5 = 37;
        int i_p6 = 38;
        int i_od6 = 39;
        int i_s6 = 40;
        int i_p7 = 41;
        int i_od7 = 42;
        int i_s7 = 43;
        int i_p8 = 44;
        int i_od8 = 45;
        int i_s8 = 46;
        int i_p9 = 47;
        int i_od9 = 48;
        int i_s9 = 49;
        int i_p10 = 50;
        int i_od10 = 51;
        int i_s10 = 52;
        int i_bn = 53;
        int i_bac = 54;
        int i_ba1 = 55;
        int i_ba2 = 56;
        int i_ba3 = 57;
        int i_ba4 = 58;
        int i_ba5 = 59;
        int i_ba6 = 60;
        int i_ba7 = 61;
        int i_ba8 = 62;
        int i_ba9 = 63;
        int i_ba10 = 64;
        int i_d_sty = 65;
        int i_d_arr = 66;
        int i_d_suff = 67;
        int i_d_decno = 68;
        int i_d_round_closest = 69;
        int i_force_dim = 70;
        int i_bm = 71;
        int i_txtf = 72;
        int i_round = 73;

        double textH = 0.08;
        string table_txt_stylename = "lgen_textstyle";
        string table_txt_fontname = "arial.ttf";
        double table_txt_oblique = 15;
        double table_txt_width = 0.8;
        bool txt_underline = false;
        bool mtxt_background_frame = false;
        bool mleader_text_frame = false;




        System.Data.DataTable dt_alias = null;
        List<string> lista_layere = null;
        string Fisier_layer_alias = "";

        bool ExcelVisible = false;
        int Start_row_layer_alias = 8;

        bool is_paperspace = false;

        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();



            lista_butoane.Add(Button_create_label);
            lista_butoane.Add(button_mleader_ne);
            lista_butoane.Add(Button_SetRotation);

            lista_butoane.Add(Combobox_scales);


            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();


            lista_butoane.Add(Button_create_label);
            lista_butoane.Add(button_mleader_ne);
            lista_butoane.Add(Button_SetRotation);

            lista_butoane.Add(Combobox_scales);




            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }
        public Lgen_label_page()
        {
            InitializeComponent();

        }


        private void button_select_table_Click(object sender, EventArgs e)
        {

            using (OpenFileDialog fbd = new OpenFileDialog())
            {
                fbd.Multiselect = false;
                fbd.Filter = "Excel files (*.xlsx)|*.xlsx";

                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    Fisier_layer_alias = fbd.FileName;
                    _SGEN_mainform.tpage_Main.LGEN_label_excel_file_green(Fisier_layer_alias);
                    dt_alias = Load_existing_Lgen_layer_alias_from_excel(Fisier_layer_alias);
                }
                else
                {
                    _SGEN_mainform.tpage_Main.LGEN_label_excel_file_red();
                    Fisier_layer_alias = "";
                }
            }
        }

        private void button_open_layer_alias_Click(object sender, EventArgs e)
        {
            if (System.IO.File.Exists(Fisier_layer_alias) == true)
            {
                Microsoft.Office.Interop.Excel.Application Excel1 = new Microsoft.Office.Interop.Excel.Application();
                if (Excel1 == null)
                {
                    MessageBox.Show("PROBLEM WITH EXCEL!");
                    return;
                }
                Excel1.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(Fisier_layer_alias);

            }
        }

        private void Lgen_label_page_load(object sender, EventArgs e)
        {
            scales = new List<string>();
            scales.Add("1' = 1\"");
            scales.Add("1' = 10\"");
            scales.Add("1' = 20\"");
            scales.Add("1' = 30\"");
            scales.Add("1' = 40\"");
            scales.Add("1' = 50\"");
            scales.Add("1' = 60\"");
            scales.Add("1' = 100\"");
            scales.Add("1' = 200\"");
            scales.Add("1' = 300\"");
            scales.Add("1' = 400\"");
            scales.Add("1' = 500\"");
            scales.Add("1' = 600\"");
            scales.Add("1' = 1000\"");
            Combobox_scales.DataSource = scales;
            Combobox_scales.SelectedIndex = 0;


            char deg_symbol = Convert.ToChar(176);
            Label_Rotation.Text = "Rotation = 0" + deg_symbol;



        }


        [System.Runtime.InteropServices.DllImport("mpr.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode, SetLastError = true)]
        public static extern int WNetGetConnection(
                    [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPTStr)] string localName,
                    [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPTStr)] StringBuilder remoteName,
                   ref int length);

        public static string GetUNCPath(string originalPath)
        {
            StringBuilder sb = new StringBuilder(512);
            int size = sb.Capacity;
            // look for the {LETTER}: combination ...
            if (originalPath.Length > 2 && originalPath[1] == ':')
            {
                // don't use char.IsLetter here - as that can be misleading
                // the only valid drive letters are a-z && A-Z.
                char c = originalPath[0];
                if ((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z'))
                {
                    int error = WNetGetConnection(originalPath.Substring(0, 2), sb, ref size);
                    if (error == 0)
                    {
                        System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(originalPath);
                        string path = System.IO.Path.GetFullPath(originalPath).Substring(System.IO.Path.GetPathRoot(originalPath).Length);
                        return System.IO.Path.Combine(sb.ToString().TrimEnd(), path);
                    }
                }
            }
            return originalPath;
        }

        private void Button_create_label_Click(object sender, EventArgs e)
        {
            if (dt_alias == null || dt_alias.Rows.Count == 0)
            {
                MessageBox.Show("no layer alias loaded\r\noperation aborted");
                return;
            }

            //End = 1,
            //Middle = 2,
            //Center = 4,
            //Node = 8,
            //Quadrant = 16,
            //Intersection = 32,
            //Insertion = 64,
            //Perpendicular = 128,
            //Tangent = 256,
            //Near = 512,
            // Quick = 1024,
            //ApparentIntersection = 2048,
            //Immediate = 65536,
            //AllowTangent = 131072,
            // DisablePerpendicular = 262144,
            //RelativeCartesian = 524288,
            //RelativePolar = 1048576,
            //NoneOverride = 2097152,  



            old_osnap = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE");
            object new_osnap = 512;


            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = Trans1.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        DimStyleTable Dim_style_table = Trans1.GetObject(ThisDrawing.Database.DimStyleTableId, OpenMode.ForWrite) as DimStyleTable;

                        Autodesk.Gis.Map.Project.ProjectModel project1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject;
                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = project1.ODTables;
                        Autodesk.Gis.Map.Project.DrawingSet drawingset1 = project1.DrawingSet;
                        Autodesk.Gis.Map.Aliases drives_aliasses = Autodesk.Gis.Map.HostMapApplicationServices.Application.Aliases;





                        double deltax = Functions.GET_deltaX_rad();
                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", new_osnap);

                        Button_SetRotation.Enabled = false;
                        button_mleader_ne.Enabled = false;


                    repeat1:


                        this.Refresh();
                        is_bdy = false;

                        string gen_description = "EMPTY";
                        string od_description = "";
                        string od_val1 = "";
                        string od_val2 = "";
                        string od_val3 = "";
                        string od_val4 = "";
                        string od_val5 = "";
                        string od_val6 = "";
                        string od_val7 = "";
                        string od_val8 = "";
                        string od_val9 = "";
                        string od_val10 = "";



                        align_to_fea = false;

                        current_labelstyle = "mt";

                        string table_mstyle_name = "lgen_mleaderstyle";
                        double table_mleader_arrow = 0.08;
                        double table_mleader_doglentgh = 0.08;
                        double table_mleader_gap = 0.08;
                        double MleadertextH = 0.08;

                        double table_dim_arrow_size = 0.08;
                        int table_dimstyle_dec_no = 0;
                        string table_dimstyle_name = "lgen_dimstyle";
                        string table_dimstyle_suffix = "' PERMANENT EASEMENT";
                        int table_dimstyle_round_closest = 5;
                        bool table_dimstyle_force_line = false;

                        table_bl_name = "";
                        string table_atr_c = "ATR1";
                        string table_atr_1 = "ATR1";
                        string table_atr_2 = "ATR1";
                        string table_atr_3 = "ATR1";
                        string table_atr_4 = "ATR1";
                        string table_atr_5 = "ATR1";
                        string table_atr_6 = "ATR1";
                        string table_atr_7 = "ATR1";
                        string table_atr_8 = "ATR1";
                        string table_atr_9 = "ATR1";
                        string table_atr_10 = "ATR1";

                        bool is_object_data = false;
                        bool is_capitalized = false;

                        string p1 = "";
                        string f1 = "";
                        string s1 = "";
                        string p2 = "";
                        string s2 = "";
                        string f2 = "";
                        string p3 = "";
                        string f3 = "";
                        string s3 = "";
                        string p4 = "";
                        string f4 = "";
                        string s4 = "";
                        string p5 = "";
                        string f5 = "";
                        string s5 = "";
                        string p6 = "";
                        string f6 = "";
                        string s6 = "";
                        string p7 = "";
                        string f7 = "";
                        string s7 = "";
                        string p8 = "";
                        string f8 = "";
                        string s8 = "";
                        string p9 = "";
                        string f9 = "";
                        string s9 = "";
                        string p10 = "";
                        string f10 = "";
                        string s10 = "";

                        string cb_txt = Combobox_scales.Text;
                        cb_txt = cb_txt.Replace("1' = ", "");

                        scale1 = Convert.ToDouble(cb_txt.Replace("\"", ""));
                        ObjectId textstyle_id = ObjectId.Null;
                        ObjectId mleader_id = ObjectId.Null;

                        double rotation1 = txt_rot;

                        bool is_contour = false;
                        int contour_rounding = 0;

                        string xref = "";

                    select0:
                        Autodesk.AutoCAD.EditorInput.PromptNestedEntityResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptNestedEntityOptions Prompt_sel_ent;
                        Prompt_sel_ent = new Autodesk.AutoCAD.EditorInput.PromptNestedEntityOptions("\nSelect the feature:");

                        Prompt_sel_ent.AllowNone = true;

                        Prompt_sel_ent.Keywords.Add("Set rotation");
                        //Prompt_sel_ent.Keywords.Add("Toggle pipeline label style");
                        Rezultat1 = ThisDrawing.Editor.GetNestedEntity(Prompt_sel_ent);

                        if (Rezultat1.Status != PromptStatus.Keyword && Rezultat1.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", old_osnap);
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            Trans1.Commit();
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            Button_SetRotation.Enabled = true;
                            button_mleader_ne.Enabled = true;
                            radioButton1.Text = "1";
                            radioButton2.Text = "2";
                            radioButton3.Text = "3";
                            radioButton1.Visible = false;
                            radioButton2.Visible = false;
                            radioButton3.Visible = false;
                            panel_type.Visible = false;

                            if (xref != "" && drawingset1.AllDrawingsCount > 0)
                            {
                                try
                                {
                                    drawingset1.DetachDrawing(xref);
                                }
                                catch (Autodesk.Gis.Map.MapException mapex)
                                {

                                    MessageBox.Show("the file\r\n" + xref + "\r\n\r\nis locked by someone else\r\n\r\n" + mapex.Message);
                                }
                            }
                            return;
                        }
                        if (Rezultat1.Status == PromptStatus.Keyword)
                        {

                            if (Rezultat1.StringResult.ToLower() == "set")
                            {
                                rotation1 = pick_rot_2pt();
                            }

                            goto select0;
                        }

                        Curve curve1 = null;

                        DBObject dbobj1 = Trans1.GetObject(Rezultat1.ObjectId, OpenMode.ForRead) as DBObject;
                        ObjectId[] ids = Rezultat1.GetContainers();

                        if (ids.Length > 0)
                        {
                            #region map attach
                            BlockReference blk1 = Trans1.GetObject(ids[0], OpenMode.ForRead) as BlockReference;
                            if (blk1 != null)
                            {
                                BlockTableRecord btr1 = Trans1.GetObject(blk1.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                                if (btr1 != null)
                                {
                                    if (btr1.IsFromExternalReference == true)
                                    {
                                        xref = btr1.GetXrefDatabase(false).Filename;
                                        string drive1 = System.IO.Path.GetPathRoot(xref).Substring(0, 1);
                                        string drive2 = System.IO.Path.GetPathRoot(xref);

                                        string network_drive_path = GetUNCPath(drive2);

                                        bool exista_alias = false;
                                        for (int i = 0; i < drives_aliasses.AliasesCount; ++i)
                                        {
                                            if (drives_aliasses[i].Name == drive1)
                                            {
                                                exista_alias = true;
                                                if (drives_aliasses[i].Path != network_drive_path)
                                                {
                                                    drives_aliasses[i].Path = network_drive_path;

                                                }
                                            }
                                        }

                                        if (exista_alias == false)
                                        {
                                            drives_aliasses.AddAlias(drive1, network_drive_path);
                                        }

                                        try
                                        {
                                            drawingset1.AttachDrawing(xref);
                                            if (dbobj1 != null)
                                            {
                                                if (dbobj1 is Curve)
                                                {
                                                    curve1 = dbobj1 as Curve;
                                                }
                                            }
                                        }
                                        catch (Autodesk.Gis.Map.MapException mapex)
                                        {

                                            MessageBox.Show("the file\r\n" + xref + "\r\n\r\nis locked by someone else\r\n\r\n" + mapex.Message);
                                        }
                                    }
                                }
                            }
                            #endregion

                        }
                        else
                        {
                            curve1 = Trans1.GetObject(Rezultat1.ObjectId, OpenMode.ForRead) as Curve;
                        }



                        if (curve1 != null)
                        {
                            string layer_curba = curve1.Layer;
                            if (layer_curba.Contains("|") == true)
                            {
                                int pos1 = layer_curba.IndexOf("|");
                                layer_curba = layer_curba.Substring(pos1 + 1, layer_curba.Length - pos1 - 1);
                            }


                            if (lista_layere.Contains(layer_curba) == true)
                            {
                                if (layer_prev != layer_curba)
                                {
                                    radioButton1.Text = "1";
                                    radioButton2.Text = "2";
                                    radioButton3.Text = "3";
                                    radioButton1.Visible = false;
                                    radioButton2.Visible = false;
                                    radioButton3.Visible = false;
                                    panel_type.Visible = false;
                                    panel_bdy.Visible = false;
                                    this.Refresh();
                                    layer_prev = layer_curba;
                                }

                                idx_alias = lista_layere.IndexOf(layer_curba);


                                panel_type.Visible = true;
                                radioButton1.Visible = true;
                                radioButton1.Checked = true;



                                #region descriptions and others

                                string layer_name_specified_in_table = "";

                                string este_bdy = "Bdy:No";
                                string algn = "Align:No";

                                if (dt_alias.Rows[idx_alias][i_desc] != DBNull.Value)
                                {
                                    gen_description = Convert.ToString(dt_alias.Rows[idx_alias][i_desc]);
                                }

                                if (dt_alias.Rows[idx_alias][i_lay] != DBNull.Value)
                                {
                                    layer_name_specified_in_table = Convert.ToString(dt_alias.Rows[idx_alias][i_lay]);
                                }

                                if (layer_name_specified_in_table.Replace(" ", "").Length == 0)
                                {
                                    layer_name_specified_in_table = "";
                                }



                                if (dt_alias.Rows[idx_alias][i_bdy] != DBNull.Value)
                                {
                                    string bdy = Convert.ToString(dt_alias.Rows[idx_alias][i_bdy]);

                                    if (bdy.ToLower() == "yes" || bdy.ToLower() == "true")
                                    {
                                        is_bdy = true;
                                        este_bdy = "Bdy:Yes";

                                    }
                                }

                                if (dt_alias.Rows[idx_alias][i_align] != DBNull.Value)
                                {
                                    string align1 = Convert.ToString(dt_alias.Rows[idx_alias][i_align]);

                                    if (align1.ToLower() == "yes" || align1.ToLower() == "true")
                                    {
                                        align_to_fea = true;
                                        algn = "Align:Yes";
                                    }
                                }


                                primary_labelstyle = "";

                                if (dt_alias.Rows[idx_alias][i_prim] != DBNull.Value)
                                {
                                    string label_val = Convert.ToString(dt_alias.Rows[idx_alias][i_prim]);
                                    if (label_val.ToLower() == "mtext")
                                    {
                                        primary_labelstyle = "mt";
                                        radioButton1.Visible = true;
                                        radioButton1.Text = "Mtext";
                                    }
                                    else if (label_val.ToLower() == "block")
                                    {
                                        primary_labelstyle = "bl";
                                        radioButton1.Visible = true;
                                        radioButton1.Text = "Block";
                                    }
                                    else if (label_val.ToLower() == "mleader")
                                    {
                                        primary_labelstyle = "ml";
                                        radioButton1.Visible = true;
                                        radioButton1.Text = "MLeader";
                                        if (is_bdy == true)
                                        {
                                            panel_bdy.Visible = true;
                                            radioButton4.Checked = true;
                                        }
                                    }
                                    else if (label_val.ToLower() == "dimension")
                                    {
                                        primary_labelstyle = "dim";
                                        radioButton1.Visible = true;
                                        radioButton1.Text = "Dimension";
                                    }
                                }

                                if (primary_labelstyle == "")
                                {
                                    MessageBox.Show("the layer " + layer_curba + " does not have specified the label type\r\nmtext, mleader ,block or dimension\r\noperation aborted");
                                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", old_osnap);
                                    this.MdiParent.WindowState = FormWindowState.Normal;
                                    Trans1.Commit();
                                    Editor1.SetImpliedSelection(Empty_array);
                                    Editor1.WriteMessage("\nCommand:");
                                    set_enable_true();
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    Button_SetRotation.Enabled = true;
                                    button_mleader_ne.Enabled = true;
                                    radioButton1.Text = "1";
                                    radioButton2.Text = "2";
                                    radioButton3.Text = "3";
                                    radioButton1.Visible = false;
                                    radioButton2.Visible = false;
                                    radioButton3.Visible = false;
                                    panel_type.Visible = false;
                                    panel_bdy.Visible = false;
                                    if (xref != "" && drawingset1.AllDrawingsCount > 0)
                                    {
                                        try
                                        {
                                            drawingset1.DetachDrawing(xref);
                                        }
                                        catch (Autodesk.Gis.Map.MapException mapex)
                                        {

                                            MessageBox.Show("the file\r\n" + xref + "\r\n\r\nis locked by someone else\r\n\r\n" + mapex.Message);
                                        }
                                    }
                                    return;
                                }

                                if (dt_alias.Rows[idx_alias][i_sec] != DBNull.Value)
                                {
                                    string label_val = Convert.ToString(dt_alias.Rows[idx_alias][i_sec]);
                                    if (label_val.ToLower() == "mtext")
                                    {

                                        radioButton2.Visible = true;
                                        radioButton2.Text = "Mtext";
                                    }
                                    else if (label_val.ToLower() == "block")
                                    {

                                        radioButton2.Visible = true;
                                        radioButton2.Text = "Block";
                                    }
                                    else if (label_val.ToLower() == "mleader")
                                    {

                                        radioButton2.Visible = true;
                                        radioButton2.Text = "MLeader";
                                    }
                                    else if (label_val.ToLower() == "dimension")
                                    {
                                        radioButton2.Visible = true;
                                        radioButton2.Text = "Dimension";
                                    }

                                }

                                if (dt_alias.Rows[idx_alias][i_ter] != DBNull.Value)
                                {
                                    string label_val = Convert.ToString(dt_alias.Rows[idx_alias][i_ter]);
                                    if (label_val.ToLower() == "mtext")
                                    {

                                        radioButton3.Visible = true;
                                        radioButton3.Text = "Mtext";
                                    }
                                    else if (label_val.ToLower() == "block")
                                    {

                                        radioButton3.Visible = true;
                                        radioButton3.Text = "Block";
                                    }
                                    else if (label_val.ToLower() == "mleader")
                                    {

                                        radioButton3.Visible = true;
                                        radioButton3.Text = "MLeader";
                                    }
                                    else if (label_val.ToLower() == "dimension")
                                    {
                                        radioButton3.Visible = true;
                                        radioButton3.Text = "Dimension";
                                    }
                                }
                                #endregion

                                #region label type
                                suff_layer = "";


                                if (primary_labelstyle == "mt")
                                {
                                    current_labelstyle = "mt";
                                    suff_layer = "_Mtext";


                                }
                                else if (primary_labelstyle == "ml")
                                {
                                    current_labelstyle = "ml";
                                    suff_layer = "_Mleader";

                                }
                                else if (primary_labelstyle == "bl")
                                {
                                    current_labelstyle = "bl";
                                    suff_layer = "_Block";

                                }
                                else if (primary_labelstyle == "dim")
                                {
                                    current_labelstyle = "dim";
                                    suff_layer = "_Dim";

                                }

                                this.Refresh();


                                #endregion

                                #region load mtext params
                                if (dt_alias.Rows[idx_alias][i_tst] != DBNull.Value)
                                {
                                    table_txt_stylename = Convert.ToString(dt_alias.Rows[idx_alias][i_tst]);
                                }
                                if (dt_alias.Rows[idx_alias][i_tf] != DBNull.Value)
                                {
                                    table_txt_fontname = Convert.ToString(dt_alias.Rows[idx_alias][i_tf]);
                                }
                                if (dt_alias.Rows[idx_alias][i_tw] != DBNull.Value)
                                {
                                    string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_tw]);
                                    if (Functions.IsNumeric(val1) == true)
                                    {
                                        table_txt_width = Convert.ToDouble(val1);
                                    }
                                }

                                if (dt_alias.Rows[idx_alias][i_to] != DBNull.Value)
                                {
                                    string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_to]);
                                    if (Functions.IsNumeric(val1) == true)
                                    {
                                        table_txt_oblique = Convert.ToDouble(val1);
                                    }
                                }

                                if (dt_alias.Rows[idx_alias][i_th] != DBNull.Value)
                                {
                                    string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_th]);
                                    if (Functions.IsNumeric(val1) == true)
                                    {
                                        textH = Convert.ToDouble(val1);
                                    }
                                }

                                if (dt_alias.Rows[idx_alias][i_tu] != DBNull.Value)
                                {
                                    string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_tu]);
                                    if (val1.ToLower() == "yes" || val1.ToLower() == "true")
                                    {
                                        txt_underline = true;
                                    }
                                }

                                if (dt_alias.Rows[idx_alias][i_bm] != DBNull.Value)
                                {
                                    string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_bm]);
                                    if (val1.ToLower() == "yes" || val1.ToLower() == "true")
                                    {
                                        mtxt_background_frame = true;
                                    }
                                }



                                #endregion

                                #region load mleader params

                                if (dt_alias.Rows[idx_alias][i_mst] != DBNull.Value)
                                {
                                    table_mstyle_name = Convert.ToString(dt_alias.Rows[idx_alias][i_mst]) + "_" + scale1.ToString();
                                }

                                if (dt_alias.Rows[idx_alias][i_ma] != DBNull.Value)
                                {
                                    string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_ma]);
                                    if (Functions.IsNumeric(val1) == true)
                                    {
                                        table_mleader_arrow = Convert.ToDouble(val1);
                                    }
                                }

                                if (dt_alias.Rows[idx_alias][i_mg] != DBNull.Value)
                                {
                                    string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_mg]);
                                    if (Functions.IsNumeric(val1) == true)
                                    {
                                        table_mleader_gap = Convert.ToDouble(val1);
                                    }
                                }

                                if (dt_alias.Rows[idx_alias][i_md] != DBNull.Value)
                                {
                                    string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_md]);
                                    if (Functions.IsNumeric(val1) == true)
                                    {
                                        table_mleader_doglentgh = Convert.ToDouble(val1);
                                    }
                                }

                                if (dt_alias.Rows[idx_alias][i_nh] != DBNull.Value)
                                {
                                    string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_nh]);
                                    if (Functions.IsNumeric(val1) == true)
                                    {
                                        MleadertextH = Convert.ToDouble(val1);
                                    }
                                }

                                if (dt_alias.Rows[idx_alias][i_txtf] != DBNull.Value)
                                {
                                    string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_txtf]);
                                    if (val1.ToLower() == "yes" || val1.ToLower() == "true")
                                    {
                                        mleader_text_frame = true;
                                    }
                                }

                                #endregion

                                #region load object data param
                                string este_od = "OD:No";
                                string sep = "[space]";
                                string upp = "Up & low";

                                if (dt_alias.Rows[idx_alias][i_od] != DBNull.Value)
                                {
                                    string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_od]);
                                    if (val1.ToLower() == "yes" || val1.ToLower() == "true")
                                    {
                                        is_object_data = true;
                                        este_od = "OD:Yes";
                                    }
                                }

                                if (dt_alias.Rows[idx_alias][i_c] != DBNull.Value)
                                {
                                    string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_c]);
                                    if (val1.ToLower() == "yes" || val1.ToLower() == "true")
                                    {
                                        is_capitalized = true;
                                        upp = "Upper";
                                    }
                                }





                                #region countour data
                                if (dt_alias.Rows[idx_alias][i_cl] != DBNull.Value)
                                {
                                    string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_cl]);
                                    if (val1.ToLower() == "yes")
                                    {
                                        is_contour = true;
                                    }
                                }

                                if (dt_alias.Rows[idx_alias][i_clp] != DBNull.Value)
                                {
                                    string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_clp]);
                                    if (Functions.IsNumeric(val1) == true)
                                    {
                                        if (Convert.ToDouble(val1) == 0)
                                        {
                                            if (val1 == "0") contour_rounding = 0;
                                            if (val1 == "0.0") contour_rounding = 1;
                                            if (val1 == "0.0") contour_rounding = 2;
                                        }
                                        else
                                        {
                                            contour_rounding = Convert.ToInt32(val1);
                                        }
                                    }
                                }
                                #endregion

                                if (dt_alias.Rows[idx_alias][i_p1] != DBNull.Value)
                                {
                                    p1 = Convert.ToString(dt_alias.Rows[idx_alias][i_p1]);
                                }
                                if (dt_alias.Rows[idx_alias][i_od1] != DBNull.Value)
                                {
                                    f1 = Convert.ToString(dt_alias.Rows[idx_alias][i_od1]);
                                }
                                if (dt_alias.Rows[idx_alias][i_s1] != DBNull.Value)
                                {
                                    s1 = Convert.ToString(dt_alias.Rows[idx_alias][i_s1]);
                                }

                                if (dt_alias.Rows[idx_alias][i_p2] != DBNull.Value)
                                {
                                    p2 = Convert.ToString(dt_alias.Rows[idx_alias][i_p2]);
                                }
                                if (dt_alias.Rows[idx_alias][i_od2] != DBNull.Value)
                                {
                                    f2 = Convert.ToString(dt_alias.Rows[idx_alias][i_od2]);
                                }
                                if (dt_alias.Rows[idx_alias][i_s2] != DBNull.Value)
                                {
                                    s2 = Convert.ToString(dt_alias.Rows[idx_alias][i_s2]);
                                }

                                if (dt_alias.Rows[idx_alias][i_p3] != DBNull.Value)
                                {
                                    p3 = Convert.ToString(dt_alias.Rows[idx_alias][i_p3]);
                                }
                                if (dt_alias.Rows[idx_alias][i_od3] != DBNull.Value)
                                {
                                    f3 = Convert.ToString(dt_alias.Rows[idx_alias][i_od3]);
                                }
                                if (dt_alias.Rows[idx_alias][i_s3] != DBNull.Value)
                                {
                                    s3 = Convert.ToString(dt_alias.Rows[idx_alias][i_s3]);
                                }

                                if (dt_alias.Rows[idx_alias][i_p4] != DBNull.Value)
                                {
                                    p4 = Convert.ToString(dt_alias.Rows[idx_alias][i_p4]);
                                }
                                if (dt_alias.Rows[idx_alias][i_od4] != DBNull.Value)
                                {
                                    f4 = Convert.ToString(dt_alias.Rows[idx_alias][i_od4]);
                                }
                                if (dt_alias.Rows[idx_alias][i_s4] != DBNull.Value)
                                {
                                    s4 = Convert.ToString(dt_alias.Rows[idx_alias][i_s4]);
                                }

                                if (dt_alias.Rows[idx_alias][i_p5] != DBNull.Value)
                                {
                                    p5 = Convert.ToString(dt_alias.Rows[idx_alias][i_p5]);
                                }
                                if (dt_alias.Rows[idx_alias][i_od5] != DBNull.Value)
                                {
                                    f5 = Convert.ToString(dt_alias.Rows[idx_alias][i_od5]);
                                }
                                if (dt_alias.Rows[idx_alias][i_s5] != DBNull.Value)
                                {
                                    s5 = Convert.ToString(dt_alias.Rows[idx_alias][i_s5]);
                                }

                                if (dt_alias.Rows[idx_alias][i_p6] != DBNull.Value)
                                {
                                    p6 = Convert.ToString(dt_alias.Rows[idx_alias][i_p6]);
                                }
                                if (dt_alias.Rows[idx_alias][i_od6] != DBNull.Value)
                                {
                                    f6 = Convert.ToString(dt_alias.Rows[idx_alias][i_od6]);
                                }
                                if (dt_alias.Rows[idx_alias][i_s6] != DBNull.Value)
                                {
                                    s6 = Convert.ToString(dt_alias.Rows[idx_alias][i_s6]);
                                }

                                if (dt_alias.Rows[idx_alias][i_p7] != DBNull.Value)
                                {
                                    p7 = Convert.ToString(dt_alias.Rows[idx_alias][i_p7]);
                                }
                                if (dt_alias.Rows[idx_alias][i_od7] != DBNull.Value)
                                {
                                    f7 = Convert.ToString(dt_alias.Rows[idx_alias][i_od7]);
                                }
                                if (dt_alias.Rows[idx_alias][i_s7] != DBNull.Value)
                                {
                                    s7 = Convert.ToString(dt_alias.Rows[idx_alias][i_s7]);
                                }

                                if (dt_alias.Rows[idx_alias][i_p8] != DBNull.Value)
                                {
                                    p8 = Convert.ToString(dt_alias.Rows[idx_alias][i_p8]);
                                }
                                if (dt_alias.Rows[idx_alias][i_od8] != DBNull.Value)
                                {
                                    f8 = Convert.ToString(dt_alias.Rows[idx_alias][i_od8]);
                                }
                                if (dt_alias.Rows[idx_alias][i_s8] != DBNull.Value)
                                {
                                    s8 = Convert.ToString(dt_alias.Rows[idx_alias][i_s8]);
                                }

                                if (dt_alias.Rows[idx_alias][i_p9] != DBNull.Value)
                                {
                                    p9 = Convert.ToString(dt_alias.Rows[idx_alias][i_p9]);
                                }
                                if (dt_alias.Rows[idx_alias][i_od9] != DBNull.Value)
                                {
                                    f9 = Convert.ToString(dt_alias.Rows[idx_alias][i_od9]);
                                }
                                if (dt_alias.Rows[idx_alias][i_s9] != DBNull.Value)
                                {
                                    s9 = Convert.ToString(dt_alias.Rows[idx_alias][i_s9]);
                                }

                                if (dt_alias.Rows[idx_alias][i_p10] != DBNull.Value)
                                {
                                    p10 = Convert.ToString(dt_alias.Rows[idx_alias][i_p10]);
                                }
                                if (dt_alias.Rows[idx_alias][i_od10] != DBNull.Value)
                                {
                                    f10 = Convert.ToString(dt_alias.Rows[idx_alias][i_od10]);
                                }
                                if (dt_alias.Rows[idx_alias][i_s10] != DBNull.Value)
                                {
                                    s10 = Convert.ToString(dt_alias.Rows[idx_alias][i_s10]);
                                }
                                #endregion

                                #region load block atribs


                                if (dt_alias.Rows[idx_alias][i_bn] != DBNull.Value)
                                {
                                    table_bl_name = Convert.ToString(dt_alias.Rows[idx_alias][i_bn]);
                                }
                                if (dt_alias.Rows[idx_alias][i_bac] != DBNull.Value)
                                {
                                    table_atr_c = Convert.ToString(dt_alias.Rows[idx_alias][i_bac]);
                                }

                                if (dt_alias.Rows[idx_alias][i_ba1] != DBNull.Value)
                                {
                                    table_atr_1 = Convert.ToString(dt_alias.Rows[idx_alias][i_ba1]);
                                }

                                if (dt_alias.Rows[idx_alias][i_ba2] != DBNull.Value)
                                {
                                    table_atr_2 = Convert.ToString(dt_alias.Rows[idx_alias][i_ba2]);
                                }

                                if (dt_alias.Rows[idx_alias][i_ba3] != DBNull.Value)
                                {
                                    table_atr_3 = Convert.ToString(dt_alias.Rows[idx_alias][i_ba3]);
                                }

                                if (dt_alias.Rows[idx_alias][i_ba4] != DBNull.Value)
                                {
                                    table_atr_4 = Convert.ToString(dt_alias.Rows[idx_alias][i_ba4]);
                                }

                                if (dt_alias.Rows[idx_alias][i_ba5] != DBNull.Value)
                                {
                                    table_atr_5 = Convert.ToString(dt_alias.Rows[idx_alias][i_ba5]);
                                }

                                if (dt_alias.Rows[idx_alias][i_ba6] != DBNull.Value)
                                {
                                    table_atr_6 = Convert.ToString(dt_alias.Rows[idx_alias][i_ba6]);
                                }

                                if (dt_alias.Rows[idx_alias][i_ba7] != DBNull.Value)
                                {
                                    table_atr_7 = Convert.ToString(dt_alias.Rows[idx_alias][i_ba7]);
                                }

                                if (dt_alias.Rows[idx_alias][i_ba8] != DBNull.Value)
                                {
                                    table_atr_8 = Convert.ToString(dt_alias.Rows[idx_alias][i_ba8]);
                                }

                                if (dt_alias.Rows[idx_alias][i_ba9] != DBNull.Value)
                                {
                                    table_atr_9 = Convert.ToString(dt_alias.Rows[idx_alias][i_ba9]);
                                }

                                if (dt_alias.Rows[idx_alias][i_ba10] != DBNull.Value)
                                {
                                    table_atr_10 = Convert.ToString(dt_alias.Rows[idx_alias][i_ba10]);
                                }
                                #endregion

                                #region dimstyle param
                                if (dt_alias.Rows[idx_alias][i_d_sty] != DBNull.Value)
                                {
                                    table_dimstyle_name = Convert.ToString(dt_alias.Rows[idx_alias][i_d_sty]);
                                    table_dimstyle_name = table_dimstyle_name + "_" + scale1.ToString();
                                }

                                if (dt_alias.Rows[idx_alias][i_d_arr] != DBNull.Value)
                                {
                                    string valtxt = Convert.ToString(dt_alias.Rows[idx_alias][i_d_arr]);
                                    if (Functions.IsNumeric(valtxt) == true)
                                    {
                                        table_dim_arrow_size = Convert.ToDouble(dt_alias.Rows[idx_alias][i_d_arr]);
                                    }
                                }
                                if (dt_alias.Rows[idx_alias][i_d_suff] != DBNull.Value)
                                {
                                    table_dimstyle_suffix = Convert.ToString(dt_alias.Rows[idx_alias][i_d_suff]);
                                }

                                if (dt_alias.Rows[idx_alias][i_d_decno] != DBNull.Value)
                                {
                                    string valtxt = Convert.ToString(dt_alias.Rows[idx_alias][i_d_decno]);
                                    if (Functions.IsNumeric(valtxt) == true)
                                    {
                                        table_dimstyle_dec_no = Convert.ToInt32(dt_alias.Rows[idx_alias][i_d_decno]);
                                    }
                                }

                                if (dt_alias.Rows[idx_alias][i_d_round_closest] != DBNull.Value)
                                {
                                    string valtxt = Convert.ToString(dt_alias.Rows[idx_alias][i_d_round_closest]);
                                    if (Functions.IsNumeric(valtxt) == true)
                                    {
                                        table_dimstyle_round_closest = Convert.ToInt32(dt_alias.Rows[idx_alias][i_d_round_closest]);
                                    }
                                }

                                if (dt_alias.Rows[idx_alias][i_force_dim] != DBNull.Value)
                                {
                                    string valtxt = Convert.ToString(dt_alias.Rows[idx_alias][i_force_dim]);
                                    if (valtxt.ToLower() == "yes" || valtxt.ToLower() == "true")
                                    {
                                        table_dimstyle_force_line = true;
                                    }
                                }

                                #endregion

                                #region define text style

                                TextStyleTable TextStyleTable1 = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, OpenMode.ForWrite) as TextStyleTable;

                                foreach (ObjectId id1 in TextStyleTable1)
                                {
                                    TextStyleTableRecord style1 = Trans1.GetObject(id1, OpenMode.ForRead) as TextStyleTableRecord;
                                    if (id1 != null)
                                    {
                                        if (style1.Name == table_txt_stylename)
                                        {
                                            style1.UpgradeOpen();
                                            style1.FileName = table_txt_fontname;
                                            style1.ObliquingAngle = table_txt_oblique * Math.PI / 180;
                                            style1.TextSize = 0;
                                            style1.XScale = table_txt_width;
                                            textstyle_id = id1;
                                        }
                                    }
                                }

                                if (textstyle_id == ObjectId.Null)
                                {
                                    TextStyleTableRecord style1 = new TextStyleTableRecord();
                                    style1.FileName = table_txt_fontname;
                                    style1.Name = table_txt_stylename;
                                    style1.TextSize = 0;
                                    style1.ObliquingAngle = table_txt_oblique * Math.PI / 180;
                                    style1.XScale = table_txt_width;
                                    TextStyleTable1.Add(style1);
                                    Trans1.AddNewlyCreatedDBObject(style1, true);
                                    textstyle_id = style1.ObjectId;
                                }


                                #endregion

                                #region define Mleader style

                                DBDictionary MleaderTable1 = Trans1.GetObject(ThisDrawing.Database.MLeaderStyleDictionaryId, OpenMode.ForWrite) as DBDictionary;
                                foreach (DBDictionaryEntry entry1 in MleaderTable1)
                                {
                                    ObjectId id1 = MleaderTable1.GetAt(entry1.Key);
                                    MLeaderStyle mstyle1 = Trans1.GetObject(id1, OpenMode.ForRead) as MLeaderStyle;
                                    if (id1 != null)
                                    {
                                        if (mstyle1.Name == table_mstyle_name)
                                        {
                                            mstyle1.UpgradeOpen();
                                            mstyle1.ArrowSize = table_mleader_arrow * scale1;
                                            mstyle1.DoglegLength = table_mleader_doglentgh * scale1;
                                            mstyle1.LandingGap = table_mleader_gap * scale1;
                                            mstyle1.TextStyleId = textstyle_id;
                                            mstyle1.EnableFrameText = mleader_text_frame;
                                            mleader_id = id1;
                                        }
                                    }
                                }

                                if (mleader_id == ObjectId.Null)
                                {
                                    MLeaderStyle mstyle1 = new MLeaderStyle();
                                    mstyle1.ArrowSize = table_mleader_arrow * scale1;
                                    mstyle1.EnableDogleg = true;
                                    mstyle1.DoglegLength = table_mleader_doglentgh * scale1;
                                    mstyle1.LandingGap = table_mleader_gap * scale1;
                                    mstyle1.TextStyleId = textstyle_id;
                                    mstyle1.EnableFrameText = mleader_text_frame;
                                    mleader_id = mstyle1.PostMLeaderStyleToDb(ThisDrawing.Database, table_mstyle_name);
                                    Trans1.AddNewlyCreatedDBObject(mstyle1, true);
                                }

                                #endregion

                                #region read object od


                                if (is_object_data == true)
                                {
                                    System.Data.DataTable dt_od = new System.Data.DataTable();

                                    using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), curve1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, true))
                                    {
                                        if (Records1 != null)
                                        {
                                            if (Records1.Count > 0)
                                            {
                                                dt_od.Rows.Add();

                                                foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                {
                                                    Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];
                                                    Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;

                                                    for (int i = 0; i < Record1.Count; ++i)
                                                    {
                                                        Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[i];
                                                        string Nume_field = Field_def1.Name;
                                                        string Valoare1 = Record1[i].StrValue;
                                                        if (dt_od.Columns.Contains(Nume_field) == false)
                                                        {
                                                            dt_od.Columns.Add(Nume_field, typeof(string));
                                                        }
                                                        dt_od.Rows[0][Nume_field] = Valoare1;
                                                    }
                                                }
                                            }
                                        }
                                    }



                                    for (int i = 0; i < dt_od.Columns.Count; ++i)
                                    {
                                        string val1 = Convert.ToString(dt_od.Rows[0][i]);

                                        if (val1 != "")
                                        {
                                            string Col_name = dt_od.Columns[i].ColumnName;
                                            string compare_val1 = f1.ToLower().Replace(" ", "");
                                            if (Col_name.ToLower().Replace(" ", "") == compare_val1)
                                            {
                                                if (is_contour == false)
                                                {

                                                    if (p1 == "")
                                                    {

                                                    }
                                                    else if (p1.Contains("{/}") == true)
                                                    {
                                                        p1 = p1.Replace("{/}", "\\P");
                                                    }
                                                    else if (p1.Contains("{\\}") == true)
                                                    {
                                                        p1 = p1.Replace("{\\}", "\\P");
                                                    }
                                                    else if (p1.Contains("{delete space}") == true)
                                                    {
                                                        p1 = p1.Replace("{delete space}", "");
                                                    }
                                                    else
                                                    {
                                                        p1 = p1 + " ";
                                                    }

                                                    if (s1.Contains("{/}") == true)
                                                    {
                                                        s1 = s1.Replace("{/}", "\\P");
                                                    }
                                                    else if (s1.Contains("{\\}") == true)
                                                    {
                                                        s1 = s1.Replace("{\\}", "\\P");
                                                    }
                                                    else if (s1.Contains("{delete space}") == true)
                                                    {
                                                        s1 = s1.Replace("{delete space}", "");
                                                    }
                                                    else
                                                    {
                                                        s1 = " " + s1;
                                                    }

                                                    od_val1 = val1;

                                                    od_val1 = od_val1.Replace("  ", "");

                                                    if (od_val1.Length > 1)
                                                        if (od_val1.Substring(0, 1) == " ")
                                                        {
                                                            do
                                                            {
                                                                od_val1 = od_val1.Substring(1, od_val1.Length - 1);

                                                            } while (od_val1.Substring(0, 1) == " ");
                                                        }

                                                    if (od_val1.Substring(od_val1.Length - 1, 1) == " ")
                                                    {
                                                        do
                                                        {
                                                            od_val1 = od_val1.Substring(0, od_val1.Length - 1);

                                                        } while (od_val1.Substring(od_val1.Length - 1, 1) == " ");
                                                    }

                                                }
                                                else
                                                {
                                                    s1 = s1.Replace("{delete space}", "");
                                                    od_val1 = val1;
                                                }


                                            }
                                            string compare_val2 = f2.ToLower().Replace(" ", "");
                                            if (Col_name.ToLower().Replace(" ", "") == compare_val2)
                                            {

                                                if (p2 == "")
                                                {
                                                    p2 = " ";
                                                }
                                                else if (p2.Contains("{/}") == true)
                                                {
                                                    p2 = p2.Replace("{/}", "\\P");
                                                }
                                                else if (p2.Contains("{\\}") == true)
                                                {
                                                    p2 = p2.Replace("{\\}", "\\P");
                                                }
                                                else if (p2.Contains("{delete space}") == true)
                                                {
                                                    p2 = p2.Replace("{delete space}", "");
                                                }
                                                else
                                                {
                                                    p2 = p2 + " ";
                                                }


                                                if (s2.Contains("{/}") == true)
                                                {
                                                    s2 = s2.Replace("{/}", "\\P");
                                                }
                                                else if (s2.Contains("{\\}") == true)
                                                {
                                                    s2 = s2.Replace("{\\}", "\\P");
                                                }
                                                else if (s2.Contains("{delete space}") == true)
                                                {
                                                    s2 = s2.Replace("{delete space}", "");
                                                }
                                                else
                                                {
                                                    s2 = " " + s2;
                                                }

                                                od_val2 = val1;

                                                od_val2 = od_val2.Replace("  ", "");

                                                if (od_val2.Length > 1)
                                                    if (od_val2.Substring(0, 1) == " ")
                                                    {
                                                        do
                                                        {
                                                            od_val2 = od_val2.Substring(1, od_val2.Length - 1);

                                                        } while (od_val2.Substring(0, 1) == " ");
                                                    }

                                                if (od_val2.Substring(od_val2.Length - 1, 1) == " ")
                                                {
                                                    do
                                                    {
                                                        od_val2 = od_val2.Substring(0, od_val2.Length - 1);

                                                    } while (od_val2.Substring(od_val2.Length - 1, 1) == " ");
                                                }
                                            }

                                            string compare_val3 = f3.ToLower().Replace(" ", "");
                                            if (Col_name.ToLower().Replace(" ", "") == compare_val3)
                                            {
                                                if (p3 == "")
                                                {
                                                    p3 = " ";
                                                }
                                                else if (p3.Contains("{/}") == true)
                                                {
                                                    p3 = p3.Replace("{/}", "\\P");
                                                }
                                                else if (p3.Contains("{\\}") == true)
                                                {
                                                    p3 = p3.Replace("{\\}", "\\P");
                                                }
                                                else if (p3.Contains("{delete space}") == true)
                                                {
                                                    p3 = p3.Replace("{delete space}", "");
                                                }
                                                else
                                                {
                                                    p3 = p3 + " ";
                                                }
                                                if (s3.Contains("{/}") == true)
                                                {
                                                    s3 = s3.Replace("{/}", "\\P");
                                                }
                                                else if (s3.Contains("{\\}") == true)
                                                {
                                                    s3 = s3.Replace("{\\}", "\\P");
                                                }
                                                else if (s3.Contains("{delete space}") == true)
                                                {
                                                    s3 = s3.Replace("{delete space}", "");
                                                }
                                                else
                                                {
                                                    s3 = " " + s3;
                                                }

                                                od_val3 = val1;

                                                od_val3 = od_val3.Replace("  ", "");

                                                if (od_val3.Length > 1)
                                                    if (od_val3.Substring(0, 1) == " ")
                                                    {
                                                        do
                                                        {
                                                            od_val3 = od_val3.Substring(1, od_val3.Length - 1);

                                                        } while (od_val3.Substring(0, 1) == " ");
                                                    }

                                                if (od_val3.Substring(od_val3.Length - 1, 1) == " ")
                                                {
                                                    do
                                                    {
                                                        od_val3 = od_val3.Substring(0, od_val3.Length - 1);

                                                    } while (od_val3.Substring(od_val3.Length - 1, 1) == " ");
                                                }
                                            }

                                            string compare_val4 = f4.ToLower().Replace(" ", "");
                                            if (Col_name.ToLower().Replace(" ", "") == compare_val4)
                                            {
                                                if (p4 == "")
                                                {
                                                    p4 = " ";
                                                }
                                                else if (p4.Contains("{/}") == true)
                                                {
                                                    p4 = p4.Replace("{/}", "\\P");
                                                }
                                                else if (p4.Contains("{\\}") == true)
                                                {
                                                    p4 = p4.Replace("{\\}", "\\P");
                                                }
                                                else if (p4.Contains("{delete space}") == true)
                                                {
                                                    p4 = p4.Replace("{delete space}", "");
                                                }
                                                else
                                                {
                                                    p4 = p4 + " ";
                                                }

                                                if (s4.Contains("{/}") == true)
                                                {
                                                    s4 = s4.Replace("{/}", "\\P");
                                                }
                                                else if (s4.Contains("{\\}") == true)
                                                {
                                                    s4 = s4.Replace("{\\}", "\\P");
                                                }
                                                else if (s4.Contains("{delete space}") == true)
                                                {
                                                    s4 = s4.Replace("{delete space}", "");
                                                }
                                                else
                                                {
                                                    s4 = " " + s4;
                                                }

                                                od_val4 = val1;

                                                od_val4 = od_val4.Replace("  ", "");

                                                if (od_val4.Length > 1)
                                                    if (od_val4.Substring(0, 1) == " ")
                                                    {
                                                        do
                                                        {
                                                            od_val4 = od_val4.Substring(1, od_val4.Length - 1);

                                                        } while (od_val4.Substring(0, 1) == " ");
                                                    }

                                                if (od_val4.Substring(od_val4.Length - 1, 1) == " ")
                                                {
                                                    do
                                                    {
                                                        od_val4 = od_val4.Substring(0, od_val4.Length - 1);

                                                    } while (od_val4.Substring(od_val4.Length - 1, 1) == " ");
                                                }
                                            }

                                            string compare_val5 = f5.ToLower().Replace(" ", "");
                                            if (Col_name.ToLower().Replace(" ", "") == compare_val5)
                                            {
                                                if (p5 == "")
                                                {
                                                    p5 = " ";
                                                }
                                                else if (p5.Contains("{/}") == true)
                                                {
                                                    p5 = p5.Replace("{/}", "\\P");
                                                }
                                                else if (p5.Contains("{\\}") == true)
                                                {
                                                    p5 = p5.Replace("{\\}", "\\P");
                                                }
                                                else if (p5.Contains("{delete space}") == true)
                                                {
                                                    p5 = p5.Replace("{delete space}", "");
                                                }
                                                else
                                                {
                                                    p5 = p5 + " ";
                                                }

                                                if (s5.Contains("{/}") == true)
                                                {
                                                    s5 = s5.Replace("{/}", "\\P");
                                                }
                                                else if (s5.Contains("{\\}") == true)
                                                {
                                                    s5 = s5.Replace("{\\}", "\\P");
                                                }
                                                else if (s5.Contains("{delete space}") == true)
                                                {
                                                    s5 = s5.Replace("{delete space}", "");
                                                }
                                                else
                                                {
                                                    s5 = " " + s5;
                                                }

                                                od_val5 = val1;

                                                od_val5 = od_val5.Replace("  ", "");

                                                if (od_val5.Length > 1)
                                                    if (od_val5.Substring(0, 1) == " ")
                                                    {
                                                        do
                                                        {
                                                            od_val5 = od_val5.Substring(1, od_val5.Length - 1);

                                                        } while (od_val5.Substring(0, 1) == " ");
                                                    }

                                                if (od_val5.Substring(od_val5.Length - 1, 1) == " ")
                                                {
                                                    do
                                                    {
                                                        od_val5 = od_val5.Substring(0, od_val5.Length - 1);

                                                    } while (od_val5.Substring(od_val5.Length - 1, 1) == " ");
                                                }
                                            }

                                            string compare_val6 = f6.ToLower().Replace(" ", "");
                                            if (Col_name.ToLower().Replace(" ", "") == compare_val6)
                                            {
                                                if (p6 == "")
                                                {
                                                    p6 = " ";
                                                }
                                                else if (p6.Contains("{/}") == true)
                                                {
                                                    p6 = p6.Replace("{/}", "\\P");
                                                }
                                                else if (p6.Contains("{\\}") == true)
                                                {
                                                    p6 = p6.Replace("{\\}", "\\P");
                                                }
                                                else if (p6.Contains("{delete space}") == true)
                                                {
                                                    p6 = p6.Replace("{delete space}", "");
                                                }
                                                else
                                                {
                                                    p6 = p6 + " ";
                                                }

                                                if (s6.Contains("{/}") == true)
                                                {
                                                    s6 = s6.Replace("{/}", "\\P");
                                                }
                                                else if (s6.Contains("{\\}") == true)
                                                {
                                                    s6 = s6.Replace("{\\}", "\\P");
                                                }
                                                else if (s6.Contains("{delete space}") == true)
                                                {
                                                    s6 = s6.Replace("{delete space}", "");
                                                }
                                                else
                                                {
                                                    s6 = " " + s6;
                                                }

                                                od_val6 = val1;

                                                od_val6 = od_val6.Replace("  ", "");

                                                if (od_val6.Length > 1)
                                                    if (od_val6.Substring(0, 1) == " ")
                                                    {
                                                        do
                                                        {
                                                            od_val6 = od_val6.Substring(1, od_val6.Length - 1);

                                                        } while (od_val6.Substring(0, 1) == " ");
                                                    }

                                                if (od_val6.Substring(od_val6.Length - 1, 1) == " ")
                                                {
                                                    do
                                                    {
                                                        od_val6 = od_val6.Substring(0, od_val6.Length - 1);

                                                    } while (od_val6.Substring(od_val6.Length - 1, 1) == " ");
                                                }
                                            }

                                            string compare_val7 = f7.ToLower().Replace(" ", "");
                                            if (Col_name.ToLower().Replace(" ", "") == compare_val7)
                                            {
                                                if (p7 == "")
                                                {
                                                    p7 = " ";
                                                }
                                                else if (p7.Contains("{/}") == true)
                                                {
                                                    p7 = p7.Replace("{/}", "\\P");
                                                }
                                                else if (p7.Contains("{\\}") == true)
                                                {
                                                    p7 = p7.Replace("{\\}", "\\P");
                                                }
                                                else if (p7.Contains("{delete space}") == true)
                                                {
                                                    p7 = p7.Replace("{delete space}", "");
                                                }
                                                else
                                                {
                                                    p7 = p7 + " ";
                                                }

                                                if (s7.Contains("{/}") == true)
                                                {
                                                    s7 = s7.Replace("{/}", "\\P");
                                                }
                                                else if (s7.Contains("{\\}") == true)
                                                {
                                                    s7 = s7.Replace("{\\}", "\\P");
                                                }
                                                else if (s7.Contains("{delete space}") == true)
                                                {
                                                    s7 = s7.Replace("{delete space}", "");
                                                }
                                                else
                                                {
                                                    s7 = " " + s7;
                                                }

                                                od_val7 = val1;

                                                od_val7 = od_val7.Replace("  ", "");

                                                if (od_val7.Length > 1)
                                                    if (od_val7.Substring(0, 1) == " ")
                                                    {
                                                        do
                                                        {
                                                            od_val7 = od_val7.Substring(1, od_val7.Length - 1);

                                                        } while (od_val7.Substring(0, 1) == " ");
                                                    }

                                                if (od_val7.Substring(od_val7.Length - 1, 1) == " ")
                                                {
                                                    do
                                                    {
                                                        od_val7 = od_val7.Substring(0, od_val7.Length - 1);

                                                    } while (od_val7.Substring(od_val7.Length - 1, 1) == " ");
                                                }
                                            }

                                            string compare_val8 = f8.ToLower().Replace(" ", "");
                                            if (Col_name.ToLower().Replace(" ", "") == compare_val8)
                                            {
                                                if (p8 == "")
                                                {
                                                    p8 = " ";
                                                }
                                                else if (p8.Contains("{/}") == true)
                                                {
                                                    p8 = p8.Replace("{/}", "\\P");
                                                }
                                                else if (p8.Contains("{\\}") == true)
                                                {
                                                    p8 = p8.Replace("{\\}", "\\P");
                                                }
                                                else if (p8.Contains("{delete space}") == true)
                                                {
                                                    p8 = p8.Replace("{delete space}", "");
                                                }
                                                else
                                                {
                                                    p8 = p8 + " ";
                                                }

                                                if (s8.Contains("{/}") == true)
                                                {
                                                    s8 = s8.Replace("{/}", "\\P");
                                                }
                                                else if (s8.Contains("{\\}") == true)
                                                {
                                                    s8 = s8.Replace("{\\}", "\\P");
                                                }
                                                else if (s8.Contains("{delete space}") == true)
                                                {
                                                    s8 = s8.Replace("{delete space}", "");
                                                }
                                                else
                                                {
                                                    s8 = " " + s8;
                                                }

                                                od_val8 = val1;

                                                od_val8 = od_val8.Replace("  ", "");

                                                if (od_val8.Length > 1)
                                                    if (od_val8.Substring(0, 1) == " ")
                                                    {
                                                        do
                                                        {
                                                            od_val8 = od_val8.Substring(1, od_val8.Length - 1);

                                                        } while (od_val8.Substring(0, 1) == " ");
                                                    }

                                                if (od_val8.Substring(od_val8.Length - 1, 1) == " ")
                                                {
                                                    do
                                                    {
                                                        od_val8 = od_val8.Substring(0, od_val8.Length - 1);

                                                    } while (od_val8.Substring(od_val8.Length - 1, 1) == " ");
                                                }
                                            }

                                            string compare_val9 = f9.ToLower().Replace(" ", "");

                                            if (Col_name.ToLower().Replace(" ", "") == compare_val9)
                                            {
                                                if (p9 == "")
                                                {
                                                    p9 = " ";
                                                }
                                                else if (p9.Contains("{/}") == true)
                                                {
                                                    p9 = p9.Replace("{/}", "\\P");
                                                }
                                                else if (p9.Contains("{\\}") == true)
                                                {
                                                    p9 = p9.Replace("{\\}", "\\P");
                                                }
                                                else if (p9.Contains("{delete space}") == true)
                                                {
                                                    p9 = p9.Replace("{delete space}", "");
                                                }
                                                else
                                                {
                                                    p9 = p9 + " ";
                                                }

                                                if (s9.Contains("{/}") == true)
                                                {
                                                    s9 = s9.Replace("{/}", "\\P");
                                                }
                                                else if (s9.Contains("{\\}") == true)
                                                {
                                                    s9 = s9.Replace("{\\}", "\\P");
                                                }
                                                else if (s9.Contains("{delete space}") == true)
                                                {
                                                    s9 = s9.Replace("{delete space}", "");
                                                }
                                                else
                                                {
                                                    s9 = " " + s9;
                                                }

                                                od_val9 = val1;

                                                od_val9 = od_val9.Replace("  ", "");

                                                if (od_val9.Length > 1)
                                                    if (od_val9.Substring(0, 1) == " ")
                                                    {
                                                        do
                                                        {
                                                            od_val9 = od_val9.Substring(1, od_val9.Length - 1);

                                                        } while (od_val9.Substring(0, 1) == " ");
                                                    }

                                                if (od_val9.Substring(od_val9.Length - 1, 1) == " ")
                                                {
                                                    do
                                                    {
                                                        od_val9 = od_val9.Substring(0, od_val9.Length - 1);

                                                    } while (od_val9.Substring(od_val9.Length - 1, 1) == " ");
                                                }
                                            }

                                            string compare_val10 = f10.ToLower().Replace(" ", "");
                                            if (Col_name.ToLower().Replace(" ", "") == compare_val10)
                                            {
                                                if (p10 == "")
                                                {
                                                    p10 = " ";
                                                }
                                                else if (p10.Contains("{/}") == true)
                                                {
                                                    p10 = p10.Replace("{/}", "\\P");
                                                }
                                                else if (p10.Contains("{\\}") == true)
                                                {
                                                    p10 = p10.Replace("{\\}", "\\P");
                                                }
                                                else if (p10.Contains("{delete space}") == true)
                                                {
                                                    p10 = p10.Replace("{delete space}", "");
                                                }
                                                else
                                                {
                                                    p10 = p10 + " ";
                                                }


                                                if (s10.Contains("{/}") == true)
                                                {
                                                    s10 = s10.Replace("{/}", "\\P");
                                                }
                                                else if (s10.Contains("{\\}") == true)
                                                {
                                                    s10 = s10.Replace("{\\}", "\\P");
                                                }
                                                else if (s10.Contains("{delete space}") == true)
                                                {
                                                    s10 = s10.Replace("{delete space}", "");
                                                }
                                                else
                                                {
                                                    s10 = " " + s10;
                                                }

                                                od_val10 = val1;

                                                od_val10 = od_val10.Replace("  ", "");

                                                if (od_val10.Length > 1)
                                                    if (od_val10.Substring(0, 1) == " ")
                                                    {
                                                        do
                                                        {
                                                            od_val10 = od_val10.Substring(1, od_val10.Length - 1);

                                                        } while (od_val10.Substring(0, 1) == " ");
                                                    }

                                                if (od_val10.Substring(od_val10.Length - 1, 1) == " ")
                                                {
                                                    do
                                                    {
                                                        od_val10 = od_val10.Substring(0, od_val10.Length - 1);

                                                    } while (od_val10.Substring(od_val10.Length - 1, 1) == " ");
                                                }
                                            }

                                        }
                                    }

                                    if (od_val1.Length > 0)
                                    {
                                        od_description = p1 + od_val1 + s1;
                                        od_description = od_description.Replace("  ", "");
                                    }

                                    if (od_val2.Length > 0)
                                    {
                                        od_description = od_description + p2 + od_val2 + s2;
                                        od_description = od_description.Replace("  ", "");
                                    }

                                    if (od_val3.Length > 0)
                                    {
                                        od_description = od_description + p3 + od_val3 + s3;
                                        od_description = od_description.Replace("  ", "");
                                    }

                                    if (od_val4.Length > 0)
                                    {
                                        od_description = od_description + p4 + od_val4 + s4;
                                        od_description = od_description.Replace("  ", "");
                                    }

                                    if (od_val5.Length > 0)
                                    {
                                        od_description = od_description + p5 + od_val5 + s5;
                                        od_description = od_description.Replace("  ", "");
                                    }

                                    if (od_val6.Length > 0)
                                    {
                                        od_description = od_description + p6 + od_val6 + s6;
                                        od_description = od_description.Replace("  ", "");
                                    }

                                    if (od_val7.Length > 0)
                                    {
                                        od_description = od_description + p7 + od_val7 + s7;
                                        od_description = od_description.Replace("  ", "");
                                    }

                                    if (od_val8.Length > 0)
                                    {
                                        od_description = od_description + p8 + od_val8 + s8;
                                        od_description = od_description.Replace("  ", "");
                                    }

                                    if (od_val9.Length > 0)
                                    {
                                        od_description = od_description + p9 + od_val9 + s9;
                                        od_description = od_description.Replace("  ", "");
                                    }

                                    if (od_val10.Length > 0)
                                    {
                                        od_description = od_description + p10 + od_val10 + s10;
                                        od_description = od_description.Replace("  ", "");
                                    }

                                    if (is_capitalized == true)
                                    {
                                        od_description = od_description.ToUpper();
                                    }
                                }
                                #endregion

                                #region dimension style

                                ObjectId dimstyle_id = ObjectId.Null;
                                bool is_dim_style = false;
                                foreach (ObjectId id1 in Dim_style_table)
                                {
                                    DimStyleTableRecord style1 = Trans1.GetObject(id1, OpenMode.ForWrite) as DimStyleTableRecord;
                                    if (style1.Name.ToLower() == table_dimstyle_name.ToLower())
                                    {
                                        dimstyle_id = style1.ObjectId;
                                        is_dim_style = true;
                                        style1.Dimasz = table_dim_arrow_size * scale1;
                                        style1.Dimdec = table_dimstyle_dec_no;
                                        style1.Dimtxt = textH * scale1;
                                        style1.Dimscale = 1;
                                        style1.Dimadec = 0;
                                        style1.Dimalt = false;
                                        style1.Dimaltd = 2;
                                        style1.Dimaltf = 25.4;
                                        style1.Dimaltrnd = 0;
                                        style1.Dimalttd = 2;
                                        style1.Dimalttz = 0;
                                        style1.Dimaltu = 2;
                                        style1.Dimaltz = 0;
                                        style1.Dimapost = "";
                                        style1.Dimarcsym = 0;
                                        style1.Dimatfit = 0;
                                        style1.Dimaunit = 0;
                                        style1.Dimazin = 0;
                                        style1.Dimcen = 0.09;
                                        style1.Dimclrd = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByBlock, 0);
                                        style1.Dimclre = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByBlock, 0);
                                        style1.Dimclrt = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByBlock, 0);
                                        style1.Dimdle = 0;
                                        style1.Dimdli = 0.38;
                                        style1.Dimexe = 0.18;
                                        style1.Dimexo = 0.0625;
                                        style1.Dimfrac = 0;
                                        style1.Dimfxlen = 0.2;
                                        style1.DimfxlenOn = false;
                                        style1.Dimgap = 0.09;
                                        style1.Dimjogang = 0.785398163397448;
                                        style1.Dimjust = 0;
                                        style1.Dimlfac = 1;
                                        style1.Dimlim = false;
                                        style1.Dimlunit = 2;
                                        style1.Dimlwd = 0;
                                        style1.Dimlwe = 0;
                                        style1.Dimpost = table_dimstyle_suffix;
                                        style1.Dimrnd = table_dimstyle_round_closest;
                                        style1.Dimsah = false;
                                        style1.Dimsd1 = false;
                                        style1.Dimsd2 = false;
                                        style1.Dimse1 = false;
                                        style1.Dimse2 = false;
                                        style1.Dimsoxd = false;
                                        style1.Dimtad = 0;
                                        style1.Dimtdec = 0;
                                        style1.Dimtfac = 1;
                                        style1.Dimtfill = 1;
                                        style1.Dimtfillclr = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByBlock, 0);
                                        style1.Dimtih = true;
                                        style1.Dimtix = false;
                                        style1.Dimtm = 0;
                                        style1.Dimtmove = 0;
                                        style1.Dimtofl = table_dimstyle_force_line;
                                        style1.Dimtoh = true;
                                        style1.Dimtol = false;
                                        style1.Dimtolj = 1;
                                        style1.Dimtp = 0;
                                        style1.Dimtsz = 0;
                                        style1.Dimtvp = 0;
                                        style1.Dimtxtdirection = false;
                                        style1.Dimtzin = 0;
                                        style1.Dimupt = false;
                                        style1.Dimzin = 0;
                                        style1.Dimtxsty = textstyle_id;


                                    }

                                }

                                if (is_dim_style == false)
                                {
                                    DimStyleTableRecord style1 = new DimStyleTableRecord();
                                    style1.Name = table_dimstyle_name;
                                    style1.Dimasz = table_dim_arrow_size * scale1;
                                    style1.Dimdec = table_dimstyle_dec_no;
                                    style1.Dimtxt = textH * scale1;
                                    style1.Dimscale = 1;
                                    style1.Dimadec = 0;
                                    style1.Dimalt = false;
                                    style1.Dimaltd = 2;
                                    style1.Dimaltf = 25.4;
                                    style1.Dimaltrnd = 0;
                                    style1.Dimalttd = 2;
                                    style1.Dimalttz = 0;
                                    style1.Dimaltu = 2;
                                    style1.Dimaltz = 0;
                                    style1.Dimapost = "";
                                    style1.Dimarcsym = 0;
                                    style1.Dimatfit = 0;
                                    style1.Dimaunit = 0;
                                    style1.Dimazin = 0;
                                    style1.Dimcen = 0.09;
                                    style1.Dimclrd = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByBlock, 0);
                                    style1.Dimclre = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByBlock, 0);
                                    style1.Dimclrt = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByBlock, 0);
                                    style1.Dimdle = 0;
                                    style1.Dimdli = 0.38;
                                    style1.Dimexe = 0.18;
                                    style1.Dimexo = 0.0625;
                                    style1.Dimfrac = 0;
                                    style1.Dimfxlen = 0.2;
                                    style1.DimfxlenOn = false;
                                    style1.Dimgap = 0.09;
                                    style1.Dimjogang = 0.785398163397448;
                                    style1.Dimjust = 0;
                                    style1.Dimlfac = 1;
                                    style1.Dimlim = false;
                                    style1.Dimlunit = 2;
                                    style1.Dimlwd = 0;
                                    style1.Dimlwe = 0;
                                    style1.Dimpost = table_dimstyle_suffix;
                                    style1.Dimrnd = table_dimstyle_round_closest;
                                    style1.Dimsah = false;
                                    style1.Dimsd1 = false;
                                    style1.Dimsd2 = false;
                                    style1.Dimse1 = false;
                                    style1.Dimse2 = false;
                                    style1.Dimsoxd = false;
                                    style1.Dimtad = 0;
                                    style1.Dimtdec = 0;
                                    style1.Dimtfac = 1;
                                    style1.Dimtfill = 1;
                                    style1.Dimtfillclr = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByBlock, 0);
                                    style1.Dimtih = true;
                                    style1.Dimtix = false;
                                    style1.Dimtm = 0;
                                    style1.Dimtmove = 0;
                                    style1.Dimtofl = table_dimstyle_force_line;
                                    style1.Dimtoh = true;
                                    style1.Dimtol = false;
                                    style1.Dimtolj = 1;
                                    style1.Dimtp = 0;
                                    style1.Dimtsz = 0;
                                    style1.Dimtvp = 0;
                                    style1.Dimtxtdirection = false;
                                    style1.Dimtzin = 0;
                                    style1.Dimupt = false;
                                    style1.Dimzin = 0;
                                    style1.Dimtxsty = textstyle_id;

                                    Dim_style_table.Add(style1);
                                    Trans1.AddNewlyCreatedDBObject(style1, true);
                                    dimstyle_id = style1.ObjectId;
                                }


                                #endregion

                                stabileste_label_current();
                            return_because_is_aligned_changed:

                                string first_point_message = "\nLabel insertion point:";
                                if (current_labelstyle == "ml")
                                {
                                    first_point_message = "\nMleader first point:";
                                }

                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1 = null;
                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions(first_point_message);
                                PP1.AllowNone = false;

                                if (is_contour == false && align_to_fea == false)
                                {
                                    Point_res1 = Editor1.GetPoint(PP1);
                                }
                                else if (align_to_fea == true && current_labelstyle != "mt")
                                {
                                    Point_res1 = Editor1.GetPoint(PP1);
                                }

                                Point3d Inspt = new Point3d();
                                Point3d base_pt = new Point3d();
                                if (Point_res1 != null)
                                {
                                    if (Point_res1.Status != PromptStatus.OK)
                                    {
                                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", old_osnap);
                                        this.MdiParent.WindowState = FormWindowState.Normal;
                                        Trans1.Commit();
                                        Editor1.SetImpliedSelection(Empty_array);
                                        Editor1.WriteMessage("\nCommand:");
                                        set_enable_true();
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        Button_SetRotation.Enabled = true;
                                        button_mleader_ne.Enabled = true;
                                        radioButton1.Text = "1";
                                        radioButton2.Text = "2";
                                        radioButton3.Text = "3";
                                        radioButton1.Visible = false;
                                        radioButton2.Visible = false;
                                        radioButton3.Visible = false;
                                        panel_type.Visible = false;
                                        panel_bdy.Visible = false;

                                        if (xref != "" && drawingset1.AllDrawingsCount > 0)
                                        {
                                            try
                                            {
                                                drawingset1.DetachDrawing(xref);
                                            }
                                            catch (Autodesk.Gis.Map.MapException mapex)
                                            {

                                                MessageBox.Show("the file\r\n" + xref + "\r\n\r\nis locked by someone else\r\n\r\n" + mapex.Message);
                                            }
                                        }
                                        return;
                                    }

                                    Inspt = Point_res1.Value.TransformBy(curent_ucs_matrix);
                                    base_pt = Point_res1.Value;

                                    if (layer_name_specified_in_table == "")
                                    {
                                        label_layer = "_Lgen_" + layer_curba + "_" + scale1.ToString() + suff_layer;
                                    }
                                    else
                                    {
                                        label_layer = layer_name_specified_in_table;
                                    }

                                }
                                else
                                {
                                    Inspt = Rezultat1.PickedPoint.TransformBy(curent_ucs_matrix);
                                    base_pt = Rezultat1.PickedPoint;
                                    if (layer_name_specified_in_table == "")
                                    {
                                        label_layer = "_Lgen_contours_" + scale1.ToString();
                                    }
                                    else
                                    {
                                        label_layer = layer_name_specified_in_table;
                                    }


                                }

                                stabileste_label_current();


                                Functions.Creaza_layer(label_layer, 2, true);

                                if (current_labelstyle == "mt")
                                {
                                    #region creaza new mtext


                                    if (is_object_data == true)
                                    {
                                        if (od_description.Length > 0)
                                        {
                                            gen_description = od_description;
                                        }
                                    }

                                    if (is_contour == true)
                                    {
                                        if (s1.Contains("{delete space}'") == true) s1 = "'";
                                        if (Functions.IsNumeric(od_val1) == true)
                                        {
                                            double Elev1 = Convert.ToDouble(od_val1);
                                            gen_description = Functions.Get_String_Rounded(Elev1, contour_rounding) + s1;
                                        }
                                    }

                                    double textrot = rotation1 - deltax;

                                    if (align_to_fea == true && is_bdy == true)
                                    {
                                        Inspt = curve1.GetClosestPointTo(Inspt, Vector3d.ZAxis, false);
                                    }

                                    if (align_to_fea == true)
                                    {


                                        textrot = pick_rot_1pt(Inspt);

                                        if (textrot == -1000)
                                        {
                                            goto return_because_is_aligned_changed;
                                        }

                                        if (Functions.IsNumeric(od_val1) == false && is_contour == true)
                                        {
                                            double Elev1 = Inspt.Z;
                                            gen_description = Functions.Get_String_Rounded(Elev1, contour_rounding) + s1;
                                        }
                                    }
                                    else
                                    {
                                        if (txt_underline == true)
                                        {
                                            gen_description = "{\\L" + gen_description + "}";
                                        }
                                    }

                                    MText mtext1 = new MText();
                                    mtext1.Contents = gen_description;
                                    mtext1.TextHeight = textH * scale1;
                                    mtext1.ColorIndex = 256;

                                    if (mtxt_background_frame == true)
                                    {
                                        mtext1.BackgroundFill = true;
                                        mtext1.UseBackgroundColor = true;
                                        mtext1.BackgroundScaleFactor = 1.2;
                                    }


                                    if (align_to_fea == true)
                                    {
                                        mtext1.Attachment = AttachmentPoint.MiddleCenter;
                                    }
                                    else if (is_bdy == true)
                                    {
                                        mtext1.Attachment = AttachmentPoint.MiddleCenter;
                                    }
                                    else if (gen_description.Contains("\\P") == true)
                                    {
                                        mtext1.Attachment = AttachmentPoint.MiddleCenter;
                                    }
                                    else
                                    {
                                        mtext1.Attachment = AttachmentPoint.BottomLeft;
                                    }



                                    mtext1.Rotation = textrot;
                                    mtext1.Layer = label_layer;
                                    mtext1.Location = Inspt;
                                    mtext1.TextStyleId = textstyle_id;
                                    BTrecord.AppendEntity(mtext1);
                                    Trans1.AddNewlyCreatedDBObject(mtext1, true);
                                    #endregion
                                }
                                else if (current_labelstyle == "ml")
                                {
                                    #region creaza new mleader

                                    if (is_object_data == true)
                                    {
                                        if (od_description.Length > 0)
                                        {
                                            gen_description = od_description;
                                        }
                                    }

                                    PromptPointResult Point_res2;
                                    PromptPointOptions PP2;
                                    PP2 = new PromptPointOptions("\nMleader second point:");
                                    PP2.AllowNone = false;
                                    PP2.UseBasePoint = true;
                                    PP2.BasePoint = base_pt;

                                    Point_res2 = Editor1.GetPoint(PP2);

                                    if (Point_res2.Status != PromptStatus.OK)
                                    {
                                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", old_osnap);
                                        this.MdiParent.WindowState = FormWindowState.Normal;
                                        Trans1.Commit();
                                        Editor1.SetImpliedSelection(Empty_array);
                                        Editor1.WriteMessage("\nCommand:");
                                        set_enable_true();
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        Button_SetRotation.Enabled = true;
                                        button_mleader_ne.Enabled = true;
                                        radioButton1.Text = "1";
                                        radioButton2.Text = "2";
                                        radioButton3.Text = "3";
                                        radioButton1.Visible = false;
                                        radioButton2.Visible = false;
                                        radioButton3.Visible = false;
                                        panel_type.Visible = false;
                                        panel_bdy.Visible = false;
                                        if (xref != "" && drawingset1.AllDrawingsCount > 0)
                                        {
                                            try
                                            {
                                                drawingset1.DetachDrawing(xref);
                                            }
                                            catch (Autodesk.Gis.Map.MapException mapex)
                                            {

                                                MessageBox.Show("the file\r\n" + xref + "\r\n\r\nis locked by someone else\r\n\r\n" + mapex.Message);
                                            }
                                        }
                                        return;
                                    }
                                    Point3d pt2 = Point_res2.Value.TransformBy(curent_ucs_matrix);

                                    MLeader mleader1 = new MLeader();
                                    mleader1.ColorIndex = 256;
                                    mleader1.MLeaderStyle = mleader_id;
                                    mleader1.Layer = label_layer;

                                    MText mt_ml = new MText();
                                    mt_ml.TextStyleId = textstyle_id;
                                    mt_ml.Contents = gen_description;
                                    mt_ml.Attachment = AttachmentPoint.MiddleLeft;

                                    if (mtxt_background_frame == true)
                                    {
                                        mt_ml.BackgroundFill = true;
                                        mt_ml.UseBackgroundColor = true;
                                        mt_ml.BackgroundScaleFactor = 1.2;
                                    }


                                    mt_ml.TextHeight = MleadertextH * scale1;
                                    mt_ml.Rotation = rotation1 - deltax;
                                    mleader1.MText = mt_ml;


                                    if (is_bdy == true && radioButton5.Checked == true)
                                    {
                                        string new_arr = "_Integral";
                                        ObjectId arrow_id = GetArrowObjectId(new_arr);
                                        mleader1.ArrowSymbolId = arrow_id;
                                    }

                                    if (radioButton4.Checked == true)
                                    {
                                        Inspt = curve1.GetClosestPointTo(Inspt, Vector3d.ZAxis, false);
                                    }


                                    int leaderline_index = mleader1.AddLeader();
                                    int leaderpoint = mleader1.AddLeaderLine(leaderline_index);
                                    mleader1.AddFirstVertex(leaderpoint, Inspt);
                                    mleader1.AddLastVertex(leaderpoint, pt2);
                                    mleader1.LeaderLineType = LeaderType.StraightLeader;
                                    mleader1.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.LeftLeader);
                                    mleader1.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.RightLeader);

                                    BTrecord.AppendEntity(mleader1);
                                    Trans1.AddNewlyCreatedDBObject(mleader1, true);

                                    #endregion
                                }
                                else if (current_labelstyle == "bl")
                                {
                                    #region creaza new block

                                    if (is_object_data == true)
                                    {
                                        if (od_description.Length > 0)
                                        {
                                            gen_description = od_description;
                                        }
                                    }

                                    System.Collections.Specialized.StringCollection col_atr = new System.Collections.Specialized.StringCollection();
                                    System.Collections.Specialized.StringCollection col_val = new System.Collections.Specialized.StringCollection();

                                    col_atr.Add(table_atr_c);
                                    col_val.Add(gen_description);


                                    List<string> atribute_partiale = new List<string>();
                                    List<string> valori_partiale = new List<string>();

                                    if (table_atr_1 != "" && od_val1.Replace(" ", "").Length > 0)
                                    {
                                        if (table_atr_1.Contains("{+}") == true)
                                        {
                                            table_atr_1 = table_atr_1.Replace("{+}", "");
                                            atribute_partiale.Add(table_atr_1);
                                            valori_partiale.Add(p1 + od_val1 + s1);
                                        }
                                        else
                                        {
                                            col_atr.Add(table_atr_1);
                                            col_val.Add(p1 + od_val1 + s1);
                                        }
                                    }




                                    if (table_atr_2 != "" && od_val2.Replace(" ", "").Length > 0)
                                    {
                                        if (table_atr_2.Contains("{+}") == true)
                                        {
                                            table_atr_2 = table_atr_2.Replace("{+}", "");
                                            if (atribute_partiale.Contains(table_atr_2) == true)
                                            {
                                                int index1 = atribute_partiale.IndexOf(table_atr_2);
                                                valori_partiale[index1] = valori_partiale[index1] + " " + p2 + od_val2 + s2;
                                            }
                                            else
                                            {
                                                atribute_partiale.Add(table_atr_2);
                                                valori_partiale.Add(p2 + od_val2 + s2);
                                            }

                                        }
                                        else
                                        {
                                            col_atr.Add(table_atr_2);
                                            col_val.Add(p2 + od_val2 + s2);
                                        }
                                    }



                                    if (table_atr_3 != "" && od_val3.Replace(" ", "").Length > 0)
                                    {
                                        if (table_atr_3.Contains("{+}") == true)
                                        {
                                            table_atr_3 = table_atr_3.Replace("{+}", "");
                                            if (atribute_partiale.Contains(table_atr_3) == true)
                                            {
                                                int index1 = atribute_partiale.IndexOf(table_atr_3);
                                                valori_partiale[index1] = valori_partiale[index1] + " " + p3 + od_val3 + s3;
                                            }
                                            else
                                            {
                                                atribute_partiale.Add(table_atr_3);
                                                valori_partiale.Add(p3 + od_val3 + s3);
                                            }

                                        }
                                        else
                                        {
                                            col_atr.Add(table_atr_3);
                                            col_val.Add(p3 + od_val3 + s3);
                                        }
                                    }

                                    if (table_atr_4 != "" && od_val4.Replace(" ", "").Length > 0)
                                    {
                                        if (table_atr_4.Contains("{+}") == true)
                                        {
                                            table_atr_4 = table_atr_4.Replace("{+}", "");
                                            if (atribute_partiale.Contains(table_atr_4) == true)
                                            {
                                                int index1 = atribute_partiale.IndexOf(table_atr_4);
                                                valori_partiale[index1] = valori_partiale[index1] + " " + p4 + od_val4 + s4;
                                            }
                                            else
                                            {
                                                atribute_partiale.Add(table_atr_4);
                                                valori_partiale.Add(p4 + od_val4 + s4);
                                            }

                                        }
                                        else
                                        {
                                            col_atr.Add(table_atr_4);
                                            col_val.Add(p4 + od_val4 + s4);
                                        }
                                    }

                                    if (table_atr_5 != "" && od_val5.Replace(" ", "").Length > 0)
                                    {
                                        if (table_atr_5.Contains("{+}") == true)
                                        {
                                            table_atr_5 = table_atr_5.Replace("{+}", "");
                                            if (atribute_partiale.Contains(table_atr_5) == true)
                                            {
                                                int index1 = atribute_partiale.IndexOf(table_atr_5);
                                                valori_partiale[index1] = valori_partiale[index1] + " " + p5 + od_val5 + s5;
                                            }
                                            else
                                            {
                                                atribute_partiale.Add(table_atr_5);
                                                valori_partiale.Add(p5 + od_val5 + s5);
                                            }

                                        }
                                        else
                                        {
                                            col_atr.Add(table_atr_5);
                                            col_val.Add(p5 + od_val5 + s5);
                                        }
                                    }

                                    if (table_atr_6 != "" && od_val6.Replace(" ", "").Length > 0)
                                    {
                                        if (table_atr_6.Contains("{+}") == true)
                                        {
                                            table_atr_6 = table_atr_6.Replace("{+}", "");
                                            if (atribute_partiale.Contains(table_atr_6) == true)
                                            {
                                                int index1 = atribute_partiale.IndexOf(table_atr_6);
                                                valori_partiale[index1] = valori_partiale[index1] + " " + p6 + od_val6 + s6;
                                            }
                                            else
                                            {
                                                atribute_partiale.Add(table_atr_6);
                                                valori_partiale.Add(p6 + od_val6 + s6);
                                            }

                                        }
                                        else
                                        {
                                            col_atr.Add(table_atr_6);
                                            col_val.Add(p6 + od_val6 + s6);
                                        }
                                    }

                                    if (table_atr_7 != "" && od_val7.Replace(" ", "").Length > 0)
                                    {
                                        if (table_atr_7.Contains("{+}") == true)
                                        {
                                            table_atr_7 = table_atr_7.Replace("{+}", "");
                                            if (atribute_partiale.Contains(table_atr_7) == true)
                                            {
                                                int index1 = atribute_partiale.IndexOf(table_atr_7);
                                                valori_partiale[index1] = valori_partiale[index1] + " " + p7 + od_val7 + s7;
                                            }
                                            else
                                            {
                                                atribute_partiale.Add(table_atr_7);
                                                valori_partiale.Add(p7 + od_val7 + s7);
                                            }

                                        }
                                        else
                                        {
                                            col_atr.Add(table_atr_7);
                                            col_val.Add(p7 + od_val7 + s7);
                                        }
                                    }

                                    if (table_atr_8 != "" && od_val8.Replace(" ", "").Length > 0)
                                    {
                                        if (table_atr_8.Contains("{+}") == true)
                                        {
                                            table_atr_8 = table_atr_8.Replace("{+}", "");
                                            if (atribute_partiale.Contains(table_atr_8) == true)
                                            {
                                                int index1 = atribute_partiale.IndexOf(table_atr_8);
                                                valori_partiale[index1] = valori_partiale[index1] + " " + p8 + od_val8 + s8;
                                            }
                                            else
                                            {
                                                atribute_partiale.Add(table_atr_8);
                                                valori_partiale.Add(p8 + od_val8 + s8);
                                            }

                                        }
                                        else
                                        {
                                            col_atr.Add(table_atr_8);
                                            col_val.Add(p8 + od_val8 + s8);
                                        }
                                    }

                                    if (table_atr_9 != "" && od_val9.Replace(" ", "").Length > 0)
                                    {
                                        if (table_atr_9.Contains("{+}") == true)
                                        {
                                            table_atr_9 = table_atr_9.Replace("{+}", "");
                                            if (atribute_partiale.Contains(table_atr_9) == true)
                                            {
                                                int index1 = atribute_partiale.IndexOf(table_atr_9);
                                                valori_partiale[index1] = valori_partiale[index1] + " " + p9 + od_val9 + s9;
                                            }
                                            else
                                            {
                                                atribute_partiale.Add(table_atr_9);
                                                valori_partiale.Add(p9 + od_val9 + s9);
                                            }

                                        }
                                        else
                                        {
                                            col_atr.Add(table_atr_9);
                                            col_val.Add(p9 + od_val9 + s9);
                                        }
                                    }

                                    if (table_atr_10 != "" && od_val10.Replace(" ", "").Length > 0)
                                    {
                                        if (table_atr_10.Contains("{+}") == true)
                                        {
                                            table_atr_10 = table_atr_10.Replace("{+}", "");
                                            if (atribute_partiale.Contains(table_atr_10) == true)
                                            {
                                                int index1 = atribute_partiale.IndexOf(table_atr_10);
                                                valori_partiale[index1] = valori_partiale[index1] + " " + p10 + od_val10 + s10;
                                            }
                                            else
                                            {
                                                atribute_partiale.Add(table_atr_10);
                                                valori_partiale.Add(p10 + od_val10 + s10);
                                            }

                                        }
                                        else
                                        {
                                            col_atr.Add(table_atr_10);
                                            col_val.Add(p10 + od_val10 + s10);
                                        }
                                    }

                                    if (atribute_partiale.Count > 0)
                                    {
                                        for (int i = 0; i < atribute_partiale.Count; ++i)
                                        {
                                            string desc_par = valori_partiale[i];
                                            desc_par = desc_par.Replace("  ", " ");
                                            if (desc_par.Length > 0)
                                            {
                                                do
                                                {
                                                    if (desc_par.Substring(0, 1) == " ")
                                                    {
                                                        desc_par = desc_par.Substring(1, desc_par.Length - 1);
                                                    }
                                                }
                                                while (desc_par.Substring(0, 1) == " ");

                                                do
                                                {
                                                    if (desc_par.Substring(desc_par.Length - 1, 1) == " ")
                                                    {
                                                        desc_par = desc_par.Substring(0, desc_par.Length - 1);
                                                    }
                                                }
                                                while (desc_par.Substring(desc_par.Length - 1, 1) == " ");
                                            }

                                            if (desc_par.Length > 0)
                                            {
                                                col_atr.Add(atribute_partiale[i]);
                                                col_val.Add(desc_par);
                                            }


                                        }
                                    }


                                    if (BlockTable1.Has(table_bl_name) == true)
                                    {
                                        BlockReference bl1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "",
                                                                    table_bl_name, Inspt, scale1, rotation1, label_layer, col_atr, col_val);
                                        bl1.ColorIndex = 256;
                                    }
                                    else
                                    {
                                        MessageBox.Show("The block defined for this layer on the LGen alias table does not exist in this drawing");
                                    }




                                    #endregion
                                }
                                else if (current_labelstyle == "dim")
                                {
                                    #region creaza new dimension
                                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", 128);
                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                                    Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                                    PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify second point");
                                    PP2.AllowNone = false;
                                    PP2.UseBasePoint = true;
                                    PP2.BasePoint = base_pt;
                                    Point_res2 = Editor1.GetPoint(PP2);

                                    if (Point_res2.Status != PromptStatus.OK)
                                    {
                                        Editor1.SetImpliedSelection(Empty_array);
                                        Editor1.WriteMessage("\nCommand:");
                                        set_enable_true();
                                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", old_osnap);
                                        if (xref != "" && drawingset1.AllDrawingsCount > 0)
                                        {
                                            try
                                            {
                                                drawingset1.DetachDrawing(xref);
                                            }
                                            catch (Autodesk.Gis.Map.MapException mapex)
                                            {

                                                MessageBox.Show("the file\r\n" + xref + "\r\n\r\nis locked by someone else\r\n\r\n" + mapex.Message);
                                            }
                                        }

                                        return;
                                    }

                                    Point3d pt1 = Point_res1.Value.TransformBy(curent_ucs_matrix);
                                    Point3d pt2 = Point_res2.Value.TransformBy(curent_ucs_matrix);
                                    Point3d pt3 = new Point3d((1 + pt1.X + pt2.X) / 2, 1 + (pt1.Y + pt2.Y) / 2, 0);

                                    RotatedDimension dim1 = new RotatedDimension();
                                    dim1.XLine1Point = pt1;
                                    dim1.XLine2Point = pt2;
                                    dim1.Rotation = Functions.GET_Bearing_rad(pt1.X, pt1.Y, pt2.X, pt2.Y);
                                    dim1.DimLinePoint = pt3;
                                    dim1.DimensionStyle = dimstyle_id;
                                    dim1.Layer = label_layer;
                                    dim1.HorizontalRotation = -deltax;

                                    BTrecord.AppendEntity(dim1);
                                    Trans1.AddNewlyCreatedDBObject(dim1, true);

                                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", old_osnap);
                                    #endregion
                                }



                                if (xref != "" && drawingset1.AllDrawingsCount > 0)
                                {
                                    try
                                    {
                                        drawingset1.DetachDrawing(xref);
                                    }
                                    catch (Autodesk.Gis.Map.MapException mapex)
                                    {

                                        MessageBox.Show("the file\r\n" + xref + "\r\n\r\nis locked by someone else\r\n\r\n" + mapex.Message);
                                    }
                                }
                            }
                            else
                            {

                                radioButton1.Text = "1";
                                radioButton2.Text = "2";
                                radioButton3.Text = "3";
                                radioButton1.Visible = false;
                                radioButton2.Visible = false;
                                radioButton3.Visible = false;
                                panel_type.Visible = false;
                                panel_bdy.Visible = false;
                                this.Refresh();

                                MessageBox.Show("The object you selected is not defined on the LGEN alias table.\r\nPlease ask the drafting lead how to label this object, or to update the table.");
                            }
                        }
                        else
                        {

                            radioButton1.Text = "1";
                            radioButton2.Text = "2";
                            radioButton3.Text = "3";
                            radioButton1.Visible = false;
                            radioButton2.Visible = false;
                            radioButton3.Visible = false;
                            panel_type.Visible = false;
                            panel_bdy.Visible = false;
                            this.Refresh();
                        }
                        Trans1.TransactionManager.QueueForGraphicsFlush();


                        goto repeat1;

                    }
                }
            }
            catch (System.Exception ex)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", old_osnap);
                this.MdiParent.WindowState = FormWindowState.Normal;
                MessageBox.Show(ex.Message);
            }

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();
            Button_SetRotation.Enabled = true;
            button_mleader_ne.Enabled = true;
            radioButton1.Text = "1";
            radioButton2.Text = "2";
            radioButton3.Text = "3";
            radioButton1.Visible = false;
            radioButton2.Visible = false;
            radioButton3.Visible = false;
            panel_type.Visible = false;
            panel_bdy.Visible = false;


        }

        private  ObjectId GetArrowObjectId(string new_arrow)
        {
            ObjectId arrObjId = ObjectId.Null;
            Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database Database1 = ThisDrawing.Database;
            string old_arrow = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("DIMBLK") as string;
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("DIMBLK", new_arrow);
            if (old_arrow.Length != 0)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("DIMBLK", old_arrow);
            }

            Transaction Trans1 = Database1.TransactionManager.StartTransaction();
            using (Trans1)
            {
                BlockTable BlockTable1 = Trans1.GetObject(Database1.BlockTableId, OpenMode.ForRead) as BlockTable;
                arrObjId = BlockTable1[new_arrow];
                Trans1.Commit();
            }
            return arrObjId;
        }

        private void Button_pick_Rotation_Click(object sender, EventArgs e)
        {
            Button_create_label.Enabled = false;
            button_mleader_ne.Enabled = false;

            old_osnap = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE");

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                set_enable_false();
                this.MdiParent.WindowState = FormWindowState.Minimized;
                txt_rot = pick_rot_2pt();
                double angle1 = txt_rot * 180 / Math.PI;
                char deg_symbol = Convert.ToChar(176);
                Label_Rotation.Text = "Rotation = " + Functions.Get_String_Rounded(angle1, 0) + deg_symbol;
                this.MdiParent.WindowState = FormWindowState.Normal;
                Button_create_label.Enabled = true;
                button_mleader_ne.Enabled = true;

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", old_osnap);

            Button_create_label.Enabled = true;
            button_mleader_ne.Enabled = true;

        }
        private double pick_rot_2pt()
        {
            //End = 1,
            //Middle = 2,
            //Center = 4,
            //Node = 8,
            //Quadrant = 16,
            //Intersection = 32,
            //Insertion = 64,
            //Perpendicular = 128,
            //Tangent = 256,
            //Near = 512,
            // Quick = 1024,
            //ApparentIntersection = 2048,
            //Immediate = 65536,
            //AllowTangent = 131072,
            // DisablePerpendicular = 262144,
            //RelativeCartesian = 524288,
            //RelativePolar = 1048576,
            //NoneOverride = 2097152,  

            old_osnap = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE");
            object new_osnap = 512;


            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", new_osnap);

            set_enable_false();
            using (DocumentLock lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {

                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                    Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                    PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nrotation pick first point");
                    PP1.AllowNone = false;
                    Point_res1 = Editor1.GetPoint(PP1);

                    if (Point_res1.Status != PromptStatus.OK)
                    {
                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", old_osnap);
                        return 0;
                    }

                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                    Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                    PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nrotation pick second point:");
                    PP2.AllowNone = false;
                    PP2.UseBasePoint = true;
                    PP2.BasePoint = Point_res1.Value;
                    Point_res2 = Editor1.GetPoint(PP2);

                    if (Point_res2.Status != PromptStatus.OK)
                    {
                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", old_osnap);
                        return 0;
                    }
                    Point3d pt1 = Point_res1.Value.TransformBy(curent_ucs_matrix);
                    Point3d pt2 = Point_res2.Value.TransformBy(curent_ucs_matrix);

                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", old_osnap);
                    return Functions.GET_Bearing_rad(pt1.X, pt1.Y, pt2.X, pt2.Y);
                }
            }
        }
        private double pick_rot_1pt(Point3d pt1)
        {
            //End = 1,
            //Middle = 2,
            //Center = 4,
            //Node = 8,
            //Quadrant = 16,
            //Intersection = 32,
            //Insertion = 64,
            //Perpendicular = 128,
            //Tangent = 256,
            //Near = 512,
            // Quick = 1024,
            //ApparentIntersection = 2048,
            //Immediate = 65536,
            //AllowTangent = 131072,
            // DisablePerpendicular = 262144,
            //RelativeCartesian = 524288,
            //RelativePolar = 1048576,
            //NoneOverride = 2097152,  


            object new_osnap = 512;


            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", new_osnap);

            set_enable_false();
            using (DocumentLock lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    stabileste_label_current();

                    if (align_to_fea == true && current_labelstyle != "mt")
                    {
                        return -1000;
                    }

                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                    Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                    PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nrotation pick second point:");
                    PP2.AllowNone = false;
                    PP2.UseBasePoint = true;
                    PP2.BasePoint = pt1.TransformBy(curent_ucs_matrix.Inverse());
                    Point_res2 = Editor1.GetPoint(PP2);

                    if (Point_res2.Status != PromptStatus.OK)
                    {
                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", old_osnap);
                        return -1000;
                    }

                    Point3d pt2 = Point_res2.Value.TransformBy(curent_ucs_matrix);

                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", old_osnap);

                    if (align_to_fea == true && current_labelstyle != "mt")
                    {
                        return -1000;
                    }

                    return Functions.GET_Bearing_rad(pt1.X, pt1.Y, pt2.X, pt2.Y);
                }
            }
        }



        private void radioButton_label_type_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton rb1 = sender as RadioButton;

            if (rb1 != null)
            {
                if (rb1.Checked == true)
                {
                    if (rb1.Text == "Mtext")
                    {
                        current_labelstyle = "mt";
                        suff_layer = "_Mtext";
                    }
                    if (rb1.Text == "MLeader")
                    {
                        current_labelstyle = "ml";
                        suff_layer = "_Mleader";
                    }
                    if (rb1.Text == "Block")
                    {
                        current_labelstyle = "bl";
                        suff_layer = "_Block";
                    }
                    if (rb1.Text == "Dimension")
                    {
                        current_labelstyle = "dim";
                        suff_layer = "_Dim";
                    }

                    if (is_bdy == true && rb1.Text == "MLeader")
                    {
                        panel_bdy.Visible = true;
                    }
                    else
                    {
                        panel_bdy.Visible = false;
                    }
                }
            }

        }

        private void stabileste_label_current()
        {

            if (radioButton1.Checked == true)
            {
                if (radioButton1.Text == "Mtext")
                {
                    current_labelstyle = "mt";
                    suff_layer = "_Mtext";
                }
                if (radioButton1.Text == "MLeader")
                {
                    current_labelstyle = "ml";
                    suff_layer = "_Mleader";
                }
                if (radioButton1.Text == "Block")
                {
                    current_labelstyle = "bl";
                    suff_layer = "_Block";
                }
                if (radioButton1.Text == "Dimension")
                {
                    current_labelstyle = "dim";
                    suff_layer = "_Dim";
                }
            }

            if (radioButton2.Checked == true)
            {
                if (radioButton2.Text == "Mtext")
                {
                    current_labelstyle = "mt";
                    suff_layer = "_Mtext";
                }
                if (radioButton2.Text == "MLeader")
                {
                    current_labelstyle = "ml";
                    suff_layer = "_Mleader";
                }
                if (radioButton2.Text == "Block")
                {
                    current_labelstyle = "bl";
                    suff_layer = "_Block";
                }
                if (radioButton2.Text == "Dimension")
                {
                    current_labelstyle = "dim";
                    suff_layer = "_Dim";
                }
            }

            if (radioButton3.Checked == true)
            {
                if (radioButton3.Text == "Mtext")
                {
                    current_labelstyle = "mt";
                    suff_layer = "_Mtext";
                }
                if (radioButton3.Text == "MLeader")
                {
                    current_labelstyle = "ml";
                    suff_layer = "_Mleader";
                }
                if (radioButton3.Text == "Block")
                {
                    current_labelstyle = "bl";
                    suff_layer = "_Block";
                }
                if (radioButton3.Text == "Dimension")
                {
                    current_labelstyle = "dim";
                    suff_layer = "_Dim";
                }
            }


        }

        private void button_mleader_ne_Click(object sender, EventArgs e)
        {
            old_osnap = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE");

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        double deltax = Functions.GET_deltaX_rad();


                        Button_SetRotation.Enabled = false;

                    repeat1:

                        string cb_txt = Combobox_scales.Text;
                        cb_txt = cb_txt.Replace("1' = ", "");
                        scale1 = Convert.ToDouble(cb_txt.Replace("\"", ""));

                        double rotation1 = txt_rot;

                        string table_txt_stylename = "LGEN_STANDARD_MTEXT";
                        string table_txt_fontname = "ARIAL.TTF";
                        double table_txt_oblique = 0;
                        double table_txt_width = 1;
                        ObjectId textstyle_id = ObjectId.Null;



                        label_layer = "_Lgen_northing_easting_" + scale1.ToString();

                        Functions.Creaza_layer(label_layer, 10, true);

                        #region define text style

                        TextStyleTable TextStyleTable1 = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, OpenMode.ForWrite) as TextStyleTable;

                        foreach (ObjectId id1 in TextStyleTable1)
                        {
                            TextStyleTableRecord style1 = Trans1.GetObject(id1, OpenMode.ForRead) as TextStyleTableRecord;
                            if (id1 != null)
                            {
                                if (style1.Name == table_txt_stylename)
                                {
                                    style1.UpgradeOpen();
                                    style1.FileName = table_txt_fontname;
                                    style1.ObliquingAngle = table_txt_oblique * Math.PI / 180;
                                    style1.TextSize = 0;
                                    style1.XScale = table_txt_width;
                                    textstyle_id = id1;
                                }
                            }
                        }

                        if (textstyle_id == ObjectId.Null)
                        {
                            TextStyleTableRecord style1 = new TextStyleTableRecord();
                            style1.FileName = table_txt_fontname;
                            style1.Name = table_txt_stylename;
                            style1.TextSize = 0;
                            style1.ObliquingAngle = table_txt_oblique * Math.PI / 180;
                            style1.XScale = table_txt_width;
                            TextStyleTable1.Add(style1);
                            Trans1.AddNewlyCreatedDBObject(style1, true);
                            textstyle_id = style1.ObjectId;
                        }


                        #endregion

                        string table_mstyle_name = "LGEN_DOT_MLEADER_" + scale1.ToString();
                        double table_mleader_gap = 0.02;
                        double table_mleader_arrow = 0.08;
                        double table_mleader_doglentgh = 0.02;
                        double MleadertextH = 0.08;
                        ObjectId mleader_id = ObjectId.Null;

                        #region define Mleader style

                        DBDictionary MleaderTable1 = Trans1.GetObject(ThisDrawing.Database.MLeaderStyleDictionaryId, OpenMode.ForWrite) as DBDictionary;
                        foreach (DBDictionaryEntry entry1 in MleaderTable1)
                        {
                            ObjectId id1 = MleaderTable1.GetAt(entry1.Key);
                            MLeaderStyle mstyle1 = Trans1.GetObject(id1, OpenMode.ForRead) as MLeaderStyle;
                            if (id1 != null)
                            {
                                if (mstyle1.Name == table_mstyle_name)
                                {
                                    mstyle1.UpgradeOpen();
                                    mstyle1.ArrowSize = table_mleader_arrow * scale1;
                                    mstyle1.DoglegLength = table_mleader_doglentgh * scale1;
                                    mstyle1.LandingGap = table_mleader_gap * scale1;
                                    mstyle1.TextStyleId = textstyle_id;
                                    mleader_id = id1;
                                }
                            }
                        }

                        if (mleader_id == ObjectId.Null)
                        {
                            MLeaderStyle mstyle1 = new MLeaderStyle();
                            mstyle1.ArrowSize = table_mleader_arrow * scale1;
                            mstyle1.EnableDogleg = true;
                            mstyle1.DoglegLength = table_mleader_doglentgh * scale1;
                            mstyle1.LandingGap = table_mleader_gap * scale1;
                            mstyle1.TextStyleId = textstyle_id;
                            mleader_id = mstyle1.PostMLeaderStyleToDb(ThisDrawing.Database, table_mstyle_name);
                            Trans1.AddNewlyCreatedDBObject(mstyle1, true);

                        }

                        #endregion

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1 = null;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                        PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify point:");
                        PP1.AllowNone = false;

                        Point_res1 = Editor1.GetPoint(PP1);

                        if (Point_res1.Status != PromptStatus.OK)
                        {
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            Trans1.Commit();
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            Button_SetRotation.Enabled = true;
                            return;
                        }

                        Point3d Inspt = Point_res1.Value.TransformBy(curent_ucs_matrix);

                        string continut = "NORTHING: " + Functions.Get_String_Rounded(Inspt.Y, 0) + "\\PEASTING: " + Functions.Get_String_Rounded(Inspt.X, 0);

                        #region creaza new mleader

                        PromptPointResult Point_res2;
                        PromptPointOptions PP2;
                        PP2 = new PromptPointOptions("\nMleader second point:");
                        PP2.AllowNone = false;
                        PP2.UseBasePoint = true;
                        PP2.BasePoint = Point_res1.Value;

                        Point_res2 = Editor1.GetPoint(PP2);

                        if (Point_res2.Status != PromptStatus.OK)
                        {

                            this.MdiParent.WindowState = FormWindowState.Normal;
                            Trans1.Commit();
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            Button_SetRotation.Enabled = true;
                            return;
                        }
                        Point3d pt2 = Point_res2.Value.TransformBy(curent_ucs_matrix);

                        MLeader mleader1 = new MLeader();
                        mleader1.ColorIndex = 256;
                        mleader1.MLeaderStyle = mleader_id;
                        mleader1.Layer = label_layer;

                        MText mt_ml = new MText();
                        mt_ml.TextStyleId = textstyle_id;
                        mt_ml.Contents = continut;
                        mt_ml.Attachment = AttachmentPoint.MiddleLeft;
                        mt_ml.BackgroundFill = true;
                        mt_ml.UseBackgroundColor = true;
                        mt_ml.BackgroundScaleFactor = 1.2;
                        mt_ml.TextHeight = MleadertextH * scale1;
                        mt_ml.Rotation = rotation1 - deltax;
                        mleader1.MText = mt_ml;


                        string new_arr = "_Dot";
                        ObjectId arrow_id = GetArrowObjectId(new_arr);
                        mleader1.ArrowSymbolId = arrow_id;


                        int leaderline_index = mleader1.AddLeader();
                        int leaderpoint = mleader1.AddLeaderLine(leaderline_index);

                        mleader1.AddFirstVertex(leaderpoint, Inspt);
                        mleader1.AddLastVertex(leaderpoint, pt2);
                        mleader1.LeaderLineType = LeaderType.StraightLeader;
                        mleader1.SetTextAttachmentType(TextAttachmentType.AttachmentBottomOfTopLine, LeaderDirectionType.LeftLeader);
                        mleader1.SetTextAttachmentType(TextAttachmentType.AttachmentBottomOfTopLine, LeaderDirectionType.RightLeader);
                        mleader1.EnableLanding = true;

                        BTrecord.AppendEntity(mleader1);
                        Trans1.AddNewlyCreatedDBObject(mleader1, true);

                        #endregion

                        Trans1.TransactionManager.QueueForGraphicsFlush();

                        goto repeat1;

                    }
                }
            }
            catch (System.Exception ex)
            {
                this.MdiParent.WindowState = FormWindowState.Normal;
                MessageBox.Show(ex.Message);
            }

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();
            Button_SetRotation.Enabled = true;


        }
        private void dimensions()
        {
            //With Dimension1
            //.Dimasz = 18 'Controls the size of dimension line and leader line arrowheads. Also controls the size of hook lines
            //            'Multiples of the arrowhead size determine whether dimension lines and text should fit between the extension lines. DIMASZ is also used to scale arrowhead blocks if set by DIMBLK. DIMASZ has no effect when DIMTSZ is other than zero

            // .Dimdec = 0
            //            'Sets the number of decimal places displayed for the primary units of a dimension
            //            'The precision is based on the units or angle format you have selected. 


            //  .Dimtxt = 8 'Specifies the height of dimension text, unless the current text style has a fixed height

            //            .TextStyleId = Text_style_romans.ObjectId

            //            .Dimtxtdirection = False
            //            'Specifies the reading direction of the dimension text. 
            //            '0 - Displays dimension text in a Left-to-Right reading style 
            //            '1 - Displays dimension text in a Right-to-Left reading style  



            //            .Dimtofl = False
            //            'Initial value: Off (imperial) or On (metric)  
            //            'Controls whether a dimension line is drawn between the extension lines even when the text is placed outside. 
            //            'For radius and diameter dimensions (when DIMTIX is off), draws a dimension line inside the circle or arc and places the text, arrowheads, and leader outside. 
            //            ' Off -  Does not draw dimension lines between the measured points when arrowheads are placed outside the measured points 
            //            ' On -  Draws dimension lines between the measured points even when arrowheads are placed outside the measured points 

            //            .Dimtoh = False
            //            'Controls the position of dimension text outside the extension lines. 
            //            ' Off -  Aligns text with the dimension line
            //            ' On -  Draws text horizontally

            //            .Dimtih = False
            //            'Initial value: On (imperial) or Off (metric)  
            //            'Controls the position of dimension text inside the extension lines for all dimension types except Ordinate. 
            //            'Off - Aligns text with the dimension line
            //            'On -  Draws text horizontally

            //            .Dimtad = 0
            //            'Controls the vertical position of text in relation to the dimension line. 
            //            '0 - Centers the dimension text between the extension lines. 
            //            '1 - Places the dimension text above the dimension line except when the dimension line is not horizontal and text inside the extension lines is forced horizontal ( DIMTIH = 1). 
            //            '    The distance from the dimension line to the baseline of the lowest line of text is the current DIMGAP value. 
            //            '2 - Places the dimension text on the side of the dimension line farthest away from the defining points. 
            //            '3 - Places the dimension text to conform to Japanese Industrial Standards (JIS). 
            //            '4 - Places the dimension text below the dimension line. 


            //            .Dimtvp = 0
            //            'Controls the vertical position of dimension text above or below the dimension line. 
            //            'The DIMTVP value is used when DIMTAD is off. The magnitude of the vertical offset of text is the product of the text height and DIMTVP. 
            //            'Setting DIMTVP to 1.0 is equivalent to setting DIMTAD to on. The dimension line splits to accommodate the text only if the absolute value of DIMTVP is less than 0.7. 


            //            .Dimsd1 = False
            //            'Controls suppression of the first dimension line and arrowhead. 
            //            'When turned on, suppresses the display of the dimension line and arrowhead between the first extension line and the text. 
            //            .Dimsd2 = False
            //            'Controls suppression of the second dimension line and arrowhead. 
            //            'When turned on, suppresses the display of the dimension line and arrowhead between the second extension line and the text. 
            //            .Dimse1 = True 'Suppresses display of the first extension line. 
            //            .Dimse2 = True 'Suppresses display of the second extension line

            //            .Dimrnd = 5
            //            'Rounds all dimensioning distances to the specified value. 
            //            'For instance, if DIMRND is set to 0.25, all distances round to the nearest 0.25 unit. 
            //            'If you set DIMRND to 1.0, all distances round to the nearest integer. 
            //            'Note that the number of digits edited after the decimal point depends on the precision set by DIMDEC. DIMRND does not apply to angular dimensions. 

            //            .Dimpost = "<>'"
            //            'Specifies a text prefix or suffix (or both) to the dimension measurement. 
            //            'For example, to establish a suffix for millimeters, set DIMPOST to mm; a distance of 19.2 units would be displayed as 19.2 mm. 
            //            'If tolerances are turned on, the suffix is applied to the tolerances as well as to the main dimension. 
            //            'Use <> to indicate placement of the text in relation to the dimension value. 
            //            'For example, enter <>mm to display a 5.0 millimeter radial dimension as "5.0mm." 
            //            'If you entered mm <>, the dimension would be displayed as "mm 5.0." 
            //            'Use the <> mechanism for angular dimensions. 

            //            .Dimjust = 0
            //            'Controls the horizontal positioning of dimension text. 
            //            '0 -  Positions the text above the dimension line and center-justifies it between the extension lines 
            //            '1 -  Positions the text next to the first extension line 
            //            '2 -  Positions the text next to the second extension line 
            //            '3 -  Positions the text above and aligned with the first extension line 
            //            '4 -  Positions the text above and aligned with the second extension line 

            //            .Dimadec = 0 'Controls the number of precision places displayed in angular dimensions. (0-8)
            //            .Dimalt = False 'Controls the display of alternate units in dimensions. Off - Disables alternate units
            //            .Dimaltd = 2 'Controls the number of decimal places in alternate units. If DIMALT is turned on, DIMALTD sets the number of digits displayed to the right of the decimal point in the alternate measurement
            //            .Dimaltf = 25.4 'Controls the multiplier for alternate units. If DIMALT is turned on, DIMALTF multiplies linear dimensions by a factor to produce a value in an alternate system of measurement. The initial value represents the number of millimeters in an inch.
            //            .Dimaltmzf = 100
            //            .Dimaltrnd = 0 'Rounds off the alternate dimension units. 
            //            .Dimalttd = 2 'Sets the number of decimal places for the tolerance values in the alternate units of a dimension. 
            //            .Dimalttz = 0 'Controls suppression of zeros in tolerance values. 
            //            .Dimaltu = 2 'Sets the units format for alternate units of all dimension substyles except Angular. (2 - Decimal)
            //            .Dimaltz = 0 'Controls the suppression of zeros for alternate unit dimension values. 
            //            .Dimapost = "" 'Specifies a text prefix or suffix (or both) to the alternate dimension measurement for all types of dimensions except angular. 
            //            'For instance, if the current units are Architectural, DIMALT is on, DIMALTF is 25.4 (the number of millimeters per inch), DIMALTD is 2, and DIMPOST is set to "mm," a distance of 10 units would be displayed as 10"[254.00mm]. 
            //            'To turn off an established prefix or suffix (or both), set it to a single period (.). 
            //            .Dimarcsym = 0 'Controls display of the arc symbol in an arc length dimension. (0- Places arc length symbols before the dimension text )
            //            '1 - Places arc length symbols above the dimension text 
            //            '2 -  Suppresses the display of arc length symbols 

            //            .Dimatfit = 3
            //            'Determines how dimension text and arrows are arranged when space is not sufficient to place both within the extension lines. 
            //            '0 -  Places both text and arrows outside extension lines 
            //            '1 -  Moves arrows first, then text
            //            '2 -  Moves text first, then arrows
            //            '3 -  Moves either text or arrows, whichever fits best 
            //            'A leader is added to moved dimension text when DIMTMOVE is set to 1. 


            //            .Dimaunit = 0 'Sets the units format for angular dimensions. (0 - Decimal degrees)
            //            .Dimazin = 0 'Suppresses zeros for angular dimensions. 


            //            .Dimsah = False
            //            'Controls the display of dimension line arrowhead blocks. 
            //            'Off - Use arrowhead blocks set by DIMBLK
            //            'On - Use arrowhead blocks set by DIMBLK1 and DIMBLK2

            //            .Dimblk = Arrowid
            //            'Sets the arrowhead block displayed at the ends of dimension lines or leader lines. 
            //            'To return to the default, closed-filled arrowhead display, enter a single period (.). Arrowhead block entries and the names used to select them in the New, Modify, and Override Dimension Style dialog boxes are shown below. You can also enter the names of user-defined arrowhead blocks. 
            //            'Note Annotative blocks cannot be used as custom arrowheads for dimensions or leaders. 
            //            '"" - Closed(filled)
            //            '"_DOT" - dot
            //            '"_DOTSMALL" - dot small
            //            '"_DOTBLANK" - dot blank
            //            '"_ORIGIN" - origin indicator
            //            '"_ORIGIN2" - origin indicator 2
            //            '"_OPEN" - open
            //            '"_OPEN90" - Right(angle)
            //            '"_OPEN30" - open 30
            //            '"_CLOSED" - Closed
            //            '"_SMALL" - dot small blank
            //            '"_NONE" - none
            //            '"_OBLIQUE" - oblique
            //            '"_BOXFILLED" - box filled
            //            '"_BOXBLANK" - box
            //            '"_CLOSEDBLANK" - Closed(blank)
            //            '"_DATUMFILLED" - datum triangle filled
            //            '"_DATUMBLANK" - datum triangle
            //            '"_INTEGRAL" - integral
            //            '"_ARCHTICK" - architectural tick


            //            .Dimblk1 = Arrowid
            //            'Sets the arrowhead for the first end of the dimension line when DIMSAH is on. 
            //            'To return to the default, closed-filled arrowhead display, enter a single period (.). For a list of arrowheads, see DIMBLK. 
            //            'Note Annotative blocks cannot be used as custom arrowheads for dimensions or leaders. 
            //            .Dimblk2 = Arrowid
            //            'Sets the arrowhead for the second end of the dimension line when DIMSAH is on. 
            //            'To return to the default, closed-filled arrowhead display, enter a single period (.). For a list of arrowhead entries, see DIMBLK. 
            //            'Note Annotative blocks cannot be used as custom arrowheads for dimensions or leaders. 
            //            .Dimldrblk = Arrowid ' Specifies the arrow type for leaders. 

            //            .Dimcen = 0.09 'Controls drawing of circle or arc center marks and centerlines by the DIMCENTER, DIMDIAMETER, and DIMRADIUS commands. 
            //            .Dimclrd = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256) ' Assigns colors to dimension lines, arrowheads, and dimension leader lines
            //            .Dimclre = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256) 'Assigns colors to dimension extension lines.
            //            .Dimclrt = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256) 'Assigns colors to dimension text


            //            .Dimdle = 0 'Sets the distance the dimension line extends beyond the extension line when oblique strokes are drawn instead of arrowheads. 
            //            .Dimdli = 0.38 'Controls the spacing of the dimension lines in baseline dimensions. 
            //            'Each dimension line is offset from the previous one by this amount, if necessary, to avoid drawing over it. Changes made with DIMDLI are not applied to existing dimensions
            //            .Dimdsep = ".c"
            //            'Specifies a single-character decimal separator to use when creating dimensions whose unit format is decimal
            //            'When prompted, enter a single character at the Command prompt. If dimension units is set to Decimal, the DIMDSEP character is used instead of the default decimal point.
            //            'If DIMDSEP is set to NULL (default value, reset by entering a period), the decimal point is used as the dimension separator
            //            .Dimexe = 0.18 'Specifies how far to extend the extension line beyond the dimension line. 
            //            .Dimexo = 0.0625 'Specifies how far extension lines are offset from origin points. 
            //            'With fixed-length extension lines, this value determines the minimum offset. 
            //            .Dimfrac = 0 'Sets the fraction format when DIMLUNIT is set to 4 (Architectural) or 5 (Fractional).
            //            '0 - Horizontal stacking
            //            '1 - Diagonal stacking
            //            '2 - Not stacked (for example, 1/2)


            //            .Dimfxlen = 1
            //            .DimfxlenOn = False

            //            .Dimgap = 0.09 'Sets the distance around the dimension text when the dimension line breaks to accommodate dimension text.
            //            .Dimjogang = 0.785398163 'Determines the angle of the transverse segment of the dimension line in a jogged radius dimension. 



            //            .Dimlfac = 1
            //            'Sets a scale factor for linear dimension measurements. 
            //            'All linear dimension distances, including radii, diameters, and coordinates, are multiplied by DIMLFAC before being converted to dimension text. Positive values of DIMLFAC are applied to dimensions in both model space and paper space; negative values are applied to paper space only. 
            //            'DIMLFAC applies primarily to nonassociative dimensions (DIMASSOC set 0 or 1). For nonassociative dimensions in paper space, DIMLFAC must be set individually for each layout viewport to accommodate viewport scaling. 
            //            'DIMLFAC has no effect on angular dimensions, and is not applied to the values held in DIMRND, DIMTM, or DIMTP. 

            //            .Dimltex1 = ThisDrawing.Database.ByBlockLinetype 'Sets the linetype of the first extension line. 
            //            .Dimltex2 = ThisDrawing.Database.ByBlockLinetype 'Sets the linetype of the second extension line. 
            //            .Dimltype = ThisDrawing.Database.ByBlockLinetype 'Sets the linetype of the dimension line.

            //            .Dimlunit = 2
            //            'Sets units for all dimension types except Angular. 
            //            '1 Scientific
            //            '2 Decimal
            //            '3 Engineering
            //            '4 Architectural (always displayed stacked)
            //            '5 Fractional (always displayed stacked)
            //            '6 Microsoft Windows Desktop (decimal format using Control Panel settings for decimal separator and number grouping symbols) 


            //            .Dimlwd = LineWeight.ByBlock
            //            'Assigns lineweight to dimension lines. 
            //            '-3 Default (the LWDEFAULT value) 
            //            '-2 BYBLOCK
            //            '-1 BYLAYER

            //            .Dimlwe = LineWeight.ByBlock
            //            'Assigns lineweight to extension  lines. 
            //            '-3 Default (the LWDEFAULT value) 
            //            '-2 BYBLOCK
            //            '-1 BYLAYER



            //            .Dimmzf = 100


            // .Dimscale = 1
            //            'Sets the overall scale factor applied to dimensioning variables that specify sizes, distances, or offsets. 
            //            'Also affects the leader objects with the LEADER command. 
            //            'Use MLEADERSCALE to scale multileader objects created with the MLEADER command. 
            //            '0.0 - A reasonable default value is computed based on the scaling between the current model space viewport and paper space. 
            //            'If you are in paper space or model space and not using the paper space feature, the scale factor is 1.0. 
            //            '>0 - A scale factor is computed that leads text sizes, arrowhead sizes, and other scaled distances to plot at their face values. 
            //            'DIMSCALE does not affect measured lengths, coordinates, or angles. 
            //            'Use DIMSCALE to control the overall scale of dimensions. However, if the current dimension style is annotative, 
            //            'DIMSCALE is automatically set to zero and the dimension scale is controlled by the CANNOSCALE system variable. DIMSCALE cannot be set to a non-zero value when using annotative dimensions. 

            //            .Dimtdec = 0
            //            'Sets the number of decimal places to display in tolerance values for the primary units in a dimension. 
            //            'This system variable has no effect unless DIMTOL is set to On. The default for DIMTOL is Off. 

            //            .Dimtfac = 1
            //            'Specifies a scale factor for the text height of fractions and tolerance values relative to the dimension text height, as set by DIMTXT. 
            //            'For example, if DIMTFAC is set to 1.0, the text height of fractions and tolerances is the same height as the dimension text. 
            //            'If DIMTFAC is set to 0.7500, the text height of fractions and tolerances is three-quarters the size of dimension text. 
            //            .Dimtfill = 1
            //            'Controls the background of dimension text. 
            //            '0 -  No Background
            //            '1 -  The background color of the drawing 
            //            '2 -  The background specified by DIMTFILLCLR
            //            .Dimtfillclr = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByBlock, 0)

            //            .Dimtix = False
            //            'Draws text between extension lines. 
            //            'Off -  Varies with the type of dimension. 
            //            '        For linear and angular dimensions, text is placed inside the extension lines if there is sufficient room. 
            //            '        For radius and diameter dimensions that don't fit inside the circle or arc, DIMTIX has no effect and always forces the text outside the circle or arc. 
            //            'On -  Draws dimension text between the extension lines even if it would ordinarily be placed outside those lines 

            //            .Dimsoxd = False
            //            'Suppresses arrowheads if not enough space is available inside the extension lines. 
            //            'Off -  Arrowheads are not suppressed
            //            'On -  Arrowheads are suppressed
            //            'If not enough space is available inside the extension lines and DIMTIX is on, setting DIMSOXD to On suppresses the arrowheads. If DIMTIX is off, DIMSOXD has no effect. 


            //            .Dimtm = 0
            //            'Sets the minimum (or lower) tolerance limit for dimension text when DIMTOL or DIMLIM is on. 
            //            'DIMTM accepts signed values. If DIMTOL is on and DIMTP and DIMTM are set to the same value, a tolerance value is drawn. 
            //            'If DIMTM and DIMTP values differ, the upper tolerance is drawn above the lower, and a plus sign is added to the DIMTP value if it is positive. 
            //            'For DIMTM, the program uses the negative of the value you enter (adding a minus sign if you specify a positive number and a plus sign if you specify a negative number). 

            //            .Dimtmove = 0
            //            'Sets dimension text movement rules. 
            //            '0 -  Moves the dimension line with dimension text
            //            '1 -  Adds a leader when dimension text is moved
            //            '2 -  Allows text to be moved freely without a leader

            //            .Dimtp = 0
            //            'Sets the maximum (or upper) tolerance limit for dimension text when DIMTOL or DIMLIM is on. 
            //            'DIMTP accepts signed values. If DIMTOL is on and DIMTP and DIMTM are set to the same value, a tolerance value is drawn. 
            //            'If DIMTM and DIMTP values differ, the upper tolerance is drawn above the lower and a plus sign is added to the DIMTP value if it is positive. 


            //            .Dimlim = False
            //            'Generates dimension limits as the default text. 
            //            'Setting DIMLIM to On turns DIMTOL off. 
            //            'Off -  Dimension limits are not generated as default text 
            //            'On -  Dimension limits are generated as default text


            //            .Dimtol = False
            //            'Appends tolerances to dimension text. 
            //            'Setting DIMTOL to on turns DIMLIM off. 

            //            .Dimtolj = 1 'Sets the vertical justification for tolerance values relative to the nominal dimension text. 



            //            .Dimtsz = 0
            //            'Specifies the size of oblique strokes drawn instead of arrowheads for linear, radius, and diameter dimensioning. 
            //            '0 -  Draws arrowheads.
            //            '>0 -  Draws oblique strokes instead of arrowheads. The size of the oblique strokes is determined by this value multiplied by the DIMSCALE value 




            //            .Dimtzin = 0 'Controls the suppression of zeros in tolerance values. 

            //            .Dimupt = False
            //            'Controls options for user-positioned text. 
            //            'Off -  Cursor controls only the dimension line location
            //            'On -  Cursor controls both the text position and the dimension line location 

            //            .Dimzin = 0
            //            'Controls the suppression of zeros in the primary unit value. 
            //            'Values 0-3 affect feet-and-inch dimensions only: 
            //            '0 -  Suppresses zero feet and precisely zero inches
            //            '1 -  Includes zero feet and precisely zero inches
            //            '2 -  Includes zero feet and suppresses zero inches
            //            '3 -  Includes zero inches and suppresses zero feet
            //            '4 -  Suppresses leading zeros in decimal dimensions (for example, 0.5000 becomes .5000) 
            //            '8 -  Suppresses trailing zeros in decimal dimensions (for example, 12.5000 becomes 12.5) 
            //            '12 -  Suppresses both leading and trailing zeros (for example, 0.5000 becomes .5) 



            //        End With
        }


        private void button1_Click_2(object sender, EventArgs e)
        {

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                        Autodesk.Gis.Map.Project.ProjectModel project1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject;
                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = project1.ODTables;
                        Autodesk.Gis.Map.Project.DrawingSet drawingset1 = project1.DrawingSet;

                        PromptNestedEntityOptions prompt1 = new PromptNestedEntityOptions("\nSelect object");
                        prompt1.AllowNone = false;
                        prompt1.UseNonInteractivePickPoint = false;
                        PromptNestedEntityResult rezultat1 = Editor1.GetNestedEntity(prompt1);
                        if (rezultat1.Status == PromptStatus.OK)
                        {
                            DBObject dbobj1 = Trans1.GetObject(rezultat1.ObjectId, OpenMode.ForRead) as DBObject;
                            BlockReference bl1 = Trans1.GetObject(rezultat1.GetContainers()[0], OpenMode.ForRead) as BlockReference;
                            if (bl1 != null)
                            {
                                BlockTableRecord btr1 = Trans1.GetObject(bl1.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                                if (btr1 != null)
                                {
                                    if (btr1.IsFromExternalReference == true)
                                    {
                                        string xref = btr1.GetXrefDatabase(false).Filename;
                                        try
                                        {
                                            drawingset1.AttachDrawing(xref);
                                            if (dbobj1 != null)
                                            {
                                                if (dbobj1 is Curve)
                                                {
                                                    Curve curva1 = dbobj1 as Curve;
                                                    Autodesk.Gis.Map.ObjectData.Records Records1;
                                                    using (Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), dbobj1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, true))
                                                    {
                                                        if (Records1 != null)
                                                        {
                                                            if (Records1.Count > 0)
                                                            {
                                                                foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                                {
                                                                    Autodesk.Gis.Map.Project.ProjectModel Proj1 = Record1.Project;
                                                                    Autodesk.Gis.Map.ObjectData.Tables Tables2 = Proj1.ODTables;

                                                                    Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables2[Record1.TableName];
                                                                    Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;

                                                                    for (int i = 0; i < Record1.Count; ++i)
                                                                    {
                                                                        Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[i];
                                                                        string Nume_field = Field_def1.Name;
                                                                        string Valoare1 = Record1[i].StrValue;
                                                                        if (Nume_field.ToLower() == "first_name") MessageBox.Show(Valoare1);
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                    drawingset1.DetachDrawing(xref);
                                                }
                                            }
                                        }
                                        catch (Autodesk.Gis.Map.MapException mapex)
                                        {

                                            MessageBox.Show("the file\r\n" + xref + "\r\n\r\nis locked by someone else\r\n\r\n" + mapex.Message);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                MessageBox.Show(ex.Message);
            }

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();


        }

        private void button_brd_mleader_Click(object sender, EventArgs e)
        {
            ObjectId[] Empty_array = null;
            idx_alias = -1;
            set_enable_false();
            old_osnap = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE");
            int round1 = 1;

            try
            {


                if (dt_alias == null || dt_alias.Rows.Count == 0)
                {
                    MessageBox.Show("NO layer alias loaded");
                    set_enable_true();
                    return;
                }

                string table_txt_stylename = "LGEN_STANDARD_MTEXT";
                string table_txt_fontname = "ARIAL.TTF";
                double table_txt_oblique = 0;
                double table_txt_width = 1;
                ObjectId textstyle_id = ObjectId.Null;

                string table_mstyle_name = "LGEN_DOT_MLEADER_" + scale1.ToString();
                double table_mleader_gap = 0.02;
                double table_mleader_arrow = 0.08;
                double table_mleader_doglentgh = 0.02;
                double MleadertextH = 0.08;
                ObjectId mleader_id = ObjectId.Null;


                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    double vp_SCALE = 1;
                    double vp_TWIST = 0;
                    ObjectId Ent_vp_id = ObjectId.Null;

                    string cb_txt = Combobox_scales.Text;
                    cb_txt = cb_txt.Replace("1' = ", "");
                    scale1 = Convert.ToDouble(cb_txt.Replace("\"", ""));


                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        if (is_paperspace==true)
                        {
                            int Tilemode1 = Convert.ToInt32(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("TILEMODE"));
                            int CVport1 = Convert.ToInt32(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CVPORT"));

                            scale1 = 1;

                            if (Tilemode1 == 0)
                            {
                                if (CVport1 != 1)
                                {
                                    Editor1.SwitchToPaperSpace();
                                }

                            }
                            else
                            {
                                MessageBox.Show("this option is meant to work in paper space\r\nPlease select Model Space from the combobox if you need the label in model space");
                                set_enable_true();
                                return;
                            }


                            Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_Viewport;
                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_viewport;
                            Prompt_viewport = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the viewport:");
                            Prompt_viewport.SetRejectMessage("\nSelect a viewport!");
                            Prompt_viewport.AllowNone = true;
                            Prompt_viewport.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Viewport), false);
                            Prompt_viewport.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);

                            Rezultat_Viewport = ThisDrawing.Editor.GetEntity(Prompt_viewport);

                            if (Rezultat_Viewport.Status != PromptStatus.OK)
                            {
                                MessageBox.Show("this requires you to select a viewport \r\n or the polyline used to clip the viewport");
                                set_enable_true();
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }





                            Viewport Ent_vp = Trans1.GetObject(Rezultat_Viewport.ObjectId, OpenMode.ForRead) as Viewport;
                            if (Ent_vp == null)
                            {
                                Polyline Ent_poly = Trans1.GetObject(Rezultat_Viewport.ObjectId, OpenMode.ForRead) as Polyline;
                                if (Ent_poly != null)
                                {
                                    ObjectId vpId = LayoutManager.Current.GetNonRectangularViewportIdFromClipId(Rezultat_Viewport.ObjectId);
                                    if (Trans1.GetObject(vpId, OpenMode.ForRead) is Viewport)
                                    {
                                        Ent_vp = Trans1.GetObject(vpId, OpenMode.ForRead) as Viewport;
                                    }

                                }
                            }

                            if (Ent_vp != null)
                            {

                                vp_SCALE = Ent_vp.CustomScale;
                                vp_TWIST = Ent_vp.TwistAngle;
                                Ent_vp_id = Ent_vp.ObjectId;

                            }



                        }
                        Trans1.Commit();
                    }

                    for (int i = 0; i < dt_alias.Rows.Count; ++i)
                    {
                        if (dt_alias.Rows[i][0] != DBNull.Value && Convert.ToString(dt_alias.Rows[i][0]).ToUpper() == "B AND D SETTINGS")
                        {
                            if (dt_alias.Rows[i][i_round] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_alias.Rows[i][i_round])) == true)
                            {
                                round1 = Convert.ToInt32(dt_alias.Rows[i][i_round]);
                            }
                            idx_alias = i;
                            i = dt_alias.Rows.Count;


                        }

                    }

                    if (idx_alias == -1)
                    {
                        MessageBox.Show("NO <B and D settings> specified\r\nYou have to specify Mtext style and Mleader style");
                        set_enable_true();
                        return;
                    }

                    //End = 1,
                    //Middle = 2,
                    //Center = 4,
                    //Node = 8,
                    //Quadrant = 16,
                    //Intersection = 32,
                    //Insertion = 64,
                    //Perpendicular = 128,
                    //Tangent = 256,
                    //Near = 512,
                    // Quick = 1024,
                    //ApparentIntersection = 2048,
                    //Immediate = 65536,
                    //AllowTangent = 131072,
                    // DisablePerpendicular = 262144,
                    //RelativeCartesian = 524288,
                    //RelativePolar = 1048576,
                    //NoneOverride = 2097152,  


                    if (dt_alias.Rows[idx_alias][i_lay] != DBNull.Value && Convert.ToString(dt_alias.Rows[idx_alias][i_lay]).Replace(" ", "").Length > 0)
                    {
                        label_layer = Convert.ToString(dt_alias.Rows[idx_alias][i_lay]);
                    }
                    else
                    {
                        label_layer = "_Lgen_x_y_" + scale1.ToString();
                    }


                    Functions.Creaza_layer(label_layer, 10, true);





                l123:

                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        BlockTable BlockTable1 = (BlockTable)Trans1.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;



                        #region load mtext params
                        if (dt_alias.Rows[idx_alias][i_tst] != DBNull.Value)
                        {
                            table_txt_stylename = Convert.ToString(dt_alias.Rows[idx_alias][i_tst]);
                        }
                        if (dt_alias.Rows[idx_alias][i_tf] != DBNull.Value)
                        {
                            table_txt_fontname = Convert.ToString(dt_alias.Rows[idx_alias][i_tf]);
                        }
                        if (dt_alias.Rows[idx_alias][i_tw] != DBNull.Value)
                        {
                            string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_tw]);
                            if (Functions.IsNumeric(val1) == true)
                            {
                                table_txt_width = Convert.ToDouble(val1);
                            }
                        }

                        if (dt_alias.Rows[idx_alias][i_to] != DBNull.Value)
                        {
                            string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_to]);
                            if (Functions.IsNumeric(val1) == true)
                            {
                                table_txt_oblique = Convert.ToDouble(val1);
                            }
                        }

                        if (dt_alias.Rows[idx_alias][i_th] != DBNull.Value)
                        {
                            string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_th]);
                            if (Functions.IsNumeric(val1) == true)
                            {
                                textH = Convert.ToDouble(val1);
                            }
                        }

                        if (dt_alias.Rows[idx_alias][i_tu] != DBNull.Value)
                        {
                            string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_tu]);
                            if (val1.ToLower() == "yes" || val1.ToLower() == "true")
                            {
                                txt_underline = true;
                            }
                        }

                        if (dt_alias.Rows[idx_alias][i_bm] != DBNull.Value)
                        {
                            string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_bm]);
                            if (val1.ToLower() == "yes" || val1.ToLower() == "true")
                            {
                                mtxt_background_frame = true;
                            }
                        }

                        #endregion
                        #region load mleader params

                        if (dt_alias.Rows[idx_alias][i_mst] != DBNull.Value)
                        {
                            table_mstyle_name = Convert.ToString(dt_alias.Rows[idx_alias][i_mst]) + "_" + scale1.ToString();
                        }

                        if (dt_alias.Rows[idx_alias][i_ma] != DBNull.Value)
                        {
                            string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_ma]);
                            if (Functions.IsNumeric(val1) == true)
                            {
                                table_mleader_arrow = Convert.ToDouble(val1);
                            }
                        }

                        if (dt_alias.Rows[idx_alias][i_mg] != DBNull.Value)
                        {
                            string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_mg]);
                            if (Functions.IsNumeric(val1) == true)
                            {
                                table_mleader_gap = Convert.ToDouble(val1);
                            }
                        }

                        if (dt_alias.Rows[idx_alias][i_md] != DBNull.Value)
                        {
                            string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_md]);
                            if (Functions.IsNumeric(val1) == true)
                            {
                                table_mleader_doglentgh = Convert.ToDouble(val1);
                            }
                        }

                        if (dt_alias.Rows[idx_alias][i_nh] != DBNull.Value)
                        {
                            string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_nh]);
                            if (Functions.IsNumeric(val1) == true)
                            {
                                MleadertextH = Convert.ToDouble(val1);
                            }
                        }

                        if (dt_alias.Rows[idx_alias][i_txtf] != DBNull.Value)
                        {
                            string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_txtf]);
                            if (val1.ToLower() == "yes" || val1.ToLower() == "true")
                            {
                                mleader_text_frame = true;
                            }
                        }

                        #endregion

                        #region define text style

                        TextStyleTable TextStyleTable1 = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, OpenMode.ForWrite) as TextStyleTable;

                        foreach (ObjectId id1 in TextStyleTable1)
                        {
                            TextStyleTableRecord style1 = Trans1.GetObject(id1, OpenMode.ForRead) as TextStyleTableRecord;
                            if (id1 != null)
                            {
                                if (style1.Name == table_txt_stylename)
                                {
                                    style1.UpgradeOpen();
                                    style1.FileName = table_txt_fontname;
                                    style1.ObliquingAngle = table_txt_oblique * Math.PI / 180;
                                    style1.TextSize = 0;
                                    style1.XScale = table_txt_width;
                                    textstyle_id = id1;
                                }
                            }
                        }

                        if (textstyle_id == ObjectId.Null)
                        {
                            TextStyleTableRecord style1 = new TextStyleTableRecord();
                            style1.FileName = table_txt_fontname;
                            style1.Name = table_txt_stylename;
                            style1.TextSize = 0;
                            style1.ObliquingAngle = table_txt_oblique * Math.PI / 180;
                            style1.XScale = table_txt_width;
                            TextStyleTable1.Add(style1);
                            Trans1.AddNewlyCreatedDBObject(style1, true);
                            textstyle_id = style1.ObjectId;
                        }


                        #endregion

                        #region define Mleader style
                        DBDictionary MleaderTable1 = Trans1.GetObject(ThisDrawing.Database.MLeaderStyleDictionaryId, OpenMode.ForWrite) as DBDictionary;
                        foreach (DBDictionaryEntry entry1 in MleaderTable1)
                        {
                            ObjectId id1 = MleaderTable1.GetAt(entry1.Key);
                            MLeaderStyle mstyle1 = Trans1.GetObject(id1, OpenMode.ForRead) as MLeaderStyle;
                            if (id1 != null)
                            {
                                if (mstyle1.Name == table_mstyle_name)
                                {
                                    mstyle1.UpgradeOpen();
                                    mstyle1.ArrowSize = table_mleader_arrow * scale1;
                                    mstyle1.DoglegLength = table_mleader_doglentgh * scale1;
                                    mstyle1.LandingGap = table_mleader_gap * scale1;
                                    mstyle1.TextStyleId = textstyle_id;
                                    mstyle1.EnableFrameText = mleader_text_frame;
                                    mleader_id = id1;
                                }
                            }
                        }

                        if (mleader_id == ObjectId.Null)
                        {
                            MLeaderStyle mstyle1 = new MLeaderStyle();
                            mstyle1.ArrowSize = table_mleader_arrow * scale1;
                            mstyle1.EnableDogleg = true;
                            mstyle1.DoglegLength = table_mleader_doglentgh * scale1;
                            mstyle1.LandingGap = table_mleader_gap * scale1;
                            mstyle1.TextStyleId = textstyle_id;
                            mstyle1.EnableFrameText = mleader_text_frame;
                            mleader_id = mstyle1.PostMLeaderStyleToDb(ThisDrawing.Database, table_mstyle_name);
                            Trans1.AddNewlyCreatedDBObject(mstyle1, true);
                        }

                        #endregion





                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", 33);

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                        PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the first point");
                        PP1.AllowNone = false;
                        Point_res1 = Editor1.GetPoint(PP1);

                        if (Point_res1.Status != PromptStatus.OK)
                        {

                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", old_osnap);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            set_enable_true();
                            Trans1.Commit();
                            return;
                        }


                        Point3d Point1 = new Point3d();
                        Point1 = Point_res1.Value.TransformBy(curent_ucs_matrix);

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                        PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the second point");
                        PP2.AllowNone = false;
                        PP2.BasePoint = Point1;
                        PP2.UseBasePoint = true;
                        Point_res2 = Editor1.GetPoint(PP2);

                        if (Point_res2.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", old_osnap);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            set_enable_true();
                            Trans1.Commit();
                            return;
                        }
                        Point3d Point2 = new Point3d();
                        Point2 = Point_res2.Value.TransformBy(curent_ucs_matrix);








                        double x1 = Point1.X;
                        double y1 = Point1.Y;
                        double x2 = Point2.X;
                        double y2 = Point2.Y;


                        string Content1 = "";
                        double Dist1 = 0;


                        Point3d Pt11 = Point1;
                        Point3d Pt22 = Point2;


                        if (is_paperspace == true)
                        {

                            Viewport Vp1 = Trans1.GetObject(Ent_vp_id, OpenMode.ForRead) as Viewport;
                            if (Vp1 != null)
                            {
                                Matrix3d TransforMatrix = Functions.PaperToModel(Vp1);
                                Pt11 = Point1.TransformBy(TransforMatrix);
                                Pt22 = Point2.TransformBy(TransforMatrix);
                                x1 = Pt11.X;
                                y1 = Pt11.Y;
                                x2 = Pt22.X;
                                y2 = Pt22.Y;
                            }
                        }

                        Dist1 = Math.Pow(Math.Pow(x1 - x2, 2) + Math.Pow(y1 - y2, 2), 0.5);
                        double Bearing1 = Functions.GET_Bearing_rad(x1, y1, x2, y2);


                        string Quadrant1 = Functions.Get_Quadrant_bearing(Bearing1);

                        Content1 = Quadrant1 + "\\P" + Functions.Get_String_Rounded_with_thousand_sep(Dist1, round1) + "'";
                        Content1 = Content1.Replace(" ", "");

                        #region creaza new mleader
                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", 512);


                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res3;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP3;
                        PP3 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the first Mleader point");
                        PP1.AllowNone = false;
                        Point_res3 = Editor1.GetPoint(PP3);

                        if (Point_res3.Status != PromptStatus.OK)
                        {

                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", old_osnap);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            set_enable_true();
                            Trans1.Commit();
                            return;
                        }


                        Point3d Point3 = new Point3d();
                        Point3 = Point_res3.Value.TransformBy(curent_ucs_matrix);

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res4;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP4;
                        PP4 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the second Mleader point");
                        PP4.AllowNone = false;
                        PP4.BasePoint = Point3;
                        PP4.UseBasePoint = true;
                        Point_res4 = Editor1.GetPoint(PP4);

                        if (Point_res4.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", old_osnap);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            set_enable_true();
                            Trans1.Commit();
                            return;
                        }
                        Point3d Point4 = new Point3d();
                        Point4 = Point_res4.Value.TransformBy(curent_ucs_matrix);

                        double Bearingml = Functions.GET_Bearing_rad(Point3.X, Point3.Y, Point4.X, Point4.Y) * 180 / Math.PI;
                        if (Bearingml >= 360)
                        {
                            do
                            {
                                Bearingml = Bearingml - 360;
                            } while (Bearingml >= 360);
                        }


                        MText Mtext1 = new MText();
                        Mtext1.Rotation = 0;

                        if ((Bearingml >= 0 && Bearingml <= 90) || (Bearingml >= 270 && Bearingml < 360))
                        {
                            Mtext1.Attachment = AttachmentPoint.MiddleLeft;
                        }
                        else
                        {
                            Mtext1.Attachment = AttachmentPoint.MiddleRight;
                        }

                        Mtext1.Contents = Content1;
                        Mtext1.TextHeight = textH * scale1;
                        Mtext1.TextStyleId = textstyle_id;

                        if (mtxt_background_frame == true)
                        {
                            Mtext1.BackgroundFill = true;
                            Mtext1.UseBackgroundColor = true;
                            Mtext1.BackgroundScaleFactor = 1.2;
                        }



                        MLeader Mleader1 = new MLeader();
                        Mleader1.ColorIndex = 256;
                        Mleader1.MLeaderStyle = mleader_id;
                        Mleader1.Layer = label_layer;

                        int Nr1 = Mleader1.AddLeader();
                        int Nr2 = Mleader1.AddLeaderLine(Nr1);
                        Mleader1.AddFirstVertex(Nr2, Point3);
                        Mleader1.AddLastVertex(Nr2, Point4);
                        Mleader1.LeaderLineType = LeaderType.StraightLeader;

                        Mleader1.ContentType = ContentType.MTextContent;

                        Mleader1.SetTextAttachmentType(TextAttachmentType.AttachmentMiddle, LeaderDirectionType.LeftLeader);
                        Mleader1.SetTextAttachmentType(TextAttachmentType.AttachmentMiddle, LeaderDirectionType.RightLeader);
                        Mleader1.Annotative = AnnotativeStates.False;

                        Mleader1.MText = Mtext1;

                        BTrecord.AppendEntity(Mleader1);
                        Trans1.AddNewlyCreatedDBObject(Mleader1, true);

                        #endregion
                        Trans1.TransactionManager.QueueForGraphicsFlush();

                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", old_osnap);

                        Trans1.Commit();
                    }
                    goto l123;

                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", old_osnap);
        }



        private void button_label_workspace_Click(object sender, EventArgs e)
        {
            ObjectId[] Empty_array = null;
            idx_alias = -1;
            set_enable_false();
            int round1 = 1;
            old_osnap = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE");
            try
            {


                if (dt_alias == null || dt_alias.Rows.Count == 0)
                {
                    MessageBox.Show("NO layer alias loaded");
                    set_enable_true();
                    return;
                }

                string table_txt_stylename = "LGEN_STANDARD_MTEXT";
                string table_txt_fontname = "ARIAL.TTF";
                double table_txt_oblique = 0;
                double table_txt_width = 1;

                string table_mstyle_name = "lgen_mleaderstyle";
                double table_mleader_arrow = 0.08;
                double table_mleader_doglentgh = 0.08;
                double table_mleader_gap = 0.08;
                double MleadertextH = 0.08;

                ObjectId textstyle_id = ObjectId.Null;
                ObjectId mleader_id = ObjectId.Null;



                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();



                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {



                    double vp_SCALE = 1;
                    double vp_TWIST = 0;
                    ObjectId Ent_vp_id = ObjectId.Null;

                    string cb_txt = Combobox_scales.Text;
                    cb_txt = cb_txt.Replace("1' = ", "");
                    scale1 = Convert.ToDouble(cb_txt.Replace("\"", ""));


                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        if (is_paperspace == true)
                        {
                            int Tilemode1 = Convert.ToInt32(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("TILEMODE"));
                            int CVport1 = Convert.ToInt32(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CVPORT"));

                            scale1 = 1;

                            if (Tilemode1 == 0)
                            {
                                if (CVport1 != 1)
                                {
                                    Editor1.SwitchToPaperSpace();
                                }

                            }
                            else
                            {
                                MessageBox.Show("this option is meant to work in paper space\r\nPlease select Model Space from the combobox if you need the label in model space");
                                set_enable_true();
                                return;
                            }


                            Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_Viewport;
                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_viewport;
                            Prompt_viewport = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the viewport:");
                            Prompt_viewport.SetRejectMessage("\nSelect a viewport!");
                            Prompt_viewport.AllowNone = true;
                            Prompt_viewport.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Viewport), false);
                            Prompt_viewport.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);

                            Rezultat_Viewport = ThisDrawing.Editor.GetEntity(Prompt_viewport);

                            if (Rezultat_Viewport.Status != PromptStatus.OK)
                            {
                                MessageBox.Show("this requires you to select a viewport \r\n or the polyline used to clip the viewport");
                                set_enable_true();
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }





                            Viewport Ent_vp = Trans1.GetObject(Rezultat_Viewport.ObjectId, OpenMode.ForRead) as Viewport;
                            if (Ent_vp == null)
                            {
                                Polyline Ent_poly = Trans1.GetObject(Rezultat_Viewport.ObjectId, OpenMode.ForRead) as Polyline;
                                if (Ent_poly != null)
                                {
                                    ObjectId vpId = LayoutManager.Current.GetNonRectangularViewportIdFromClipId(Rezultat_Viewport.ObjectId);
                                    if (Trans1.GetObject(vpId, OpenMode.ForRead) is Viewport)
                                    {
                                        Ent_vp = Trans1.GetObject(vpId, OpenMode.ForRead) as Viewport;
                                    }

                                }
                            }

                            if (Ent_vp != null)
                            {

                                vp_SCALE = Ent_vp.CustomScale;
                                vp_TWIST = Ent_vp.TwistAngle;
                                Ent_vp_id = Ent_vp.ObjectId;

                            }



                        }
                        Trans1.Commit();
                    }




                    for (int i = 0; i < dt_alias.Rows.Count; ++i)
                    {
                        if (dt_alias.Rows[i][0] != DBNull.Value && Convert.ToString(dt_alias.Rows[i][0]).ToUpper() == "LABEL WORKSPACE SETTINGS")
                        {
                            if (dt_alias.Rows[i][i_round] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_alias.Rows[i][i_round])) == true)
                            {
                                round1 = Convert.ToInt32(dt_alias.Rows[i][i_round]);
                            }
                            idx_alias = i;
                            i = dt_alias.Rows.Count;
                        }

                    }

                    if (idx_alias == -1)
                    {
                        MessageBox.Show("NO <Label Easement settings> specified\r\nYou have to specify Mtext style");
                        set_enable_true();
                        return;
                    }

                    //End = 1,
                    //Middle = 2,
                    //Center = 4,
                    //Node = 8,
                    //Quadrant = 16,
                    //Intersection = 32,
                    //Insertion = 64,
                    //Perpendicular = 128,
                    //Tangent = 256,
                    //Near = 512,
                    // Quick = 1024,
                    //ApparentIntersection = 2048,
                    //Immediate = 65536,
                    //AllowTangent = 131072,
                    // DisablePerpendicular = 262144,
                    //RelativeCartesian = 524288,
                    //RelativePolar = 1048576,
                    //NoneOverride = 2097152,  



                    if (dt_alias.Rows[idx_alias][i_lay] != DBNull.Value && Convert.ToString(dt_alias.Rows[idx_alias][i_lay]).Replace(" ", "").Length > 0)
                    {
                        label_layer = Convert.ToString(dt_alias.Rows[idx_alias][i_lay]);
                    }
                    else
                    {
                        label_layer = "_Lgen_wspace_" + scale1.ToString();
                    }

                    Functions.Creaza_layer(label_layer, 10, true);


                    primary_labelstyle = "";

                    if (dt_alias.Rows[idx_alias][i_prim] != DBNull.Value)
                    {
                        string label_val = Convert.ToString(dt_alias.Rows[idx_alias][i_prim]);
                        if (label_val.ToLower() == "mtext")
                        {
                            primary_labelstyle = "mt";
                            radioButton1.Visible = true;
                            radioButton1.Text = "Mtext";
                        }
                        else if (label_val.ToLower() == "mleader")
                        {
                            primary_labelstyle = "ml";
                            radioButton1.Visible = true;
                            radioButton1.Text = "MLeader";
                            if (is_bdy == true)
                            {
                                panel_bdy.Visible = true;
                                radioButton4.Checked = true;
                            }
                        }
                    }

                    if (primary_labelstyle == "")
                    {
                        MessageBox.Show("work space does not have specified the label type\r\nmtext or mleader\r\noperation aborted");
                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", old_osnap);
                        this.MdiParent.WindowState = FormWindowState.Normal;
                        Editor1.SetImpliedSelection(Empty_array);
                        Editor1.WriteMessage("\nCommand:");
                        set_enable_true();
                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                        radioButton1.Text = "1";
                        radioButton2.Text = "2";
                        radioButton3.Text = "3";
                        radioButton1.Visible = false;
                        radioButton2.Visible = false;
                        radioButton3.Visible = false;
                        panel_type.Visible = false;
                        panel_bdy.Visible = false;

                        return;
                    }

                    if (dt_alias.Rows[idx_alias][i_sec] != DBNull.Value)
                    {
                        string label_val = Convert.ToString(dt_alias.Rows[idx_alias][i_sec]);
                        if (label_val.ToLower() == "mtext")
                        {

                            radioButton2.Visible = true;
                            radioButton2.Text = "Mtext";
                        }
                        else if (label_val.ToLower() == "mleader")
                        {

                            radioButton2.Visible = true;
                            radioButton2.Text = "MLeader";
                        }


                    }

                    #region label type
                    suff_layer = "";


                    if (primary_labelstyle == "mt")
                    {
                        current_labelstyle = "mt";
                        suff_layer = "_Mtext";


                    }
                    else if (primary_labelstyle == "ml")
                    {
                        current_labelstyle = "ml";
                        suff_layer = "_Mleader";

                    }

                    panel_type.Visible = true;

                    this.Refresh();


                #endregion

                l123:

                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        BlockTable BlockTable1 = (BlockTable)Trans1.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;



                        #region load mtext params
                        if (dt_alias.Rows[idx_alias][i_tst] != DBNull.Value)
                        {
                            table_txt_stylename = Convert.ToString(dt_alias.Rows[idx_alias][i_tst]);
                        }
                        if (dt_alias.Rows[idx_alias][i_tf] != DBNull.Value)
                        {
                            table_txt_fontname = Convert.ToString(dt_alias.Rows[idx_alias][i_tf]);
                        }
                        if (dt_alias.Rows[idx_alias][i_tw] != DBNull.Value)
                        {
                            string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_tw]);
                            if (Functions.IsNumeric(val1) == true)
                            {
                                table_txt_width = Convert.ToDouble(val1);
                            }
                        }

                        if (dt_alias.Rows[idx_alias][i_to] != DBNull.Value)
                        {
                            string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_to]);
                            if (Functions.IsNumeric(val1) == true)
                            {
                                table_txt_oblique = Convert.ToDouble(val1);
                            }
                        }

                        if (dt_alias.Rows[idx_alias][i_th] != DBNull.Value)
                        {
                            string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_th]);
                            if (Functions.IsNumeric(val1) == true)
                            {
                                textH = Convert.ToDouble(val1);
                            }
                        }

                        if (dt_alias.Rows[idx_alias][i_tu] != DBNull.Value)
                        {
                            string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_tu]);
                            if (val1.ToLower() == "yes" || val1.ToLower() == "true")
                            {
                                txt_underline = true;
                            }
                        }

                        if (dt_alias.Rows[idx_alias][i_bm] != DBNull.Value)
                        {
                            string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_bm]);
                            if (val1.ToLower() == "yes" || val1.ToLower() == "true")
                            {
                                mtxt_background_frame = true;
                            }
                        }

                        #endregion


                        #region load mleader params

                        if (dt_alias.Rows[idx_alias][i_mst] != DBNull.Value)
                        {
                            table_mstyle_name = Convert.ToString(dt_alias.Rows[idx_alias][i_mst]) + "_" + scale1.ToString();
                        }

                        if (dt_alias.Rows[idx_alias][i_ma] != DBNull.Value)
                        {
                            string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_ma]);
                            if (Functions.IsNumeric(val1) == true)
                            {
                                table_mleader_arrow = Convert.ToDouble(val1);
                            }
                        }

                        if (dt_alias.Rows[idx_alias][i_mg] != DBNull.Value)
                        {
                            string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_mg]);
                            if (Functions.IsNumeric(val1) == true)
                            {
                                table_mleader_gap = Convert.ToDouble(val1);
                            }
                        }

                        if (dt_alias.Rows[idx_alias][i_md] != DBNull.Value)
                        {
                            string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_md]);
                            if (Functions.IsNumeric(val1) == true)
                            {
                                table_mleader_doglentgh = Convert.ToDouble(val1);
                            }
                        }

                        if (dt_alias.Rows[idx_alias][i_nh] != DBNull.Value)
                        {
                            string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_nh]);
                            if (Functions.IsNumeric(val1) == true)
                            {
                                MleadertextH = Convert.ToDouble(val1);
                            }
                        }

                        if (dt_alias.Rows[idx_alias][i_txtf] != DBNull.Value)
                        {
                            string val1 = Convert.ToString(dt_alias.Rows[idx_alias][i_txtf]);
                            if (val1.ToLower() == "yes" || val1.ToLower() == "true")
                            {
                                mleader_text_frame = true;
                            }
                        }

                        #endregion

                        #region define text style

                        TextStyleTable TextStyleTable1 = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, OpenMode.ForWrite) as TextStyleTable;

                        foreach (ObjectId id1 in TextStyleTable1)
                        {
                            TextStyleTableRecord style1 = Trans1.GetObject(id1, OpenMode.ForRead) as TextStyleTableRecord;
                            if (id1 != null)
                            {
                                if (style1.Name == table_txt_stylename)
                                {
                                    style1.UpgradeOpen();
                                    style1.FileName = table_txt_fontname;
                                    style1.ObliquingAngle = table_txt_oblique * Math.PI / 180;
                                    style1.TextSize = 0;
                                    style1.XScale = table_txt_width;
                                    textstyle_id = id1;
                                }
                            }
                        }



                        if (textstyle_id == ObjectId.Null)
                        {
                            TextStyleTableRecord style1 = new TextStyleTableRecord();
                            style1.FileName = table_txt_fontname;
                            style1.Name = table_txt_stylename;
                            style1.TextSize = 0;
                            style1.ObliquingAngle = table_txt_oblique * Math.PI / 180;
                            style1.XScale = table_txt_width;
                            TextStyleTable1.Add(style1);
                            Trans1.AddNewlyCreatedDBObject(style1, true);
                            textstyle_id = style1.ObjectId;
                        }


                        #endregion

                        #region define Mleader style

                        DBDictionary MleaderTable1 = Trans1.GetObject(ThisDrawing.Database.MLeaderStyleDictionaryId, OpenMode.ForWrite) as DBDictionary;
                        foreach (DBDictionaryEntry entry1 in MleaderTable1)
                        {
                            ObjectId id1 = MleaderTable1.GetAt(entry1.Key);
                            MLeaderStyle mstyle1 = Trans1.GetObject(id1, OpenMode.ForRead) as MLeaderStyle;
                            if (id1 != null)
                            {
                                if (mstyle1.Name == table_mstyle_name)
                                {
                                    mstyle1.UpgradeOpen();
                                    mstyle1.ArrowSize = table_mleader_arrow * scale1;
                                    mstyle1.DoglegLength = table_mleader_doglentgh * scale1;
                                    mstyle1.LandingGap = table_mleader_gap * scale1;
                                    mstyle1.TextStyleId = textstyle_id;
                                    mstyle1.EnableFrameText = mleader_text_frame;
                                    mleader_id = id1;
                                }
                            }
                        }

                        if (mleader_id == ObjectId.Null)
                        {
                            MLeaderStyle mstyle1 = new MLeaderStyle();
                            mstyle1.ArrowSize = table_mleader_arrow * scale1;
                            mstyle1.EnableDogleg = true;
                            mstyle1.DoglegLength = table_mleader_doglentgh * scale1;
                            mstyle1.LandingGap = table_mleader_gap * scale1;
                            mstyle1.TextStyleId = textstyle_id;
                            mstyle1.EnableFrameText = mleader_text_frame;
                            mleader_id = mstyle1.PostMLeaderStyleToDb(ThisDrawing.Database, table_mstyle_name);
                            Trans1.AddNewlyCreatedDBObject(mstyle1, true);
                        }

                        #endregion


                        double d1 = -1;

                        List<Point3d> lista_puncte = new List<Point3d>();
                        List<Point3d> lista_puncte_ps = new List<Point3d>();


                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", 33);

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                        PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the first point [Length]:");
                        PP1.AllowNone = false;
                        Point_res1 = Editor1.GetPoint(PP1);

                        if (Point_res1.Status != PromptStatus.OK)
                        {

                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", old_osnap);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            set_enable_true();
                            Trans1.Commit();
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            radioButton1.Text = "1";
                            radioButton2.Text = "2";
                            radioButton3.Text = "3";
                            radioButton1.Visible = false;
                            radioButton2.Visible = false;
                            radioButton3.Visible = false;
                            panel_type.Visible = false;
                            panel_bdy.Visible = false;
                            return;
                        }

                        Point3d Point1 = Point_res1.Value.TransformBy(curent_ucs_matrix);
                        if (is_paperspace == true)
                        {

                            Viewport Vp1 = Trans1.GetObject(Ent_vp_id, OpenMode.ForRead) as Viewport;
                            if (Vp1 != null)
                            {
                                Matrix3d TransforMatrix = Functions.PaperToModel(Vp1);
                                lista_puncte.Add(Point1.TransformBy(TransforMatrix));
                                lista_puncte_ps.Add(Point1);
                            }
                        }
                        else
                        {
                            lista_puncte.Add(Point1);
                            lista_puncte_ps.Add(Point1);

                        }




                        Point3d pp = Point_res1.Value;
                        Polyline poly1 = new Polyline();
                        Polyline polyps = new Polyline();



                        bool repeta = true;
                        do
                        {
                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                            Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                            PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the next point [Length]:");
                            PP2.AllowNone = true;
                            PP2.UseBasePoint = true;
                            PP2.BasePoint = pp;
                            Point_res2 = Editor1.GetPoint(PP2);

                            if (Point_res2.Status != PromptStatus.OK)
                            {
                                if (lista_puncte.Count <= 1)
                                {
                                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", old_osnap);
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    set_enable_true();
                                    Trans1.Commit();
                                    Editor1.SetImpliedSelection(Empty_array);
                                    Editor1.WriteMessage("\nCommand:");
                                    set_enable_true();
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    radioButton1.Text = "1";
                                    radioButton2.Text = "2";
                                    radioButton3.Text = "3";
                                    radioButton1.Visible = false;
                                    radioButton2.Visible = false;
                                    radioButton3.Visible = false;
                                    panel_type.Visible = false;
                                    panel_bdy.Visible = false;
                                    return;
                                }
                                else
                                {

                                    for (int i = 0; i < lista_puncte.Count; ++i)
                                    {
                                        poly1.AddVertexAt(i, new Point2d(lista_puncte[i].X, lista_puncte[i].Y), 0, 0, 0);
                                        polyps.AddVertexAt(i, new Point2d(lista_puncte_ps[i].X, lista_puncte_ps[i].Y), 0, 0, 0);
                                    }
                                    d1 = poly1.Length;

                                    repeta = false;
                                }
                            }
                            else
                            {
                                Point3d Point2 = Point_res2.Value.TransformBy(curent_ucs_matrix);


                                if (is_paperspace == true)
                                {

                                    Viewport Vp1 = Trans1.GetObject(Ent_vp_id, OpenMode.ForRead) as Viewport;
                                    if (Vp1 != null)
                                    {
                                        Matrix3d TransforMatrix = Functions.PaperToModel(Vp1);
                                        lista_puncte.Add(Point2.TransformBy(TransforMatrix));
                                        lista_puncte_ps.Add(Point2);

                                    }
                                }
                                else
                                {
                                    lista_puncte.Add(Point2);
                                    lista_puncte_ps.Add(Point2);
                                }


                                pp = Point_res2.Value;
                            }
                        } while (repeta == true);








                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res4;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP4;
                        PP4 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify start point [width]");
                        PP4.AllowNone = false;

                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", 512);

                        Point_res4 = Editor1.GetPoint(PP4);

                        if (Point_res4.Status != PromptStatus.OK)
                        {

                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", old_osnap);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            set_enable_true();
                            Trans1.Commit();
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            radioButton1.Text = "1";
                            radioButton2.Text = "2";
                            radioButton3.Text = "3";
                            radioButton1.Visible = false;
                            radioButton2.Visible = false;
                            radioButton3.Visible = false;
                            panel_type.Visible = false;
                            panel_bdy.Visible = false;
                            return;
                        }


                        Point3d Point4 = new Point3d();
                        Point4 = Point_res4.Value.TransformBy(curent_ucs_matrix);



                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res5;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP5;
                        PP5 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify end point [width]");
                        PP5.AllowNone = false;
                        PP5.UseBasePoint = true;
                        PP5.BasePoint = Point_res4.Value;
                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", 128);
                        Point_res5 = Editor1.GetPoint(PP5);

                        if (Point_res5.Status != PromptStatus.OK)
                        {

                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", old_osnap);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            set_enable_true();
                            Trans1.Commit();
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            radioButton1.Text = "1";
                            radioButton2.Text = "2";
                            radioButton3.Text = "3";
                            radioButton1.Visible = false;
                            radioButton2.Visible = false;
                            radioButton3.Visible = false;
                            panel_type.Visible = false;
                            panel_bdy.Visible = false;
                            return;
                        }


                        Point3d Point5 = new Point3d();
                        Point5 = Point_res5.Value.TransformBy(curent_ucs_matrix);


                        double w1 = Math.Pow(Math.Pow(Point4.X - Point5.X, 2) + Math.Pow(Point4.Y - Point5.Y, 2), 0.5);

                        if (is_paperspace == true)
                        {

                            Viewport Vp1 = Trans1.GetObject(Ent_vp_id, OpenMode.ForRead) as Viewport;
                            if (Vp1 != null)
                            {
                                Matrix3d TransforMatrix = Functions.PaperToModel(Vp1);
                                w1 = Math.Pow(Math.Pow(Point4.TransformBy(TransforMatrix).X - Point5.TransformBy(TransforMatrix).X, 2) + Math.Pow(Point4.TransformBy(TransforMatrix).Y - Point5.TransformBy(TransforMatrix).Y, 2), 0.5);
                            }
                        }


                        string Content1 = Functions.Get_String_Rounded(w1, round1) + "'X" + Functions.Get_String_Rounded(d1, round1) + "'";



                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", 512);


                        stabileste_label_current();

                        string first_point_message = "\nLabel position:";
                        if (current_labelstyle == "ml")
                        {
                            first_point_message = "\nMleader first point:";
                        }

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res3;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP3;
                        PP3 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions(first_point_message);
                        PP3.AllowNone = false;
                        Point_res3 = Editor1.GetPoint(PP3);

                        if (Point_res3.Status != PromptStatus.OK)
                        {

                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", old_osnap);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            set_enable_true();
                            Trans1.Commit();
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            radioButton1.Text = "1";
                            radioButton2.Text = "2";
                            radioButton3.Text = "3";
                            radioButton1.Visible = false;
                            radioButton2.Visible = false;
                            radioButton3.Visible = false;
                            panel_type.Visible = false;
                            panel_bdy.Visible = false;
                            return;
                        }


                        Point3d Point3 = new Point3d();
                        Point3 = Point_res3.Value.TransformBy(curent_ucs_matrix);



                        Point3d point_on_poly3 = new Point3d();

                        if (is_paperspace == true)
                        {
                            point_on_poly3 = polyps.GetClosestPointTo(Point3, Vector3d.ZAxis, false);
                        }
                        else
                        {
                            point_on_poly3 = poly1.GetClosestPointTo(Point3, Vector3d.ZAxis, false);
                        }

                        double rot1 = Functions.GET_Bearing_rad(Point3.X, Point3.Y, point_on_poly3.X, point_on_poly3.Y) + Math.PI / 2;

                        if (rot1 > Math.PI / 2 && rot1 <= 3 * Math.PI / 2)
                        {
                            rot1 = rot1 + Math.PI;
                        }

                        if (current_labelstyle == "mt")
                        {
                            #region creaza new mtext

                            MText Mtext1 = new MText();
                            Mtext1.Rotation = rot1;
                            Mtext1.Attachment = AttachmentPoint.MiddleCenter;
                            Mtext1.Contents = Content1;
                            Mtext1.TextHeight = textH * scale1;
                            Mtext1.TextStyleId = textstyle_id;

                            if (mtxt_background_frame == true)
                            {
                                Mtext1.BackgroundFill = true;
                                Mtext1.UseBackgroundColor = true;
                                Mtext1.BackgroundScaleFactor = 1.2;
                            }

                            Mtext1.Location = Point3;
                            Mtext1.Layer = label_layer;


                            BTrecord.AppendEntity(Mtext1);
                            Trans1.AddNewlyCreatedDBObject(Mtext1, true);

                            #endregion
                        }

                        if (current_labelstyle == "ml")
                        {
                            #region creaza new mleader



                            PromptPointResult Point_res2;
                            PromptPointOptions PP2;
                            PP2 = new PromptPointOptions("\nMleader second point:");
                            PP2.AllowNone = false;
                            PP2.UseBasePoint = true;
                            PP2.BasePoint = Point_res3.Value;

                            Point_res2 = Editor1.GetPoint(PP2);

                            if (Point_res2.Status != PromptStatus.OK)
                            {
                                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", old_osnap);
                                this.MdiParent.WindowState = FormWindowState.Normal;
                                Trans1.Commit();
                                Editor1.SetImpliedSelection(Empty_array);
                                Editor1.WriteMessage("\nCommand:");
                                set_enable_true();
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                radioButton1.Text = "1";
                                radioButton2.Text = "2";
                                radioButton3.Text = "3";
                                radioButton1.Visible = false;
                                radioButton2.Visible = false;
                                radioButton3.Visible = false;
                                panel_type.Visible = false;
                                panel_bdy.Visible = false;
                                return;
                            }
                            Point3d pt2 = Point_res2.Value.TransformBy(curent_ucs_matrix);

                            MLeader mleader1 = new MLeader();
                            mleader1.ColorIndex = 256;
                            mleader1.MLeaderStyle = mleader_id;
                            mleader1.Layer = label_layer;

                            MText mt_ml = new MText();
                            mt_ml.TextStyleId = textstyle_id;
                            mt_ml.Contents = Content1;
                            mt_ml.Attachment = AttachmentPoint.MiddleLeft;

                            if (mtxt_background_frame == true)
                            {
                                mt_ml.BackgroundFill = true;
                                mt_ml.UseBackgroundColor = true;
                                mt_ml.BackgroundScaleFactor = 1.2;
                            }


                            mt_ml.TextHeight = MleadertextH * scale1;
                            mt_ml.Rotation = txt_rot - Functions.GET_deltaX_rad();
                            mleader1.MText = mt_ml;



                            int leaderline_index = mleader1.AddLeader();
                            int leaderpoint = mleader1.AddLeaderLine(leaderline_index);
                            mleader1.AddFirstVertex(leaderpoint, Point3);
                            mleader1.AddLastVertex(leaderpoint, pt2);
                            mleader1.LeaderLineType = LeaderType.StraightLeader;
                            mleader1.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.LeftLeader);
                            mleader1.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.RightLeader);

                            BTrecord.AppendEntity(mleader1);
                            Trans1.AddNewlyCreatedDBObject(mleader1, true);
                            #endregion
                        }

                        Trans1.TransactionManager.QueueForGraphicsFlush();

                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", old_osnap);

                        Trans1.Commit();
                    }
                    goto l123;

                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", old_osnap);
        }





        private void button_refresh_Click(object sender, EventArgs e)
        {
            if (System.IO.File.Exists(Fisier_layer_alias) == true) dt_alias = Load_existing_Lgen_layer_alias_from_excel(Fisier_layer_alias);
        }

        public  System.Data.DataTable Load_existing_Lgen_layer_alias_from_excel(string File1)
        {

            if (System.IO.File.Exists(File1) == false)
            {
                MessageBox.Show("the layer alias data file does not exist");
                return null;
            }


            System.Data.DataTable dt1 = new System.Data.DataTable();

            try
            {
                Microsoft.Office.Interop.Excel.Application Excel1 = new Microsoft.Office.Interop.Excel.Application();
                if (Excel1 == null)
                {
                    MessageBox.Show("PROBLEM WITH EXCEL!");
                    return null;
                }


                Excel1.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(File1);
                Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];

                try
                {
                    dt1 = Build_Lgen_Data_table_layer_alias_from_excel(W1, Start_row_layer_alias + 1);
                    if (dt1.Rows.Count > 0)
                    {
                        lista_layere = new List<string>();
                        for (int i = 0; i < dt1.Rows.Count; ++i)
                        {
                            if (dt1.Rows[i][0] != DBNull.Value)
                            {
                                string layer1 = Convert.ToString(dt1.Rows[i][0]);
                                if (lista_layere.Contains(layer1) == false)
                                {
                                    lista_layere.Add(layer1);
                                }
                                else
                                {
                                    MessageBox.Show("the layer " + layer1 + "already exist in layer alias\r\nlayer not added to the layer alias");
                                }
                            }
                        }
                    }

                    Workbook1.Close();
                    Excel1.Quit();

                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);

                }
                finally
                {
                    if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (Excel1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }
            return dt1;

        }

        public System.Data.DataTable Build_Lgen_Data_table_layer_alias_from_excel(Microsoft.Office.Interop.Excel.Worksheet W1, int Start_row)
        {


            System.Data.DataTable Data_table_alias = Creaza_lgen_alias_datatable_structure();
            int NrR = 0;
            int NrC = Data_table_alias.Columns.Count;

            string col1 = "A";



            for (int i = Start_row; i < 30000; ++i)
            {
                if (i == Start_row)
                {
                    if (W1.Range[col1 + i.ToString()].Value2 == null)
                    {
                        return Data_table_alias;
                    }
                }

                if (W1.Range[col1 + i.ToString()].Value2 == null)
                {
                    NrR = i - Start_row;
                    i = 31000;
                }
                else
                {
                    Data_table_alias.Rows.Add();
                }
            }


            Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[Start_row, 1], W1.Cells[NrR + Start_row - 1, NrC]];





            object[,] values = new object[NrR - 1, NrC - 1];

            values = range1.Value2;

            for (int i = 0; i < Data_table_alias.Rows.Count; ++i)
            {
                for (int j = 0; j < Data_table_alias.Columns.Count; ++j)
                {
                    object Valoare = values[i + 1, j + 1];
                    if (Valoare == null) Valoare = DBNull.Value;

                    Data_table_alias.Rows[i][j] = Valoare;
                }
            }




            return Data_table_alias;


        }

        public  System.Data.DataTable Creaza_lgen_alias_datatable_structure()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("Layer name", typeof(string));
            dt.Columns.Add("Description", typeof(string));
            dt.Columns.Add("Label Layer", typeof(string));
            dt.Columns.Add("Boundary_Layer (Yes/No)", typeof(string));
            dt.Columns.Add("Align_To_Feature (Yes/No)", typeof(string));
            dt.Columns.Add("Primary_Label_Type", typeof(string));
            dt.Columns.Add("Secondary_Label_Type", typeof(string));
            dt.Columns.Add("Tertiary_Label_Type", typeof(string));
            dt.Columns.Add("Mtext_Style_Name", typeof(string));
            dt.Columns.Add("Mtext_Style_Font", typeof(string));
            dt.Columns.Add("Mtext_Style_Width_Factor", typeof(string));
            dt.Columns.Add("Mtext_Style_Oblique_Angle", typeof(string));
            dt.Columns.Add("Mtext_Style_Height_1:1", typeof(string));
            dt.Columns.Add("Mtext_Underline (Yes/No)", typeof(string));
            dt.Columns.Add("Mleader Style Name", typeof(string));
            dt.Columns.Add("Mleader Arrow size", typeof(string));
            dt.Columns.Add("Mleader Gap", typeof(string));
            dt.Columns.Add("Mleader Dog Length", typeof(string));
            dt.Columns.Add("Mleader Text height at 1:1", typeof(string));
            dt.Columns.Add("Use_Object_Data (Yes/No)", typeof(string));
            dt.Columns.Add("Break lines (Yes/No)", typeof(string));
            dt.Columns.Add("Force_Caps (Yes/No)", typeof(string));
            dt.Columns.Add("Contour_layer (Yes/No)", typeof(string));
            dt.Columns.Add("Contour_Label_precision", typeof(string));
            dt.Columns.Add("Prefix1", typeof(string));
            dt.Columns.Add("Object Data Field1", typeof(string));
            dt.Columns.Add("Suffix1", typeof(string));
            dt.Columns.Add("Prefix2", typeof(string));
            dt.Columns.Add("Object Data Field2", typeof(string));
            dt.Columns.Add("Suffix2", typeof(string));
            dt.Columns.Add("Prefix3", typeof(string));
            dt.Columns.Add("Object Data Field3", typeof(string));
            dt.Columns.Add("Suffix3", typeof(string));
            dt.Columns.Add("Prefix4", typeof(string));
            dt.Columns.Add("Object Data Field4", typeof(string));
            dt.Columns.Add("Suffix4", typeof(string));
            dt.Columns.Add("Prefix5", typeof(string));
            dt.Columns.Add("Object Data Field5", typeof(string));
            dt.Columns.Add("Suffix5", typeof(string));
            dt.Columns.Add("Prefix6", typeof(string));
            dt.Columns.Add("Object Data Field6", typeof(string));
            dt.Columns.Add("Suffix6", typeof(string));
            dt.Columns.Add("Prefix7", typeof(string));
            dt.Columns.Add("Object Data Field7", typeof(string));
            dt.Columns.Add("Suffix7", typeof(string));
            dt.Columns.Add("Prefix8", typeof(string));
            dt.Columns.Add("Object Data Field8", typeof(string));
            dt.Columns.Add("Suffix8", typeof(string));
            dt.Columns.Add("Prefix9", typeof(string));
            dt.Columns.Add("Object Data Field9", typeof(string));
            dt.Columns.Add("Suffix9", typeof(string));
            dt.Columns.Add("Prefix10", typeof(string));
            dt.Columns.Add("Object Data Field10", typeof(string));
            dt.Columns.Add("Suffix10", typeof(string));
            dt.Columns.Add("Block Name", typeof(string));
            dt.Columns.Add("Block Attribute using concatenated description", typeof(string));
            dt.Columns.Add("Block_Attribute_Source_OD_Field_1", typeof(string));
            dt.Columns.Add("Block_Attribute_Source_OD_Field_2", typeof(string));
            dt.Columns.Add("Block_Attribute_Source_OD_Field_3", typeof(string));
            dt.Columns.Add("Block_Attribute_Source_OD_Field_4", typeof(string));
            dt.Columns.Add("Block_Attribute_Source_OD_Field_5", typeof(string));
            dt.Columns.Add("Block_Attribute_Source_OD_Field_6", typeof(string));
            dt.Columns.Add("Block_Attribute_Source_OD_Field_7", typeof(string));
            dt.Columns.Add("Block_Attribute_Source_OD_Field_8", typeof(string));
            dt.Columns.Add("Block_Attribute_Source_OD_Field_9", typeof(string));
            dt.Columns.Add("Block_Attribute_Source_OD_Field_10", typeof(string));
            dt.Columns.Add("DimStyle_name", typeof(string));
            dt.Columns.Add("DimStyle_ArrowSize", typeof(string));
            dt.Columns.Add("DimStyle_Suffix", typeof(string));
            dt.Columns.Add("DimStyle_Decimals_no", typeof(string));
            dt.Columns.Add("DimStyle_Round_to_closest", typeof(string));
            dt.Columns.Add("DimStyle_force_dimline", typeof(string));
            dt.Columns.Add("Background Mask (Yes/No)", typeof(string));
            dt.Columns.Add("Text Frame Mleaders Only (Yes/No)", typeof(string));
            dt.Columns.Add("Precision (Decimal Places)", typeof(string));
            return dt;
        }
       
    }
}
