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

namespace Alignment_mdi
{
    public partial class AGEN_band_analyze : Form
    {
        bool Freeze_operations = false;

        System.Data.DataTable dt_ownership_block_from_dwg = null;
        System.Data.DataTable dt_crossing_mtext_from_dwg = null;
        System.Data.DataTable dt_crossing_data_from_xl = null;

        System.Data.DataTable dt_err0 = null;

        int extra1 = 6;
        int row_start = 1;
        private ContextMenuStrip ContextMenuStrip_set_combo_values;
        public AGEN_band_analyze()
        {
            InitializeComponent();
            var toolStripMenuItem1 = new ToolStripMenuItem { Text = "Set Table Value" };
            toolStripMenuItem1.Click += set_combo_to_table_value_Click;

            var toolStripMenuItem2 = new ToolStripMenuItem { Text = "Set DWG Value" };
            toolStripMenuItem2.Click += set_combo_to_DWG_value_Click;


            ContextMenuStrip_set_combo_values = new ContextMenuStrip();
            ContextMenuStrip_set_combo_values.Items.AddRange(new ToolStripItem[] { toolStripMenuItem1, toolStripMenuItem2 });
        }

        private void set_combo_to_table_value_Click(object sender, EventArgs e)
        {
            foreach (Control ctrl1 in panel_err.Controls)
            {
                ComboBox combo1 = ctrl1 as ComboBox;
                if (combo1 != null)
                {
                    if (combo1.Visible == true) combo1.SelectedIndex = 1;
                }
            }
        }

        private void set_combo_to_DWG_value_Click(object sender, EventArgs e)
        {
            foreach (Control ctrl1 in panel_err.Controls)
            {
                ComboBox combo1 = ctrl1 as ComboBox;
                if (combo1 != null)
                {
                    if (combo1.Visible == true) combo1.SelectedIndex = 2;
                }
            }
        }

        private void label_correct_value_Click(object sender, EventArgs e)
        {
            Type t = e.GetType();
            if (t.Equals(typeof(MouseEventArgs)))
            {
                MouseEventArgs mouse = (MouseEventArgs)e;
                if (mouse.Button == MouseButtons.Right)
                {
                    ContextMenuStrip_set_combo_values.Show(Cursor.Position);
                    ContextMenuStrip_set_combo_values.Visible = true;
                }
            }
            else
            {
                ContextMenuStrip_set_combo_values.Visible = false;
            }
        }

        private void set_enable_false(object sender)
        {
            List<System.Windows.Forms.Button> lista_butoane = new List<Button>();
            lista_butoane.Add(button_analyze);
            lista_butoane.Add(button_output_mat_lin_to_excel);
            lista_butoane.Add(button_reconcile_data);
            lista_butoane.Add(button_z1);


            foreach (System.Windows.Forms.Button bt1 in lista_butoane)
            {
                if (sender as System.Windows.Forms.Button != bt1)
                {
                    bt1.Enabled = false;
                }
            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Button> lista_butoane = new List<Button>();
            lista_butoane.Add(button_analyze);
            lista_butoane.Add(button_output_mat_lin_to_excel);
            lista_butoane.Add(button_reconcile_data);
            lista_butoane.Add(button_z1);
            foreach (System.Windows.Forms.Button bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }


        private void button_analyze_Click(object sender, EventArgs e)
        {
            if (checkBox_ownership.Checked == false && checkBox_crossing.Checked == false)
            {
                MessageBox.Show("please select a band type!");
                return;
            }

            make_first_line_invisible();
            _AGEN_mainform.tpage_processing.Show();
            if (Freeze_operations == false)
            {
                _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;

                string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();

                if (System.IO.Directory.Exists(ProjF) == false)
                {
                    Freeze_operations = false;
                    MessageBox.Show("the project database folder does not exist");
                    _AGEN_mainform.tpage_processing.Hide();
                    return;
                }

                if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                {
                    ProjF = ProjF + "\\";
                }

                _AGEN_mainform.owner_sta1_atr = "";
                _AGEN_mainform.owner_sta2_atr = "";
                _AGEN_mainform.owner_len_atr = "";
                _AGEN_mainform.owner_linelist_atr = "";
                _AGEN_mainform.owner_owner_atr = "";

                if (System.IO.File.Exists(_AGEN_mainform.config_path) == false)
                {
                    Freeze_operations = false;
                    MessageBox.Show("no config file loaded");
                    _AGEN_mainform.tpage_processing.Hide();
                    return;
                }


                string fisier_si = ProjF + _AGEN_mainform.sheet_index_excel_name;

                if (System.IO.File.Exists(fisier_si) == false)
                {
                    Freeze_operations = false;
                    MessageBox.Show("the sheet index data file does not exist");
                    _AGEN_mainform.tpage_processing.Hide();
                    return;
                }


                _AGEN_mainform.dt_sheet_index = _AGEN_mainform.tpage_setup.Load_existing_sheet_index(fisier_si);

                if (_AGEN_mainform.dt_sheet_index == null)
                {
                    Freeze_operations = false;
                    MessageBox.Show("no sheet index data");
                    _AGEN_mainform.tpage_processing.Hide();
                    return;
                }
                if (_AGEN_mainform.dt_sheet_index.Rows.Count == 0)
                {
                    Freeze_operations = false;
                    MessageBox.Show("no sheet index data");
                    _AGEN_mainform.tpage_processing.Hide();
                    return;
                }

                string fisier_cl = ProjF + _AGEN_mainform.cl_excel_name;

                if (System.IO.File.Exists(fisier_cl) == false)
                {
                    Freeze_operations = false;
                    MessageBox.Show("the centerline data file does not exist");
                    _AGEN_mainform.tpage_processing.Hide();
                    return;
                }


                _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);
                if (_AGEN_mainform.dt_centerline == null)
                {
                    Freeze_operations = false;
                    MessageBox.Show("no centerline data");
                    _AGEN_mainform.tpage_processing.Hide();
                    return;
                }
                if (_AGEN_mainform.dt_centerline.Rows.Count == 0)
                {
                    Freeze_operations = false;
                    MessageBox.Show("no centerline data");
                    _AGEN_mainform.tpage_processing.Hide();
                    return;
                }


                if (checkBox_ownership.Checked == true || checkBox_crossing.Checked == true)
                {
                    Functions.Load_entities_records_from_config_file(_AGEN_mainform.config_path);
                }


                if (_AGEN_mainform.owner_sta1_atr == "")
                {
                    Freeze_operations = false;
                    MessageBox.Show("no block attribute information found");
                    _AGEN_mainform.tpage_processing.Hide();
                    return;
                }

                if (_AGEN_mainform.dt_config_ownership == null)
                {
                    Freeze_operations = false;
                    MessageBox.Show("no block attribute information found");
                    _AGEN_mainform.tpage_processing.Hide();
                    return;
                }

                if (_AGEN_mainform.dt_config_ownership.Rows.Count == 0)
                {
                    Freeze_operations = false;
                    MessageBox.Show("no block attribute information found");
                    _AGEN_mainform.tpage_processing.Hide();
                    return;
                }



                ObjectId[] Empty_array = null;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();



                try
                {
                    Freeze_operations = true;
                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            BlockTable BlockTable1 = Trans1.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as BlockTable;

                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                            List<string> handles_of_new_items = new List<string>();


                            dt_ownership_block_from_dwg = new System.Data.DataTable();
                            dt_ownership_block_from_dwg.Columns.Add("objectid", typeof(string));
                            dt_ownership_block_from_dwg.Columns.Add("blockname", typeof(string));
                            dt_ownership_block_from_dwg.Columns.Add("layer", typeof(string));
                            dt_ownership_block_from_dwg.Columns.Add("x", typeof(double));
                            dt_ownership_block_from_dwg.Columns.Add("y", typeof(double));
                            dt_ownership_block_from_dwg.Columns.Add("visibility", typeof(string));
                            dt_ownership_block_from_dwg.Columns.Add("stretch", typeof(double));

                            dt_crossing_mtext_from_dwg = new System.Data.DataTable();
                            dt_crossing_mtext_from_dwg.Columns.Add("objectid", typeof(string));
                            dt_crossing_mtext_from_dwg.Columns.Add("layer", typeof(string));
                            dt_crossing_mtext_from_dwg.Columns.Add("stationprefix", typeof(string));
                            dt_crossing_mtext_from_dwg.Columns.Add("station", typeof(string));
                            dt_crossing_mtext_from_dwg.Columns.Add("descriptionprefix", typeof(string));
                            dt_crossing_mtext_from_dwg.Columns.Add("description", typeof(string));
                            dt_crossing_mtext_from_dwg.Columns.Add("textheight", typeof(double));
                            dt_crossing_mtext_from_dwg.Columns.Add("rotation", typeof(double));
                            dt_crossing_mtext_from_dwg.Columns.Add("underline", typeof(bool));
                            dt_crossing_mtext_from_dwg.Columns.Add("widthfactor", typeof(double));
                            dt_crossing_mtext_from_dwg.Columns.Add("x", typeof(double));
                            dt_crossing_mtext_from_dwg.Columns.Add("y", typeof(double));
                            dt_crossing_mtext_from_dwg.Columns.Add("crossingposition", typeof(string));
                            //dt_crossing_mtext_from_dwg.TableName = "dwg";

                            dt_crossing_data_from_xl = new System.Data.DataTable();
                            dt_crossing_data_from_xl.Columns.Add("objectid", typeof(string));
                            dt_crossing_data_from_xl.Columns.Add("layer", typeof(string));
                            dt_crossing_data_from_xl.Columns.Add("stationprefix", typeof(string));
                            dt_crossing_data_from_xl.Columns.Add("station", typeof(string));
                            dt_crossing_data_from_xl.Columns.Add("descriptionprefix", typeof(string));
                            dt_crossing_data_from_xl.Columns.Add("description", typeof(string));
                            dt_crossing_data_from_xl.Columns.Add("textheight", typeof(double));
                            dt_crossing_data_from_xl.Columns.Add("rotation", typeof(double));
                            dt_crossing_data_from_xl.Columns.Add("underline", typeof(bool));
                            dt_crossing_data_from_xl.Columns.Add("widthfactor", typeof(double));
                            dt_crossing_data_from_xl.Columns.Add("x", typeof(double));
                            dt_crossing_data_from_xl.Columns.Add("y", typeof(double));
                            dt_crossing_data_from_xl.Columns.Add("crossingposition", typeof(string));
                            //dt_crossing_data_from_xl.TableName = "xls";

                            #region load blocks from drawing
                            foreach (ObjectId id1 in BTrecord)
                            {
                                Entity ent1 = Trans1.GetObject(id1, OpenMode.ForRead) as Entity;

                                #region ownership analyze
                                if (checkBox_ownership.Checked == true && ent1 is BlockReference && Functions.select_entity_segment_name(_AGEN_mainform.layer_ownership_band, "Agen_owner", ent1.ObjectId) == true)
                                {
                                    BlockReference block1 = ent1 as BlockReference;

                                    if (block1.AttributeCollection.Count > 0)
                                    {
                                        string blockname = Functions.get_block_name(block1);

                                        dt_ownership_block_from_dwg.Rows.Add();
                                        dt_ownership_block_from_dwg.Rows[dt_ownership_block_from_dwg.Rows.Count - 1][0] = block1.ObjectId.Handle.Value.ToString();
                                        dt_ownership_block_from_dwg.Rows[dt_ownership_block_from_dwg.Rows.Count - 1][1] = blockname;
                                        dt_ownership_block_from_dwg.Rows[dt_ownership_block_from_dwg.Rows.Count - 1][2] = block1.Layer;
                                        dt_ownership_block_from_dwg.Rows[dt_ownership_block_from_dwg.Rows.Count - 1][3] = block1.Position.X;
                                        dt_ownership_block_from_dwg.Rows[dt_ownership_block_from_dwg.Rows.Count - 1][4] = block1.Position.Y;

                                        Autodesk.AutoCAD.DatabaseServices.AttributeCollection attColl = block1.AttributeCollection;

                                        foreach (ObjectId odid in attColl)
                                        {
                                            AttributeReference atr1 = Trans1.GetObject(odid, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as AttributeReference;
                                            if (atr1 != null)
                                            {
                                                string Tag1 = atr1.Tag;
                                                if (dt_ownership_block_from_dwg.Columns.Contains(Tag1) == false)
                                                {
                                                    dt_ownership_block_from_dwg.Columns.Add(Tag1, typeof(string));
                                                }
                                                string val1 = atr1.TextString;
                                                if (val1 != "")
                                                {
                                                    dt_ownership_block_from_dwg.Rows[dt_ownership_block_from_dwg.Rows.Count - 1][Tag1] = val1;
                                                }
                                            }
                                        }

                                        if (block1.IsDynamicBlock == true)
                                        {
                                            using (DynamicBlockReferencePropertyCollection pc = block1.DynamicBlockReferencePropertyCollection)
                                            {
                                                foreach (DynamicBlockReferenceProperty prop in pc)
                                                {
                                                    if (prop.PropertyName == "Visibility1")
                                                    {
                                                        dt_ownership_block_from_dwg.Rows[dt_ownership_block_from_dwg.Rows.Count - 1][5] = Convert.ToString(prop.Value);
                                                    }
                                                    if (prop.PropertyName == "Distance1")
                                                    {
                                                        dt_ownership_block_from_dwg.Rows[dt_ownership_block_from_dwg.Rows.Count - 1][6] = Convert.ToDouble(prop.Value);
                                                    }
                                                }
                                            }
                                        }
                                    }

                                }
                                #endregion

                                #region crossing analyze
                                if (checkBox_crossing.Checked == true && ent1 is MText &&
                                    (Functions.select_entity_segment_name(_AGEN_mainform.layer_crossing_band_text, "Agen_crossing", ent1.ObjectId) == true ||
                                    Functions.select_entity_segment_name(_AGEN_mainform.layer_crossing_band_pi, "Agen_crossing", ent1.ObjectId) == true))
                                {
                                    MText Mtext1 = ent1 as MText;
                                    string continut = Mtext1.Contents;

                                    dt_crossing_mtext_from_dwg.Rows.Add();
                                    dt_crossing_mtext_from_dwg.Rows[dt_crossing_mtext_from_dwg.Rows.Count - 1][0] = id1.Handle.Value.ToString();
                                    dt_crossing_mtext_from_dwg.Rows[dt_crossing_mtext_from_dwg.Rows.Count - 1][1] = ent1.Layer;

                                    string prefix = Functions.extract_stationprefix_from_mtext(Mtext1.Text);
                                    if (prefix != "") dt_crossing_mtext_from_dwg.Rows[dt_crossing_mtext_from_dwg.Rows.Count - 1][2] = prefix;
                                    dt_crossing_mtext_from_dwg.Rows[dt_crossing_mtext_from_dwg.Rows.Count - 1][3] = Functions.extract_station_from_mtext(Mtext1.Text);
                                    dt_crossing_mtext_from_dwg.Rows[dt_crossing_mtext_from_dwg.Rows.Count - 1][5] = Functions.extract_description_from_mtext(Mtext1.Text);
                                    dt_crossing_mtext_from_dwg.Rows[dt_crossing_mtext_from_dwg.Rows.Count - 1][6] = Mtext1.TextHeight;
                                    dt_crossing_mtext_from_dwg.Rows[dt_crossing_mtext_from_dwg.Rows.Count - 1][7] = Mtext1.Rotation;
                                    if (Mtext1.Contents.Contains("\\L") == true)
                                    {
                                        dt_crossing_mtext_from_dwg.Rows[dt_crossing_mtext_from_dwg.Rows.Count - 1][8] = true;
                                    }
                                    else
                                    {
                                        dt_crossing_mtext_from_dwg.Rows[dt_crossing_mtext_from_dwg.Rows.Count - 1][8] = false;
                                    }
                                    dt_crossing_mtext_from_dwg.Rows[dt_crossing_mtext_from_dwg.Rows.Count - 1][9] = Functions.get_width_factor_of_an_mtext(Mtext1.Contents);
                                    dt_crossing_mtext_from_dwg.Rows[dt_crossing_mtext_from_dwg.Rows.Count - 1][10] = Mtext1.Location.X;
                                    dt_crossing_mtext_from_dwg.Rows[dt_crossing_mtext_from_dwg.Rows.Count - 1][11] = Mtext1.Location.Y;
                                }
                                #endregion


                            }
                            #endregion


                            if (dt_ownership_block_from_dwg.Rows.Count == 0 && checkBox_ownership.Checked == true)
                            {
                                Freeze_operations = false;
                                MessageBox.Show("no drawing information found");
                                _AGEN_mainform.tpage_processing.Hide();
                                return;
                            }

                            if (dt_crossing_mtext_from_dwg.Rows.Count == 0 && checkBox_crossing.Checked == true)
                            {
                                Freeze_operations = false;
                                MessageBox.Show("no drawing information found");
                                _AGEN_mainform.tpage_processing.Hide();
                                return;
                            }

                            dt_err0 = new System.Data.DataTable();
                            dt_err0.Columns.Add("handle", typeof(string));
                            dt_err0.Columns.Add("band (layer)", typeof(string));
                            dt_err0.Columns.Add("block name", typeof(string));
                            dt_err0.Columns.Add("tag name", typeof(string));
                            dt_err0.Columns.Add("dwg value", typeof(string));
                            dt_err0.Columns.Add("xl value", typeof(string));
                            dt_err0.Columns.Add("type of discrepancy", typeof(string));

                            _AGEN_mainform.Poly3D = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                            if (checkBox_ownership.Checked == true)
                            {
                                #region ownership

                                DataSet dataset1 = new DataSet();
                                dataset1.Tables.Add(dt_ownership_block_from_dwg);
                                dataset1.Tables.Add(_AGEN_mainform.dt_config_ownership);
                                DataRelation relation1 = new DataRelation("xxx", dt_ownership_block_from_dwg.Columns[0], _AGEN_mainform.dt_config_ownership.Columns[0], false);
                                dataset1.Relations.Add(relation1);
                                for (int i = 0; i < dt_ownership_block_from_dwg.Rows.Count; ++i)
                                {
                                    if (dt_ownership_block_from_dwg.Rows[i].GetChildRows(relation1).Length == 0)
                                    {
                                        string hh1 = Convert.ToString(dt_ownership_block_from_dwg.Rows[i][0]);

                                        dt_err0.Rows.Add();
                                        dt_err0.Rows[dt_err0.Rows.Count - 1][0] = hh1;
                                        dt_err0.Rows[dt_err0.Rows.Count - 1][1] = Convert.ToString(dt_ownership_block_from_dwg.Rows[i][2]);
                                        dt_err0.Rows[dt_err0.Rows.Count - 1][2] = Convert.ToString(dt_ownership_block_from_dwg.Rows[i][1]);
                                        dt_err0.Rows[dt_err0.Rows.Count - 1][4] = "New Item";
                                        dt_err0.Rows[dt_err0.Rows.Count - 1][5] = "Missing";
                                        dt_err0.Rows[dt_err0.Rows.Count - 1][6] = "Extra Basefile Features";
                                        handles_of_new_items.Add(hh1);
                                    }
                                }

                                dataset1.Relations.Remove(relation1);
                                dataset1.Tables.Remove(dt_ownership_block_from_dwg);
                                dataset1.Tables.Remove(_AGEN_mainform.dt_config_ownership);


                                dataset1 = new DataSet();
                                dataset1.Tables.Add(_AGEN_mainform.dt_config_ownership);
                                dataset1.Tables.Add(dt_ownership_block_from_dwg);
                                relation1 = new DataRelation("xxx", _AGEN_mainform.dt_config_ownership.Columns[0], dt_ownership_block_from_dwg.Columns[0], false);
                                dataset1.Relations.Add(relation1);

                                for (int i = 0; i < _AGEN_mainform.dt_config_ownership.Rows.Count; ++i)
                                {
                                    if (_AGEN_mainform.dt_config_ownership.Rows[i].GetChildRows(relation1).Length == 0)
                                    {
                                        #region find the closest block to zoom at
                                        string h1 = "";
                                        bool found = false;
                                        if (i > 0)
                                        {
                                            for (int j = i - 1; j >= 0; --j)
                                            {
                                                if (_AGEN_mainform.dt_config_ownership.Rows[j][0] != DBNull.Value)
                                                {
                                                    h1 = Convert.ToString(_AGEN_mainform.dt_config_ownership.Rows[j][0]);
                                                    try
                                                    {
                                                        ObjectId odid = Functions.GetObjectId(ThisDrawing.Database, h1);
                                                        Entity ent1 = Trans1.GetObject(odid, OpenMode.ForRead) as Entity;
                                                        if (ent1 != null)
                                                        {
                                                            j = -1;
                                                            found = true;
                                                        }
                                                    }
                                                    catch (System.Exception ex)
                                                    {

                                                    }
                                                }
                                            }
                                        }

                                        if (found == false)
                                        {
                                            for (int j = i + 1; j < _AGEN_mainform.dt_config_ownership.Rows.Count; ++j)
                                            {
                                                if (_AGEN_mainform.dt_config_ownership.Rows[j][0] != DBNull.Value)
                                                {
                                                    h1 = Convert.ToString(_AGEN_mainform.dt_config_ownership.Rows[j][0]);
                                                    try
                                                    {
                                                        ObjectId odid = Functions.GetObjectId(ThisDrawing.Database, h1);
                                                        Entity ent1 = Trans1.GetObject(odid, OpenMode.ForRead) as Entity;
                                                        if (ent1 != null)
                                                        {
                                                            j = _AGEN_mainform.dt_config_ownership.Rows.Count;
                                                        }
                                                    }
                                                    catch (System.Exception ex)
                                                    {

                                                    }
                                                }
                                            }
                                        }
                                        #endregion

                                        dt_err0.Rows.Add();
                                        if (h1 != "") dt_err0.Rows[dt_err0.Rows.Count - 1][0] = h1;
                                        dt_err0.Rows[dt_err0.Rows.Count - 1][1] = Convert.ToString(_AGEN_mainform.dt_config_ownership.Rows[i][2]);
                                        dt_err0.Rows[dt_err0.Rows.Count - 1][2] = Convert.ToString(_AGEN_mainform.dt_config_ownership.Rows[i][1]);
                                        dt_err0.Rows[dt_err0.Rows.Count - 1][4] = "Missing";
                                        dt_err0.Rows[dt_err0.Rows.Count - 1][5] = "New Item";
                                        dt_err0.Rows[dt_err0.Rows.Count - 1][6] = "Missing Basefile Features";

                                    }
                                }

                                dataset1.Relations.Remove(relation1);
                                dataset1.Tables.Remove(dt_ownership_block_from_dwg);
                                dataset1.Tables.Remove(_AGEN_mainform.dt_config_ownership);


                                string fisier_prop = ProjF + _AGEN_mainform.property_excel_name;
                                if (System.IO.File.Exists(fisier_prop) == false)
                                {
                                    Freeze_operations = false;
                                    MessageBox.Show("the property data file does not exist");
                                    _AGEN_mainform.tpage_processing.Hide();
                                    return;
                                }


                                _AGEN_mainform.Data_Table_property = _AGEN_mainform.tpage_setup.Load_existing_property(fisier_prop);
                                System.Data.DataTable dtpxl = _AGEN_mainform.Data_Table_property;
                                System.Data.DataTable dtpxl2 = new System.Data.DataTable();
                                dtpxl2 = _AGEN_mainform.Data_Table_property.Clone();
                                System.Data.DataTable dtsi = _AGEN_mainform.dt_sheet_index;

                                int index_sta1 = 1;
                                int index_sta2 = 3;

                                if (_AGEN_mainform.tpage_sheetindex.get_radioButton_use3D_stations() == true)
                                {
                                    index_sta1 = 2;
                                    index_sta2 = 4;
                                }

                                for (int i = 0; i < dtpxl.Rows.Count; ++i)
                                {
                                    if (dtpxl.Rows[i][index_sta1] != DBNull.Value && dtpxl.Rows[i][index_sta2] != DBNull.Value)
                                    {
                                        double sta1 = Convert.ToDouble(dtpxl.Rows[i][index_sta1]);
                                        double sta2 = Convert.ToDouble(dtpxl.Rows[i][index_sta2]);

                                        double eqsta1 = Functions.Station_equation_of(sta1, _AGEN_mainform.dt_station_equation);
                                        double eqsta2 = Functions.Station_equation_of(sta2, _AGEN_mainform.dt_station_equation);

                                        double Len1 = sta2 - sta1;
                                        string Owner1 = "";

                                        if (dtpxl.Rows[i][7] != DBNull.Value)
                                        {
                                            Owner1 = Convert.ToString(dtpxl.Rows[i][7]);
                                        }

                                        string Parcelid1 = "";
                                        if (dtpxl.Rows[i][8] != DBNull.Value)
                                        {
                                            Parcelid1 = Convert.ToString(dtpxl.Rows[i][8]);
                                        }

                                        dtpxl.Rows[i][5] = eqsta1;
                                        dtpxl.Rows[i][6] = eqsta2;
                                        dtpxl.Rows[i][9] = Len1;

                                        string handle_all = "";

                                        if (dtpxl.Rows[i]["BlockHandle"] != DBNull.Value)
                                        {
                                            handle_all = Convert.ToString(dtpxl.Rows[i]["BlockHandle"]);
                                        }

                                        char comma = Convert.ToChar(",");
                                        string[] handles_array = handle_all.Split(comma);
                                        int index_h = 1;

                                        for (int j = 0; j < dtsi.Rows.Count; ++j)
                                        {
                                            if (dtsi.Rows[j][3] != DBNull.Value && dtsi.Rows[j][4] != DBNull.Value)
                                            {
                                                if (Functions.IsNumeric(dtsi.Rows[j][3].ToString()) == true && Functions.IsNumeric(dtsi.Rows[j][4].ToString()) == true)
                                                {
                                                    double M1 = Convert.ToDouble(dtsi.Rows[j][3]);
                                                    double M2 = Convert.ToDouble(dtsi.Rows[j][4]);

                                                    if (sta1 >= M1 && sta2 <= M2)
                                                    {
                                                        System.Data.DataRow row1 = dtpxl.Rows[i];
                                                        dtpxl2.ImportRow(row1);
                                                        j = dtsi.Rows.Count;
                                                    }
                                                    else if (sta1 < M1 && sta2 <= M2 && sta2 > M1)
                                                    {
                                                        System.Data.DataRow row1 = dtpxl.Rows[i];
                                                        dtpxl2.ImportRow(row1);
                                                        dtpxl2.Rows[dtpxl2.Rows.Count - 1][index_sta1] = M1;
                                                        double eqm1 = Functions.Station_equation_of(M1, _AGEN_mainform.dt_station_equation);
                                                        double Len2 = sta2 - M1;
                                                        dtpxl2.Rows[dtpxl2.Rows.Count - 1][5] = eqm1;
                                                        dtpxl2.Rows[dtpxl2.Rows.Count - 1][9] = Len2;
                                                        dtpxl2.Rows[dtpxl2.Rows.Count - 1][11] = handles_array[handles_array.Length - 1];

                                                        j = dtsi.Rows.Count;
                                                    }
                                                    else if (sta1 < M2 && sta2 > M2 && sta1 > M1)
                                                    {
                                                        System.Data.DataRow row1 = dtpxl.Rows[i];
                                                        dtpxl2.ImportRow(row1);
                                                        dtpxl2.Rows[dtpxl2.Rows.Count - 1][index_sta2] = M2;
                                                        double eqm2 = Functions.Station_equation_of(M2, _AGEN_mainform.dt_station_equation);
                                                        double Len2 = M2 - sta1;
                                                        dtpxl2.Rows[dtpxl2.Rows.Count - 1][6] = eqm2;
                                                        dtpxl2.Rows[dtpxl2.Rows.Count - 1][9] = Len2;
                                                        dtpxl2.Rows[dtpxl2.Rows.Count - 1][11] = handles_array[0];
                                                    }
                                                    else if (sta1 < M1 && sta2 > M2)
                                                    {
                                                        System.Data.DataRow row1 = dtpxl.Rows[i];
                                                        dtpxl2.ImportRow(row1);
                                                        dtpxl2.Rows[dtpxl2.Rows.Count - 1][index_sta1] = M1;
                                                        dtpxl2.Rows[dtpxl2.Rows.Count - 1][index_sta2] = M2;
                                                        double eqm1 = Functions.Station_equation_of(M1, _AGEN_mainform.dt_station_equation);
                                                        double eqm2 = Functions.Station_equation_of(M2, _AGEN_mainform.dt_station_equation);
                                                        double Len2 = M2 - M1;
                                                        dtpxl2.Rows[dtpxl2.Rows.Count - 1][5] = eqm1;
                                                        dtpxl2.Rows[dtpxl2.Rows.Count - 1][6] = eqm2;
                                                        dtpxl2.Rows[dtpxl2.Rows.Count - 1][9] = Len2;
                                                        dtpxl2.Rows[dtpxl2.Rows.Count - 1][11] = handles_array[index_h];
                                                        ++index_h;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        Freeze_operations = false;
                                        MessageBox.Show("the excel data station issues on row " + (i + 1).ToString() + "\r\n" + fisier_prop);
                                        _AGEN_mainform.tpage_processing.Hide();
                                        return;
                                    }
                                }

                                dtpxl2.Columns.Add("x", typeof(double));
                                dtpxl2.Columns.Add("y", typeof(double));
                                dtpxl2.Columns.Add("visibility", typeof(string));
                                dtpxl2.Columns.Add("stretch", typeof(double));

                                DataSet dataset2 = new DataSet();
                                dataset2.Tables.Add(dtpxl2);
                                dataset2.Tables.Add(_AGEN_mainform.dt_config_ownership);

                                DataRelation relation2 = new DataRelation("xxx", dtpxl2.Columns[11], _AGEN_mainform.dt_config_ownership.Columns[0], false);
                                dataset2.Relations.Add(relation2);

                                for (int i = 0; i < dtpxl2.Rows.Count; ++i)
                                {

                                    if (dtpxl2.Rows[i].GetChildRows(relation2).Length == 1)
                                    {
                                        dtpxl2.Rows[i]["x"] = dtpxl2.Rows[i].GetChildRows(relation2)[0]["x"];
                                        dtpxl2.Rows[i]["y"] = dtpxl2.Rows[i].GetChildRows(relation2)[0]["y"];
                                        dtpxl2.Rows[i]["visibility"] = dtpxl2.Rows[i].GetChildRows(relation2)[0]["visibility"];
                                        dtpxl2.Rows[i]["stretch"] = dtpxl2.Rows[i].GetChildRows(relation2)[0]["stretch"];
                                    }
                                    else
                                    {
                                        string handle1 = Convert.ToString(dtpxl2.Rows[i][11]);
                                        if (handles_of_new_items.Contains(handle1) == false)
                                        {
                                            handles_of_new_items.Add(handle1);
                                            dt_err0.Rows.Add();
                                            dt_err0.Rows[dt_err0.Rows.Count - 1][0] = handle1;
                                            dt_err0.Rows[dt_err0.Rows.Count - 1][1] = _AGEN_mainform.layer_ownership_band;
                                            dt_err0.Rows[dt_err0.Rows.Count - 1][4] = "New Item";
                                            dt_err0.Rows[dt_err0.Rows.Count - 1][5] = "Missing";
                                            dt_err0.Rows[dt_err0.Rows.Count - 1][6] = "Extra Basefile Features";
                                        }

                                    }
                                }

                                dataset2.Relations.Remove(relation2);
                                dataset2.Tables.Remove(dtpxl2);
                                dataset2.Tables.Remove(_AGEN_mainform.dt_config_ownership);


                                dataset2.Tables.Add(_AGEN_mainform.dt_config_ownership);
                                dataset2.Tables.Add(dtpxl2);

                                relation2 = new DataRelation("xxx", _AGEN_mainform.dt_config_ownership.Columns[0], dtpxl2.Columns[11], false);
                                dataset2.Relations.Add(relation2);

                                for (int i = 0; i < _AGEN_mainform.dt_config_ownership.Rows.Count; ++i)
                                {

                                    if (_AGEN_mainform.dt_config_ownership.Rows[i].GetChildRows(relation2).Length == 0)
                                    {

                                        string handle1 = Convert.ToString(_AGEN_mainform.dt_config_ownership.Rows[i][0]);
                                        if (handles_of_new_items.Contains(handle1) == false)
                                        {
                                            handles_of_new_items.Add(handle1);
                                            dt_err0.Rows.Add();
                                            dt_err0.Rows[dt_err0.Rows.Count - 1][0] = handle1;
                                            dt_err0.Rows[dt_err0.Rows.Count - 1][1] = _AGEN_mainform.layer_ownership_band;
                                            dt_err0.Rows[dt_err0.Rows.Count - 1][4] = "New Item";
                                            dt_err0.Rows[dt_err0.Rows.Count - 1][5] = "Missing";
                                            dt_err0.Rows[dt_err0.Rows.Count - 1][6] = "Extra Basefile Features";
                                        }

                                    }
                                }

                                dataset2.Relations.Remove(relation2);
                                dataset2.Tables.Remove(dtpxl2);
                                dataset2.Tables.Remove(_AGEN_mainform.dt_config_ownership);

                                for (int i = 0; i < dtpxl2.Rows.Count; ++i)
                                {
                                    if (dtpxl2.Rows[i]["x"] != DBNull.Value && dtpxl2.Rows[i]["y"] != DBNull.Value && dtpxl2.Rows[i]["stretch"] != DBNull.Value && dtpxl2.Rows[i][11] != DBNull.Value)
                                    {
                                        string handle0 = Convert.ToString(dtpxl2.Rows[i][11]);
                                        string visib0 = Convert.ToString(dtpxl2.Rows[i]["visibility"]);
                                        double x0 = Math.Round(Convert.ToDouble(dtpxl2.Rows[i]["x"]), 3);
                                        double y0 = Math.Round(Convert.ToDouble(dtpxl2.Rows[i]["y"]), 3);
                                        double str0 = Math.Round(Convert.ToDouble(dtpxl2.Rows[i]["stretch"]), 3);

                                        for (int j = dt_ownership_block_from_dwg.Rows.Count - 1; j >= 0; --j)
                                        {
                                            string handle1 = Convert.ToString(dt_ownership_block_from_dwg.Rows[j][0]);
                                            double x1 = Math.Round(Convert.ToDouble(dt_ownership_block_from_dwg.Rows[j][3]), 3);
                                            double y1 = Math.Round(Convert.ToDouble(dt_ownership_block_from_dwg.Rows[j][4]), 3);
                                            string visib1 = Convert.ToString(dt_ownership_block_from_dwg.Rows[j][5]);
                                            double str1 = Math.Round(Convert.ToDouble(dt_ownership_block_from_dwg.Rows[j][6]), 3);

                                            if (handle0 == handle1)
                                            {
                                                if (_AGEN_mainform.owner_sta1_atr != "")
                                                {
                                                    double EqSta1 = Convert.ToDouble(dtpxl2.Rows[i][5]);
                                                    string disp_sta1 = Convert.ToString(dt_ownership_block_from_dwg.Rows[j][_AGEN_mainform.owner_sta1_atr]);
                                                    string disp_sta0 = Functions.Get_chainage_from_double(EqSta1, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);

                                                    if (disp_sta0 != disp_sta1)
                                                    {
                                                        dt_err0.Rows.Add();
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][0] = handle0;
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][1] = Convert.ToString(dt_ownership_block_from_dwg.Rows[j][2]);
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][2] = Convert.ToString(dt_ownership_block_from_dwg.Rows[j][1]);
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][3] = _AGEN_mainform.owner_sta1_atr;
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][4] = disp_sta1;
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][5] = disp_sta0;
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][6] = "Data Discrepancies";
                                                    }
                                                }

                                                if (_AGEN_mainform.owner_sta2_atr != "")
                                                {
                                                    double EqSta2 = Convert.ToDouble(dtpxl2.Rows[i][6]);
                                                    string disp_sta2 = Convert.ToString(dt_ownership_block_from_dwg.Rows[j][_AGEN_mainform.owner_sta2_atr]);
                                                    string disp_sta0 = Functions.Get_chainage_from_double(EqSta2, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);

                                                    if (disp_sta0 != disp_sta2)
                                                    {
                                                        dt_err0.Rows.Add();
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][0] = handle0;
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][1] = Convert.ToString(dt_ownership_block_from_dwg.Rows[j][2]);
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][2] = Convert.ToString(dt_ownership_block_from_dwg.Rows[j][1]);
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][3] = _AGEN_mainform.owner_sta2_atr;
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][4] = disp_sta2;
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][5] = disp_sta0;
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][6] = "Data Discrepancies";
                                                    }
                                                }


                                                if (_AGEN_mainform.owner_len_atr != "")
                                                {
                                                    double len0 = -1;
                                                    string l_string = "";
                                                    if (dtpxl2.Rows[i][9] != DBNull.Value)
                                                    {
                                                        l_string = Convert.ToString(dtpxl2.Rows[i][9]);
                                                        if (Functions.IsNumeric(l_string) == true)
                                                        {
                                                            len0 = Math.Round(Convert.ToDouble(l_string), _AGEN_mainform.round1);
                                                        }
                                                    }

                                                    string len1 = Convert.ToString(dt_ownership_block_from_dwg.Rows[j][_AGEN_mainform.owner_len_atr]);

                                                    string len2 = len1.Replace("'", "");
                                                    if (Functions.IsNumeric(len2) == true)
                                                    {
                                                        double len11 = Math.Round(Convert.ToDouble(len2), _AGEN_mainform.round1);
                                                        if (len0 != len11)
                                                        {
                                                            dt_err0.Rows.Add();
                                                            dt_err0.Rows[dt_err0.Rows.Count - 1][0] = handle0;
                                                            dt_err0.Rows[dt_err0.Rows.Count - 1][1] = Convert.ToString(dt_ownership_block_from_dwg.Rows[j][2]);
                                                            dt_err0.Rows[dt_err0.Rows.Count - 1][2] = Convert.ToString(dt_ownership_block_from_dwg.Rows[j][1]);
                                                            dt_err0.Rows[dt_err0.Rows.Count - 1][3] = _AGEN_mainform.owner_len_atr;
                                                            dt_err0.Rows[dt_err0.Rows.Count - 1][4] = len1;
                                                            dt_err0.Rows[dt_err0.Rows.Count - 1][5] = l_string;
                                                            dt_err0.Rows[dt_err0.Rows.Count - 1][6] = "Data Discrepancies";
                                                        }

                                                    }
                                                    else
                                                    {
                                                        dt_err0.Rows.Add();
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][0] = handle0;
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][1] = Convert.ToString(dt_ownership_block_from_dwg.Rows[j][2]);
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][2] = Convert.ToString(dt_ownership_block_from_dwg.Rows[j][1]);
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][3] = _AGEN_mainform.owner_len_atr;
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][4] = len1;
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][5] = l_string;
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][6] = "Data Discrepancies";
                                                    }


                                                }

                                                if (_AGEN_mainform.owner_linelist_atr != "")
                                                {
                                                    string linelist0 = "";
                                                    if (dtpxl2.Rows[i][8] != DBNull.Value)
                                                    {
                                                        linelist0 = Convert.ToString(dtpxl2.Rows[i][8]);
                                                    }

                                                    string linelist1 = Convert.ToString(dt_ownership_block_from_dwg.Rows[j][_AGEN_mainform.owner_linelist_atr]);
                                                    if (linelist1 != linelist0)
                                                    {
                                                        dt_err0.Rows.Add();
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][0] = handle0;
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][1] = Convert.ToString(dt_ownership_block_from_dwg.Rows[j][2]);
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][2] = Convert.ToString(dt_ownership_block_from_dwg.Rows[j][1]);
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][3] = _AGEN_mainform.owner_linelist_atr;
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][4] = linelist1;
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][5] = linelist0;
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][6] = "Data Discrepancies";
                                                    }
                                                }

                                                if (_AGEN_mainform.owner_owner_atr != "")
                                                {
                                                    string owner0 = "";
                                                    if (dtpxl2.Rows[i][7] != DBNull.Value)
                                                    {
                                                        owner0 = Convert.ToString(dtpxl2.Rows[i][7]);
                                                    }
                                                    string owner1 = Convert.ToString(dt_ownership_block_from_dwg.Rows[j][_AGEN_mainform.owner_owner_atr]);
                                                    if (owner1 != owner0)
                                                    {
                                                        dt_err0.Rows.Add();
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][0] = handle0;
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][1] = Convert.ToString(dt_ownership_block_from_dwg.Rows[j][2]);
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][2] = Convert.ToString(dt_ownership_block_from_dwg.Rows[j][1]);
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][3] = _AGEN_mainform.owner_owner_atr;
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][4] = owner1;
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][5] = owner0;
                                                        dt_err0.Rows[dt_err0.Rows.Count - 1][6] = "Data Discrepancies";
                                                    }
                                                }

                                                if (visib0 != visib1)
                                                {
                                                    dt_err0.Rows.Add();
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][0] = handle0;
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][1] = Convert.ToString(dt_ownership_block_from_dwg.Rows[j][2]);
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][2] = Convert.ToString(dt_ownership_block_from_dwg.Rows[j][1]);
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][3] = "Visibility";
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][4] = visib1;
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][5] = visib0;
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][6] = "Manual Drafting Adjustments";
                                                }

                                                if (x0 != x1)
                                                {
                                                    dt_err0.Rows.Add();
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][0] = handle0;
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][1] = Convert.ToString(dt_ownership_block_from_dwg.Rows[j][2]);
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][2] = Convert.ToString(dt_ownership_block_from_dwg.Rows[j][1]);
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][3] = "X Position";
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][4] = Convert.ToString(x1);
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][5] = Convert.ToString(x0);
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][6] = "Manual Drafting Adjustments";
                                                }

                                                if (y0 != y1)
                                                {
                                                    dt_err0.Rows.Add();
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][0] = handle0;
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][1] = Convert.ToString(dt_ownership_block_from_dwg.Rows[j][2]);
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][2] = Convert.ToString(dt_ownership_block_from_dwg.Rows[j][1]);
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][3] = "Y Position";
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][4] = Convert.ToString(y1);
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][5] = Convert.ToString(y0);
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][6] = "Manual Drafting Adjustments";
                                                }

                                                if (str0 != str1)
                                                {

                                                    dt_err0.Rows.Add();
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][0] = handle0;
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][1] = Convert.ToString(dt_ownership_block_from_dwg.Rows[j][2]);
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][2] = Convert.ToString(dt_ownership_block_from_dwg.Rows[j][1]);
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][3] = "Stretch";
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][4] = Convert.ToString(str1);
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][5] = Convert.ToString(str0);
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][6] = "Manual Drafting Adjustments";
                                                }
                                                dt_ownership_block_from_dwg.Rows[j].Delete();
                                                j = -1;
                                            }
                                        }
                                    }
                                }



                                #endregion
                            }


                            #region crossing
                            if (checkBox_crossing.Checked == true)
                            {
                                DataSet dataset1 = new DataSet();
                                dataset1.Tables.Add(dt_crossing_mtext_from_dwg);
                                dataset1.Tables.Add(_AGEN_mainform.dt_config_crossing);

                                DataRelation relation1 = new DataRelation("xxx", dt_crossing_mtext_from_dwg.Columns[0], _AGEN_mainform.dt_config_crossing.Columns[0], false);
                                dataset1.Relations.Add(relation1);
                                for (int i = 0; i < dt_crossing_mtext_from_dwg.Rows.Count; ++i)
                                {
                                    string idnew = Convert.ToString(dt_crossing_mtext_from_dwg.Rows[i][0]);

                                    #region extra crossings detections
                                    if (dt_crossing_mtext_from_dwg.Rows[i].GetChildRows(relation1).Length == 0)
                                    {


                                        dt_err0.Rows.Add();
                                        dt_err0.Rows[dt_err0.Rows.Count - 1][0] = idnew;
                                        dt_err0.Rows[dt_err0.Rows.Count - 1][1] = Convert.ToString(dt_crossing_mtext_from_dwg.Rows[i][1]);

                                        string prefix_sta = "";
                                        if (dt_crossing_mtext_from_dwg.Rows[i][2] != DBNull.Value)
                                        {
                                            prefix_sta = Convert.ToString(dt_crossing_mtext_from_dwg.Rows[i][2]);
                                        }

                                        string sta = "";
                                        if (dt_crossing_mtext_from_dwg.Rows[i][3] != DBNull.Value)
                                        {
                                            sta = Convert.ToString(dt_crossing_mtext_from_dwg.Rows[i][3]);
                                        }

                                        string disp_sta = sta;
                                        if (prefix_sta != "")
                                        {
                                            disp_sta = prefix_sta + " " + sta;
                                        }

                                        dt_err0.Rows[dt_err0.Rows.Count - 1][2] = disp_sta;

                                        if (dt_crossing_mtext_from_dwg.Rows[i][5] != DBNull.Value)
                                        {
                                            dt_err0.Rows[dt_err0.Rows.Count - 1][3] = Convert.ToString(dt_crossing_mtext_from_dwg.Rows[i][5]);
                                        }

                                        dt_err0.Rows[dt_err0.Rows.Count - 1][4] = "New Item";
                                        dt_err0.Rows[dt_err0.Rows.Count - 1][5] = "Missing";
                                        dt_err0.Rows[dt_err0.Rows.Count - 1][6] = "Extra Basefile Features";
                                        handles_of_new_items.Add(idnew);
                                    }
                                    #endregion

                                    if (dt_crossing_mtext_from_dwg.Rows[i].GetChildRows(relation1).Length > 0)
                                    {
                                        int j = _AGEN_mainform.dt_config_crossing.Rows.IndexOf(dt_crossing_mtext_from_dwg.Rows[i].GetChildRows(relation1)[0]);

                                        #region detection of positional changes
                                        double x1 = Convert.ToDouble(dt_crossing_mtext_from_dwg.Rows[i][10]);
                                        double y1 = Convert.ToDouble(dt_crossing_mtext_from_dwg.Rows[i][11]);

                                        double x2 = Convert.ToDouble(_AGEN_mainform.dt_config_crossing.Rows[j][12]);
                                        double y2 = Convert.ToDouble(_AGEN_mainform.dt_config_crossing.Rows[j][13]);

                                        string layer1 = Convert.ToString(dt_crossing_mtext_from_dwg.Rows[i][1]);

                                        double dist1 = new Point3d(x1, y1, 0).DistanceTo(new Point3d(x2, y2, 0));
                                        if (dist1 > 0.01)
                                        {
                                            if (Math.Round(x1, 3) != Math.Round(x2, 3))
                                            {
                                                dt_err0.Rows.Add();
                                                dt_err0.Rows[dt_err0.Rows.Count - 1][0] = idnew;
                                                dt_err0.Rows[dt_err0.Rows.Count - 1][1] = layer1;

                                                dt_err0.Rows[dt_err0.Rows.Count - 1][3] = "X Position";
                                                dt_err0.Rows[dt_err0.Rows.Count - 1][4] = Convert.ToString(x1);
                                                dt_err0.Rows[dt_err0.Rows.Count - 1][5] = Convert.ToString(x2);
                                                dt_err0.Rows[dt_err0.Rows.Count - 1][6] = "Manual Drafting Adjustments";
                                            }

                                            if (Math.Round(y1, 3) != Math.Round(y2, 3))
                                            {
                                                dt_err0.Rows.Add();
                                                dt_err0.Rows[dt_err0.Rows.Count - 1][0] = idnew;
                                                dt_err0.Rows[dt_err0.Rows.Count - 1][1] = layer1;

                                                dt_err0.Rows[dt_err0.Rows.Count - 1][3] = "Y Position";
                                                dt_err0.Rows[dt_err0.Rows.Count - 1][4] = Convert.ToString(y1);
                                                dt_err0.Rows[dt_err0.Rows.Count - 1][5] = Convert.ToString(y2);
                                                dt_err0.Rows[dt_err0.Rows.Count - 1][6] = "Manual Drafting Adjustments";
                                            }
                                        }
                                        #endregion

                                        double xcrossing = Math.Round(Convert.ToDouble(_AGEN_mainform.dt_config_crossing.Rows[j][6]), 3);
                                        double ycrossing = Math.Round(Convert.ToDouble(_AGEN_mainform.dt_config_crossing.Rows[j][7]), 3);

                                        dt_crossing_mtext_from_dwg.Rows[i][12] = Convert.ToString(xcrossing.ToString() + "-" + ycrossing.ToString());
                                    }
                                }

                                dataset1.Relations.Remove(relation1);
                                dataset1.Tables.Remove(dt_crossing_mtext_from_dwg);
                                dataset1.Tables.Remove(_AGEN_mainform.dt_config_crossing);


                                dataset1 = new DataSet();
                                dataset1.Tables.Add(_AGEN_mainform.dt_config_crossing);
                                dataset1.Tables.Add(dt_crossing_mtext_from_dwg);
                                relation1 = new DataRelation("xxx", _AGEN_mainform.dt_config_crossing.Columns[0], dt_crossing_mtext_from_dwg.Columns[0], false);
                                dataset1.Relations.Add(relation1);

                                #region missing crossings listed in the config file
                                for (int i = 0; i < _AGEN_mainform.dt_config_crossing.Rows.Count; ++i)
                                {
                                    if (_AGEN_mainform.dt_config_crossing.Rows[i].GetChildRows(relation1).Length == 0)
                                    {
                                        string h1 = "";
                                        bool found = false;
                                        if (i > 0)
                                        {
                                            for (int j = i - 1; j >= 0; --j)
                                            {
                                                if (_AGEN_mainform.dt_config_crossing.Rows[j][0] != DBNull.Value)
                                                {
                                                    h1 = Convert.ToString(_AGEN_mainform.dt_config_crossing.Rows[j][0]);
                                                    try
                                                    {
                                                        ObjectId odid = Functions.GetObjectId(ThisDrawing.Database, h1);
                                                        Entity ent1 = Trans1.GetObject(odid, OpenMode.ForRead) as Entity;
                                                        if (ent1 != null)
                                                        {
                                                            j = -1;
                                                            found = true;
                                                        }
                                                    }
                                                    catch (System.Exception ex)
                                                    {

                                                    }
                                                }
                                            }
                                        }

                                        if (found == false)
                                        {
                                            for (int j = i + 1; j < _AGEN_mainform.dt_config_crossing.Rows.Count; ++j)
                                            {
                                                if (_AGEN_mainform.dt_config_crossing.Rows[j][0] != DBNull.Value)
                                                {
                                                    h1 = Convert.ToString(_AGEN_mainform.dt_config_crossing.Rows[j][0]);
                                                    try
                                                    {
                                                        ObjectId odid = Functions.GetObjectId(ThisDrawing.Database, h1);
                                                        Entity ent1 = Trans1.GetObject(odid, OpenMode.ForRead) as Entity;
                                                        if (ent1 != null)
                                                        {
                                                            j = _AGEN_mainform.dt_config_crossing.Rows.Count;
                                                        }
                                                    }
                                                    catch (System.Exception ex)
                                                    {

                                                    }
                                                }
                                            }
                                        }

                                        dt_err0.Rows.Add();
                                        if (h1 != "") dt_err0.Rows[dt_err0.Rows.Count - 1][0] = h1;
                                        dt_err0.Rows[dt_err0.Rows.Count - 1][1] = Convert.ToString(_AGEN_mainform.dt_config_crossing.Rows[i][1]);

                                        string prefix_sta = "";
                                        if (_AGEN_mainform.dt_config_crossing.Rows[i][2] != DBNull.Value)
                                        {
                                            prefix_sta = Convert.ToString(_AGEN_mainform.dt_config_crossing.Rows[i][2]);
                                        }

                                        string sta = "";
                                        if (_AGEN_mainform.dt_config_crossing.Rows[i][3] != DBNull.Value)
                                        {
                                            sta = Convert.ToString(_AGEN_mainform.dt_config_crossing.Rows[i][3]);
                                        }

                                        string disp_sta = sta;
                                        if (prefix_sta != "")
                                        {
                                            disp_sta = prefix_sta + " " + sta;
                                        }

                                        dt_err0.Rows[dt_err0.Rows.Count - 1][2] = disp_sta;

                                        if (_AGEN_mainform.dt_config_crossing.Rows[i][5] != DBNull.Value)
                                        {
                                            dt_err0.Rows[dt_err0.Rows.Count - 1][3] = Convert.ToString(_AGEN_mainform.dt_config_crossing.Rows[i][5]);
                                        }

                                        dt_err0.Rows[dt_err0.Rows.Count - 1][4] = "Missing";
                                        dt_err0.Rows[dt_err0.Rows.Count - 1][5] = "New Item";
                                        dt_err0.Rows[dt_err0.Rows.Count - 1][6] = "Missing Basefile Features";
                                    }
                                }
                                #endregion

                                dataset1.Relations.Remove(relation1);
                                dataset1.Tables.Remove(dt_crossing_mtext_from_dwg);
                                dataset1.Tables.Remove(_AGEN_mainform.dt_config_crossing);

                                #region build dt_crossing_data_from_xl
                                string fisier_prop = ProjF + _AGEN_mainform.property_excel_name;
                                if (System.IO.File.Exists(fisier_prop) == false)
                                {
                                    Freeze_operations = false;
                                    MessageBox.Show("the property data file does not exist");
                                    _AGEN_mainform.tpage_processing.Hide();
                                    return;
                                }
                                _AGEN_mainform.Data_Table_property = _AGEN_mainform.tpage_setup.Load_existing_property(fisier_prop);
                                System.Data.DataTable dtpxl = _AGEN_mainform.Data_Table_property;

                                string fisier_crossing = ProjF + _AGEN_mainform.crossing_excel_name;
                                if (System.IO.File.Exists(fisier_crossing) == false)
                                {
                                    Freeze_operations = false;
                                    MessageBox.Show("the crossing data file does not exist");
                                    _AGEN_mainform.tpage_processing.Hide();
                                    return;
                                }

                                _AGEN_mainform.Data_Table_crossings = _AGEN_mainform.tpage_crossing_draw.Load_existing_crossing(fisier_crossing);

                                System.Data.DataTable dt_crossing = _AGEN_mainform.Data_Table_crossings.Clone();
                                if (_AGEN_mainform.Data_Table_crossings != null)
                                {
                                    if (_AGEN_mainform.Data_Table_crossings.Rows.Count > 0)
                                    {
                                        for (int i = 0; i < _AGEN_mainform.Data_Table_crossings.Rows.Count; ++i)
                                        {
                                            bool add_xing = true;

                                            if (_AGEN_mainform.Data_Table_crossings.Rows[i][12] != DBNull.Value)
                                            {
                                                if (_AGEN_mainform.Data_Table_crossings.Rows[i][12].ToString().ToUpper() == "NO") add_xing = false;
                                            }

                                            if (_AGEN_mainform.tpage_sheetindex.get_radioButton_use3D_stations() == false)
                                            {
                                                if (_AGEN_mainform.Data_Table_crossings.Rows[i][1] != DBNull.Value && _AGEN_mainform.Data_Table_crossings.Rows[i][4] != DBNull.Value && add_xing == true)
                                                {
                                                    dt_crossing.ImportRow(_AGEN_mainform.Data_Table_crossings.Rows[i]);
                                                }
                                            }

                                            else
                                            {
                                                if (_AGEN_mainform.Data_Table_crossings.Rows[i][2] != DBNull.Value && _AGEN_mainform.Data_Table_crossings.Rows[i][4] != DBNull.Value && add_xing == true)
                                                {
                                                    dt_crossing.ImportRow(_AGEN_mainform.Data_Table_crossings.Rows[i]);
                                                }
                                            }
                                        }

                                        if (_AGEN_mainform.dt_centerline != null)
                                        {
                                            if (_AGEN_mainform.dt_centerline.Rows.Count > 2)
                                            {
                                                for (int i = 1; i < _AGEN_mainform.dt_centerline.Rows.Count - 1; ++i)
                                                {
                                                    if (_AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_DeflAngDMS] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_DeflAngDMS])) == true)
                                                    {
                                                        double defl1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_DeflAngDMS]);
                                                        double min_ang = 0;
                                                        if (Functions.IsNumeric(_AGEN_mainform.tpage_crossing_draw.get_textBox_pi_min_angle()) == true)
                                                        {
                                                            min_ang = Convert.ToDouble(_AGEN_mainform.tpage_crossing_draw.get_textBox_pi_min_angle());
                                                        }

                                                        if (min_ang <= defl1)
                                                        {
                                                            dt_crossing.Rows.Add();
                                                            dt_crossing.Rows[dt_crossing.Rows.Count - 1][_AGEN_mainform.Col_Type] = _AGEN_mainform.crossing_type_pi;
                                                            dt_crossing.Rows[dt_crossing.Rows.Count - 1][_AGEN_mainform.Col_DeflAng] = _AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_DeflAng];
                                                            dt_crossing.Rows[dt_crossing.Rows.Count - 1]["Desc"] = defl1;

                                                            if (_AGEN_mainform.tpage_sheetindex.get_radioButton_use3D_stations() == false)
                                                            {
                                                                dt_crossing.Rows[dt_crossing.Rows.Count - 1][_AGEN_mainform.Col_2DSta] = _AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_2DSta];

                                                            }
                                                            else
                                                            {
                                                                dt_crossing.Rows[dt_crossing.Rows.Count - 1][_AGEN_mainform.Col_3DSta] = _AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_3DSta];
                                                            }
                                                            dt_crossing.Rows[dt_crossing.Rows.Count - 1]["X"] = _AGEN_mainform.dt_centerline.Rows[i]["X"];
                                                            dt_crossing.Rows[dt_crossing.Rows.Count - 1]["Y"] = _AGEN_mainform.dt_centerline.Rows[i]["Y"];
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        if (_AGEN_mainform.tpage_crossing_draw.get_checkBox_include_property_lines() == true)
                                        {
                                            if (_AGEN_mainform.Data_Table_property != null)
                                            {
                                                if (_AGEN_mainform.Data_Table_property.Rows.Count > 1)
                                                {
                                                    for (int i = 1; i < _AGEN_mainform.Data_Table_property.Rows.Count; ++i)
                                                    {

                                                        if (_AGEN_mainform.tpage_sheetindex.get_radioButton_use3D_stations() == false)
                                                        {
                                                            if (_AGEN_mainform.Data_Table_property.Rows[i][1] != DBNull.Value)
                                                            {
                                                                dt_crossing.Rows.Add();
                                                                dt_crossing.Rows[dt_crossing.Rows.Count - 1][1] = _AGEN_mainform.Data_Table_property.Rows[i][1];
                                                                dt_crossing.Rows[dt_crossing.Rows.Count - 1][4] = "ownership";
                                                                dt_crossing.Rows[dt_crossing.Rows.Count - 1][6] = "PROPERTY LINE";

                                                            }
                                                        }
                                                        else
                                                        {

                                                            if (_AGEN_mainform.Data_Table_property.Rows[i][1] != DBNull.Value)
                                                            {
                                                                dt_crossing.Rows.Add();
                                                                dt_crossing.Rows[dt_crossing.Rows.Count - 1][2] = _AGEN_mainform.Data_Table_property.Rows[i][2];
                                                                dt_crossing.Rows[dt_crossing.Rows.Count - 1][4] = "ownership";
                                                                dt_crossing.Rows[dt_crossing.Rows.Count - 1][6] = "PROPERTY LINE";
                                                            }

                                                        }
                                                        dt_crossing.Rows[dt_crossing.Rows.Count - 1]["X"] = _AGEN_mainform.Data_Table_property.Rows[i]["X_Beg"];
                                                        dt_crossing.Rows[dt_crossing.Rows.Count - 1]["Y"] = _AGEN_mainform.Data_Table_property.Rows[i]["Y_Beg"];
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }


                                ObjectId Txt_style_id_crossing = Functions.Get_textstyle_id(_AGEN_mainform.tpage_crossing_draw.get_comboBox_crossing_textstyle());
                                ObjectId Txt_style_id_pi_crossing = Functions.Get_textstyle_id(_AGEN_mainform.tpage_crossing_draw.get_comboBox_crossing_pi_textstyle());
                                if (Txt_style_id_crossing != null)
                                {
                                    TextStyleTableRecord TextStyle1 = Trans1.GetObject(Txt_style_id_crossing, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as TextStyleTableRecord;
                                    TextStyleTableRecord TextStyle2 = Trans1.GetObject(Txt_style_id_pi_crossing, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as TextStyleTableRecord;
                                    if (TextStyle1 != null && TextStyle2 != null)
                                    {
                                        double TextH_crossing = TextStyle1.TextSize;
                                        double TextH_pi_crossing = TextStyle2.TextSize;
                                        int nr_rand = -1;
                                        double min_dist = 2.75 * TextH_crossing;


                                        double TextR = Math.PI / 2;
                                        if (Functions.IsNumeric(_AGEN_mainform.tpage_crossing_draw.get_textBox_crossing_text_rotation()) == true)
                                        {
                                            TextR = Convert.ToDouble(_AGEN_mainform.tpage_crossing_draw.get_textBox_crossing_text_rotation()) * Math.PI / 180;
                                        }

                                        string pi_prefix = _AGEN_mainform.tpage_crossing_draw.get_textBox_pi_prefix();

                                        double Wfactor = Functions.Get_text_width_factor_from_textstyle(_AGEN_mainform.tpage_crossing_draw.get_comboBox_crossing_textstyle());

                                        for (int j = 0; j < _AGEN_mainform.dt_sheet_index.Rows.Count; ++j)
                                        {
                                            nr_rand = nr_rand + 1;

                                            if (_AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_M1] != DBNull.Value &&
                                                _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_M2] != DBNull.Value)
                                            {

                                                double M1 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_M1]);
                                                double M2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_M2]);

                                                if (M2 <= M1)
                                                {
                                                    _AGEN_mainform.tpage_processing.Hide();
                                                    Freeze_operations = false;
                                                    MessageBox.Show("End Station is smaller than Start Station on row " + (j).ToString() + "\r\n" + _AGEN_mainform.sheet_index_excel_name);
                                                    return;
                                                }

                                                if (M2 > _AGEN_mainform.Poly3D.Length)
                                                {
                                                    if (Math.Abs(M2 - _AGEN_mainform.Poly3D.Length) < 0.99)
                                                    {
                                                        M2 = _AGEN_mainform.Poly3D.Length - 0.001;
                                                    }
                                                    else
                                                    {
                                                        _AGEN_mainform.tpage_processing.Hide();
                                                        Freeze_operations = false;
                                                        MessageBox.Show("End Station is bigger than poly length on row " + (j).ToString() + "\r\n" + _AGEN_mainform.sheet_index_excel_name);
                                                        return;
                                                    }
                                                }


                                                Point3d pm1 = _AGEN_mainform.Poly3D.GetPointAtDist(M1);
                                                Point3d pm2 = _AGEN_mainform.Poly3D.GetPointAtDist(M2);

                                                double ml_len = new Point3d(pm1.X, pm1.Y, 0).DistanceTo(new Point3d(pm2.X, pm2.Y, 0));
                                                double ScaleF = ml_len / (M2 - M1);

                                                Point3d PtM1 = new Point3d(_AGEN_mainform.Point0_cross.X - ml_len * _AGEN_mainform.Vw_scale / 2,
                                                    _AGEN_mainform.Point0_cross.Y - nr_rand * _AGEN_mainform.Band_Separation - _AGEN_mainform.Vw_cross_height + 2 * TextH_crossing, 0);
                                                Point3d Prevpt = PtM1;


                                                for (int i = 0; i < dt_crossing.Rows.Count; ++i)
                                                {
                                                    double Sta3d = -1;
                                                    if (_AGEN_mainform.tpage_sheetindex.get_radioButton_use3D_stations() == false)
                                                    {
                                                        Sta3d = Convert.ToDouble(dt_crossing.Rows[i][_AGEN_mainform.Col_2DSta]);
                                                    }
                                                    else
                                                    {
                                                        Sta3d = Convert.ToDouble(dt_crossing.Rows[i][_AGEN_mainform.Col_3DSta]);
                                                    }

                                                    string Type1 = Convert.ToString(dt_crossing.Rows[i][_AGEN_mainform.Col_Type]);
                                                    string Desc1 = Convert.ToString(dt_crossing.Rows[i][_AGEN_mainform.col_desc]);

                                                    string station_prefix = "";

                                                    double xcrossing = Math.Round(Convert.ToDouble(dt_crossing.Rows[i]["X"]), 3);
                                                    double ycrossing = Math.Round(Convert.ToDouble(dt_crossing.Rows[i]["Y"]), 3);

                                                    if (Sta3d >= M1 && Sta3d <= M2)
                                                    {
                                                        if (Type1 == _AGEN_mainform.crossing_type_pi)
                                                        {
                                                            if (dt_crossing.Rows[i][_AGEN_mainform.Col_DeflAng] != DBNull.Value)
                                                            {

                                                                Point3d P1 = _AGEN_mainform.Poly3D.GetPointAtDist(Sta3d);
                                                                Line LineM1M2 = new Line(pm1, pm2);
                                                                Point3d PP1 = LineM1M2.GetClosestPointTo(P1, Vector3d.ZAxis, false);

                                                                double Deltax = pm1.DistanceTo(PP1) * _AGEN_mainform.Vw_scale;

                                                                Point3d Inspt = new Point3d(PtM1.X + Deltax, Prevpt.Y, 0);
                                                                if (Inspt.X - Prevpt.X < min_dist)
                                                                {
                                                                    Inspt = new Point3d(Prevpt.X + min_dist, Prevpt.Y, 0);
                                                                }

                                                                string sta_string = "";
                                                                if (_AGEN_mainform.tpage_crossing_draw.get_checkBox_display_station_value() == true)
                                                                {
                                                                    sta_string = Functions.Get_chainage_from_double(Functions.Station_equation_ofV2(Sta3d, _AGEN_mainform.dt_station_equation),
                                                                        _AGEN_mainform.units_of_measurement,
                                                                        _AGEN_mainform.round1) + " ";

                                                                    station_prefix = _AGEN_mainform.tpage_crossing_draw.get_textBox_station_prefix();


                                                                }

                                                                string a = "";
                                                                string c = "";
                                                                if (Wfactor == 1)
                                                                {
                                                                    if (_AGEN_mainform.tpage_crossing_draw.get_checkBox_pi_underline_value() == true)
                                                                    {
                                                                        a = "\\L{";
                                                                        c = "}";
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    if (_AGEN_mainform.tpage_crossing_draw.get_checkBox_pi_underline_value() == true)
                                                                    {
                                                                        a = "{\\W" + Wfactor.ToString() + ";\\L";
                                                                        c = "}";
                                                                    }
                                                                    else
                                                                    {
                                                                        a = "{\\W" + Wfactor.ToString() + ";";
                                                                        c = "}";
                                                                    }
                                                                }

                                                                station_prefix = Functions.remove_space_from_start_and_end_of_a_string(station_prefix);
                                                                sta_string = Functions.remove_space_from_start_and_end_of_a_string(sta_string);
                                                                pi_prefix = Functions.remove_space_from_start_and_end_of_a_string(pi_prefix);
                                                                Desc1 = Functions.remove_space_from_start_and_end_of_a_string(Desc1);

                                                                dt_crossing_data_from_xl.Rows.Add();
                                                                dt_crossing_data_from_xl.Rows[dt_crossing_data_from_xl.Rows.Count - 1][1] = _AGEN_mainform.layer_crossing_band_pi;
                                                                if (station_prefix != "") dt_crossing_data_from_xl.Rows[dt_crossing_data_from_xl.Rows.Count - 1][2] = station_prefix;
                                                                if (sta_string != "") dt_crossing_data_from_xl.Rows[dt_crossing_data_from_xl.Rows.Count - 1][3] = sta_string;
                                                                if (Desc1 != "")
                                                                {
                                                                    if (pi_prefix != "")
                                                                    {
                                                                        dt_crossing_data_from_xl.Rows[dt_crossing_data_from_xl.Rows.Count - 1][5] = pi_prefix + " " + Desc1;
                                                                    }
                                                                    else
                                                                    {
                                                                        dt_crossing_data_from_xl.Rows[dt_crossing_data_from_xl.Rows.Count - 1][5] = Desc1;
                                                                    }
                                                                }
                                                                dt_crossing_data_from_xl.Rows[dt_crossing_data_from_xl.Rows.Count - 1][6] = TextH_pi_crossing;
                                                                dt_crossing_data_from_xl.Rows[dt_crossing_data_from_xl.Rows.Count - 1][7] = TextR;
                                                                if (_AGEN_mainform.tpage_crossing_draw.get_checkBox_pi_underline_value() == true)
                                                                {
                                                                    dt_crossing_data_from_xl.Rows[dt_crossing_data_from_xl.Rows.Count - 1][8] = true;
                                                                }
                                                                else
                                                                {
                                                                    dt_crossing_data_from_xl.Rows[dt_crossing_data_from_xl.Rows.Count - 1][8] = false;
                                                                }

                                                                dt_crossing_data_from_xl.Rows[dt_crossing_data_from_xl.Rows.Count - 1][9] = Wfactor;
                                                                dt_crossing_data_from_xl.Rows[dt_crossing_data_from_xl.Rows.Count - 1][10] = Inspt.X;
                                                                dt_crossing_data_from_xl.Rows[dt_crossing_data_from_xl.Rows.Count - 1][11] = Inspt.Y;
                                                                dt_crossing_data_from_xl.Rows[dt_crossing_data_from_xl.Rows.Count - 1][12] = xcrossing.ToString() + "-" + ycrossing.ToString();

                                                                Prevpt = Inspt;
                                                            }
                                                        }
                                                        else
                                                        {
                                                            Point3d P1 = _AGEN_mainform.Poly3D.GetPointAtDist(Sta3d);
                                                            Line LineM1M2 = new Line(pm1, pm2);
                                                            Point3d PP1 = LineM1M2.GetClosestPointTo(P1, Vector3d.ZAxis, false);

                                                            double Deltax = pm1.DistanceTo(PP1) * _AGEN_mainform.Vw_scale;

                                                            Point3d Inspt = new Point3d(PtM1.X + Deltax, Prevpt.Y, 0);
                                                            if (Inspt.X - Prevpt.X < min_dist)
                                                            {
                                                                Inspt = new Point3d(Prevpt.X + min_dist, Prevpt.Y, 0);
                                                            }

                                                            string sta_string = "";
                                                            if (_AGEN_mainform.tpage_crossing_draw.get_checkBox_display_station_value() == true)
                                                            {
                                                                sta_string = Functions.Get_chainage_from_double(Functions.Station_equation_ofV2(Sta3d, _AGEN_mainform.dt_station_equation),
                                                                    _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1) + " ";

                                                                station_prefix = _AGEN_mainform.tpage_crossing_draw.get_textBox_station_prefix();
                                                            }

                                                            sta_string = Functions.remove_space_from_start_and_end_of_a_string(sta_string);
                                                            station_prefix = Functions.remove_space_from_start_and_end_of_a_string(station_prefix);
                                                            Desc1 = Functions.remove_space_from_start_and_end_of_a_string(Desc1);


                                                            dt_crossing_data_from_xl.Rows.Add();
                                                            dt_crossing_data_from_xl.Rows[dt_crossing_data_from_xl.Rows.Count - 1][1] = _AGEN_mainform.layer_crossing_band_pi;
                                                            if (station_prefix != "") dt_crossing_data_from_xl.Rows[dt_crossing_data_from_xl.Rows.Count - 1][2] = station_prefix;
                                                            if (sta_string != "") dt_crossing_data_from_xl.Rows[dt_crossing_data_from_xl.Rows.Count - 1][3] = sta_string;
                                                            if (Desc1 != "")
                                                            {
                                                                dt_crossing_data_from_xl.Rows[dt_crossing_data_from_xl.Rows.Count - 1][5] = Desc1;
                                                            }
                                                            dt_crossing_data_from_xl.Rows[dt_crossing_data_from_xl.Rows.Count - 1][6] = TextH_crossing;
                                                            dt_crossing_data_from_xl.Rows[dt_crossing_data_from_xl.Rows.Count - 1][7] = TextR;
                                                            dt_crossing_data_from_xl.Rows[dt_crossing_data_from_xl.Rows.Count - 1][8] = false;

                                                            dt_crossing_data_from_xl.Rows[dt_crossing_data_from_xl.Rows.Count - 1][9] = Wfactor;
                                                            dt_crossing_data_from_xl.Rows[dt_crossing_data_from_xl.Rows.Count - 1][10] = Inspt.X;
                                                            dt_crossing_data_from_xl.Rows[dt_crossing_data_from_xl.Rows.Count - 1][11] = Inspt.Y;
                                                            dt_crossing_data_from_xl.Rows[dt_crossing_data_from_xl.Rows.Count - 1][12] = xcrossing.ToString() + "-" + ycrossing.ToString();
                                                            Prevpt = Inspt;
                                                        }
                                                    }
                                                }
                                            }
                                        }


                                    }
                                }
                                #endregion

                                if (dt_crossing_data_from_xl != null && dt_crossing_data_from_xl.Rows.Count > 0)
                                {
                                    DataSet dataset2 = new DataSet();
                                    dataset2.Tables.Add(dt_crossing_mtext_from_dwg);
                                    dataset2.Tables.Add(dt_crossing_data_from_xl);

                                    DataRelation relation2 = new DataRelation("xyx", dt_crossing_mtext_from_dwg.Columns[12], dt_crossing_data_from_xl.Columns[12], false);
                                    dataset2.Relations.Add(relation2);
                                    DataRelation relation3 = new DataRelation("xnx", dt_crossing_data_from_xl.Columns[12], dt_crossing_mtext_from_dwg.Columns[12], false);
                                    dataset2.Relations.Add(relation3);

                                    for (int i = 0; i < dt_crossing_mtext_from_dwg.Rows.Count; ++i)
                                    {
                                        string handle1 = Convert.ToString(dt_crossing_mtext_from_dwg.Rows[i][0]);

                                        #region content check between excel crossing xls and the dwg
                                        if (dt_crossing_mtext_from_dwg.Rows[i].GetChildRows(relation2).Length > 0)
                                        {

                                            string Sta_xl = "";
                                            string Sta_dwg = "";
                                            string Descr_dwg = "";
                                            string Descr_xl = "";
                                            bool is_found = false;
                                            int j = -1;

                                            string new_line_xl = "\\P";
                                            string new_line_xl1 = "\n";
                                            string new_line_dwg = "\r\n";

                                            for (int k = 0; k < dt_crossing_mtext_from_dwg.Rows[i].GetChildRows(relation2).Length; ++k)
                                            {
                                                j = dt_crossing_data_from_xl.Rows.IndexOf(dt_crossing_mtext_from_dwg.Rows[i].GetChildRows(relation2)[k]);

                                                Descr_xl = Convert.ToString(dt_crossing_data_from_xl.Rows[j][5]).Replace(new_line_xl, "");
                                                Descr_xl = Descr_xl.Replace(new_line_xl1, "");

                                                string Station_prefix_xl = "";
                                                if (dt_crossing_data_from_xl.Rows[j][2] != DBNull.Value)
                                                {
                                                    Station_prefix_xl = Convert.ToString(dt_crossing_data_from_xl.Rows[j][2]);
                                                }
                                                string Station_xl = "";
                                                if (dt_crossing_data_from_xl.Rows[j][3] != DBNull.Value)
                                                {
                                                    Station_xl = Convert.ToString(dt_crossing_data_from_xl.Rows[j][3]);
                                                }
                                                Descr_dwg = Convert.ToString(dt_crossing_mtext_from_dwg.Rows[i][5]).Replace(new_line_dwg, "");
                                                string Station_prefix_dwg = "";
                                                if (dt_crossing_mtext_from_dwg.Rows[i][2] != DBNull.Value)
                                                {
                                                    Station_prefix_dwg = Convert.ToString(dt_crossing_mtext_from_dwg.Rows[i][2]);
                                                }
                                                string Station_dwg = "";
                                                if (dt_crossing_mtext_from_dwg.Rows[i][3] != DBNull.Value)
                                                {
                                                    Station_dwg = Convert.ToString(dt_crossing_mtext_from_dwg.Rows[i][3]);
                                                }
                                                Sta_xl = Station_xl;
                                                if (Station_prefix_xl != "")
                                                {
                                                    Sta_xl = Station_prefix_xl + " " + Station_xl;
                                                }
                                                Sta_dwg = Station_dwg;
                                                if (Station_prefix_dwg != "")
                                                {
                                                    Sta_dwg = Station_prefix_dwg + " " + Station_dwg;
                                                }


                                                if (Descr_xl == Descr_dwg && Sta_xl == Sta_dwg)
                                                {
                                                    is_found = true;
                                                    k = dt_crossing_mtext_from_dwg.Rows[i].GetChildRows(relation2).Length;
                                                }

                                            }

                                            if (is_found == false)
                                            {
                                                dt_err0.Rows.Add();
                                                dt_err0.Rows[dt_err0.Rows.Count - 1][0] = handle1;
                                                dt_err0.Rows[dt_err0.Rows.Count - 1][1] = Convert.ToString(dt_crossing_mtext_from_dwg.Rows[j][1]);
                                                dt_err0.Rows[dt_err0.Rows.Count - 1][3] = "Content";
                                                dt_err0.Rows[dt_err0.Rows.Count - 1][4] = Sta_dwg + " " + Descr_dwg;
                                                dt_err0.Rows[dt_err0.Rows.Count - 1][5] = Sta_xl + " " + Descr_xl;
                                                dt_err0.Rows[dt_err0.Rows.Count - 1][6] = "Data Discrepancies";
                                            }
                                        }
                                        #endregion

                                        //here i have to see why at pi's i have a children count of 0
                                        #region extra crossings detected
                                        if (dt_crossing_mtext_from_dwg.Rows[i].GetChildRows(relation2).Length == 0 && Convert.ToString(dt_crossing_mtext_from_dwg.Rows[i][1]) == _AGEN_mainform.layer_crossing_band_text)
                                        {
                                            bool adauga_in_dt_err = true;
                                            if (dt_err0.Rows.Count > 0)
                                            {
                                                for (int n = 0; n < dt_crossing_mtext_from_dwg.Rows.Count; ++n)
                                                {
                                                    if (Convert.ToString(dt_err0.Rows[n][0]) == handle1 && Convert.ToString(dt_err0.Rows[n][6]) == "Extra Basefile Features")
                                                    {
                                                        adauga_in_dt_err = false;
                                                        n = dt_crossing_mtext_from_dwg.Rows.Count;
                                                    }
                                                }
                                            }
                                            if (adauga_in_dt_err == true)
                                            {
                                                dt_err0.Rows.Add();
                                                dt_err0.Rows[dt_err0.Rows.Count - 1][0] = handle1;
                                                dt_err0.Rows[dt_err0.Rows.Count - 1][1] = Convert.ToString(dt_crossing_mtext_from_dwg.Rows[i][1]);

                                                string prefix_sta = "";
                                                if (dt_crossing_mtext_from_dwg.Rows[i][2] != DBNull.Value)
                                                {
                                                    prefix_sta = Convert.ToString(dt_crossing_mtext_from_dwg.Rows[i][2]);
                                                }

                                                string sta = "";
                                                if (dt_crossing_mtext_from_dwg.Rows[i][3] != DBNull.Value)
                                                {
                                                    sta = Convert.ToString(dt_crossing_mtext_from_dwg.Rows[i][3]);
                                                }

                                                string disp_sta = sta;
                                                if (prefix_sta != "")
                                                {
                                                    disp_sta = prefix_sta + " " + sta;
                                                }

                                                if (dt_crossing_mtext_from_dwg.Rows[i][5] != DBNull.Value)
                                                {
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][3] = disp_sta + " " + Convert.ToString(dt_crossing_mtext_from_dwg.Rows[i][5]);
                                                }

                                                dt_err0.Rows[dt_err0.Rows.Count - 1][4] = "New Item";
                                                dt_err0.Rows[dt_err0.Rows.Count - 1][5] = "Missing";
                                                dt_err0.Rows[dt_err0.Rows.Count - 1][6] = "Extra Basefile Features";
                                                handles_of_new_items.Add(handle1);
                                            }

                                        }
                                        #endregion
                                    }

                                    #region missing crossings listed in the excel file
                                    for (int i = 0; i < dt_crossing_data_from_xl.Rows.Count; ++i)
                                    {
                                        if (dt_crossing_data_from_xl.Rows[i].GetChildRows(relation3).Length == 0)
                                        {
                                            string layer1 = Convert.ToString(dt_crossing_data_from_xl.Rows[i][1]);
                                            if (layer1 == _AGEN_mainform.layer_crossing_band_text)
                                            {
                                                dt_err0.Rows.Add();

                                                dt_err0.Rows[dt_err0.Rows.Count - 1][1] = layer1;

                                                if (dt_crossing_data_from_xl.Rows[i][5] != DBNull.Value)
                                                {
                                                    dt_err0.Rows[dt_err0.Rows.Count - 1][3] = "Content";
                                                }

                                                string sta = "";
                                                if (dt_crossing_data_from_xl.Rows[i][3] != DBNull.Value)
                                                {
                                                    sta = Convert.ToString(dt_crossing_data_from_xl.Rows[i][3]);
                                                }
                                                dt_err0.Rows[dt_err0.Rows.Count - 1][4] = "Missing";
                                                dt_err0.Rows[dt_err0.Rows.Count - 1][5] = sta + " " + Convert.ToString(dt_crossing_data_from_xl.Rows[i][5]);
                                                dt_err0.Rows[dt_err0.Rows.Count - 1][6] = "Missing Basefile Features";
                                            }
                                        }
                                    }
                                    #endregion

                                    dataset2.Relations.Remove(relation2);
                                    dataset2.Relations.Remove(relation3);
                                    dataset2.Tables.Remove(dt_crossing_data_from_xl);
                                    dataset2.Tables.Remove(dt_crossing_mtext_from_dwg);
                                }
                            }
                            #endregion

                            calculate_total_errors();
                            transfer_data_to_panel(dt_err0);

                            //Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_block_dwg_data);
                            _AGEN_mainform.Poly3D.Erase();
                            Trans1.Commit();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                Editor1.SetImpliedSelection(Empty_array);
                Editor1.WriteMessage("\nCommand:");
                Freeze_operations = false;
                _AGEN_mainform.tpage_processing.Hide();
            }

        }

        private void calculate_total_errors()
        {
            int d_err = 0;
            int g_err = 0;
            int ndwg_err = 0;
            int nxl_err = 0;

            if (dt_err0 != null)
            {
                if (dt_err0.Rows.Count > 0)
                {
                    for (int i = 0; i < dt_err0.Rows.Count; ++i)
                    {
                        string val = Convert.ToString(dt_err0.Rows[i][6]);
                        switch (val)
                        {
                            case "Data Discrepancies":
                                ++d_err;
                                break;
                            case "Manual Drafting Adjustments":
                                ++g_err;
                                break;
                            case "Extra Basefile Features":
                                ++ndwg_err;
                                break;
                            case "Missing Basefile Features":
                                ++nxl_err;
                                break;
                            default:

                                break;

                        }
                    }

                }
            }

            textBox_data_changes.Text = d_err.ToString();
            textBox_geom_changes.Text = g_err.ToString();
            textBox_new_dwg_features.Text = ndwg_err.ToString();
            textBox_new_data_table_features.Text = nxl_err.ToString();

        }



        private void transfer_data_to_panel(System.Data.DataTable dt1)
        {
            if (dt1.Rows.Count > 0)
            {
                textBox1.Visible = true;
                textBox2.Visible = true;
                textBox3.Visible = true;
                textBox4.Visible = true;
                textBox5.Visible = true;
                comboBox1.Visible = true;
                button_z1.Visible = true;

                for (int i = panel_err.Controls.Count - 1; i >= 0; --i)
                {
                    Control ctrl1 = panel_err.Controls[i] as Control;
                    if (ctrl1.Location.Y > textBox1.Location.Y + extra1)
                    {
                        panel_err.Controls.Remove(ctrl1);
                        ctrl1.Dispose();
                    }
                }

                string text1 = "";
                string text2 = "";
                string text3 = "";
                string text4 = "";
                string text5 = "";

                if (dt1.Rows[0][1] != DBNull.Value)
                {
                    text1 = Convert.ToString(dt1.Rows[0][1]);
                }
                if (dt1.Rows[0][3] != DBNull.Value)
                {
                    text2 = Convert.ToString(dt1.Rows[0][3]);
                }
                if (dt1.Rows[0][4] != DBNull.Value)
                {
                    text3 = Convert.ToString(dt1.Rows[0][4]);
                }
                if (dt1.Rows[0][5] != DBNull.Value)
                {
                    text4 = Convert.ToString(dt1.Rows[0][5]);
                }
                if (dt1.Rows[0][6] != DBNull.Value)
                {
                    text5 = Convert.ToString(dt1.Rows[0][6]);
                }

                if (text1 == _AGEN_mainform.layer_ownership_band) text1 = "Ownership";
                if (text1 == _AGEN_mainform.layer_crossing_band_text) text1 = "Crossing";
                if (text1 == _AGEN_mainform.layer_crossing_band_pi) text1 = "Crossing";

                textBox1.Text = text1;
                textBox2.Text = text2;
                textBox3.Text = text3;
                textBox4.Text = text4;
                textBox5.Text = text5;

                if (dt1.Rows.Count > 1)
                {
                    for (int i = 1; i < dt1.Rows.Count; ++i)
                    {
                        string text11 = "";
                        string text21 = "";
                        string text31 = "";
                        string text41 = "";
                        string text51 = "";

                        if (dt1.Rows[i][1] != DBNull.Value)
                        {
                            text11 = Convert.ToString(dt1.Rows[i][1]);
                        }
                        if (dt1.Rows[i][3] != DBNull.Value)
                        {
                            text21 = Convert.ToString(dt1.Rows[i][3]);
                        }
                        if (dt1.Rows[i][4] != DBNull.Value)
                        {
                            text31 = Convert.ToString(dt1.Rows[i][4]);
                        }
                        if (dt1.Rows[i][5] != DBNull.Value)
                        {
                            text41 = Convert.ToString(dt1.Rows[i][5]);
                        }
                        if (dt1.Rows[i][6] != DBNull.Value)
                        {
                            text51 = Convert.ToString(dt1.Rows[i][6]);
                        }

                        if (text11 == _AGEN_mainform.layer_ownership_band) text11 = "Ownership";
                        if (text11 == _AGEN_mainform.layer_crossing_band_text) text11 = "Crossing";
                        if (text11 == _AGEN_mainform.layer_crossing_band_pi) text11 = "Crossing";

                        ComboBox combo1 = new ComboBox();
                        combo1.Location = new Point(comboBox1.Location.X, comboBox1.Location.Y + i * (comboBox1.Height + extra1));
                        combo1.BackColor = comboBox1.BackColor;
                        combo1.ForeColor = comboBox1.ForeColor;
                        combo1.Font = comboBox1.Font;
                        combo1.Size = comboBox1.Size;
                        combo1.FlatStyle = comboBox1.FlatStyle;
                        if (comboBox1.Items.Count > 0)
                        {
                            foreach (string item in comboBox1.Items)
                            {
                                combo1.Items.Add(item);
                            }
                        }
                        panel_err.Controls.Add(combo1);

                        Button bt1 = new Button();
                        bt1.Location = new Point(button_z1.Location.X, button_z1.Location.Y + i * (button_z1.Height + extra1));
                        bt1.BackColor = button_z1.BackColor;
                        bt1.ForeColor = button_z1.ForeColor;
                        bt1.Font = button_z1.Font;
                        bt1.Size = button_z1.Size;
                        bt1.FlatStyle = button_z1.FlatStyle;
                        bt1.FlatAppearance.BorderColor = button_z1.FlatAppearance.BorderColor;
                        bt1.FlatAppearance.BorderSize = button_z1.FlatAppearance.BorderSize;
                        bt1.FlatAppearance.MouseDownBackColor = button_z1.FlatAppearance.MouseDownBackColor;
                        bt1.FlatAppearance.MouseOverBackColor = button_z1.FlatAppearance.MouseOverBackColor;
                        bt1.BackgroundImage = button_z1.BackgroundImage;
                        bt1.BackgroundImageLayout = button_z1.BackgroundImageLayout;
                        panel_err.Controls.Add(bt1);

                        bt1.Click += delegate (object s, EventArgs e1)
                        {
                            button_zoom_click(bt1, e1);
                        };

                        TextBox tb1 = new TextBox();
                        tb1.Location = new Point(textBox1.Location.X, textBox1.Location.Y + i * (textBox1.Height + extra1));
                        tb1.BackColor = textBox1.BackColor;
                        tb1.ForeColor = textBox1.ForeColor;
                        tb1.Font = textBox1.Font;
                        tb1.Size = textBox1.Size;
                        tb1.ReadOnly = textBox1.ReadOnly;
                        tb1.BorderStyle = textBox1.BorderStyle;
                        tb1.Text = text11;
                        panel_err.Controls.Add(tb1);

                        TextBox tb2 = new TextBox();
                        tb2.Location = new Point(textBox2.Location.X, textBox2.Location.Y + i * (textBox2.Height + extra1));
                        tb2.BackColor = textBox2.BackColor;
                        tb2.ForeColor = textBox2.ForeColor;
                        tb2.Font = textBox2.Font;
                        tb2.Size = textBox2.Size;
                        tb2.ReadOnly = textBox2.ReadOnly;
                        tb2.BorderStyle = textBox2.BorderStyle;
                        tb2.Text = text21;
                        panel_err.Controls.Add(tb2);

                        TextBox tb3 = new TextBox();
                        tb3.Location = new Point(textBox3.Location.X, textBox3.Location.Y + i * (textBox3.Height + extra1));
                        tb3.BackColor = textBox3.BackColor;
                        tb3.ForeColor = textBox3.ForeColor;
                        tb3.Font = textBox3.Font;
                        tb3.Size = textBox3.Size;
                        tb3.ReadOnly = textBox3.ReadOnly;
                        tb3.BorderStyle = textBox3.BorderStyle;
                        tb3.Text = text31;
                        panel_err.Controls.Add(tb3);

                        TextBox tb4 = new TextBox();
                        tb4.Location = new Point(textBox4.Location.X, textBox4.Location.Y + i * (textBox4.Height + extra1));
                        tb4.BackColor = textBox4.BackColor;
                        tb4.ForeColor = textBox4.ForeColor;
                        tb4.Font = textBox4.Font;
                        tb4.Size = textBox4.Size;
                        tb4.ReadOnly = textBox4.ReadOnly;
                        tb4.BorderStyle = textBox4.BorderStyle;
                        tb4.Text = text41;
                        panel_err.Controls.Add(tb4);

                        TextBox tb5 = new TextBox();
                        tb5.Location = new Point(textBox5.Location.X, textBox5.Location.Y + i * (textBox5.Height + extra1));
                        tb5.BackColor = textBox5.BackColor;
                        tb5.ForeColor = textBox5.ForeColor;
                        tb5.Font = textBox5.Font;
                        tb5.Size = textBox5.Size;
                        tb5.BorderStyle = textBox5.BorderStyle;
                        tb5.ReadOnly = textBox5.ReadOnly;
                        tb5.Text = text51;
                        panel_err.Controls.Add(tb5);
                    }
                }
            }
        }
        private void make_first_line_invisible()
        {


            textBox1.Visible = false;
            textBox2.Visible = false;
            textBox3.Visible = false;
            textBox4.Visible = false;
            textBox5.Visible = false;
            comboBox1.SelectedIndex = 0;
            comboBox1.Visible = false;
            button_z1.Visible = false;

            for (int i = panel_err.Controls.Count - 1; i >= 0; --i)
            {
                Control ctrl1 = panel_err.Controls[i] as Control;
                if (ctrl1.Location.Y > textBox1.Location.Y + extra1)
                {
                    panel_err.Controls.Remove(ctrl1);
                    ctrl1.Dispose();
                }
            }
            textBox_data_changes.Text = "";
            textBox_geom_changes.Text = "";
            textBox_new_data_table_features.Text = "";
            textBox_new_dwg_features.Text = "";


        }

        private void button_zoom_click(object ob, EventArgs e)
        {
            Control ctrl1 = ob as Control;
            if (dt_err0 == null) return;
            if (dt_err0.Rows.Count == 0) return;

            if (Freeze_operations == false && ctrl1 != null)
            {
                int Y = ctrl1.Location.Y;

                int index1 = (Y - textBox1.Location.Y) / (textBox1.Height + extra1);

                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                Freeze_operations = true;
                try
                {
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite);

                            if (dt_err0.Rows[index1][0] != DBNull.Value)
                            {
                                string handle1 = Convert.ToString(dt_err0.Rows[index1][0]);


                                ObjectId ObjID1 = Functions.GetObjectId(ThisDrawing.Database, handle1);
                                Entity Ent1 = Trans1.GetObject(ObjID1, OpenMode.ForRead) as Entity;

                                if (Ent1 != null)
                                {

                                    using (Autodesk.AutoCAD.GraphicsSystem.Manager GraphicsManager = ThisDrawing.GraphicsManager)
                                    {

                                        int Cvport = Convert.ToInt32(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CVPORT"));

                                        #region 2015 dlls:
                                        Autodesk.AutoCAD.GraphicsSystem.KernelDescriptor kd = new Autodesk.AutoCAD.GraphicsSystem.KernelDescriptor();
                                        kd.addRequirement(Autodesk.AutoCAD.UniqueString.Intern("3D Drawing"));
                                        Autodesk.AutoCAD.GraphicsSystem.View view = GraphicsManager.ObtainAcGsView(Cvport, kd);
                                        #endregion

                                        #region 2013 dlls:
                                        //Autodesk.AutoCAD.GraphicsSystem.View view = GraphicsManager.GetGsView(Cvport, true);
                                        #endregion

                                        if (view != null)
                                        {
                                            using (view)
                                            {
                                                {
                                                    view.ZoomExtents(Ent1.GeometricExtents.MaxPoint, Ent1.GeometricExtents.MinPoint);
                                                }
                                                view.Zoom(0.95);//<--optional 
                                                GraphicsManager.SetViewportFromView(Cvport, view, true, true, false);
                                                Trans1.Commit();

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
                    MessageBox.Show(ex.Message);
                }

                Freeze_operations = false;
            }

        }


        private string get_combobox_value(int i)
        {

            int Y = comboBox1.Location.Y + i * (comboBox1.Height + extra1);

            foreach (Control ctrl1 in panel_err.Controls)
            {
                ComboBox combo1 = ctrl1 as ComboBox;
                if (combo1 != null)
                {
                    if (Y == combo1.Location.Y)
                    {
                        return combo1.Text;
                    }
                }
            }
            return "";
        }

        private void button_reconcile_data_Click(object sender, EventArgs e)
        {
            if (dt_err0 == null)
            {
                MessageBox.Show("no errors loaded");
                return;
            }
            if (dt_err0.Rows.Count == 0)
            {
                MessageBox.Show("no errors loaded");
                return;
            }

            bool transfer_to_config = false;
            bool is_sta1_or_len = false;
            bool transfer_to_xl = false;

            bool is_new_ownership_block = false;
            bool is_new_crossing_mtext = false;
            bool is_new_xl = false;

            List<int> Lista_del = new List<int>();

            if (Freeze_operations == false)
            {
                ObjectId[] Empty_array = null;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                try
                {
                    Freeze_operations = true;
                    _AGEN_mainform.tpage_processing.Show();
                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                            bool new_description_in_crossing_xls = false;
                            for (int i = 0; i < dt_err0.Rows.Count; ++i)
                            {
                                string err_type = Convert.ToString(dt_err0.Rows[i][6]);
                                string correct_value = get_combobox_value(i);
                                string layer_name = Convert.ToString(dt_err0.Rows[i][1]);

                                switch (err_type)
                                {
                                    case "Data Discrepancies":
                                        switch (correct_value)
                                        {

                                            case "Table Value":
                                                #region  Table Value ownership and crossing
                                                try
                                                {
                                                    string id_err = Convert.ToString(dt_err0.Rows[i][0]);
                                                    ObjectId id1 = Functions.GetObjectId(ThisDrawing.Database, id_err);
                                                    Entity ent1 = Trans1.GetObject(id1, OpenMode.ForWrite) as Entity;

                                                    if (ent1 is BlockReference)
                                                    {
                                                        BlockReference block1 = ent1 as BlockReference;
                                                        if (block1 != null)
                                                        {
                                                            if (block1.AttributeCollection.Count > 0)
                                                            {
                                                                foreach (ObjectId id2 in block1.AttributeCollection)
                                                                {
                                                                    AttributeReference atr1 = Trans1.GetObject(id2, OpenMode.ForRead) as AttributeReference;
                                                                    string tag_err = Convert.ToString(dt_err0.Rows[i][3]);
                                                                    if (atr1.Tag == tag_err)
                                                                    {
                                                                        atr1.UpgradeOpen();
                                                                        for (int j = 0; j < _AGEN_mainform.dt_config_ownership.Rows.Count; ++j)
                                                                        {
                                                                            string id_config = Convert.ToString(_AGEN_mainform.dt_config_ownership.Rows[j][0]);
                                                                            if (id_config == id_err)
                                                                            {

                                                                                string config_val = Convert.ToString(_AGEN_mainform.dt_config_ownership.Rows[j][tag_err]);
                                                                                string err_val = Convert.ToString(dt_err0.Rows[i][5]);

                                                                                string updated_val = err_val;

                                                                                if (tag_err == _AGEN_mainform.owner_len_atr)
                                                                                {
                                                                                    if (config_val.Contains("'") == true)
                                                                                    {
                                                                                        updated_val = updated_val + "'";
                                                                                        config_val = config_val.Replace("'", "");
                                                                                    }
                                                                                }

                                                                                if (config_val != err_val)
                                                                                {
                                                                                    _AGEN_mainform.dt_config_ownership.Rows[j][tag_err] = updated_val;
                                                                                    transfer_to_config = true;
                                                                                }


                                                                                atr1.TextString = updated_val;
                                                                                if (Lista_del.Contains(i) == false) Lista_del.Add(i);
                                                                                j = _AGEN_mainform.dt_config_ownership.Rows.Count;
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }

                                                    if (ent1 is MText)
                                                    {
                                                        MText mtext1 = ent1 as MText;
                                                        if (mtext1 != null)
                                                        {
                                                            string handle0 = Convert.ToString(dt_err0.Rows[i][0]);
                                                            string valoarea_corecta = Convert.ToString(dt_err0.Rows[i][5]);
                                                            string station_prefix = _AGEN_mainform.tpage_crossing_draw.get_textBox_station_prefix();

                                                            valoarea_corecta = Functions.remove_space_from_start_and_end_of_a_string(valoarea_corecta);

                                                            mtext1.Contents = valoarea_corecta;

                                                            string Sta_correct = Functions.extract_station_from_mtext(valoarea_corecta);
                                                            string Descr_correct = Functions.extract_description_from_mtext(valoarea_corecta);

                                                            if (station_prefix.Replace(" ", "") != "")
                                                            {
                                                                Sta_correct = station_prefix + " " + Sta_correct;
                                                            }

                                                            for (int j = 0; j < _AGEN_mainform.dt_config_crossing.Rows.Count; ++j)
                                                            {
                                                                string id_config = Convert.ToString(_AGEN_mainform.dt_config_crossing.Rows[j][0]);
                                                                string layer1 = Convert.ToString(_AGEN_mainform.dt_config_crossing.Rows[j][1]);
                                                                if (handle0 == id_config && layer1 == _AGEN_mainform.layer_crossing_band_text)
                                                                {
                                                                    _AGEN_mainform.dt_config_crossing.Rows[j][3] = Sta_correct;
                                                                    _AGEN_mainform.dt_config_crossing.Rows[j][5] = Descr_correct;
                                                                    transfer_to_config = true;
                                                                }

                                                            }


                                                            if (Lista_del.Contains(i) == false) Lista_del.Add(i);


                                                        }

                                                    }
                                                }
                                                catch (System.Exception ex)
                                                {
                                                    MessageBox.Show(ex.Message + "\r\n" + "No object id found");
                                                }
                                                break;
                                            #endregion


                                            case "DWG Value":
                                                #region DWG Value for ownership with crossing
                                                try
                                                {
                                                    string id_err = Convert.ToString(dt_err0.Rows[i][0]);
                                                    ObjectId id1 = Functions.GetObjectId(ThisDrawing.Database, id_err);
                                                    Entity ent1 = Trans1.GetObject(id1, OpenMode.ForRead) as Entity;
                                                    if (ent1 is BlockReference && ent1.Layer == _AGEN_mainform.layer_ownership_band)
                                                    {
                                                        #region ownership 
                                                        BlockReference block1 = ent1 as BlockReference;
                                                        if (block1 != null)
                                                        {
                                                            if (block1.AttributeCollection.Count > 0)
                                                            {
                                                                foreach (ObjectId id2 in block1.AttributeCollection)
                                                                {
                                                                    AttributeReference atr1 = Trans1.GetObject(id2, OpenMode.ForRead) as AttributeReference;
                                                                    string tag_err = Convert.ToString(dt_err0.Rows[i][3]);

                                                                    if (atr1.Tag == tag_err)
                                                                    {
                                                                        if (checkBox_ownership.Checked == true)
                                                                        {
                                                                            if (layer_name == _AGEN_mainform.layer_ownership_band)
                                                                            {
                                                                                if (tag_err != _AGEN_mainform.owner_len_atr && tag_err != _AGEN_mainform.owner_sta1_atr && tag_err != _AGEN_mainform.owner_sta2_atr)
                                                                                {
                                                                                    for (int j = 0; j < _AGEN_mainform.dt_config_ownership.Rows.Count; ++j)
                                                                                    {
                                                                                        string id_config = Convert.ToString(_AGEN_mainform.dt_config_ownership.Rows[j][0]);
                                                                                        if (id_config == id_err)
                                                                                        {
                                                                                            _AGEN_mainform.dt_config_ownership.Rows[j][tag_err] = atr1.TextString;
                                                                                            transfer_to_config = true;
                                                                                            if (Lista_del.Contains(i) == false) Lista_del.Add(i);
                                                                                            j = _AGEN_mainform.dt_config_ownership.Rows.Count;
                                                                                        }
                                                                                    }

                                                                                    for (int j = 0; j < _AGEN_mainform.Data_Table_property.Rows.Count; ++j)
                                                                                    {
                                                                                        string handles = Convert.ToString(_AGEN_mainform.Data_Table_property.Rows[j][11]);
                                                                                        if (handles.Contains(id_err) == true)
                                                                                        {
                                                                                            transfer_to_xl = true;

                                                                                            int col1 = 0;
                                                                                            if (tag_err == _AGEN_mainform.owner_linelist_atr) col1 = 8;
                                                                                            if (tag_err == _AGEN_mainform.owner_owner_atr) col1 = 7;

                                                                                            _AGEN_mainform.Data_Table_property.Rows[j][col1] = atr1.TextString;
                                                                                        }
                                                                                    }
                                                                                }
                                                                                else
                                                                                {
                                                                                    is_sta1_or_len = true;
                                                                                }
                                                                            }
                                                                        }

                                                                    }
                                                                }
                                                            }
                                                        }
                                                        #endregion
                                                    }

                                                    if (ent1 is MText && (ent1.Layer == _AGEN_mainform.layer_crossing_band_text || ent1.Layer == _AGEN_mainform.layer_crossing_band_pi))
                                                    {
                                                        #region Crossing
                                                        MText mtext1 = ent1 as MText;
                                                        if (mtext1 != null)
                                                        {
                                                            string valoarea_corecta = Functions.remove_space_from_start_and_end_of_a_string(mtext1.Contents);
                                                            string Descr_correct = Functions.extract_description_from_mtext(valoarea_corecta);
                                                            string Sta_correct = Functions.extract_station_from_mtext(valoarea_corecta);

                                                            double x1 = -1.234;
                                                            double y1 = -1.234;

                                                            for (int j = 0; j < _AGEN_mainform.dt_config_crossing.Rows.Count; ++j)
                                                            {
                                                                string id_config = Convert.ToString(_AGEN_mainform.dt_config_crossing.Rows[j][0]);
                                                                string layer1 = Convert.ToString(_AGEN_mainform.dt_config_crossing.Rows[j][1]);
                                                                if (id_err == id_config && layer1 == _AGEN_mainform.layer_crossing_band_text)
                                                                {
                                                                    _AGEN_mainform.dt_config_crossing.Rows[j][3] = Sta_correct;
                                                                    _AGEN_mainform.dt_config_crossing.Rows[j][5] = Descr_correct;
                                                                    transfer_to_config = true;
                                                                    if (Lista_del.Contains(i) == false) Lista_del.Add(i);
                                                                    x1 = Convert.ToDouble(_AGEN_mainform.dt_config_crossing.Rows[j][6]);
                                                                    y1 = Convert.ToDouble(_AGEN_mainform.dt_config_crossing.Rows[j][7]);
                                                                    j = _AGEN_mainform.dt_config_crossing.Rows.Count;
                                                                }
                                                            }

                                                            if (_AGEN_mainform.Data_Table_crossings != null && _AGEN_mainform.Data_Table_crossings.Rows.Count > 0 && x1 != -1.234 && y1 != -1.234)
                                                            {
                                                                List<string> lista_descr_existing = new List<string>();
                                                                List<int> lista_index = new List<int>();

                                                                for (int k = 0; k < _AGEN_mainform.Data_Table_crossings.Rows.Count; ++k)
                                                                {
                                                                    double x2 = Convert.ToDouble(_AGEN_mainform.Data_Table_crossings.Rows[k]["X"]);
                                                                    double y2 = Convert.ToDouble(_AGEN_mainform.Data_Table_crossings.Rows[k]["Y"]);
                                                                    string descr_existing = Convert.ToString(_AGEN_mainform.Data_Table_crossings.Rows[k]["Desc"]);
                                                                    if (Math.Round(x1, 3) == Math.Round(x2, 3) && Math.Round(y1, 3) == Math.Round(y2, 3))
                                                                    {
                                                                        lista_descr_existing.Add(descr_existing);
                                                                        lista_index.Add(k);
                                                                    }
                                                                }

                                                                if (lista_index.Count == 1)
                                                                {
                                                                    _AGEN_mainform.Data_Table_crossings.Rows[lista_index[0]]["Desc"] = Descr_correct;
                                                                    new_description_in_crossing_xls = true;
                                                                }

                                                                if (lista_index.Count > 1)
                                                                {
                                                                    for (int k = 0; k < lista_index.Count; ++k)
                                                                    {
                                                                        string messagebox_dialog = "There is more than one crossing with the same coordinate," +
                                                                                           "\r\nPress yes if this line is the one you want to update:\r\n" +
                                                                              Convert.ToString(_AGEN_mainform.Data_Table_crossings.Rows[lista_index[k]]["Desc"]);

                                                                        if (MessageBox.Show(messagebox_dialog, "agen", MessageBoxButtons.YesNo) == DialogResult.Yes)
                                                                        {
                                                                            _AGEN_mainform.Data_Table_crossings.Rows[lista_index[k]]["Desc"] = Descr_correct;
                                                                            new_description_in_crossing_xls = true;
                                                                            k = lista_index.Count;
                                                                        }
                                                                    }
                                                                }

                                                            }
                                                        }
                                                        #endregion


                                                    }


                                                }
                                                catch (System.Exception ex)
                                                {
                                                    MessageBox.Show(ex.Message + "\r\n" + "No object id found");
                                                }

                                                break;
                                            #endregion

                                            default:

                                                break;
                                        }

                                        break;
                                    case "Manual Drafting Adjustments":
                                        switch (correct_value)
                                        {
                                            #region Table Value ownership and crossing
                                            case "Table Value":
                                                try
                                                {
                                                    string id_err = Convert.ToString(dt_err0.Rows[i][0]);
                                                    ObjectId id1 = Functions.GetObjectId(ThisDrawing.Database, id_err);
                                                    Entity Ent1 = Trans1.GetObject(id1, OpenMode.ForWrite) as Entity;
                                                    string tag_err = Convert.ToString(dt_err0.Rows[i][3]);

                                                    if (Ent1 is BlockReference)
                                                    {
                                                        BlockReference block1 = Ent1 as BlockReference;
                                                        if (block1 != null)
                                                        {
                                                            string visib1 = "";
                                                            double stretch1 = 0;

                                                            if (tag_err == "Visibility") visib1 = Convert.ToString(dt_err0.Rows[i][5]);
                                                            if (tag_err == "Stretch") stretch1 = Convert.ToDouble(dt_err0.Rows[i][5]);

                                                            if (block1.IsDynamicBlock == true)
                                                            {
                                                                using (DynamicBlockReferencePropertyCollection pc = block1.DynamicBlockReferencePropertyCollection)
                                                                {
                                                                    foreach (DynamicBlockReferenceProperty prop in pc)
                                                                    {
                                                                        if (tag_err == "Visibility")
                                                                        {
                                                                            if (prop.PropertyName == "Visibility1")
                                                                            {
                                                                                prop.Value = visib1;
                                                                            }
                                                                        }
                                                                        if (tag_err == "Stretch")
                                                                        {
                                                                            if (prop.PropertyName == "Distance1")
                                                                            {
                                                                                prop.Value = stretch1;
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }

                                                            if (tag_err == "X Position")
                                                            {
                                                                double x1 = Convert.ToDouble(dt_err0.Rows[i][5]);
                                                                double y1 = block1.Position.Y;
                                                                block1.Position = new Point3d(x1, y1, 0);
                                                            }
                                                            if (tag_err == "Y Position")
                                                            {
                                                                double x1 = block1.Position.X;
                                                                double y1 = Convert.ToDouble(dt_err0.Rows[i][5]);
                                                                block1.Position = new Point3d(x1, y1, 0);
                                                            }



                                                            if (Lista_del.Contains(i) == false) Lista_del.Add(i);
                                                        }
                                                    }

                                                    if (Ent1 is MText)
                                                    {
                                                        MText mtext1 = Ent1 as MText;
                                                        if (mtext1 != null)
                                                        {
                                                            if (tag_err == "X Position")
                                                            {
                                                                double x1 = Convert.ToDouble(dt_err0.Rows[i][5]);
                                                                double y1 = mtext1.Location.Y;
                                                                mtext1.Location = new Point3d(x1, y1, 0);
                                                            }
                                                            if (tag_err == "Y Position")
                                                            {
                                                                double x1 = mtext1.Location.X;
                                                                double y1 = Convert.ToDouble(dt_err0.Rows[i][5]);
                                                                mtext1.Location = new Point3d(x1, y1, 0);
                                                            }



                                                            if (Lista_del.Contains(i) == false) Lista_del.Add(i);
                                                        }
                                                    }

                                                }
                                                catch (System.Exception ex)
                                                {
                                                    MessageBox.Show(ex.Message + "\r\n" + "No object id found");
                                                }

                                                break;
                                            #endregion

                                            #region DWG Value ownership and crossing
                                            case "DWG Value":

                                                try
                                                {
                                                    string id_err = Convert.ToString(dt_err0.Rows[i][0]);
                                                    ObjectId id1 = Functions.GetObjectId(ThisDrawing.Database, id_err);
                                                    Entity ent1 = Trans1.GetObject(id1, OpenMode.ForRead) as Entity;
                                                    string tag_err = Convert.ToString(dt_err0.Rows[i][3]);
                                                    if (ent1 is BlockReference)
                                                    {
                                                        BlockReference block1 = ent1 as BlockReference;
                                                        if (block1 != null)
                                                        {
                                                            string visib1 = "";
                                                            double stretch1 = 0;
                                                            if (block1.IsDynamicBlock == true)
                                                            {
                                                                using (DynamicBlockReferencePropertyCollection pc = block1.DynamicBlockReferencePropertyCollection)
                                                                {
                                                                    foreach (DynamicBlockReferenceProperty prop in pc)
                                                                    {
                                                                        if (tag_err == "Visibility")
                                                                        {
                                                                            if (prop.PropertyName == "Visibility1")
                                                                            {
                                                                                visib1 = Convert.ToString(prop.Value);
                                                                            }
                                                                        }
                                                                        if (tag_err == "Stretch")
                                                                        {
                                                                            if (prop.PropertyName == "Distance1")
                                                                            {
                                                                                stretch1 = Convert.ToDouble(prop.Value);
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }

                                                            double x1 = -123.456;
                                                            double y1 = -123.456;
                                                            if (tag_err == "X Position")
                                                            {
                                                                x1 = block1.Position.X;
                                                            }
                                                            if (tag_err == "Y Position")
                                                            {
                                                                y1 = block1.Position.Y;
                                                            }

                                                            for (int j = 0; j < _AGEN_mainform.dt_config_ownership.Rows.Count; ++j)
                                                            {
                                                                string id_config = Convert.ToString(_AGEN_mainform.dt_config_ownership.Rows[j][0]);
                                                                if (id_config == id_err)
                                                                {
                                                                    if (visib1 != "") _AGEN_mainform.dt_config_ownership.Rows[j][5] = visib1;
                                                                    if (stretch1 > 0) _AGEN_mainform.dt_config_ownership.Rows[j][6] = stretch1;
                                                                    if (x1 != -123.456) _AGEN_mainform.dt_config_ownership.Rows[j][3] = x1;
                                                                    if (y1 != -123.456) _AGEN_mainform.dt_config_ownership.Rows[j][4] = y1;
                                                                    transfer_to_config = true;
                                                                    if (Lista_del.Contains(i) == false) Lista_del.Add(i);
                                                                    j = _AGEN_mainform.dt_config_ownership.Rows.Count;
                                                                }
                                                            }
                                                        }
                                                    }

                                                    if (ent1 is MText)
                                                    {
                                                        MText Mtext1 = ent1 as MText;
                                                        if (Mtext1 != null)
                                                        {


                                                            double x1 = -123.456;
                                                            double y1 = -123.456;
                                                            if (tag_err == "X Position")
                                                            {
                                                                x1 = Mtext1.Location.X;
                                                            }
                                                            if (tag_err == "Y Position")
                                                            {
                                                                y1 = Mtext1.Location.Y;
                                                            }

                                                            for (int j = 0; j < _AGEN_mainform.dt_config_crossing.Rows.Count; ++j)
                                                            {
                                                                string id_config = Convert.ToString(_AGEN_mainform.dt_config_crossing.Rows[j][0]);
                                                                if (id_config == id_err)
                                                                {
                                                                    if (x1 != -123.456) _AGEN_mainform.dt_config_crossing.Rows[j][12] = x1;
                                                                    if (y1 != -123.456) _AGEN_mainform.dt_config_crossing.Rows[j][13] = y1;
                                                                    transfer_to_config = true;
                                                                    if (Lista_del.Contains(i) == false) Lista_del.Add(i);
                                                                    j = _AGEN_mainform.dt_config_crossing.Rows.Count;
                                                                }
                                                            }
                                                        }
                                                    }

                                                }
                                                catch (System.Exception ex)
                                                {
                                                    MessageBox.Show(ex.Message + "\r\n" + "No object id found");
                                                }

                                                break;
                                            #endregion

                                            default:

                                                break;
                                        }
                                        break;
                                    case "Extra Basefile Features":
                                        switch (correct_value)
                                        {
                                            #region DWG Value
                                            case "DWG Value":

                                                try
                                                {
                                                    string id_err = Convert.ToString(dt_err0.Rows[i][0]);
                                                    ObjectId id1 = Functions.GetObjectId(ThisDrawing.Database, id_err);
                                                    Entity ent1 = Trans1.GetObject(id1, OpenMode.ForRead) as Entity;

                                                    #region ownerhip band
                                                    if (ent1 is BlockReference && ent1.Layer == _AGEN_mainform.layer_ownership_band)
                                                    {
                                                        BlockReference block1 = ent1 as BlockReference;
                                                        if (block1 != null)
                                                        {

                                                            for (int k = _AGEN_mainform.dt_config_ownership.Rows.Count - 1; k >= 0; --k)
                                                            {
                                                                string id0 = Convert.ToString(_AGEN_mainform.dt_config_ownership.Rows[k][0]);
                                                                if (id0 == id_err)
                                                                {
                                                                    _AGEN_mainform.dt_config_ownership.Rows[k].Delete();
                                                                }
                                                            }

                                                            _AGEN_mainform.dt_config_ownership.Rows.Add();
                                                            _AGEN_mainform.dt_config_ownership.Rows[_AGEN_mainform.dt_config_ownership.Rows.Count - 1][0] = id_err;
                                                            _AGEN_mainform.dt_config_ownership.Rows[_AGEN_mainform.dt_config_ownership.Rows.Count - 1][1] = Functions.get_block_name(block1);
                                                            _AGEN_mainform.dt_config_ownership.Rows[_AGEN_mainform.dt_config_ownership.Rows.Count - 1][2] = block1.Layer;
                                                            _AGEN_mainform.dt_config_ownership.Rows[_AGEN_mainform.dt_config_ownership.Rows.Count - 1][3] = block1.Position.X;
                                                            _AGEN_mainform.dt_config_ownership.Rows[_AGEN_mainform.dt_config_ownership.Rows.Count - 1][4] = block1.Position.Y;


                                                            if (block1.IsDynamicBlock == true)
                                                            {
                                                                using (DynamicBlockReferencePropertyCollection pc = block1.DynamicBlockReferencePropertyCollection)
                                                                {
                                                                    foreach (DynamicBlockReferenceProperty prop in pc)
                                                                    {
                                                                        if (prop.PropertyName == "Visibility1")
                                                                        {
                                                                            _AGEN_mainform.dt_config_ownership.Rows[_AGEN_mainform.dt_config_ownership.Rows.Count - 1][5] = Convert.ToString(prop.Value);

                                                                        }
                                                                        if (prop.PropertyName == "Distance1")
                                                                        {
                                                                            _AGEN_mainform.dt_config_ownership.Rows[_AGEN_mainform.dt_config_ownership.Rows.Count - 1][6] = Convert.ToDouble(prop.Value);
                                                                        }
                                                                    }
                                                                }
                                                            }

                                                            if (block1.AttributeCollection.Count > 0)
                                                            {
                                                                foreach (ObjectId id2 in block1.AttributeCollection)
                                                                {
                                                                    AttributeReference atr1 = Trans1.GetObject(id2, OpenMode.ForRead) as AttributeReference;
                                                                    if (atr1 != null)
                                                                    {
                                                                        string tag1 = atr1.Tag;
                                                                        if (_AGEN_mainform.dt_config_ownership.Columns.Contains(tag1) == false)
                                                                        {
                                                                            _AGEN_mainform.dt_config_ownership.Columns.Add(tag1, typeof(string));
                                                                        }
                                                                        if (atr1.TextString != "") _AGEN_mainform.dt_config_ownership.Rows[_AGEN_mainform.dt_config_ownership.Rows.Count - 1][tag1] = atr1.TextString;
                                                                    }
                                                                }
                                                            }

                                                            transfer_to_config = true;
                                                            is_new_ownership_block = true;
                                                        }
                                                    }

                                                    #endregion

                                                    #region crossing band
                                                    if (ent1 is MText && (ent1.Layer == _AGEN_mainform.layer_crossing_band_text || ent1.Layer == _AGEN_mainform.layer_crossing_band_pi))
                                                    {
                                                        is_new_crossing_mtext = true;



                                                    }

                                                    #endregion

                                                }
                                                catch (System.Exception ex)
                                                {
                                                    MessageBox.Show(ex.Message + "\r\n" + "No object id found");
                                                }

                                                break;
                                            #endregion

                                            default:

                                                break;
                                        }



                                        break;
                                    case "Missing Basefile Features":
                                        is_new_xl = true;
                                        break;
                                    default:
                                        break;
                                }
                            }

                            if (transfer_to_config == true)
                            {
                                _AGEN_mainform.tpage_setup.transfera_band_settings_to_config_excel(_AGEN_mainform.dt_config_ownership, _AGEN_mainform.dt_config_crossing);
                            }

                            if (is_sta1_or_len == true)
                            {
                                MessageBox.Show("All data changes containing stationing or length will not be updated in the data tables.\r\nPlease adjust your data accordingly.");
                            }

                            if (is_new_ownership_block == true)
                            {
                                MessageBox.Show("You have a new ownership block in the drawing.\r\nPlease adjust your data accordingly.");
                            }

                            if (is_new_crossing_mtext == true)
                            {
                                MessageBox.Show("You have a new crossing text in the drawing.\r\nPlease adjust your data accordingly.");
                            }

                            if (is_new_xl == true)
                            {
                                MessageBox.Show("You have a new entry in the data table.\r\nPlease redraft the band");
                            }

                            if (transfer_to_xl == true)
                            {
                                string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();

                                if (System.IO.Directory.Exists(ProjF) == false)
                                {
                                    Freeze_operations = false;
                                    MessageBox.Show("the project database folder does not exist");
                                    _AGEN_mainform.tpage_processing.Hide();
                                    return;
                                }

                                if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                                {
                                    ProjF = ProjF + "\\";
                                }

                                if (checkBox_ownership.Checked == true)
                                {
                                    string fisier_prop = ProjF + _AGEN_mainform.property_excel_name;
                                    if (System.IO.File.Exists(fisier_prop) == false)
                                    {
                                        Freeze_operations = false;
                                        MessageBox.Show("the property data file does not exist");
                                        _AGEN_mainform.tpage_processing.Hide();
                                        return;
                                    }
                                    Functions.create_backup(fisier_prop);
                                    _AGEN_mainform.tpage_owner_scan.Populate_property_file(fisier_prop);
                                }

                            }

                            if (Lista_del.Count > 0)
                            {
                                for (int i = Lista_del.Count - 1; i >= 0; --i)
                                {
                                    dt_err0.Rows[i].Delete();
                                }
                            }

                            if (new_description_in_crossing_xls == true)
                            {
                                string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();

                                if (System.IO.Directory.Exists(ProjF) == false)
                                {
                                    Freeze_operations = false;
                                    MessageBox.Show("the project database folder does not exist");
                                    _AGEN_mainform.tpage_processing.Hide();
                                    return;
                                }
                                string fisier_cs = ProjF + _AGEN_mainform.crossing_excel_name;
                                Functions.create_backup(fisier_cs);
                                _AGEN_mainform.tpage_crossing_scan.Populate_crossing_file(fisier_cs);
                            }


                            make_first_line_invisible();
                            calculate_total_errors();
                            transfer_data_to_panel(dt_err0);

                            Trans1.Commit();

                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                Editor1.SetImpliedSelection(Empty_array);
                Editor1.WriteMessage("\nCommand:");
                Freeze_operations = false;
                _AGEN_mainform.tpage_processing.Hide();
            }

        }
        private void label_band_analyze_Click(object sender, EventArgs e)
        {
            if (Functions.is_dan_popescu() == true)
            {
                if (panel_dan.Visible == false)
                {
                    panel_dan.Visible = true;

                }
                else
                {
                    panel_dan.Visible = false;
                }
            }
        }
        private void button_output_mat_lin_to_excel_Click(object sender, EventArgs e)
        {

            if (Freeze_operations == false)
            {


                ObjectId[] Empty_array = null;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                if (Functions.IsNumeric(textBox_row_start.Text) == true)
                {
                    row_start = Convert.ToInt32(textBox_row_start.Text);


                    try
                    {
                        Freeze_operations = true;
                        using (DocumentLock lock1 = ThisDrawing.LockDocument())
                        {
                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                            {
                                BlockTable BlockTable1 = Trans1.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as BlockTable;

                                BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                                List<string> handles_of_new_items = new List<string>();


                                dt_ownership_block_from_dwg = new System.Data.DataTable();
                                dt_ownership_block_from_dwg.Columns.Add("objectid", typeof(string));
                                dt_ownership_block_from_dwg.Columns.Add("blockname", typeof(string));
                                dt_ownership_block_from_dwg.Columns.Add("layer", typeof(string));
                                dt_ownership_block_from_dwg.Columns.Add("x", typeof(double));
                                dt_ownership_block_from_dwg.Columns.Add("y", typeof(double));
                                dt_ownership_block_from_dwg.Columns.Add("visibility", typeof(string));
                                dt_ownership_block_from_dwg.Columns.Add("stretch", typeof(double));
                                dt_ownership_block_from_dwg.Columns.Add("dwg_name", typeof(string));
                                #region load blocks from drawing
                                foreach (ObjectId id1 in BTrecord)
                                {
                                    BlockReference block1 = Trans1.GetObject(id1, OpenMode.ForRead) as BlockReference;

                                    if (block1 != null)
                                    {
                                        bool add_it = false;
                                        if (block1.AttributeCollection.Count > 0)
                                        {
                                            Autodesk.AutoCAD.DatabaseServices.AttributeCollection attColl = block1.AttributeCollection;

                                            foreach (ObjectId odid in attColl)
                                            {
                                                AttributeReference atr1 = Trans1.GetObject(odid, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as AttributeReference;
                                                if (atr1 != null)
                                                {
                                                    string Tag1 = atr1.Tag;
                                                    if (Tag1.ToUpper() == "MAT")
                                                    {
                                                        add_it = true;
                                                    }
                                                }
                                            }
                                        }

                                        if (add_it == true)
                                        {
                                            if (block1.AttributeCollection.Count > 0)
                                            {
                                                string blockname = Functions.get_block_name(block1);

                                                dt_ownership_block_from_dwg.Rows.Add();
                                                dt_ownership_block_from_dwg.Rows[dt_ownership_block_from_dwg.Rows.Count - 1][0] = block1.ObjectId.Handle.Value.ToString();
                                                dt_ownership_block_from_dwg.Rows[dt_ownership_block_from_dwg.Rows.Count - 1][1] = blockname;
                                                dt_ownership_block_from_dwg.Rows[dt_ownership_block_from_dwg.Rows.Count - 1][2] = block1.Layer;
                                                dt_ownership_block_from_dwg.Rows[dt_ownership_block_from_dwg.Rows.Count - 1][3] = block1.Position.X;
                                                dt_ownership_block_from_dwg.Rows[dt_ownership_block_from_dwg.Rows.Count - 1][4] = block1.Position.Y;
                                                dt_ownership_block_from_dwg.Rows[dt_ownership_block_from_dwg.Rows.Count - 1][7] = ThisDrawing.Name;
                                                Autodesk.AutoCAD.DatabaseServices.AttributeCollection attColl = block1.AttributeCollection;

                                                foreach (ObjectId odid in attColl)
                                                {
                                                    AttributeReference atr1 = Trans1.GetObject(odid, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as AttributeReference;
                                                    if (atr1 != null)
                                                    {
                                                        string Tag1 = atr1.Tag;
                                                        if (dt_ownership_block_from_dwg.Columns.Contains(Tag1) == false)
                                                        {
                                                            dt_ownership_block_from_dwg.Columns.Add(Tag1, typeof(string));
                                                        }
                                                        string val1 = atr1.TextString;
                                                        if (val1 != "")
                                                        {
                                                            dt_ownership_block_from_dwg.Rows[dt_ownership_block_from_dwg.Rows.Count - 1][Tag1] = val1;
                                                        }
                                                    }
                                                }

                                                if (block1.IsDynamicBlock == true)
                                                {
                                                    using (DynamicBlockReferencePropertyCollection pc = block1.DynamicBlockReferencePropertyCollection)
                                                    {
                                                        foreach (DynamicBlockReferenceProperty prop in pc)
                                                        {
                                                            if (prop.PropertyName == "Visibility1")
                                                            {
                                                                dt_ownership_block_from_dwg.Rows[dt_ownership_block_from_dwg.Rows.Count - 1][5] = Convert.ToString(prop.Value);
                                                            }
                                                            if (prop.PropertyName == "Distance1")
                                                            {
                                                                dt_ownership_block_from_dwg.Rows[dt_ownership_block_from_dwg.Rows.Count - 1][6] = Convert.ToDouble(prop.Value);
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }


                                    }
                                }
                                #endregion






                                Functions.Transfer_datatable_to_existing_excel_spreadsheet(dt_ownership_block_from_dwg, row_start);

                                row_start = row_start + dt_ownership_block_from_dwg.Rows.Count + 1;
                                textBox_row_start.Text = row_start.ToString();
                                Trans1.Commit();
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                Editor1.SetImpliedSelection(Empty_array);
                Editor1.WriteMessage("\nCommand:");
                Freeze_operations = false;
                _AGEN_mainform.tpage_processing.Hide();
            }

        }

        private void button_dwg_to_excel_Click(object sender, EventArgs e)
        {
            if (dt_err0 == null)
            {
                MessageBox.Show("no errors loaded");
                return;
            }
            if (dt_err0.Rows.Count == 0)
            {
                MessageBox.Show("no errors loaded");
                return;
            }

            bool transfer_to_config = false;
            bool is_sta1_or_len = false;
            bool transfer_to_xl = false;

            bool is_new_block = false;
            bool is_new_xl = false;

            List<int> Lista_del = new List<int>();

            if (Freeze_operations == false)
            {
                ObjectId[] Empty_array = null;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                try
                {
                    Freeze_operations = true;
                    _AGEN_mainform.tpage_processing.Show();
                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                            for (int i = 0; i < dt_err0.Rows.Count; ++i)
                            {
                                try
                                {
                                    string id_err = Convert.ToString(dt_err0.Rows[i][0]);
                                    ObjectId id1 = Functions.GetObjectId(ThisDrawing.Database, id_err);
                                    Entity Ent1 = Trans1.GetObject(id1, OpenMode.ForRead) as Entity;
                                    string tag_err = Convert.ToString(dt_err0.Rows[i][3]);
                                    string err_type = Convert.ToString(dt_err0.Rows[i][6]);

                                    if (err_type == "Manual Drafting Adjustments")
                                    {
                                        if (Ent1 is BlockReference)
                                        {
                                            BlockReference block1 = Ent1 as BlockReference;
                                            if (block1 != null)
                                            {
                                                string visib1 = "";
                                                double stretch1 = 0;
                                                if (block1.IsDynamicBlock == true)
                                                {
                                                    using (DynamicBlockReferencePropertyCollection pc = block1.DynamicBlockReferencePropertyCollection)
                                                    {
                                                        foreach (DynamicBlockReferenceProperty prop in pc)
                                                        {
                                                            if (tag_err == "Visibility")
                                                            {
                                                                if (prop.PropertyName == "Visibility1")
                                                                {
                                                                    visib1 = Convert.ToString(prop.Value);
                                                                }
                                                            }
                                                            if (tag_err == "Stretch")
                                                            {
                                                                if (prop.PropertyName == "Distance1")
                                                                {
                                                                    stretch1 = Convert.ToDouble(prop.Value);
                                                                }
                                                            }
                                                        }
                                                    }
                                                }

                                                double x1 = -123.456;
                                                double y1 = -123.456;
                                                if (tag_err == "X Position")
                                                {
                                                    x1 = block1.Position.X;
                                                }
                                                if (tag_err == "Y Position")
                                                {
                                                    y1 = block1.Position.Y;
                                                }

                                                for (int j = 0; j < _AGEN_mainform.dt_config_ownership.Rows.Count; ++j)
                                                {
                                                    string id_config = Convert.ToString(_AGEN_mainform.dt_config_ownership.Rows[j][0]);
                                                    if (id_config == id_err)
                                                    {
                                                        if (visib1 != "") _AGEN_mainform.dt_config_ownership.Rows[j][5] = visib1;
                                                        if (stretch1 > 0) _AGEN_mainform.dt_config_ownership.Rows[j][6] = stretch1;
                                                        if (x1 != -123.456) _AGEN_mainform.dt_config_ownership.Rows[j][3] = x1;
                                                        if (y1 != -123.456) _AGEN_mainform.dt_config_ownership.Rows[j][4] = y1;
                                                        transfer_to_config = true;
                                                        if (Lista_del.Contains(i) == false) Lista_del.Add(i);
                                                        j = _AGEN_mainform.dt_config_ownership.Rows.Count;
                                                    }
                                                }
                                            }
                                        }

                                        if (Ent1 is MText)
                                        {
                                            MText Mtext1 = Ent1 as MText;
                                            if (Mtext1 != null)
                                            {
                                                double x1 = -123.456;
                                                double y1 = -123.456;
                                                if (tag_err == "X Position")
                                                {
                                                    x1 = Mtext1.Location.X;
                                                }
                                                if (tag_err == "Y Position")
                                                {
                                                    y1 = Mtext1.Location.Y;
                                                }

                                                for (int j = 0; j < _AGEN_mainform.dt_config_crossing.Rows.Count; ++j)
                                                {
                                                    string id_config = Convert.ToString(_AGEN_mainform.dt_config_crossing.Rows[j][0]);
                                                    if (id_config == id_err)
                                                    {
                                                        if (x1 != -123.456) _AGEN_mainform.dt_config_crossing.Rows[j][12] = x1;
                                                        if (y1 != -123.456) _AGEN_mainform.dt_config_crossing.Rows[j][13] = y1;
                                                        transfer_to_config = true;
                                                        if (Lista_del.Contains(i) == false) Lista_del.Add(i);
                                                        j = _AGEN_mainform.dt_config_crossing.Rows.Count;
                                                    }
                                                }
                                            }
                                        }
                                    }

                                }
                                catch (System.Exception ex)
                                {
                                    MessageBox.Show(ex.Message + "\r\n" + "No object id found");
                                }
                            }

                            if (transfer_to_config == true)
                            {
                                _AGEN_mainform.tpage_setup.transfera_band_settings_to_config_excel(_AGEN_mainform.dt_config_ownership, _AGEN_mainform.dt_config_crossing);
                            }

                            if (is_sta1_or_len == true)
                            {
                                MessageBox.Show("All data changes containing stationing or length will not be updated in the data tables.\r\nPlease adjust your data accordingly.");
                            }

                            if (is_new_block == true)
                            {
                                MessageBox.Show("You have a new block in the drawing.\r\nPlease adjust your data accordingly.");
                            }

                            if (is_new_xl == true)
                            {
                                MessageBox.Show("You have a new entry in the data table.\r\nPlease redraft the band");
                            }

                            if (transfer_to_xl == true)
                            {
                                string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();

                                if (System.IO.Directory.Exists(ProjF) == false)
                                {
                                    Freeze_operations = false;
                                    MessageBox.Show("the project database folder does not exist");
                                    _AGEN_mainform.tpage_processing.Hide();
                                    return;
                                }

                                if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                                {
                                    ProjF = ProjF + "\\";
                                }

                                if (checkBox_ownership.Checked == true)
                                {
                                    string fisier_prop = ProjF + _AGEN_mainform.property_excel_name;
                                    if (System.IO.File.Exists(fisier_prop) == false)
                                    {
                                        Freeze_operations = false;
                                        MessageBox.Show("the property data file does not exist");
                                        _AGEN_mainform.tpage_processing.Hide();
                                        return;
                                    }

                                    _AGEN_mainform.tpage_owner_scan.Populate_property_file(fisier_prop);
                                }

                            }

                            if (Lista_del.Count > 0)
                            {
                                for (int i = Lista_del.Count - 1; i >= 0; --i)
                                {
                                    dt_err0.Rows[i].Delete();
                                }
                            }

                            make_first_line_invisible();
                            calculate_total_errors();
                            transfer_data_to_panel(dt_err0);

                            Trans1.Commit();

                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                Editor1.SetImpliedSelection(Empty_array);
                Editor1.WriteMessage("\nCommand:");
                Freeze_operations = false;
                _AGEN_mainform.tpage_processing.Hide();
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Functions.extract_description_from_mtext("LXX; 200+12.111 asdfghj");
        }

        private void button_show_tools_Click(object sender, EventArgs e)
        {
            _AGEN_mainform.tpage_processing.Hide();
            _AGEN_mainform.tpage_blank.Hide();
            _AGEN_mainform.tpage_setup.Hide();
            _AGEN_mainform.tpage_viewport_settings.Hide();
            _AGEN_mainform.tpage_tblk_attrib.Hide();
            _AGEN_mainform.tpage_band_analize.Hide();
            _AGEN_mainform.tpage_sheetindex.Hide();
            _AGEN_mainform.tpage_layer_alias.Hide();
            _AGEN_mainform.tpage_crossing_scan.Hide();
            _AGEN_mainform.tpage_crossing_draw.Hide();
            _AGEN_mainform.tpage_profilescan.Hide();
            _AGEN_mainform.tpage_profdraw.Hide();
            _AGEN_mainform.tpage_owner_scan.Hide();
            _AGEN_mainform.tpage_owner_draw.Hide();
            _AGEN_mainform.tpage_mat.Hide();
            _AGEN_mainform.tpage_cust_scan.Hide();
            _AGEN_mainform.tpage_cust_draw.Hide();
            _AGEN_mainform.tpage_sheet_gen.Hide();
          
            _AGEN_mainform.tpage_tools.Show();
        }
    }
}
