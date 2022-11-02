using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System.Windows.Forms;

namespace Alignment_mdi
{
    public class Functions
    {
        public static bool is_dan_popescu()
        {
            if (Environment.UserName.ToUpper() == "POP70694" || Environment.UserName.ToUpper() == "SPI81600")
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
                System.Windows.Forms.MessageBox.Show(ex.Message);
                return 0;
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
            catch (System.Exception)
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


        static public bool IsNumeric(string s)
        {
            double myNum = 0;
            if (Double.TryParse(s, out myNum))
            {
                if (s.Contains(",")) return false;
                return true;
            }
            else
            {
                return false;
            }
        }

        static public void Incarca_existing_layers_to_combobox(System.Windows.Forms.ComboBox Combo_layer)
        {

            {
                try
                {
                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.LayerTable layer_table = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        Combo_layer.Items.Clear();

                        string[] Array1 = null;

                        int idx1 = 1;
                        foreach (ObjectId Layer_id in layer_table)
                        {
                            LayerTableRecord Layer1 = (LayerTableRecord)Trans1.GetObject(Layer_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                            string Name_of_layer = Layer1.Name;
                            if (Name_of_layer.Contains("|") == false & Name_of_layer.Contains("$") == false)
                            {
                                Array.Resize(ref Array1, idx1);
                                Array1[idx1 - 1] = Name_of_layer;
                                idx1 = idx1 + 1;
                            }

                        }
                        if (Array1 != null)
                        {
                            System.Array.Sort(Array1);
                            for (int i = 0; i < Array1.Length; ++i)
                            {
                                Combo_layer.Items.Add(Array1[i]);
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

        }


        static public void Incarca_existing_Blocks_with_attributes_to_combobox(System.Windows.Forms.ComboBox Combo_blockname)
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
                    foreach (ObjectId Block_id in BlockTable_data1)
                    {
                        BlockTableRecord Block1 = (BlockTableRecord)Trans1.GetObject(Block_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        if (Block1.HasAttributeDefinitions == true)
                        {
                            string nume1 = Block1.Name;
                            if (nume1.Contains("*") == false) Combo_blockname.Items.Add(nume1);
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



        static public string Get_String_Rounded(Double Numar, int Nr_dec)
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




        static public double GET_Bearing_rad(Double x1, double y1, double x2, double y2)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;
            CoordinateSystem3d CurentUCS = CurentUCSmatrix.CoordinateSystem3d;
            Plane Planul_curent = new Plane(new Point3d(0, 0, 0), Vector3d.ZAxis);
            return new Point3d(x1, y1, 0).GetVectorTo(new Point3d(x2, y2, 0)).AngleOnPlane(Planul_curent);
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

        static public void add_extra_param_to_dim(RotatedDimension dimension1, Document ThisDrawing)
        {

            dimension1.Dimpost = "<>";//prefix
            dimension1.Dimrnd = 0;
            //Rounds all dimensioning distances to the specified value. 

            dimension1.Dimtxtdirection = false;
            //Specifies the reading direction of the dimension text. 
            dimension1.Dimtofl = false;
            //Initial value: Off (imperial) or On (metric)  
            dimension1.Dimtoh = false;
            //Controls the position of dimension text outside the extension lines. 
            dimension1.Dimtih = false;
            //Initial value: On (imperial) or Off (metric)  
            dimension1.Dimtad = 0;
            //Controls the vertical position of text in relation to the dimension line. 
            dimension1.Dimtvp = 0;
            //Controls the vertical position of dimension text above or below the dimension line. 
            dimension1.Dimsd1 = false;
            //Controls suppression of the first dimension line and arrowhead. 
            dimension1.Dimsd2 = false;
            //Controls suppression of the second dimension line and arrowhead. 
            dimension1.Dimse1 = true; //Suppresses display of the first extension line. 
            dimension1.Dimse2 = true;//Suppresses display of the second extension line
            dimension1.Dimjust = 0;
            //Controls the horizontal positioning of dimension text. 
            dimension1.Dimadec = 0;//Controls the number of precision places displayed in angular dimensions. (0-8)
            dimension1.Dimalt = false; //Controls the display of alternate units in dimensions. Off - Disables alternate units
            dimension1.Dimaltd = 2; //Controls the number of decimal places in alternate units. If DIMALT is turned on, DIMALTD sets the number of digits displayed to the right of the decimal point in the alternate measurement
            dimension1.Dimaltf = 25.4; //Controls the multiplier for alternate units. If DIMALT is turned on, DIMALTF multiplies linear dimensions by a factor to produce a value in an alternate system of measurement. The initial value represents the number of millimeters in an inch.
            dimension1.Dimaltmzf = 100;
            dimension1.Dimaltrnd = 0; //Rounds off the alternate dimension units. 
            dimension1.Dimalttd = 2; //Sets the number of decimal places for the tolerance values in the alternate units of a dimension. 
            dimension1.Dimalttz = 0; //Controls suppression of zeros in tolerance values. 
            dimension1.Dimaltu = 2;//Sets the units format for alternate units of all dimension substyles except Angular. (2 - Decimal)
            dimension1.Dimaltz = 0;//Controls the suppression of zeros for alternate unit dimension values. 
            dimension1.Dimapost = ""; //Specifies a text prefix or suffix (or both) to the alternate dimension measurement for all types of dimensions except angular. 
            dimension1.Dimarcsym = 0; //Controls display of the arc symbol in an arc length dimension. (0- Places arc length symbols before the dimension text )
            dimension1.Dimatfit = 3;
            //Determines how dimension text and arrows are arranged when space is not sufficient to place both within the extension lines. 
            dimension1.Dimaunit = 0;//Sets the units format for angular dimensions. (0 - Decimal degrees)
            dimension1.Dimazin = 0;//Suppresses zeros for angular dimensions. 
            dimension1.Dimsah = false;
            //Controls the display of dimension line arrowhead blocks. 
            dimension1.Dimcen = 0.09; //Controls drawing of circle or arc center marks and centerlines by the DIMCENTER, DIMDIAMETER, and DIMRADIUS commands. 
            dimension1.Dimclrd = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256); // Assigns colors to dimension lines, arrowheads, and dimension leader lines
            dimension1.Dimclre = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256); //Assigns colors to dimension extension lines.
            dimension1.Dimclrt = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256); //Assigns colors to dimension text
            dimension1.Dimdle = 0;//Sets the distance the dimension line extends beyond the extension line when oblique strokes are drawn instead of arrowheads. 
            dimension1.Dimdli = 0.38; //Controls the spacing of the dimension lines in baseline dimensions. 
            //Each dimension line is offset from the previous one by this amount, if necessary, to avoid drawing over it. Changes made with DIMDLI are not applied to existing dimensions
            dimension1.Dimdsep = Convert.ToChar(".");
            //Specifies a single-character decimal separator to use when creating dimensions whose unit format is decimal
            dimension1.Dimexe = 0.18; //Specifies how far to extend the extension line beyond the dimension line. 
            dimension1.Dimexo = 0.0625;//Specifies how far extension lines are offset from origin points. 
            //With fixed-length extension lines, this value determines the minimum offset. 
            dimension1.Dimfrac = 0;//Sets the fraction format when DIMLUNIT is set to 4 (Architectural) or 5 (Fractional).

            dimension1.Dimfxlen = 1;
            dimension1.DimfxlenOn = false;
            dimension1.Dimgap = 0.09;//Sets the distance around the dimension text when the dimension line breaks to accommodate dimension text.
            dimension1.Dimjogang = 0.785398163; //Determines the angle of the transverse segment of the dimension line in a jogged radius dimension. 
            dimension1.Dimlfac = 1;
            //Sets a scale factor for linear dimension measurements. 
            dimension1.Dimltex1 = ThisDrawing.Database.ByBlockLinetype; //Sets the linetype of the first extension line. 
            dimension1.Dimltex2 = ThisDrawing.Database.ByBlockLinetype; //Sets the linetype of the second extension line. 
            dimension1.Dimltype = ThisDrawing.Database.ByBlockLinetype; //Sets the linetype of the dimension line.
            dimension1.Dimlunit = 2;
            //Sets units for all dimension types except Angular. 
            dimension1.Dimlwd = LineWeight.ByBlock;
            //Assigns lineweight to dimension lines. 
            dimension1.Dimlwe = LineWeight.ByBlock;
            //Assigns lineweight to extension  lines. 
            dimension1.Dimmzf = 100;
            dimension1.Dimscale = 1;
            //Sets the overall scale factor applied to dimensioning variables that specify sizes, distances, or offsets. 
            dimension1.Dimtdec = 0;
            //Sets the number of decimal places to display in tolerance values for the primary units in a dimension. 
            dimension1.Dimtfac = 1;
            //Specifies a scale factor for the text height of fractions and tolerance values relative to the dimension text height, as set by DIMTXT. 
            dimension1.Dimtfill = 1;
            //Controls the background of dimension text. 
            dimension1.Dimtfillclr = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByBlock, 0);
            dimension1.Dimtix = false;
            //Draws text between extension lines. 
            dimension1.Dimsoxd = false;
            //Suppresses arrowheads if not enough space is available inside the extension lines. 
            dimension1.Dimtm = 0;
            //Sets the minimum (or lower) tolerance limit for dimension text when DIMTOL or DIMLIM is on. 
            dimension1.Dimtmove = 0;
            //Sets dimension text movement rules. 
            dimension1.Dimtp = 0;
            //Sets the maximum (or upper) tolerance limit for dimension text when DIMTOL or DIMLIM is on. 
            dimension1.Dimlim = false;
            //Generates dimension limits as the default text. 
            dimension1.Dimtol = false;
            //Appends tolerances to dimension text. 
            dimension1.Dimtolj = 1;//Sets the vertical justification for tolerance values relative to the nominal dimension text. 
            dimension1.Dimtsz = 0;
            //Specifies the size of oblique strokes drawn instead of arrowheads for linear, radius, and diameter dimensioning. 
            dimension1.Dimtzin = 0;//Controls the suppression of zeros in tolerance values. 
            dimension1.Dimupt = false; ;
            //Controls options for user-positioned text. 
            dimension1.Dimzin = 0;
            //Controls the suppression of zeros in the primary unit value. 







        }

        static public void add_OD_fieds_to_combobox(ComboBox Combobox_table_name, ComboBox Combobox1)
        {

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;



            Combobox1.Items.Clear();

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
                        Combobox_table_name.Items.Clear();
                    }
                    Trans1.Commit();
                }
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


        static public System.Data.DataTable Populate_data_table_from_excel(System.Data.DataTable dt1, Worksheet W1, int start_row,
          string checkColumn1, string checkColumn2, string checkColumn3, string checkColumn4, string checkColumn5, string checkColumn6, string checkColumn7, string checkColumn8, string checkColumn9, string checkColumn10, string checkColumn11,
          bool show_message)
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
                if (show_message == true)
                {
                    MessageBox.Show("no data found in the file");
                }

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

        public static void Create_od_table(string nume_table, List<string> Lista_field_name, List<Autodesk.Gis.Map.Constants.DataType> Lista_field_type, List<string> Lista_field_description)
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


                        for (int i = 0; i < Lista_field_name.Count; ++i)
                        {
                            List1.Add(Lista_field_name[i]);
                            List2.Add(Lista_field_description[i]);
                            List3.Add(Lista_field_type[i]);
                        }
                        Functions.Get_object_data_table(nume_table, "Generated by OD attach", List1, List2, List3);
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
                                        string Valoare = List_value[i].ToString();
                                        Val.Assign(Valoare);
                                    }
                                    if (List_types[i] == Autodesk.Gis.Map.Constants.DataType.Real)
                                    {
                                        double Valoare = Convert.ToDouble(List_value[i]);
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

        public static Polyline3d Build_3d_poly_for_scanning(System.Data.DataTable dt_cl)
        {

            string Col_x = "X";
            string Col_y = "Y";
            string Col_z = "Z";

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;


            Polyline3d Poly3D = new Polyline3d();
            using (DocumentLock lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {

                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                    BTrecord.AppendEntity(Poly3D);
                    Trans1.AddNewlyCreatedDBObject(Poly3D, true);



                    Poly3D.SetDatabaseDefaults();

                    for (int i = 0; i < dt_cl.Rows.Count; ++i)
                    {
                        double x = 0;
                        double y = 0;
                        double z = 0;

                        if (dt_cl.Rows[i][Col_x] != DBNull.Value)
                        {
                            x = (double)dt_cl.Rows[i][Col_x];
                        }

                        if (dt_cl.Rows[i][Col_y] != DBNull.Value)
                        {
                            y = (double)dt_cl.Rows[i][Col_y];
                        }

                        if (dt_cl.Rows[i][Col_z] != DBNull.Value)
                        {
                            z = (double)dt_cl.Rows[i][Col_z];
                        }


                        PolylineVertex3d Vertex_new = new PolylineVertex3d(new Point3d(x, y, z));
                        Poly3D.AppendVertex(Vertex_new);
                        Trans1.AddNewlyCreatedDBObject(Vertex_new, true);

                    }

                    Trans1.Commit();
                }
            }
            return Poly3D;

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
    }
}
