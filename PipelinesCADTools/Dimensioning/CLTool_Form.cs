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
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.BoundaryRepresentation;


namespace Dimensioning
{
    public partial class CLTool_Form : Form
    {

        private bool clickdragdown;
        private Point lastLocation;
        bool Freeze_operations = false;


        System.Data.DataTable dt_frozen_xref_layers = null;
        List<string> Lista_layers_from_xrefs = null;


        System.Data.DataTable dtc = null;
        string layer_name = "0";
        short layer_color = 7;
        double Ltscale = 1;
        string linetype_name = "Continuous";
        double Extension_value = 0;


        public CLTool_Form()
        {
            InitializeComponent();
            Load_clients_to_combobox();
            creaza_line_types_file();
        }

        private void clickmove_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            clickdragdown = true;
            lastLocation = e.Location;
        }

        private void clickmove_MouseMove(object sender, MouseEventArgs e)
        {
            if (clickdragdown)
            {
                this.Location = new Point(
                  (this.Location.X - lastLocation.X) + e.X, (this.Location.Y - lastLocation.Y) + e.Y);

                this.Update();
            }
        }

        private void clickmove_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            clickdragdown = false;

        }
        private void button_Exit_Click(object sender, EventArgs e)
        {
            if (dtc != null)
            {
                if (dtc.Rows.Count > 0)
                {

                    for (int i = 0; i < dtc.Rows.Count; ++i)
                    {
                        if (dtc.Rows[i][0] != DBNull.Value)
                        {
                            String client1 = dtc.Rows[i][0].ToString();
                            if (String.IsNullOrEmpty(client1) == false)
                            {
                                string Fisier = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + client1 + ".lin";
                                if (System.IO.File.Exists(Fisier) == true)
                                {
                                    System.IO.File.Delete(Fisier);
                                }
                            }
                        }
                    }

                }
            }

            this.Close();
        }

        private DBObjectCollection get_xrefs_centerlines()
        {

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                BlockTableRecord BTrecord = Trans1.GetObject(BlockTable1[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                ObjectIdCollection Col_btr_id = new ObjectIdCollection();
                ObjectIdCollection Col_xref_id = new ObjectIdCollection();

                foreach (ObjectId id1 in BTrecord)
                {
                    BlockReference Xr1 = Trans1.GetObject(id1, OpenMode.ForRead) as BlockReference;
                    if (Xr1 != null)
                    {
                        BlockTableRecord btr1 = Trans1.GetObject(Xr1.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                        if (btr1 != null)
                        {
                            if (btr1.IsFromExternalReference == true)
                            {
                                Col_btr_id.Add(btr1.ObjectId);
                                Col_xref_id.Add(Xr1.ObjectId);
                            }
                        }
                    }
                }

                //System.Collections.Specialized.StringCollection col_tempfiles = new System.Collections.Specialized.StringCollection();

                DBObjectCollection Col_xref = new DBObjectCollection();

                using (XrefGraph XRgraph1 = ThisDrawing.Database.GetHostDwgXrefGraph(false))
                {

                    Lista_layers_from_xrefs = new List<string>();

                    for (int i = 0; i < XRgraph1.NumNodes; ++i)
                    {
                        XrefGraphNode xNode = XRgraph1.GetXrefNode(i) as XrefGraphNode;

                        if (xNode != null)
                        {
                            if (xNode.Database != null)
                            {
                                if (xNode.Database.Filename.Equals(ThisDrawing.Database.Filename) == false)
                                {

                                    if (xNode.XrefStatus == XrefStatus.Resolved)
                                    {
                                        if (Col_btr_id.Contains(xNode.BlockTableRecordId) == true)
                                        {

                                            int incr = 0;

                                            int index1 = Col_btr_id.IndexOf(xNode.BlockTableRecordId);
                                            BlockReference blockXref1 = Trans1.GetObject(Col_xref_id[index1], OpenMode.ForRead) as BlockReference;
                                            BlockTableRecord btr_xr1 = Trans1.GetObject(blockXref1.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;

                                            foreach (ObjectId xrefid in btr_xr1)
                                            {
                                                BlockReference Block_child = Trans1.GetObject(xrefid, OpenMode.ForRead) as BlockReference;
                                                if (Block_child != null)
                                                {
                                                    BlockTableRecord btr_xrC = Trans1.GetObject(Block_child.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                                                    if (btr_xrC.IsFromExternalReference == true && btr_xrC.IsResolved == true)
                                                    {
                                                        XrefGraphNode xNode1 = XRgraph1.GetXrefNode(btr_xrC.Name) as XrefGraphNode;
                                                        string XrefName_nested = xNode1.Database.Filename;

                                                        //string XrefName_nested_copy = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\temp" + incr.ToString() + ".dwg";

                                                        // do
                                                        // {
                                                        //incr = incr + 1;
                                                        //XrefName_nested_copy = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\temp" + incr.ToString() + ".dwg";
                                                        // } while (System.IO.File.Exists(XrefName_nested_copy) == true);

                                                        // System.IO.File.Copy(XrefName_nested, XrefName_nested_copy);

                                                        using (Database database_c2 = new Database(false, true))
                                                        {


                                                            database_c2.ReadDwgFile(XrefName_nested, FileOpenMode.OpenForReadAndAllShare, false, null);
                                                            HostApplicationServices.WorkingDatabase = database_c2;

                                                            string Layer_prefix = System.IO.Path.GetFileNameWithoutExtension(XrefName_nested) + "|";

                                                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Transc2 = database_c2.TransactionManager.StartTransaction())
                                                            {
                                                                Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data_c2 = (BlockTable)database_c2.BlockTableId.GetObject(OpenMode.ForRead);
                                                                Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecordMS_c2 = (BlockTableRecord)Transc2.GetObject(BlockTable_data_c2[BlockTableRecord.ModelSpace], Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                                                                List<Circle> Col_circ_id = new List<Circle>();
                                                                List<string> lista_layers_circle = new List<string>();

                                                                foreach (ObjectId id1 in BTrecordMS_c2)
                                                                {
                                                                    DBObjectCollection Col_exploded = new DBObjectCollection();
                                                                    Solid3d SSolid3d1 = Transc2.GetObject(id1, OpenMode.ForRead) as Solid3d;
                                                                    Line LLine1 = Transc2.GetObject(id1, OpenMode.ForRead) as Line;
                                                                    Polyline PolyLine1 = Transc2.GetObject(id1, OpenMode.ForRead) as Polyline;
                                                                    Arc AArc1 = Transc2.GetObject(id1, OpenMode.ForRead) as Arc;
                                                                    try
                                                                    {
                                                                        if (SSolid3d1 != null)
                                                                        {

                                                                            SSolid3d1.Explode(Col_exploded);

                                                                            if (Col_exploded.Count > 0)
                                                                            {
                                                                                foreach (DBObject Obj1 in Col_exploded)
                                                                                {
                                                                                    Line l1 = Obj1 as Line;
                                                                                    if (l1 != null)
                                                                                    {
                                                                                        Line l2 = new Line();
                                                                                        l2 = l1.Clone() as Line;
                                                                                        if (l2 != null)
                                                                                        {
                                                                                            l2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(Block_child.Position)));
                                                                                            l2.TransformBy(Matrix3d.Rotation(Block_child.Rotation, Vector3d.ZAxis, Block_child.Position));
                                                                                            l2.TransformBy(Matrix3d.Scaling(Block_child.ScaleFactors.X, Block_child.Position));

                                                                                            l2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(blockXref1.Position)));
                                                                                            l2.TransformBy(Matrix3d.Rotation(blockXref1.Rotation, Vector3d.ZAxis, blockXref1.Position));
                                                                                            l2.TransformBy(Matrix3d.Scaling(blockXref1.ScaleFactors.X, blockXref1.Position));

                                                                                            Col_xref.Add(l2);
                                                                                            Lista_layers_from_xrefs.Add(Layer_prefix + SSolid3d1.Layer);

                                                                                        }
                                                                                    }

                                                                                    Polyline pl1 = Obj1 as Polyline;
                                                                                    if (pl1 != null)
                                                                                    {
                                                                                        Line pl2 = new Line(PolyLine1.GetPointAtParameter(0), PolyLine1.GetPointAtParameter(PolyLine1.NumberOfVertices - 1));

                                                                                        if (pl2 != null)
                                                                                        {

                                                                                            pl2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(Block_child.Position)));
                                                                                            pl2.TransformBy(Matrix3d.Rotation(Block_child.Rotation, Vector3d.ZAxis, Block_child.Position));
                                                                                            pl2.TransformBy(Matrix3d.Scaling(Block_child.ScaleFactors.X, Block_child.Position));

                                                                                            pl2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(blockXref1.Position)));
                                                                                            pl2.TransformBy(Matrix3d.Rotation(blockXref1.Rotation, Vector3d.ZAxis, blockXref1.Position));
                                                                                            pl2.TransformBy(Matrix3d.Scaling(blockXref1.ScaleFactors.X, blockXref1.Position));

                                                                                            Col_xref.Add(pl2);
                                                                                            Lista_layers_from_xrefs.Add(Layer_prefix + SSolid3d1.Layer);
                                                                                        }
                                                                                    }

                                                                                    Arc a1 = Obj1 as Arc;
                                                                                    if (a1 != null)
                                                                                    {
                                                                                        Arc a2 = new Arc();
                                                                                        a2 = a1.Clone() as Arc;
                                                                                        if (a2 != null)
                                                                                        {
                                                                                            a2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(Block_child.Position)));
                                                                                            a2.TransformBy(Matrix3d.Rotation(Block_child.Rotation, Vector3d.ZAxis, Block_child.Position));
                                                                                            a2.TransformBy(Matrix3d.Scaling(Block_child.ScaleFactors.X, Block_child.Position));

                                                                                            a2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(blockXref1.Position)));
                                                                                            a2.TransformBy(Matrix3d.Rotation(blockXref1.Rotation, Vector3d.ZAxis, blockXref1.Position));
                                                                                            a2.TransformBy(Matrix3d.Scaling(blockXref1.ScaleFactors.X, blockXref1.Position));
                                                                                            Col_xref.Add(a2);
                                                                                            Lista_layers_from_xrefs.Add(Layer_prefix + SSolid3d1.Layer);
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }



                                                                            IntPtr pSubentityIdPE = SSolid3d1.QueryX(AssocPersSubentityIdPE.GetClass(typeof(AssocPersSubentityIdPE)));
                                                                            if (pSubentityIdPE != IntPtr.Zero)
                                                                            {
                                                                                AssocPersSubentityIdPE subentityIdPE = AssocPersSubentityIdPE.Create(pSubentityIdPE, false) as AssocPersSubentityIdPE;
                                                                                SubentityId[] edgeIds = subentityIdPE.GetAllSubentities(SSolid3d1, SubentityType.Edge);

                                                                                ObjectId[] ids = new ObjectId[] { SSolid3d1.ObjectId };
                                                                                foreach (SubentityId subentId in edgeIds)
                                                                                {
                                                                                    FullSubentityPath path = new FullSubentityPath(ids, subentId);

                                                                                    Entity edgeEntity = SSolid3d1.GetSubentity(path);

                                                                                    if (edgeEntity != null)
                                                                                    {
                                                                                        Circle cc1 = new Circle();
                                                                                        cc1 = edgeEntity as Circle;

                                                                                        if (cc1 != null)
                                                                                        {
                                                                                            Col_circ_id.Add(cc1);

                                                                                            lista_layers_circle.Add(Layer_prefix + SSolid3d1.Layer);

                                                                                        }
                                                                                    }
                                                                                }
                                                                            }
                                                                        }

                                                                        Col_exploded = new DBObjectCollection();

                                                                        if (LLine1 != null)
                                                                        {
                                                                            if (LLine1.Layer.ToLower() == "cl")
                                                                            {
                                                                                Col_exploded.Add(LLine1);

                                                                            }
                                                                        }

                                                                        if (PolyLine1 != null)
                                                                        {
                                                                            if (PolyLine1.Layer.ToLower() == "cl")
                                                                            {
                                                                                Col_exploded.Add(PolyLine1);
                                                                            }
                                                                        }

                                                                        if (AArc1 != null)
                                                                        {
                                                                            if (AArc1.Layer.ToLower() == "cl")
                                                                            {
                                                                                Col_exploded.Add(AArc1);
                                                                            }
                                                                        }

                                                                        if (Col_exploded.Count > 0)
                                                                        {
                                                                            foreach (DBObject Obj1 in Col_exploded)
                                                                            {
                                                                                Line l1 = Obj1 as Line;
                                                                                if (l1 != null)
                                                                                {
                                                                                    Line l2 = new Line();
                                                                                    l2 = l1.Clone() as Line;
                                                                                    if (l2 != null)
                                                                                    {
                                                                                        l2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(Block_child.Position)));
                                                                                        l2.TransformBy(Matrix3d.Rotation(Block_child.Rotation, Vector3d.ZAxis, Block_child.Position));
                                                                                        l2.TransformBy(Matrix3d.Scaling(Block_child.ScaleFactors.X, Block_child.Position));

                                                                                        l2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(blockXref1.Position)));
                                                                                        l2.TransformBy(Matrix3d.Rotation(blockXref1.Rotation, Vector3d.ZAxis, blockXref1.Position));
                                                                                        l2.TransformBy(Matrix3d.Scaling(blockXref1.ScaleFactors.X, blockXref1.Position));

                                                                                        Col_xref.Add(l2);
                                                                                        Lista_layers_from_xrefs.Add(Layer_prefix + l1.Layer);

                                                                                    }
                                                                                }

                                                                                Polyline pl1 = Obj1 as Polyline;
                                                                                if (pl1 != null)
                                                                                {
                                                                                    Line pl2 = new Line(PolyLine1.GetPointAtParameter(0), PolyLine1.GetPointAtParameter(PolyLine1.NumberOfVertices - 1));

                                                                                    if (pl2 != null)
                                                                                    {

                                                                                        pl2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(Block_child.Position)));
                                                                                        pl2.TransformBy(Matrix3d.Rotation(Block_child.Rotation, Vector3d.ZAxis, Block_child.Position));
                                                                                        pl2.TransformBy(Matrix3d.Scaling(Block_child.ScaleFactors.X, Block_child.Position));

                                                                                        pl2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(blockXref1.Position)));
                                                                                        pl2.TransformBy(Matrix3d.Rotation(blockXref1.Rotation, Vector3d.ZAxis, blockXref1.Position));
                                                                                        pl2.TransformBy(Matrix3d.Scaling(blockXref1.ScaleFactors.X, blockXref1.Position));

                                                                                        Col_xref.Add(pl2);
                                                                                        Lista_layers_from_xrefs.Add(Layer_prefix + pl1.Layer);
                                                                                    }
                                                                                }

                                                                                Arc a1 = Obj1 as Arc;
                                                                                if (a1 != null)
                                                                                {
                                                                                    Arc a2 = new Arc();
                                                                                    a2 = a1.Clone() as Arc;
                                                                                    if (a2 != null)
                                                                                    {
                                                                                        a2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(Block_child.Position)));
                                                                                        a2.TransformBy(Matrix3d.Rotation(Block_child.Rotation, Vector3d.ZAxis, Block_child.Position));
                                                                                        a2.TransformBy(Matrix3d.Scaling(Block_child.ScaleFactors.X, Block_child.Position));

                                                                                        a2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(blockXref1.Position)));
                                                                                        a2.TransformBy(Matrix3d.Rotation(blockXref1.Rotation, Vector3d.ZAxis, blockXref1.Position));
                                                                                        a2.TransformBy(Matrix3d.Scaling(blockXref1.ScaleFactors.X, blockXref1.Position));
                                                                                        Col_xref.Add(a2);
                                                                                        Lista_layers_from_xrefs.Add(Layer_prefix + a1.Layer);
                                                                                    }
                                                                                }
                                                                            }
                                                                        }

                                                                    }
                                                                    catch (System.Exception ex)
                                                                    {

                                                                    }
                                                                }
                                                                if (Col_circ_id.Count > 0)
                                                                {
                                                                    for (int j = 0; j < Col_circ_id.Count; ++j)
                                                                    {

                                                                        Circle cerc1 = Col_circ_id[j] as Circle;
                                                                        if (cerc1 != null)
                                                                        {
                                                                            if (cerc1 != null)
                                                                            {
                                                                                cerc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(Block_child.Position)));
                                                                                cerc1.TransformBy(Matrix3d.Rotation(Block_child.Rotation, Vector3d.ZAxis, Block_child.Position));
                                                                                cerc1.TransformBy(Matrix3d.Scaling(Block_child.ScaleFactors.X, Block_child.Position));

                                                                                cerc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(blockXref1.Position)));
                                                                                cerc1.TransformBy(Matrix3d.Rotation(blockXref1.Rotation, Vector3d.ZAxis, blockXref1.Position));
                                                                                cerc1.TransformBy(Matrix3d.Scaling(blockXref1.ScaleFactors.X, blockXref1.Position));
                                                                                Col_xref.Add(cerc1);
                                                                                Lista_layers_from_xrefs.Add(lista_layers_circle[j]);

                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }

                                                            HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                                                        }

                                                        //col_tempfiles.Add(XrefName_nested_copy);
                                                    }
                                                }
                                            }

                                            string XrefName_original = xNode.Database.Filename;

                                            // string XrefName_copy = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\temp" + incr.ToString() + ".dwg";
                                            //do
                                            //  {
                                            //incr = incr + 1;
                                            // XrefName_copy = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\temp" + incr.ToString() + ".dwg";
                                            //  } while (System.IO.File.Exists(XrefName_copy) == true);

                                            //System.IO.File.Copy(XrefName_original, XrefName_copy);

                                            using (Database database2 = new Database(false, true))
                                            {
                                                database2.ReadDwgFile(XrefName_original, FileOpenMode.OpenForReadAndAllShare, false, null);
                                                HostApplicationServices.WorkingDatabase = database2;

                                                string Layer_prefix = System.IO.Path.GetFileNameWithoutExtension(XrefName_original) + "|";

                                                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans2 = database2.TransactionManager.StartTransaction())
                                                {
                                                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data2 = (BlockTable)database2.BlockTableId.GetObject(OpenMode.ForRead);
                                                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecordMS2 = (BlockTableRecord)Trans2.GetObject(BlockTable_data2[BlockTableRecord.ModelSpace], Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                                                    List<Circle> Col_circ_id = new List<Circle>();
                                                    List<string> lista_layers_circle = new List<string>();


                                                    foreach (ObjectId id1 in BTrecordMS2)
                                                    {
                                                        DBObjectCollection Col_exploded = new DBObjectCollection();

                                                        Solid3d SSolid3d1 = Trans2.GetObject(id1, OpenMode.ForRead) as Solid3d;
                                                        Line LLine1 = Trans2.GetObject(id1, OpenMode.ForRead) as Line;
                                                        Polyline PolyLine1 = Trans2.GetObject(id1, OpenMode.ForRead) as Polyline;
                                                        Arc AArc1 = Trans2.GetObject(id1, OpenMode.ForRead) as Arc;
                                                        try
                                                        {
                                                            if (SSolid3d1 != null)
                                                            {
                                                                SSolid3d1.Explode(Col_exploded);

                                                                if (Col_exploded.Count > 0)
                                                                {
                                                                    foreach (DBObject Obj1 in Col_exploded)
                                                                    {
                                                                        Line l1 = Obj1 as Line;
                                                                        if (l1 != null)
                                                                        {
                                                                            Line l2 = new Line();
                                                                            l2 = l1.Clone() as Line;
                                                                            if (l2 != null)
                                                                            {
                                                                                l2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(blockXref1.Position)));
                                                                                l2.TransformBy(Matrix3d.Rotation(blockXref1.Rotation, Vector3d.ZAxis, blockXref1.Position));
                                                                                l2.TransformBy(Matrix3d.Scaling(blockXref1.ScaleFactors.X, blockXref1.Position));
                                                                                Col_xref.Add(l2);
                                                                                Lista_layers_from_xrefs.Add(Layer_prefix + SSolid3d1.Layer);
                                                                            }
                                                                        }

                                                                        Polyline pl1 = Obj1 as Polyline;
                                                                        if (pl1 != null)
                                                                        {
                                                                            Line pl2 = new Line(PolyLine1.GetPointAtParameter(0), PolyLine1.GetPointAtParameter(PolyLine1.NumberOfVertices - 1));

                                                                            if (pl2 != null)
                                                                            {
                                                                                pl2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(blockXref1.Position)));
                                                                                pl2.TransformBy(Matrix3d.Rotation(blockXref1.Rotation, Vector3d.ZAxis, blockXref1.Position));
                                                                                pl2.TransformBy(Matrix3d.Scaling(blockXref1.ScaleFactors.X, blockXref1.Position));
                                                                                Col_xref.Add(pl2);
                                                                                Lista_layers_from_xrefs.Add(Layer_prefix + SSolid3d1.Layer);
                                                                            }
                                                                        }

                                                                        Arc a1 = Obj1 as Arc;
                                                                        if (a1 != null)
                                                                        {
                                                                            Arc a2 = new Arc();
                                                                            a2 = a1.Clone() as Arc;
                                                                            if (a2 != null)
                                                                            {
                                                                                a2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(blockXref1.Position)));
                                                                                a2.TransformBy(Matrix3d.Rotation(blockXref1.Rotation, Vector3d.ZAxis, blockXref1.Position));
                                                                                a2.TransformBy(Matrix3d.Scaling(blockXref1.ScaleFactors.X, blockXref1.Position));
                                                                                Col_xref.Add(a2);
                                                                                Lista_layers_from_xrefs.Add(Layer_prefix + SSolid3d1.Layer);
                                                                            }
                                                                        }
                                                                    }
                                                                }


                                                                IntPtr pSubentityIdPE = SSolid3d1.QueryX(AssocPersSubentityIdPE.GetClass(typeof(AssocPersSubentityIdPE)));
                                                                if (pSubentityIdPE != IntPtr.Zero)
                                                                {
                                                                    AssocPersSubentityIdPE subentityIdPE = AssocPersSubentityIdPE.Create(pSubentityIdPE, false) as AssocPersSubentityIdPE;
                                                                    SubentityId[] edgeIds = subentityIdPE.GetAllSubentities(SSolid3d1, SubentityType.Edge);

                                                                    ObjectId[] ids = new ObjectId[] { SSolid3d1.ObjectId };
                                                                    foreach (SubentityId subentId in edgeIds)
                                                                    {
                                                                        FullSubentityPath path = new FullSubentityPath(ids, subentId);

                                                                        Entity edgeEntity = SSolid3d1.GetSubentity(path);

                                                                        if (edgeEntity != null)
                                                                        {
                                                                            Circle cc1 = new Circle();
                                                                            cc1 = edgeEntity as Circle;

                                                                            if (cc1 != null)
                                                                            {
                                                                                Col_circ_id.Add(cc1);
                                                                                lista_layers_circle.Add(Layer_prefix + SSolid3d1.Layer);
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }

                                                            Col_exploded = new DBObjectCollection();

                                                            if (LLine1 != null)
                                                            {
                                                                if (LLine1.Layer.ToLower() == "cl")
                                                                {
                                                                    Col_exploded.Add(LLine1);
                                                                }
                                                            }

                                                            if (PolyLine1 != null)
                                                            {
                                                                if (PolyLine1.Layer.ToLower() == "cl")
                                                                {
                                                                    Col_exploded.Add(PolyLine1);
                                                                }
                                                            }

                                                            if (AArc1 != null)
                                                            {
                                                                if (AArc1.Layer.ToLower() == "cl")
                                                                {
                                                                    Col_exploded.Add(AArc1);
                                                                }
                                                            }

                                                            if (Col_exploded.Count > 0)
                                                            {
                                                                foreach (DBObject Obj1 in Col_exploded)
                                                                {

                                                                    Line l1 = Obj1 as Line;
                                                                    if (l1 != null)
                                                                    {
                                                                        Line l2 = new Line();
                                                                        l2 = l1.Clone() as Line;
                                                                        if (l2 != null)
                                                                        {
                                                                            l2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(blockXref1.Position)));
                                                                            l2.TransformBy(Matrix3d.Rotation(blockXref1.Rotation, Vector3d.ZAxis, blockXref1.Position));
                                                                            l2.TransformBy(Matrix3d.Scaling(blockXref1.ScaleFactors.X, blockXref1.Position));
                                                                            Col_xref.Add(l2);
                                                                            Lista_layers_from_xrefs.Add(Layer_prefix + l1.Layer);
                                                                        }
                                                                    }

                                                                    Polyline pl1 = Obj1 as Polyline;
                                                                    if (pl1 != null)
                                                                    {
                                                                        Line l2 = new Line(PolyLine1.GetPointAtParameter(0), PolyLine1.GetPointAtParameter(PolyLine1.NumberOfVertices - 1));

                                                                        if (l2 != null)
                                                                        {
                                                                            l2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(blockXref1.Position)));
                                                                            l2.TransformBy(Matrix3d.Rotation(blockXref1.Rotation, Vector3d.ZAxis, blockXref1.Position));
                                                                            l2.TransformBy(Matrix3d.Scaling(blockXref1.ScaleFactors.X, blockXref1.Position));
                                                                            Col_xref.Add(l2);
                                                                            Lista_layers_from_xrefs.Add(Layer_prefix + pl1.Layer);
                                                                        }
                                                                    }

                                                                    Arc a1 = Obj1 as Arc;
                                                                    if (a1 != null)
                                                                    {
                                                                        Arc a2 = new Arc();
                                                                        a2 = a1.Clone() as Arc;
                                                                        if (a2 != null)
                                                                        {
                                                                            a2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(blockXref1.Position)));
                                                                            a2.TransformBy(Matrix3d.Rotation(blockXref1.Rotation, Vector3d.ZAxis, blockXref1.Position));
                                                                            a2.TransformBy(Matrix3d.Scaling(blockXref1.ScaleFactors.X, blockXref1.Position));
                                                                            Col_xref.Add(a2);
                                                                            Lista_layers_from_xrefs.Add(Layer_prefix + a1.Layer);
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        catch (System.Exception ex)
                                                        {

                                                        }
                                                    }

                                                    if (Col_circ_id.Count > 0)
                                                    {
                                                        for (int j = 0; j < Col_circ_id.Count; ++j)
                                                        {
                                                            Circle cerc1 = Col_circ_id[j] as Circle;
                                                            if (cerc1 != null)
                                                            {
                                                                if (cerc1 != null)
                                                                {
                                                                    cerc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(blockXref1.Position)));
                                                                    cerc1.TransformBy(Matrix3d.Rotation(blockXref1.Rotation, Vector3d.ZAxis, blockXref1.Position));
                                                                    cerc1.TransformBy(Matrix3d.Scaling(blockXref1.ScaleFactors.X, blockXref1.Position));
                                                                    Col_xref.Add(cerc1);
                                                                    Lista_layers_from_xrefs.Add(lista_layers_circle[j]);
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                                            }
                                            //col_tempfiles.Add(XrefName_copy);
                                            // Col_xref_id.Add(xNode.BlockTableRecordId);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                return Col_xref;
            }
        }



        private void button_GenerateCL_Click(object sender, EventArgs e)

        {
            if (dtc == null)
            {
                MessageBox.Show("no clients file found \r\naborting operation");
                return;
            }
            if (dtc.Rows.Count == 0)
            {
                MessageBox.Show("no clients file found \r\naborting operation");
                return;
            }
            ObjectId[] Empty_array = null;

            if (Freeze_operations == false)
            {
                Freeze_operations = true;
                try
                {
                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    Editor1.SetImpliedSelection(Empty_array);
                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {

                        int Tilemode1 = Convert.ToInt32(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("TILEMODE"));
                        int CVport1 = Convert.ToInt32(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CVPORT"));

                        if (Tilemode1 == 0)
                        {
                            if (CVport1 != 1)
                            {
                                Editor1.SwitchToPaperSpace();
                            }
                        }
                        else
                        {
                            MessageBox.Show("this option is meant to work in paper space");
                            Freeze_operations = false;
                            return;
                        }

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_Viewport;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_viewport;
                        Prompt_viewport = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the viewport:");
                        Prompt_viewport.SetRejectMessage("\nSelect a viewport!");
                        Prompt_viewport.AllowNone = true;
                        Prompt_viewport.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Viewport), false);
                        Prompt_viewport.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);

                        Polyline clip_poly = null;

                        Rezultat_Viewport = ThisDrawing.Editor.GetEntity(Prompt_viewport);

                        if (Rezultat_Viewport.Status != PromptStatus.OK)
                        {
                            MessageBox.Show("this requires you to select a viewport \r\nor to select the clipping polyline of the viewport");
                            Freeze_operations = false;
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        string client1 = comboBox_Client.Text;
                        bool isfound = false;
                        if (String.IsNullOrEmpty(client1) == false)
                        {
                            for (int i = 0; i < dtc.Rows.Count; ++i)
                            {
                                if (isfound == false)
                                {
                                    string client0 = "";
                                    if (dtc.Rows[i][0] != DBNull.Value)
                                    {
                                        client0 = dtc.Rows[i][0].ToString();
                                        if (client0 == client1)
                                        {

                                            if (dtc.Rows[i][7] != DBNull.Value)
                                            {
                                                linetype_name = dtc.Rows[i][7].ToString();
                                            }
                                            if (dtc.Rows[i][1] != DBNull.Value)
                                            {
                                                layer_name = dtc.Rows[i][1].ToString();
                                            }
                                            if (dtc.Rows[i][2] != DBNull.Value)
                                            {
                                                string val0 = dtc.Rows[i][2].ToString();
                                                if (Functions.IsNumeric(val0) == true)
                                                {
                                                    layer_color = Convert.ToInt16(val0);
                                                }
                                            }
                                            if (dtc.Rows[i][3] != DBNull.Value)
                                            {
                                                string val0 = dtc.Rows[i][3].ToString();
                                                if (Functions.IsNumeric(val0) == true)
                                                {
                                                    Ltscale = Convert.ToDouble(val0);
                                                }
                                            }

                                            if (linetype_name != "Continuous")
                                            {
                                                Incarca_line_type(linetype_name, Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + comboBox_Client.Text + ".lin", true);
                                            }

                                            if (dtc.Rows[i][6] != DBNull.Value)
                                            {

                                                Extension_value = Convert.ToDouble(dtc.Rows[i][6]);
                                            }

                                            isfound = true;
                                            i = dtc.Rows.Count;
                                        }
                                    }

                                }

                            }
                        }

                        if (isfound == false)
                        {
                            MessageBox.Show("linetype issue\r\naborting operation");
                            Freeze_operations = false;
                            return;
                        }



                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                            Viewport Ent_vp = Trans1.GetObject(Rezultat_Viewport.ObjectId, OpenMode.ForRead) as Viewport;
                            if (Ent_vp == null)
                            {
                                clip_poly = Trans1.GetObject(Rezultat_Viewport.ObjectId, OpenMode.ForRead) as Polyline;
                                if (clip_poly != null)
                                {
                                    ObjectId vpId1 = LayoutManager.Current.GetNonRectangularViewportIdFromClipId(Rezultat_Viewport.ObjectId);
                                    if (Trans1.GetObject(vpId1, OpenMode.ForRead) is Viewport)
                                    {
                                        Ent_vp = Trans1.GetObject(vpId1, OpenMode.ForRead) as Viewport;
                                        List<string> Lista_frozen = get_frozen_or_off_layers_from_viewport(Ent_vp.ObjectId);

                                        dt_frozen_xref_layers = new System.Data.DataTable();
                                        dt_frozen_xref_layers.Columns.Add("ln", typeof(string));
                                        dt_frozen_xref_layers.Columns.Add("id", typeof(string));

                                        foreach (string layer1 in Lista_frozen)
                                        {
                                            dt_frozen_xref_layers.Rows.Add();
                                            dt_frozen_xref_layers.Rows[dt_frozen_xref_layers.Rows.Count - 1]["ln"] = layer1;
                                            dt_frozen_xref_layers.Rows[dt_frozen_xref_layers.Rows.Count - 1]["id"] = Ent_vp.ObjectId.Handle.Value.ToString();
                                        }
                                    }
                                }
                            }
                            else
                            {
                                List<string> Lista_frozen = get_frozen_or_off_layers_from_viewport(Ent_vp.ObjectId);

                                dt_frozen_xref_layers = new System.Data.DataTable();
                                dt_frozen_xref_layers.Columns.Add("ln", typeof(string));
                                dt_frozen_xref_layers.Columns.Add("id", typeof(string));

                                foreach (string layer1 in Lista_frozen)
                                {
                                    dt_frozen_xref_layers.Rows.Add();
                                    dt_frozen_xref_layers.Rows[dt_frozen_xref_layers.Rows.Count - 1]["ln"] = layer1;
                                    dt_frozen_xref_layers.Rows[dt_frozen_xref_layers.Rows.Count - 1]["id"] = Ent_vp.ObjectId.Handle.Value.ToString();
                                }
                            }


                            if (clip_poly == null)
                            {
                                clip_poly = new Polyline();
                                clip_poly.AddVertexAt(0, new Point2d(Ent_vp.CenterPoint.X - Ent_vp.Width / 2, Ent_vp.CenterPoint.Y - Ent_vp.Height / 2), 0, 0, 0);
                                clip_poly.AddVertexAt(1, new Point2d(Ent_vp.CenterPoint.X - Ent_vp.Width / 2, Ent_vp.CenterPoint.Y + Ent_vp.Height / 2), 0, 0, 0);
                                clip_poly.AddVertexAt(2, new Point2d(Ent_vp.CenterPoint.X + Ent_vp.Width / 2, Ent_vp.CenterPoint.Y + Ent_vp.Height / 2), 0, 0, 0);
                                clip_poly.AddVertexAt(3, new Point2d(Ent_vp.CenterPoint.X + Ent_vp.Width / 2, Ent_vp.CenterPoint.Y - Ent_vp.Height / 2), 0, 0, 0);
                                clip_poly.Closed = true;
                            }



                            if (Ent_vp != null)
                            {
                                Functions.Creaza_layer_with_linetype(layer_name, layer_color, linetype_name, true);
                                DBObjectCollection Col_xref = get_xrefs_centerlines();
                                GenerateCL(Ent_vp.ObjectId, clip_poly, Col_xref);
                                Editor1.SetImpliedSelection(Empty_array);
                                Trans1.Commit();
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

        private List<string> get_frozen_or_off_layers_from_viewport(ObjectId vpid)
        {
            List<string> lista1 = new List<string>();

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                Viewport Vp = Trans1.GetObject(vpid, OpenMode.ForRead) as Viewport;
                ResultBuffer Buffer1 = Vp.XData;
                if (Buffer1 != null)
                {
                    foreach (TypedValue tv in Buffer1)
                    {
                        Int16 typecode1 = tv.TypeCode;
                        if (typecode1 == 1003)
                        {
                            lista1.Add(tv.Value.ToString());
                        }
                    }
                }
                Trans1.Dispose();
            }
            return lista1;
        }


        private void button_cl_from_all_viewports_Click(object sender, EventArgs e)
        {
            if (dtc == null)
            {
                MessageBox.Show("no clients file found \r\naborting operation");
                return;
            }
            if (dtc.Rows.Count == 0)
            {
                MessageBox.Show("no clients file found \r\naborting operation");
                return;
            }
            ObjectId[] Empty_array = null;

            if (Freeze_operations == false)
            {
                Freeze_operations = true;
                try
                {
                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    Editor1.SetImpliedSelection(Empty_array);
                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {

                        int Tilemode1 = Convert.ToInt32(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("TILEMODE"));
                        int CVport1 = Convert.ToInt32(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CVPORT"));

                        if (Tilemode1 == 0)
                        {
                            if (CVport1 != 1)
                            {
                                Editor1.SwitchToPaperSpace();
                            }
                        }
                        else
                        {
                            MessageBox.Show("this option is meant to work in paper space");
                            Freeze_operations = false;
                            return;
                        }

                        string client1 = comboBox_Client.Text;
                        bool isfound = false;
                        if (String.IsNullOrEmpty(client1) == false)
                        {
                            for (int i = 0; i < dtc.Rows.Count; ++i)
                            {
                                if (isfound == false)
                                {
                                    string client0 = "";
                                    if (dtc.Rows[i][0] != DBNull.Value)
                                    {
                                        client0 = dtc.Rows[i][0].ToString();
                                        if (client0 == client1)
                                        {

                                            if (dtc.Rows[i][7] != DBNull.Value)
                                            {
                                                linetype_name = dtc.Rows[i][7].ToString();
                                            }
                                            if (dtc.Rows[i][1] != DBNull.Value)
                                            {
                                                layer_name = dtc.Rows[i][1].ToString();
                                            }
                                            if (dtc.Rows[i][2] != DBNull.Value)
                                            {
                                                string val0 = dtc.Rows[i][2].ToString();
                                                if (Functions.IsNumeric(val0) == true)
                                                {
                                                    layer_color = Convert.ToInt16(val0);
                                                }
                                            }
                                            if (dtc.Rows[i][3] != DBNull.Value)
                                            {
                                                string val0 = dtc.Rows[i][3].ToString();
                                                if (Functions.IsNumeric(val0) == true)
                                                {
                                                    Ltscale = Convert.ToDouble(val0);
                                                }
                                            }

                                            if (linetype_name != "Continuous")
                                            {
                                                Incarca_line_type(linetype_name, Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + comboBox_Client.Text + ".lin", true);
                                            }

                                            if (dtc.Rows[i][6] != DBNull.Value)
                                            {

                                                Extension_value = Convert.ToDouble(dtc.Rows[i][6]);
                                            }

                                            isfound = true;
                                            i = dtc.Rows.Count;
                                        }
                                    }

                                }

                            }
                        }

                        if (isfound == false)
                        {
                            MessageBox.Show("linetype issue\r\naborting operation");
                            Freeze_operations = false;
                            return;
                        }
                        List<ObjectId> Lista_vp = new List<ObjectId>();
                        List<Polyline> Lista_poly = new List<Polyline>();

                        ObjectIdCollection Colectie_vp = new ObjectIdCollection();

                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {

                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                            Functions.Creaza_layer_with_linetype(layer_name, layer_color, linetype_name, true);

                            foreach (ObjectId polyid in BTrecord)
                            {

                                Polyline poly1 = Trans1.GetObject(polyid, OpenMode.ForRead) as Polyline;
                                if (poly1 != null)
                                {
                                    ObjectId vpid = LayoutManager.Current.GetNonRectangularViewportIdFromClipId(polyid);
                                    if (vpid != ObjectId.Null)
                                    {
                                        if (Trans1.GetObject(vpid, OpenMode.ForRead) is Viewport)
                                        {

                                            Lista_vp.Add(vpid);
                                            Lista_poly.Add(poly1);
                                        }
                                    }
                                }
                            }

                            DBDictionary Layoutdict = (DBDictionary)Trans1.GetObject(ThisDrawing.Database.LayoutDictionaryId, OpenMode.ForRead);

                            LayoutManager LayoutManager1 = (LayoutManager)Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;

                            foreach (DBDictionaryEntry entry in Layoutdict)
                            {
                                Layout Layout0 = (Layout)Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForRead);
                                if (Layout0.LayoutName == LayoutManager1.CurrentLayout)
                                {
                                    Colectie_vp = Layout0.GetViewports();
                                    Colectie_vp.RemoveAt(0); // asta este de la layoutul curent este un viewport
                                }
                            }


                            foreach (ObjectId vpid in Colectie_vp)
                            {
                                if (Lista_vp.Contains(vpid) == false)
                                {
                                    Viewport Ent_vp = Trans1.GetObject(vpid, OpenMode.ForRead) as Viewport;
                                    if (Ent_vp != null)
                                    {
                                        Polyline poly1 = new Polyline();
                                        poly1.AddVertexAt(0, new Point2d(Ent_vp.CenterPoint.X - Ent_vp.Width / 2, Ent_vp.CenterPoint.Y - Ent_vp.Height / 2), 0, 0, 0);
                                        poly1.AddVertexAt(1, new Point2d(Ent_vp.CenterPoint.X - Ent_vp.Width / 2, Ent_vp.CenterPoint.Y + Ent_vp.Height / 2), 0, 0, 0);
                                        poly1.AddVertexAt(2, new Point2d(Ent_vp.CenterPoint.X + Ent_vp.Width / 2, Ent_vp.CenterPoint.Y + Ent_vp.Height / 2), 0, 0, 0);
                                        poly1.AddVertexAt(3, new Point2d(Ent_vp.CenterPoint.X + Ent_vp.Width / 2, Ent_vp.CenterPoint.Y - Ent_vp.Height / 2), 0, 0, 0);
                                        poly1.Closed = true;

                                        Lista_vp.Add(vpid);
                                        Lista_poly.Add(poly1);



                                    }
                                }
                            }


                            Trans1.Commit();
                        }

                        if (Lista_vp.Count > 0)
                        {

                            DBObjectCollection Col_xref = get_xrefs_centerlines();

                            for (int i = 0; i < Lista_vp.Count; ++i)
                            {
                                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                                {
                                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecordMS = (BlockTableRecord)Trans1.GetObject(BlockTable_data1[BlockTableRecord.ModelSpace], Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                                    List<string> Lista_frozen = get_frozen_or_off_layers_from_viewport(Lista_vp[i]);

                                    dt_frozen_xref_layers = new System.Data.DataTable();
                                    dt_frozen_xref_layers.Columns.Add("ln", typeof(string));
                                    dt_frozen_xref_layers.Columns.Add("id", typeof(string));

                                    foreach (string layer1 in Lista_frozen)
                                    {
                                        dt_frozen_xref_layers.Rows.Add();
                                        dt_frozen_xref_layers.Rows[dt_frozen_xref_layers.Rows.Count - 1]["ln"] = layer1;
                                        dt_frozen_xref_layers.Rows[dt_frozen_xref_layers.Rows.Count - 1]["id"] = Lista_vp[i].Handle.Value.ToString();
                                    }

                                    GenerateCL(Lista_vp[i], Lista_poly[i], Col_xref);

                                    Trans1.Commit();
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

        private void GenerateCL(ObjectId vp_id, Polyline clip_poly, DBObjectCollection Col_xref)

        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecordMS = (BlockTableRecord)Trans1.GetObject(BlockTable_data1[BlockTableRecord.ModelSpace], Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);


                Viewport Ent_vp = Trans1.GetObject(vp_id, OpenMode.ForRead) as Viewport;

                Solid3d Cub3d = new Solid3d();
                double ft = 0;
                double bk = 0;

                Autodesk.AutoCAD.GraphicsSystem.View view1 = null;
                DBObjectCollection Poly_Colection1 = new DBObjectCollection();
                Poly_Colection1.Add(clip_poly);
                DBObjectCollection Region_Colection1 = new DBObjectCollection();
                Region_Colection1 = Autodesk.AutoCAD.DatabaseServices.Region.CreateFromCurves(Poly_Colection1);

                Autodesk.AutoCAD.DatabaseServices.Region RegionBK = null;
                Autodesk.AutoCAD.DatabaseServices.Region Regionft = null;
                Matrix3d FCS = new Matrix3d();
                Matrix3d BCS = new Matrix3d();
                DBPoint pptft = new DBPoint();
                DBPoint pptVT = new DBPoint();
                DBPoint pptbk = new DBPoint();

                Line[] Lines = null;
                int idx1 = 1;

                DBObjectCollection Col_delete = new DBObjectCollection();
                ObjectIdCollection Col_cercuriMS = new ObjectIdCollection();
                ObjectIdCollection Col_cercuriPS = new ObjectIdCollection();
                Polyline3d Poly3D_vt = new Polyline3d();


                using (Autodesk.AutoCAD.DatabaseServices.Region RegionPS = (Autodesk.AutoCAD.DatabaseServices.Region)Region_Colection1[0])
                {
                    if (Ent_vp != null && RegionPS != null)
                    {

                        Matrix3d TransforMatrixM2PS = Functions.ModelToPaper(Ent_vp);
                        Matrix3d TransforMatrixpPS2m = Functions.PaperToModel(Ent_vp);

                        Autodesk.AutoCAD.GraphicsSystem.KernelDescriptor kd = new Autodesk.AutoCAD.GraphicsSystem.KernelDescriptor();
                        kd.addRequirement(Autodesk.AutoCAD.UniqueString.Intern("3D Drawing"));
                        Autodesk.AutoCAD.GraphicsSystem.Manager GraphicsManager = ThisDrawing.GraphicsManager;

                        using (GraphicsManager)
                        {
                            view1 = GraphicsManager.ObtainAcGsView(Ent_vp.Number, kd);


                            if (view1 != null)
                            {
                                if (view1.EnableFrontClip == true)
                                {
                                    ft = view1.FrontClip;
                                }
                                if (view1.EnableBackClip == true)
                                {
                                    bk = view1.BackClip;
                                }


                                Matrix3d UCStoWCS = Matrix3d.PlaneToWorld(new Plane(new Point3d(0, 0, 0), Vector3d.ZAxis));
                                Point3d VT = view1.Target.TransformBy(UCStoWCS);

                                Vector3d zAxis = Ent_vp.ViewDirection.GetNormal();

                                Vector3d yAxis = view1.UpVector.GetNormal();

                                Vector3d xAxis = zAxis.CrossProduct(yAxis);

                                pptVT = new DBPoint(VT);
                                pptVT.ColorIndex = 7;
                                // BTrecordMS.AppendEntity(pptVT);
                                // Trans1.AddNewlyCreatedDBObject(pptVT, true);

                                pptbk = new DBPoint(VT);
                                pptbk.TransformBy(Matrix3d.Displacement(zAxis * bk));
                                pptbk.ColorIndex = 1;

                                if (bk != 0)
                                {
                                    // BTrecordMS.AppendEntity(pptbk);
                                    // Trans1.AddNewlyCreatedDBObject(pptbk, true);
                                }


                                pptft = new DBPoint(VT);
                                pptft.TransformBy(Matrix3d.Displacement(zAxis * ft));
                                pptft.ColorIndex = 2;

                                if (ft != 0)
                                {
                                    // BTrecordMS.AppendEntity(pptft);
                                    //Trans1.AddNewlyCreatedDBObject(pptft, true);
                                }


                                Matrix3d WCStoUCS = Matrix3d.WorldToPlane(new Plane(VT, zAxis));
                                Matrix3d WCS = new Matrix3d();
                                WCS = Matrix3d.AlignCoordinateSystem(Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis, VT, xAxis, yAxis, zAxis);


                                Poly3D_vt.SetDatabaseDefaults();
                                Poly3D_vt.Closed = true;
                                Poly3D_vt.ColorIndex = 7;
                                BTrecordMS.AppendEntity(Poly3D_vt);
                                Trans1.AddNewlyCreatedDBObject(Poly3D_vt, true);


                                for (int i = 0; i < clip_poly.NumberOfVertices; ++i)
                                {
                                    Point3d Ms1 = clip_poly.GetPointAtParameter(i).TransformBy(Functions.PaperToModel(Ent_vp));
                                    PolylineVertex3d Vertex_new = new PolylineVertex3d(Ms1);
                                    Poly3D_vt.AppendVertex(Vertex_new);
                                    Trans1.AddNewlyCreatedDBObject(Vertex_new, true);
                                }


                                Polyline3d Poly3D_ft = new Polyline3d();
                                Polyline3d Poly3D_bk = new Polyline3d();

                                if (bk != 0)
                                {
                                    Poly3D_bk = Poly3D_vt.Clone() as Polyline3d;
                                    Poly3D_bk.ColorIndex = 1;
                                    Poly3D_bk.TransformBy(Matrix3d.Displacement(VT.GetVectorTo(pptbk.Position)));
                                    BTrecordMS.AppendEntity(Poly3D_bk);
                                    Trans1.AddNewlyCreatedDBObject(Poly3D_bk, true);
                                }

                                if (ft != 0)
                                {
                                    Poly3D_ft = Poly3D_vt.Clone() as Polyline3d;
                                    Poly3D_ft.ColorIndex = 2;
                                    Poly3D_ft.TransformBy(Matrix3d.Displacement(pptVT.Position.GetVectorTo(pptft.Position)));
                                    BTrecordMS.AppendEntity(Poly3D_ft);
                                    Trans1.AddNewlyCreatedDBObject(Poly3D_ft, true);
                                }


                                if (bk != 0)
                                {
                                    DBObjectCollection Poly_ColectionBK = new DBObjectCollection();
                                    Poly_ColectionBK.Add(Poly3D_bk);
                                    DBObjectCollection Region_ColectionBK = new DBObjectCollection();
                                    Region_ColectionBK = Autodesk.AutoCAD.DatabaseServices.Region.CreateFromCurves(Poly_ColectionBK);
                                    RegionBK = Region_ColectionBK[0] as Autodesk.AutoCAD.DatabaseServices.Region;

                                    if (ft == 0)
                                    {
                                        RegionBK.ColorIndex = 1;
                                        BTrecordMS.AppendEntity(RegionBK);
                                        Trans1.AddNewlyCreatedDBObject(RegionBK, true);
                                        BCS = Matrix3d.AlignCoordinateSystem(pptbk.Position, xAxis, yAxis, zAxis, Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis);
                                    }


                                    Poly3D_bk.Erase();
                                }


                                if (ft != 0)
                                {

                                    DBObjectCollection Poly_Colectionft = new DBObjectCollection();
                                    Poly_Colectionft.Add(Poly3D_ft);
                                    DBObjectCollection Region_Colectionft = new DBObjectCollection();
                                    Region_Colectionft = Autodesk.AutoCAD.DatabaseServices.Region.CreateFromCurves(Poly_Colectionft);

                                    if (bk == 0)
                                    {
                                        Regionft = Region_Colectionft[0] as Autodesk.AutoCAD.DatabaseServices.Region;
                                        Regionft.ColorIndex = 2;
                                        BTrecordMS.AppendEntity(Regionft);
                                        Trans1.AddNewlyCreatedDBObject(Regionft, true);

                                        FCS = Matrix3d.AlignCoordinateSystem(pptft.Position, xAxis, yAxis, zAxis, Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis);

                                    }

                                    Poly3D_ft.Erase();


                                    if (bk != 0)
                                    {
                                        Line Line_extrude = new Line(pptbk.Position, pptft.Position);
                                        Cub3d.ExtrudeAlongPath(RegionBK, Line_extrude, 0);
                                        Cub3d.ColorIndex = 5;
                                        BTrecordMS.AppendEntity(Cub3d);
                                        Trans1.AddNewlyCreatedDBObject(Cub3d, true);
                                    }
                                }
                                // MessageBox.Show("back=" + bk.ToString() + "\r\nfront=" + ft.ToString());
                            }
                        }


                        System.Data.DataTable dt_layers_obj = new System.Data.DataTable();
                        dt_layers_obj.Columns.Add("ln", typeof(string));
                        dt_layers_obj.Columns.Add("id", typeof(string));


                        //ThisDrawing.Database.BindXrefs(Col_xref_id, true);

                        if (Col_xref != null)
                        {
                            if (Col_xref.Count > 0)
                            {
                                Functions.Creaza_layer("cl", 1, false);

                                int jj = 0;
                                foreach (DBObject Obj1 in Col_xref)
                                {
                                    Line L2 = Obj1 as Line;
                                    // L2 = Obj1 as Line;
                                    if (L2 != null)
                                    {
                                        Line L3 = new Line();
                                        L3.StartPoint = L2.StartPoint;
                                        L3.EndPoint = L2.EndPoint;
                                        L3.Layer = "cl";
                                        BTrecordMS.AppendEntity(L3);
                                        Trans1.AddNewlyCreatedDBObject(L3, true);

                                        dt_layers_obj.Rows.Add();
                                        dt_layers_obj.Rows[dt_layers_obj.Rows.Count - 1]["ln"] = Lista_layers_from_xrefs[jj];
                                        dt_layers_obj.Rows[dt_layers_obj.Rows.Count - 1]["id"] = L3.ObjectId.Handle.Value.ToString();
                                        Col_delete.Add(L3);
                                    }

                                    Arc a2 = new Arc();
                                    a2 = Obj1 as Arc;
                                    if (a2 != null)
                                    {
                                        Arc a3 = new Arc();
                                        a3.Center = a2.Center;
                                        a3.StartAngle = a2.StartAngle;
                                        a3.EndAngle = a2.EndAngle;
                                        a3.Radius = a2.Radius;
                                        a3.Normal = a2.Normal;
                                        a3.Layer = "cl";
                                        BTrecordMS.AppendEntity(a3);
                                        Trans1.AddNewlyCreatedDBObject(a3, true);

                                        dt_layers_obj.Rows.Add();
                                        dt_layers_obj.Rows[dt_layers_obj.Rows.Count - 1]["ln"] = Lista_layers_from_xrefs[jj];
                                        dt_layers_obj.Rows[dt_layers_obj.Rows.Count - 1]["id"] = a3.ObjectId.Handle.Value.ToString();

                                        Col_delete.Add(a3);
                                    }

                                    Circle c2 = new Circle();
                                    c2 = Obj1 as Circle;
                                    if (c2 != null)
                                    {
                                        Circle c3 = new Circle();
                                        c3.Center = c2.Center;
                                        c3.Normal = c2.Normal;
                                        c3.Radius = c2.Radius;
                                        c3.Layer = "cl";
                                        BTrecordMS.AppendEntity(c3);
                                        Trans1.AddNewlyCreatedDBObject(c3, true);

                                        Col_cercuriMS.Add(c3.ObjectId);
                                        dt_layers_obj.Rows.Add();
                                        dt_layers_obj.Rows[dt_layers_obj.Rows.Count - 1]["ln"] = Lista_layers_from_xrefs[jj];
                                        dt_layers_obj.Rows[dt_layers_obj.Rows.Count - 1]["id"] = c3.ObjectId.Handle.Value.ToString();

                                        Col_delete.Add(c3);
                                    }

                                    jj = jj + 1;
                                }
                            }
                        }


                        foreach (ObjectId id1 in BTrecordMS)
                        {
                            Solid3d SSolid3d1 = Trans1.GetObject(id1, OpenMode.ForRead) as Solid3d;
                            if (SSolid3d1 != null)
                            {
                                try
                                {
                                    if (SSolid3d1 != null)
                                    {

                                        IntPtr pSubentityIdPE = SSolid3d1.QueryX(AssocPersSubentityIdPE.GetClass(typeof(AssocPersSubentityIdPE)));

                                        if (pSubentityIdPE != IntPtr.Zero)
                                        {
                                            AssocPersSubentityIdPE subentityIdPE = AssocPersSubentityIdPE.Create(pSubentityIdPE, false) as AssocPersSubentityIdPE;
                                            SubentityId[] edgeIds = subentityIdPE.GetAllSubentities(SSolid3d1, SubentityType.Edge);

                                            ObjectId[] ids = new ObjectId[] { SSolid3d1.ObjectId };
                                            foreach (SubentityId subentId in edgeIds)
                                            {
                                                FullSubentityPath path = new FullSubentityPath(ids, subentId);

                                                Entity edgeEntity = SSolid3d1.GetSubentity(path);

                                                if (edgeEntity != null)
                                                {
                                                    Circle cc1 = new Circle();
                                                    cc1 = edgeEntity as Circle;

                                                    if (cc1 != null)
                                                    {
                                                        cc1.Layer = "cl";
                                                        BTrecordMS.AppendEntity(cc1);
                                                        Trans1.AddNewlyCreatedDBObject(cc1, true);

                                                        Col_cercuriMS.Add(cc1.ObjectId);

                                                        dt_layers_obj.Rows.Add();
                                                        dt_layers_obj.Rows[dt_layers_obj.Rows.Count - 1]["ln"] = SSolid3d1.Layer;
                                                        dt_layers_obj.Rows[dt_layers_obj.Rows.Count - 1]["id"] = cc1.ObjectId.Handle.Value.ToString();
                                                    }
                                                    edgeEntity.Dispose();
                                                }
                                            }
                                        }
                                    }

                                }
                                catch (System.Exception ex)
                                {
                                    ThisDrawing.Editor.WriteMessage("\n" + ex.Message);
                                }

                            }

                        }


                        if (Col_cercuriMS.Count > 0)
                        {
                            foreach (ObjectId Objid_1 in Col_cercuriMS)
                            {
                                Circle C1 = Trans1.GetObject(Objid_1, OpenMode.ForRead) as Circle;

                                Plane Plan_region = new Plane(Poly3D_vt.GetPointAtParameter(0), Poly3D_vt.GetPointAtParameter(1), Poly3D_vt.GetPointAtParameter(2));

                                Curve Cprojected = C1.GetProjectedCurve(Plan_region, Plan_region.Normal) as Curve;

                                if (C1 != null && Cprojected != null)
                                {
                                    Point3d Pt1 = C1.Center;

                                    bool Process_cerc_in_PS = false;
                                    if (ft != 0 && bk != 0)
                                    {
                                        PointContainment pc1 = PointContainment.Outside;

                                        using (Brep Brep_obj = new Brep(Cub3d))
                                        {
                                            if (Brep_obj != null)
                                            {
                                                using (BrepEntity ent1 = Brep_obj.GetPointContainment(Pt1, out pc1))
                                                {
                                                    if (ent1 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                                    {
                                                        pc1 = PointContainment.Inside;
                                                    }
                                                }
                                            }
                                        }


                                        if (pc1 == PointContainment.Inside)
                                        {
                                            Process_cerc_in_PS = true;
                                        }
                                    }

                                    else
                                    {
                                        Point3d Pt11 = Pt1.TransformBy(TransforMatrixM2PS);
                                        Pt11 = new Point3d(Pt11.X, Pt11.Y, clip_poly.Elevation);
                                        PointContainment Pc11 = Get_Point_Containment(RegionPS, Pt11);
                                        if (Pc11 == PointContainment.Inside)
                                        {
                                            Process_cerc_in_PS = true;
                                        }
                                    }




                                    if (Process_cerc_in_PS == true)
                                    {
                                        Process_cerc_in_PS = Verifica_daca_trebuie_procesat(vp_id, dt_layers_obj, Objid_1);
                                    }

                                    if (Process_cerc_in_PS == true)
                                    {
                                        Point3d Pt11 = Pt1.TransformBy(TransforMatrixM2PS);

                                        Pt11 = new Point3d(Pt11.X, Pt11.Y, clip_poly.Elevation);

                                        PointContainment Pc11 = Get_Point_Containment(RegionPS, Pt11);

                                        Cprojected.TransformBy(TransforMatrixM2PS);

                                        if (Pc11 == PointContainment.Inside || Pc11 == PointContainment.OnBoundary)
                                        {
                                            if (C1.Radius * Ent_vp.CustomScale > 0.0001)
                                            {
                                                Circle cext1 = Cprojected as Circle;
                                                Ellipse ellext1 = Cprojected as Ellipse;
                                                if (cext1 != null)
                                                {
                                                    cext1.Layer = "0";
                                                    cext1.ColorIndex = 3;
                                                    cext1.Center = new Point3d(cext1.Center.X, cext1.Center.Y, clip_poly.Elevation);
                                                    BTrecord.AppendEntity(cext1);
                                                    Trans1.AddNewlyCreatedDBObject(cext1, true);
                                                    Col_cercuriPS.Add(cext1.ObjectId);
                                                }
                                                if (ellext1 != null)
                                                {
                                                    ellext1.Layer = "0";
                                                    ellext1.ColorIndex = 5;
                                                    ellext1.Center = new Point3d(ellext1.Center.X, ellext1.Center.Y, clip_poly.Elevation);
                                                    BTrecord.AppendEntity(ellext1);
                                                    Trans1.AddNewlyCreatedDBObject(ellext1, true);
                                                    Col_cercuriPS.Add(ellext1.ObjectId);
                                                }

                                            }
                                        }
                                    }
                                }
                            }
                        }

                        foreach (ObjectId id1 in BTrecordMS)
                        {

                            DBObjectCollection Col_exploded = new DBObjectCollection();

                            Solid3d SSolid3d1 = Trans1.GetObject(id1, OpenMode.ForRead) as Solid3d;
                            Line LLine1 = Trans1.GetObject(id1, OpenMode.ForRead) as Line;
                            Arc AArc1 = Trans1.GetObject(id1, OpenMode.ForRead) as Arc;
                            Polyline PLLine1 = Trans1.GetObject(id1, OpenMode.ForRead) as Polyline;

                            if (SSolid3d1 != null || LLine1 != null || AArc1 != null)
                            {
                                try
                                {
                                    if (SSolid3d1 != null)
                                    {
                                        if (Verifica_layerul_obj(SSolid3d1.Layer) == true)
                                        {
                                            SSolid3d1.Explode(Col_exploded);
                                        }

                                    }

                                    if (LLine1 != null)
                                    {
                                        if (LLine1.Layer.ToLower() == "cl")
                                        {
                                            if (Verifica_daca_trebuie_procesat(vp_id, dt_layers_obj, id1) == true) Col_exploded.Add(LLine1);

                                        }
                                    }

                                    if (AArc1 != null)
                                    {
                                        if (AArc1.Layer.ToLower() == "cl")
                                        {
                                            if (Verifica_daca_trebuie_procesat(vp_id, dt_layers_obj, id1) == true) Col_exploded.Add(AArc1);
                                        }
                                    }

                                    if (PLLine1 != null)
                                    {
                                        if (PLLine1.Layer.ToLower() == "cl")
                                        {
                                            if (Verifica_daca_trebuie_procesat(vp_id, dt_layers_obj, id1) == true) Col_exploded.Add(new Line(PLLine1.GetPointAtParameter(0), PLLine1.GetPointAtParameter(PLLine1.NumberOfVertices - 1)));
                                        }
                                    }

                                    if (Col_exploded.Count > 0)
                                    {
                                        foreach (DBObject Obj1 in Col_exploded)
                                        {

                                            Line l1 = Obj1 as Line;
                                            if (l1 != null)
                                            {
                                                Point3d Pt1 = l1.StartPoint;
                                                Point3d Pt2 = l1.EndPoint;
                                                bool Process_line_in_PS = false;
                                                if (ft != 0 && bk != 0)
                                                {
                                                    PointContainment pc1 = PointContainment.Outside;
                                                    PointContainment pc2 = PointContainment.Outside;
                                                    using (Brep Brep_obj = new Brep(Cub3d))
                                                    {
                                                        if (Brep_obj != null)
                                                        {
                                                            using (BrepEntity ent1 = Brep_obj.GetPointContainment(Pt1, out pc1))
                                                            {
                                                                if (ent1 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                                                {
                                                                    pc1 = PointContainment.Inside;
                                                                }
                                                            }
                                                            using (BrepEntity ent2 = Brep_obj.GetPointContainment(Pt2, out pc2))
                                                            {
                                                                if (ent2 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                                                {
                                                                    pc2 = PointContainment.Inside;
                                                                }
                                                            }
                                                        }
                                                    }

                                                    if (pc1 == PointContainment.Inside && pc2 == PointContainment.Inside)
                                                    {
                                                        Process_line_in_PS = true;
                                                    }

                                                    else
                                                    {


                                                        Point3d pint1 = Pt1;
                                                        Point3d pint2 = Pt2;


                                                        Line line_test = new Line(Pt1, Pt2);
                                                        line_test.TransformBy(TransforMatrixM2PS);
                                                        line_test.StartPoint = new Point3d(line_test.StartPoint.X, line_test.StartPoint.Y, clip_poly.Elevation);
                                                        line_test.EndPoint = new Point3d(line_test.EndPoint.X, line_test.EndPoint.Y, clip_poly.Elevation);
                                                        Point3dCollection col_pt_int = new Point3dCollection();

                                                        line_test.IntersectWith(clip_poly, Intersect.OnBothOperands, col_pt_int, IntPtr.Zero, IntPtr.Zero);
                                                        if (col_pt_int.Count > 0)
                                                        {
                                                            ObjectId[] ids = new ObjectId[] { Cub3d.ObjectId };
                                                            FullSubentityPath path = new FullSubentityPath(ids, new SubentityId(SubentityType.Null, IntPtr.Zero));

                                                            using (Brep Brep_obj = new Brep(path))
                                                            {
                                                                if (Brep_obj != null)
                                                                {
                                                                    foreach (Autodesk.AutoCAD.BoundaryRepresentation.Face Face1 in Brep_obj.Faces)
                                                                    {
                                                                        Point3dCollection Pt3d_col = new Point3dCollection();
                                                                        foreach (BoundaryLoop Loop1 in Face1.Loops)
                                                                        {
                                                                            foreach (Autodesk.AutoCAD.BoundaryRepresentation.Vertex Vtx in Loop1.Vertices)
                                                                            {
                                                                                Pt3d_col.Add(Vtx.Point);
                                                                            }
                                                                        }

                                                                        if (Pt3d_col.Count > 2)
                                                                        {
                                                                            Plane Plan_face = new Plane(Pt3d_col[0], Pt3d_col[1], Pt3d_col[2]);


                                                                            Curve Lprojected = l1.GetProjectedCurve(Plan_face, Plan_face.Normal);

                                                                            Point3dCollection Coltest = new Point3dCollection();
                                                                            l1.IntersectWith(Lprojected, Intersect.OnBothOperands, Coltest, IntPtr.Zero, IntPtr.Zero);
                                                                            if (Coltest.Count > 0)
                                                                            {

                                                                                PointContainment pct1 = PointContainment.Outside;
                                                                                using (BrepEntity ent1 = Brep_obj.GetPointContainment(Coltest[0], out pct1))
                                                                                {
                                                                                    if (ent1 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                                                                    {
                                                                                        pct1 = PointContainment.Inside;
                                                                                    }
                                                                                }

                                                                                PointContainment pct01 = PointContainment.Outside;
                                                                                using (BrepEntity ent1 = Brep_obj.GetPointContainment(Pt1, out pct01))
                                                                                {
                                                                                    if (ent1 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                                                                    {
                                                                                        pct01 = PointContainment.Inside;
                                                                                    }
                                                                                }

                                                                                PointContainment pct02 = PointContainment.Outside;
                                                                                using (BrepEntity ent1 = Brep_obj.GetPointContainment(Pt2, out pct02))
                                                                                {
                                                                                    if (ent1 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                                                                    {
                                                                                        pct02 = PointContainment.Inside;
                                                                                    }
                                                                                }

                                                                                if (pct1 == PointContainment.Inside)
                                                                                {
                                                                                    if (pct01 == PointContainment.Outside && pct02 == PointContainment.Inside)
                                                                                    {
                                                                                        pint1 = Coltest[0];
                                                                                    }

                                                                                    if (pct02 == PointContainment.Outside && pct01 == PointContainment.Inside)
                                                                                    {
                                                                                        pint2 = Coltest[0];
                                                                                    }

                                                                                    if (pct02 == PointContainment.Outside && pct01 == PointContainment.Outside)
                                                                                    {
                                                                                        if (pint1 == Pt1)
                                                                                        {
                                                                                            pint1 = Coltest[0];
                                                                                        }
                                                                                        else if (pint2 == Pt2)
                                                                                        {
                                                                                            pint2 = Coltest[0];
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                    }

                                                                }
                                                            }
                                                        }
                                                        if (pint1 != Pt1)
                                                        {
                                                            Pt1 = pint1;
                                                            Process_line_in_PS = true;
                                                        }

                                                        if (pint2 != Pt2)
                                                        {
                                                            Pt2 = pint2;
                                                            Process_line_in_PS = true;
                                                        }
                                                    }
                                                }

                                                if (ft == 0 && bk == 0)
                                                {
                                                    Process_line_in_PS = true;
                                                }

                                                if (ft != 0 && bk == 0)
                                                {
                                                    double pc1 = Pt1.TransformBy(FCS).Z;
                                                    double pc2 = Pt2.TransformBy(FCS).Z;
                                                    double ftz = pptft.Position.TransformBy(FCS).Z;
                                                    double vtz = pptVT.Position.TransformBy(FCS).Z;
                                                    Boolean Displ = false;

                                                    if (vtz <= pc1 && pc1 <= ftz)
                                                    {
                                                        Displ = true;
                                                    }

                                                    if (ftz <= pc1 && pc1 <= vtz)
                                                    {
                                                        Displ = true;
                                                    }

                                                    if (vtz <= pc2 && pc2 <= ftz)
                                                    {
                                                        Displ = true;
                                                    }

                                                    if (ftz <= pc2 && pc2 <= vtz)
                                                    {
                                                        Displ = true;
                                                    }


                                                    if (Displ == true)
                                                    {
                                                        Process_line_in_PS = true;

                                                    }

                                                }
                                                if (ft == 0 && bk != 0)
                                                {
                                                    double pc1 = Pt1.TransformBy(BCS).Z;
                                                    double pc2 = Pt2.TransformBy(BCS).Z;
                                                    double bkz = pptbk.Position.TransformBy(BCS).Z;
                                                    double vtz = pptVT.Position.TransformBy(BCS).Z;
                                                    Boolean Displ = false;

                                                    if (pc1 >= bkz)
                                                    {
                                                        Displ = true;
                                                    }

                                                    if (pc2 >= bkz)
                                                    {
                                                        Displ = true;
                                                    }

                                                    if (Displ == true)
                                                    {
                                                        Process_line_in_PS = true;
                                                    }
                                                }




                                                if (Process_line_in_PS == true)
                                                {
                                                    Point3d Pt11 = Pt1.TransformBy(TransforMatrixM2PS);
                                                    Point3d Pt22 = Pt2.TransformBy(TransforMatrixM2PS);

                                                    Pt11 = new Point3d(Pt11.X, Pt11.Y, clip_poly.Elevation);
                                                    Pt22 = new Point3d(Pt22.X, Pt22.Y, clip_poly.Elevation);

                                                    PointContainment Pc11 = Get_Point_Containment(RegionPS, Pt11);
                                                    PointContainment Pc22 = Get_Point_Containment(RegionPS, Pt22);

                                                    Line Line2 = new Line(Pt11, Pt22);

                                                    if ((Pc11 == PointContainment.Inside || Pc11 == PointContainment.OnBoundary) && (Pc22 == PointContainment.Inside || Pc22 == PointContainment.OnBoundary))
                                                    {
                                                        if (Line2.Length > 0.0001)
                                                        {
                                                            Line2.Layer = l1.Layer;
                                                            Line2.ColorIndex = 1;
                                                            //BTrecord.AppendEntity(Line2);
                                                            //Trans1.AddNewlyCreatedDBObject(Line2, true);
                                                            Array.Resize(ref Lines, idx1);
                                                            Lines[idx1 - 1] = Line2;
                                                            idx1 = idx1 + 1;
                                                        }
                                                    }

                                                    if (Pc11 == PointContainment.Outside && (Pc22 == PointContainment.Inside || Pc22 == PointContainment.OnBoundary))
                                                    {
                                                        Point3dCollection ColPS = Functions.Intersect_on_both_operands(Line2, clip_poly);
                                                        if (ColPS.Count > 0)
                                                        {
                                                            Line2.StartPoint = ColPS[0];
                                                            if (Line2.Length > 0.0001)
                                                            {
                                                                Line2.Layer = l1.Layer;
                                                                Line2.ColorIndex = 1;
                                                                //BTrecord.AppendEntity(Line2);
                                                                //Trans1.AddNewlyCreatedDBObject(Line2, true);
                                                                Array.Resize(ref Lines, idx1);
                                                                Lines[idx1 - 1] = Line2;
                                                                idx1 = idx1 + 1;
                                                            }
                                                        }

                                                    }

                                                    if ((Pc11 == PointContainment.Inside || Pc11 == PointContainment.OnBoundary) && Pc22 == PointContainment.Outside)
                                                    {
                                                        Point3dCollection ColPS = Functions.Intersect_on_both_operands(Line2, clip_poly);
                                                        if (ColPS.Count > 0)
                                                        {
                                                            Line2.EndPoint = ColPS[0];
                                                            if (Line2.Length > 0.0001)
                                                            {
                                                                Line2.Layer = l1.Layer;
                                                                Line2.ColorIndex = 1;
                                                                //BTrecord.AppendEntity(Line2);
                                                                //Trans1.AddNewlyCreatedDBObject(Line2, true);
                                                                Array.Resize(ref Lines, idx1);
                                                                Lines[idx1 - 1] = Line2;
                                                                idx1 = idx1 + 1;
                                                            }
                                                        }

                                                    }


                                                    if (Pc11 == PointContainment.Outside && Pc22 == PointContainment.Outside)
                                                    {
                                                        Point3dCollection ColPS = Functions.Intersect_on_both_operands(Line2, clip_poly);
                                                        if (ColPS.Count == 2)
                                                        {
                                                            Line2.StartPoint = ColPS[0];
                                                            Line2.EndPoint = ColPS[1];
                                                            if (Line2.Length > 0.0001)
                                                            {
                                                                Line2.Layer = l1.Layer;
                                                                Line2.ColorIndex = 1;
                                                                //BTrecord.AppendEntity(Line2);
                                                                //Trans1.AddNewlyCreatedDBObject(Line2, true);
                                                                Array.Resize(ref Lines, idx1);
                                                                Lines[idx1 - 1] = Line2;
                                                                idx1 = idx1 + 1;
                                                            }
                                                        }

                                                    }

                                                }


                                            }

                                            Arc A1 = Obj1 as Arc;
                                            if (A1 != null)
                                            {

                                                Matrix3d Arc_cs = GetArcMatrix(A1);

                                                Point3d Pt1 = A1.StartPoint;
                                                Point3d Pt2 = A1.EndPoint;
                                                Point3d Pt3 = A1.Center;

                                                bool Process_line_in_PS1 = false;
                                                bool Process_line_in_PS2 = false;

                                                Plane Plan1 = new Plane(Pt1, Pt2, Pt3);

                                                Line Line1 = new Line(Pt1, Pt3);
                                                Line1.TransformBy(Matrix3d.Rotation(Math.PI / 2, Arc_cs.CoordinateSystem3d.Zaxis, Pt1));

                                                Line Line2 = new Line(Pt2, Pt3);
                                                Line2.TransformBy(Matrix3d.Rotation(Math.PI / 2, Arc_cs.CoordinateSystem3d.Zaxis, Pt2));

                                                Point3dCollection col1 = new Point3dCollection();
                                                Line1.IntersectWith(Line2, Intersect.ExtendBoth, col1, IntPtr.Zero, IntPtr.Zero);

                                                if (col1.Count == 1)
                                                {
                                                    Point3d Pt31 = col1[0];
                                                    Point3d Pt32 = col1[0];
                                                    Line Line01 = new Line(Pt1, Pt31);
                                                    Line Line02 = new Line(Pt2, Pt32);

                                                    if (ft == 0 && bk == 0)
                                                    {
                                                        Process_line_in_PS1 = true;
                                                        Process_line_in_PS2 = true;
                                                    }

                                                    if (ft != 0 && bk != 0)
                                                    {
                                                        PointContainment pc1 = PointContainment.Outside;
                                                        PointContainment pc2 = PointContainment.Outside;
                                                        PointContainment pc3 = PointContainment.Outside;

                                                        using (Brep Brep_obj = new Brep(Cub3d))
                                                        {
                                                            if (Brep_obj != null)
                                                            {
                                                                using (BrepEntity ent1 = Brep_obj.GetPointContainment(Pt1, out pc1))
                                                                {
                                                                    if (ent1 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                                                    {
                                                                        pc1 = PointContainment.Inside;
                                                                    }
                                                                }

                                                                using (BrepEntity ent2 = Brep_obj.GetPointContainment(Pt2, out pc2))
                                                                {
                                                                    if (ent2 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                                                    {
                                                                        pc2 = PointContainment.Inside;
                                                                    }
                                                                }

                                                                using (BrepEntity ent3 = Brep_obj.GetPointContainment(Pt31, out pc3))
                                                                {
                                                                    if (ent3 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                                                    {
                                                                        pc3 = PointContainment.Inside;
                                                                    }
                                                                }
                                                            }
                                                        }

                                                        if (pc1 == PointContainment.Inside && pc3 == PointContainment.Inside)
                                                        {
                                                            Process_line_in_PS1 = true;
                                                        }
                                                        if (pc2 == PointContainment.Inside && pc3 == PointContainment.Inside)
                                                        {
                                                            Process_line_in_PS2 = true;
                                                        }
                                                        else
                                                        {
                                                            Point3d pint1 = Pt1;
                                                            Point3d pint2 = Pt2;
                                                            Point3d pint31 = Pt31;
                                                            Point3d pint32 = Pt32;

                                                            ObjectId[] ids = new ObjectId[] { Cub3d.ObjectId };
                                                            FullSubentityPath path = new FullSubentityPath(ids, new SubentityId(SubentityType.Null, IntPtr.Zero));

                                                            using (Brep Brep_obj = new Brep(path))
                                                            {
                                                                if (Brep_obj != null)
                                                                {
                                                                    foreach (Autodesk.AutoCAD.BoundaryRepresentation.Face Face1 in Brep_obj.Faces)
                                                                    {
                                                                        Point3dCollection Pt3d_col = new Point3dCollection();
                                                                        foreach (BoundaryLoop Loop1 in Face1.Loops)
                                                                        {
                                                                            foreach (Autodesk.AutoCAD.BoundaryRepresentation.Vertex Vtx in Loop1.Vertices)
                                                                            {
                                                                                Pt3d_col.Add(Vtx.Point);
                                                                            }
                                                                        }

                                                                        if (Pt3d_col.Count > 2)
                                                                        {
                                                                            Plane Plan_face = new Plane(Pt3d_col[0], Pt3d_col[1], Pt3d_col[2]);

                                                                            Curve Lprojected1 = Line01.GetProjectedCurve(Plan_face, Plan_face.Normal);
                                                                            Curve Lprojected2 = Line02.GetProjectedCurve(Plan_face, Plan_face.Normal);

                                                                            Point3dCollection Coltest1 = new Point3dCollection();
                                                                            Line01.IntersectWith(Lprojected1, Intersect.OnBothOperands, Coltest1, IntPtr.Zero, IntPtr.Zero);

                                                                            Point3dCollection Coltest2 = new Point3dCollection();
                                                                            Line02.IntersectWith(Lprojected2, Intersect.OnBothOperands, Coltest2, IntPtr.Zero, IntPtr.Zero);

                                                                            if (Coltest1.Count > 0)
                                                                            {
                                                                                PointContainment pct1 = PointContainment.Outside;
                                                                                using (BrepEntity ent1 = Brep_obj.GetPointContainment(Coltest1[0], out pct1))
                                                                                {
                                                                                    if (ent1 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                                                                    {
                                                                                        pct1 = PointContainment.Inside;
                                                                                    }
                                                                                }

                                                                                PointContainment pct01 = PointContainment.Outside;
                                                                                using (BrepEntity ent1 = Brep_obj.GetPointContainment(Pt1, out pct01))
                                                                                {
                                                                                    if (ent1 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                                                                    {
                                                                                        pct01 = PointContainment.Inside;
                                                                                    }
                                                                                }

                                                                                PointContainment pct03 = PointContainment.Outside;
                                                                                using (BrepEntity ent1 = Brep_obj.GetPointContainment(Pt31, out pct03))
                                                                                {
                                                                                    if (ent1 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                                                                    {
                                                                                        pct03 = PointContainment.Inside;
                                                                                    }
                                                                                }

                                                                                if (pct1 == PointContainment.Inside)
                                                                                {
                                                                                    if (pct01 == PointContainment.Outside && pct03 == PointContainment.Inside)
                                                                                    {
                                                                                        pint1 = Coltest1[0];
                                                                                    }

                                                                                    if (pct03 == PointContainment.Outside && pct01 == PointContainment.Inside)
                                                                                    {
                                                                                        pint31 = Coltest1[0];
                                                                                    }

                                                                                    if (pct03 == PointContainment.Outside && pct01 == PointContainment.Outside)
                                                                                    {
                                                                                        if (pint1 == Pt1)
                                                                                        {
                                                                                            pint1 = Coltest1[0];

                                                                                        }
                                                                                        else if (pint31 == Pt31)
                                                                                        {
                                                                                            pint31 = Coltest1[0];
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }

                                                                            if (Coltest2.Count > 0)
                                                                            {
                                                                                PointContainment pct2 = PointContainment.Outside;
                                                                                using (BrepEntity ent1 = Brep_obj.GetPointContainment(Coltest2[0], out pct2))
                                                                                {
                                                                                    if (ent1 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                                                                    {
                                                                                        pct2 = PointContainment.Inside;
                                                                                    }
                                                                                }

                                                                                PointContainment pct02 = PointContainment.Outside;
                                                                                using (BrepEntity ent1 = Brep_obj.GetPointContainment(Pt2, out pct02))
                                                                                {
                                                                                    if (ent1 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                                                                    {
                                                                                        pct02 = PointContainment.Inside;
                                                                                    }
                                                                                }

                                                                                PointContainment pct03 = PointContainment.Outside;
                                                                                using (BrepEntity ent1 = Brep_obj.GetPointContainment(Pt32, out pct03))
                                                                                {
                                                                                    if (ent1 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                                                                    {
                                                                                        pct03 = PointContainment.Inside;
                                                                                    }
                                                                                }

                                                                                if (pct2 == PointContainment.Inside)
                                                                                {
                                                                                    if (pct02 == PointContainment.Outside && pct03 == PointContainment.Inside)
                                                                                    {
                                                                                        pint2 = Coltest2[0];
                                                                                    }

                                                                                    if (pct03 == PointContainment.Outside && pct02 == PointContainment.Inside)
                                                                                    {
                                                                                        pint32 = Coltest2[0];
                                                                                    }

                                                                                    if (pct03 == PointContainment.Outside && pct02 == PointContainment.Outside)
                                                                                    {
                                                                                        if (pint2 == Pt2)
                                                                                        {
                                                                                            pint2 = Coltest2[0];

                                                                                        }
                                                                                        else if (pint32 == Pt32)
                                                                                        {
                                                                                            pint32 = Coltest2[0];
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }


                                                            if (pint1 != Pt1)
                                                            {
                                                                Pt1 = pint1;
                                                                Process_line_in_PS1 = true;
                                                            }

                                                            if (pint31 != Pt31)
                                                            {
                                                                Pt31 = pint31;
                                                                Process_line_in_PS1 = true;
                                                            }

                                                            if (pint2 != Pt2)
                                                            {
                                                                Pt2 = pint2;
                                                                Process_line_in_PS2 = true;
                                                            }


                                                            if (pint32 != Pt32)
                                                            {
                                                                Pt32 = pint32;
                                                                Process_line_in_PS2 = true;
                                                            }

                                                        }

                                                    }


                                                    if (ft != 0 && bk == 0)
                                                    {
                                                        double pc1 = Pt1.TransformBy(FCS).Z;
                                                        double pc2 = Pt2.TransformBy(FCS).Z;
                                                        double pc31 = Pt31.TransformBy(FCS).Z;
                                                        double pc32 = pc31;

                                                        double ftz = pptft.Position.TransformBy(FCS).Z;
                                                        double vtz = pptVT.Position.TransformBy(FCS).Z;
                                                        Boolean Displ1 = false;

                                                        if (vtz <= pc1 && pc1 <= ftz)
                                                        {
                                                            Displ1 = true;
                                                        }

                                                        if (ftz <= pc1 && pc1 <= vtz)
                                                        {
                                                            Displ1 = true;
                                                        }

                                                        if (vtz <= pc31 && pc31 <= ftz)
                                                        {
                                                            Displ1 = true;
                                                        }

                                                        if (ftz <= pc31 && pc31 <= vtz)
                                                        {
                                                            Displ1 = true;
                                                        }

                                                        if (Displ1 == true)
                                                        {
                                                            Process_line_in_PS1 = true;
                                                        }

                                                        Boolean Displ2 = false;

                                                        if (vtz <= pc32 && pc32 <= ftz)
                                                        {
                                                            Displ2 = true;
                                                        }

                                                        if (ftz <= pc32 && pc32 <= vtz)
                                                        {
                                                            Displ2 = true;
                                                        }

                                                        if (vtz <= pc2 && pc2 <= ftz)
                                                        {
                                                            Displ2 = true;
                                                        }

                                                        if (ftz <= pc2 && pc2 <= vtz)
                                                        {
                                                            Displ2 = true;
                                                        }

                                                        if (Displ2 == true)
                                                        {
                                                            Process_line_in_PS2 = true;
                                                        }

                                                    }

                                                    if (ft == 0 && bk != 0)
                                                    {
                                                        double pc1 = Pt1.TransformBy(BCS).Z;
                                                        double pc2 = Pt2.TransformBy(BCS).Z;
                                                        double pc31 = Pt31.TransformBy(BCS).Z;
                                                        double pc32 = pc31;

                                                        double bkz = pptbk.Position.TransformBy(BCS).Z;
                                                        double vtz = pptVT.Position.TransformBy(BCS).Z;
                                                        Boolean Displ1 = false;

                                                        if (pc1 >= bkz)
                                                        {
                                                            Displ1 = true;
                                                        }


                                                        if (pc31 >= bkz)
                                                        {
                                                            Displ1 = true;
                                                        }


                                                        if (Displ1 == true)
                                                        {
                                                            Process_line_in_PS1 = true;
                                                        }

                                                        Boolean Displ2 = false;

                                                        if (pc2 >= bkz)
                                                        {
                                                            Displ2 = true;
                                                        }


                                                        if (pc32 >= bkz)
                                                        {
                                                            Displ2 = true;
                                                        }

                                                        if (Displ2 == true)
                                                        {
                                                            Process_line_in_PS2 = true;
                                                        }
                                                    }


                                                    if (Process_line_in_PS1 == true)
                                                    {

                                                        Point3d Pt11 = Pt1.TransformBy(TransforMatrixM2PS);
                                                        Point3d Pt33 = Pt31.TransformBy(TransforMatrixM2PS);

                                                        Pt11 = new Point3d(Pt11.X, Pt11.Y, clip_poly.Elevation);
                                                        Pt33 = new Point3d(Pt33.X, Pt33.Y, clip_poly.Elevation);
                                                        PointContainment Pc11 = Get_Point_Containment(RegionPS, Pt11);
                                                        PointContainment Pc33 = Get_Point_Containment(RegionPS, Pt33);

                                                        Line Line_arc1 = new Line(Pt11, Pt33);
                                                        if ((Pc11 == PointContainment.Inside || Pc11 == PointContainment.OnBoundary) && (Pc33 == PointContainment.Inside || Pc33 == PointContainment.OnBoundary))
                                                        {
                                                            if (Line_arc1.Length > 0.0001)
                                                            {
                                                                Line_arc1.Layer = A1.Layer;
                                                                Line_arc1.ColorIndex = 1;
                                                                //BTrecord.AppendEntity(Line_arc1);
                                                                //Trans1.AddNewlyCreatedDBObject(Line_arc1, true);
                                                                Array.Resize(ref Lines, idx1);
                                                                Lines[idx1 - 1] = Line_arc1;
                                                                idx1 = idx1 + 1;
                                                            }
                                                        }

                                                        if (Pc11 == PointContainment.Outside && (Pc33 == PointContainment.Inside || Pc33 == PointContainment.OnBoundary))
                                                        {
                                                            Point3dCollection ColPS = Functions.Intersect_on_both_operands(Line_arc1, clip_poly);
                                                            if (ColPS.Count > 0)
                                                            {
                                                                Line_arc1.StartPoint = ColPS[0];
                                                                if (Line_arc1.Length > 0.0001)
                                                                {
                                                                    Line_arc1.Layer = A1.Layer;
                                                                    Line_arc1.ColorIndex = 1;
                                                                    // BTrecord.AppendEntity(Line_arc1);
                                                                    //Trans1.AddNewlyCreatedDBObject(Line_arc1, true);
                                                                    Array.Resize(ref Lines, idx1);
                                                                    Lines[idx1 - 1] = Line_arc1;
                                                                    idx1 = idx1 + 1;
                                                                }
                                                            }
                                                        }

                                                        if ((Pc11 == PointContainment.Inside || Pc11 == PointContainment.OnBoundary) && Pc33 == PointContainment.Outside)
                                                        {
                                                            Point3dCollection ColPS = Functions.Intersect_on_both_operands(Line_arc1, clip_poly);
                                                            if (ColPS.Count > 0)
                                                            {
                                                                Line_arc1.EndPoint = ColPS[0];
                                                                if (Line_arc1.Length > 0.0001)
                                                                {
                                                                    Line_arc1.Layer = A1.Layer;
                                                                    Line_arc1.ColorIndex = 1;
                                                                    //BTrecord.AppendEntity(Line_arc1);
                                                                    //Trans1.AddNewlyCreatedDBObject(Line_arc1, true);
                                                                    Array.Resize(ref Lines, idx1);
                                                                    Lines[idx1 - 1] = Line_arc1;
                                                                    idx1 = idx1 + 1;
                                                                }
                                                            }
                                                        }

                                                    }

                                                    if (Process_line_in_PS2 == true)
                                                    {

                                                        Point3d Pt22 = Pt2.TransformBy(TransforMatrixM2PS);
                                                        Point3d Pt33 = Pt32.TransformBy(TransforMatrixM2PS);

                                                        Pt22 = new Point3d(Pt22.X, Pt22.Y, clip_poly.Elevation);
                                                        Pt33 = new Point3d(Pt33.X, Pt33.Y, clip_poly.Elevation);
                                                        PointContainment Pc22 = Get_Point_Containment(RegionPS, Pt22);
                                                        PointContainment Pc33 = Get_Point_Containment(RegionPS, Pt33);

                                                        Line Line_arc2 = new Line(Pt22, Pt33);
                                                        if ((Pc22 == PointContainment.Inside || Pc22 == PointContainment.OnBoundary) && (Pc33 == PointContainment.Inside || Pc33 == PointContainment.OnBoundary))
                                                        {
                                                            if (Line_arc2.Length > 0.0001)
                                                            {
                                                                Line_arc2.Layer = A1.Layer;
                                                                Line_arc2.ColorIndex = 1;
                                                                //BTrecord.AppendEntity(Line_arc2);
                                                                // Trans1.AddNewlyCreatedDBObject(Line_arc2, true);
                                                                Array.Resize(ref Lines, idx1);
                                                                Lines[idx1 - 1] = Line_arc2;
                                                                idx1 = idx1 + 1;
                                                            }
                                                        }

                                                        if (Pc22 == PointContainment.Outside && (Pc33 == PointContainment.Inside || Pc33 == PointContainment.OnBoundary))
                                                        {
                                                            Point3dCollection ColPS = Functions.Intersect_on_both_operands(Line_arc2, clip_poly);
                                                            if (ColPS.Count > 0)
                                                            {
                                                                Line_arc2.StartPoint = ColPS[0];
                                                                if (Line_arc2.Length > 0.0001)
                                                                {
                                                                    Line_arc2.Layer = A1.Layer;
                                                                    Line_arc2.ColorIndex = 1;
                                                                    // BTrecord.AppendEntity(Line_arc2);
                                                                    // Trans1.AddNewlyCreatedDBObject(Line_arc2, true);
                                                                    Array.Resize(ref Lines, idx1);
                                                                    Lines[idx1 - 1] = Line_arc2;
                                                                    idx1 = idx1 + 1;
                                                                }
                                                            }
                                                        }

                                                        if ((Pc22 == PointContainment.Inside || Pc22 == PointContainment.OnBoundary) && Pc33 == PointContainment.Outside)
                                                        {
                                                            Point3dCollection ColPS = Functions.Intersect_on_both_operands(Line_arc2, clip_poly);
                                                            if (ColPS.Count > 0)
                                                            {
                                                                Line_arc2.EndPoint = ColPS[0];
                                                                if (Line_arc2.Length > 0.0001)
                                                                {
                                                                    Line_arc2.Layer = A1.Layer;
                                                                    Line_arc2.ColorIndex = 1;
                                                                    //BTrecord.AppendEntity(Line_arc2);
                                                                    //Trans1.AddNewlyCreatedDBObject(Line_arc2, true);
                                                                    Array.Resize(ref Lines, idx1);
                                                                    Lines[idx1 - 1] = Line_arc2;
                                                                    idx1 = idx1 + 1;
                                                                }
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

                                }
                            }
                        }
                    }
                }

                if (ft != 0 && bk != 0) Cub3d.Erase();
                if (ft == 0 && bk != 0) RegionBK.Erase();
                if (ft != 0 && bk == 0) Regionft.Erase();
                Poly3D_vt.Erase();

                if (Col_delete != null)
                {
                    foreach (DBObject obj1 in Col_delete)
                    {
                        Entity ent1 = obj1 as Entity;
                        if (ent1 != null)
                        {
                            ent1.Erase();
                        }
                    }
                }



                if (Lines != null)
                {
                    if (Lines.Length > 0)
                    {
                        for (int i = 0; i < Lines.Length; ++i)
                        {
                            Lines[i].TransformBy(Matrix3d.Scaling((Lines[i].Length + Extension_value) / Lines[i].Length, Lines[i].StartPoint));
                            Lines[i].TransformBy(Matrix3d.Scaling((Lines[i].Length + Extension_value) / Lines[i].Length, Lines[i].EndPoint));
                        }


                        int index1 = 0;
                        do
                        {
                            Line l1 = Lines[index1];
                            double b1 = l1.Angle;
                            Point3d p11 = l1.StartPoint;
                            Point3d p12 = l1.EndPoint;

                            for (int index2 = 0; index2 < Lines.Length; ++index2)
                            {
                                if (index1 != index2)
                                {
                                    Line l2 = Lines[index2];

                                    if (Math.Round(l2.Length, 4) == 0)
                                    {
                                        Lines[index2] = Lines[Lines.Length - 1];
                                        Lines[index1] = l1;
                                        index2 = Lines.Length;
                                        Array.Resize(ref Lines, Lines.Length - 1);
                                        index1 = index1 - 1;
                                    }

                                    double b2 = l2.Angle;
                                    Point3d p21 = l2.StartPoint;
                                    Point3d p22 = l2.EndPoint;

                                    if (Math.Round(b1, 3) == Math.Round(b2, 3) ||
                                        Math.Round(b1, 3) == Math.Round(b2 + Math.PI, 3) ||
                                        Math.Round(b1 + Math.PI, 3) == Math.Round(b2, 3) ||
                                        Math.Round(b1, 3) == Math.Round(b2 - Math.PI, 3) ||
                                        Math.Round(b1 - Math.PI, 3) == Math.Round(b2, 3) ||
                                        Math.Round(b1 - Math.PI, 3) == Math.Round(b2 + Math.PI, 3) ||
                                        Math.Round(b1 + Math.PI, 3) == Math.Round(b2 - Math.PI, 3))
                                    {

                                        bool Changed = false;

                                        if (pointpelinie(l1, p21) == true && pointpelinie(l1, p22) == true)
                                        {
                                            Lines[index2] = Lines[Lines.Length - 1];
                                            Array.Resize(ref Lines, Lines.Length - 1);
                                            index2 = Lines.Length;
                                        }
                                        else
                                        {

                                            if (pointpelinie(l2, p11) == false)
                                            {
                                                if (pointpelinie(l1, p21) == true && pointpelinie(l1, p22) == false)
                                                {
                                                    l1.EndPoint = p22;
                                                    Changed = true;
                                                }

                                                if (pointpelinie(l1, p22) == true && pointpelinie(l1, p21) == false)
                                                {
                                                    l1.EndPoint = p21;
                                                    Changed = true;
                                                }
                                            }


                                            if (pointpelinie(l2, p11) == true)
                                            {
                                                if (pointpelinie(l1, p21) == true && pointpelinie(l1, p22) == false)
                                                {
                                                    l1.StartPoint = p22;
                                                    Changed = true;
                                                }

                                                if (pointpelinie(l1, p22) == true && pointpelinie(l1, p21) == false)
                                                {
                                                    l1.StartPoint = p21;
                                                    Changed = true;
                                                }

                                                if (pointpelinie(l2, p12) == true)
                                                {
                                                    l1.StartPoint = p21;
                                                    l1.EndPoint = p22;
                                                    Changed = true;
                                                }
                                            }

                                        }
                                        if (Changed == true)
                                        {
                                            Lines[index1] = l1;
                                            index1 = index1 - 1;

                                            Lines[index2] = Lines[Lines.Length - 1];
                                            Array.Resize(ref Lines, Lines.Length - 1);
                                            index2 = Lines.Length;
                                        }

                                    }
                                }
                            }

                            index1 = index1 + 1;

                        } while (index1 < Lines.Length);





                        if (Col_cercuriPS != null)
                        {
                            if (Col_cercuriPS.Count > 0)
                            {


                                for (int i = 0; i < Lines.Length; ++i)
                                {
                                    Line L1 = Lines[i];

                                    ObjectIdCollection Colectie_cercuri_online = new ObjectIdCollection();
                                    ObjectIdCollection Colectie_elipse_online = new ObjectIdCollection();

                                    foreach (ObjectId obj_id1 in Col_cercuriPS)
                                    {
                                        Circle Circ1 = Trans1.GetObject(obj_id1, OpenMode.ForRead) as Circle;
                                        if (Circ1 != null)
                                        {
                                            Point3dCollection Coll1 = new Point3dCollection();
                                            L1.IntersectWith(Circ1, Intersect.OnBothOperands, Coll1, IntPtr.Zero, IntPtr.Zero);
                                            if (Coll1.Count == 1)
                                            {
                                                Point3d cen1 = Circ1.Center;


                                                if (L1.GetClosestPointTo(cen1, Vector3d.ZAxis, false).DistanceTo(cen1) < 0.01)
                                                {

                                                    Colectie_cercuri_online.Add(obj_id1);

                                                }
                                            }
                                        }

                                        Ellipse el1 = Trans1.GetObject(obj_id1, OpenMode.ForRead) as Ellipse;
                                        if (el1 != null)
                                        {
                                            Point3dCollection Coll2 = new Point3dCollection();
                                            L1.IntersectWith(el1, Intersect.OnBothOperands, Coll2, IntPtr.Zero, IntPtr.Zero);
                                            if (Coll2.Count == 1)
                                            {
                                                Point3d cen1 = el1.Center;


                                                if (L1.GetClosestPointTo(cen1, Vector3d.ZAxis, false).DistanceTo(cen1) < 0.01)
                                                {

                                                    Colectie_elipse_online.Add(obj_id1);

                                                }
                                            }

                                        }

                                    }

                                    if (Colectie_cercuri_online.Count > 0)
                                    {



                                        foreach (ObjectId ent_id1 in Colectie_cercuri_online)
                                        {
                                            Circle Circ1 = Trans1.GetObject(ent_id1, OpenMode.ForRead) as Circle;
                                            Point3dCollection Coll1 = new Point3dCollection();
                                            L1.IntersectWith(Circ1, Intersect.OnBothOperands, Coll1, IntPtr.Zero, IntPtr.Zero);
                                            if (Coll1.Count == 1)
                                            {
                                                Point3d Int_on_line = L1.GetClosestPointTo(Coll1[0], Vector3d.ZAxis, false);
                                                DBObjectCollection cercuricol = new DBObjectCollection();
                                                cercuricol.Add(Circ1);
                                                DBObjectCollection Region_cercuri = new DBObjectCollection();
                                                Region_cercuri = Autodesk.AutoCAD.DatabaseServices.Region.CreateFromCurves(cercuricol);
                                                Autodesk.AutoCAD.DatabaseServices.Region Regioncerc = Region_cercuri[0] as Autodesk.AutoCAD.DatabaseServices.Region;

                                                PointContainment Pc_start = Get_Point_Containment(Regioncerc, L1.StartPoint);
                                                PointContainment Pc_end = Get_Point_Containment(Regioncerc, L1.EndPoint);
                                                if (Pc_start == PointContainment.Inside && Pc_end == PointContainment.Outside)
                                                {
                                                    Line Lt = new Line(L1.StartPoint, Int_on_line);
                                                    double SCfac = (Circ1.Diameter + Extension_value) / Lt.Length;
                                                    Lt.TransformBy(Matrix3d.Scaling(SCfac, Int_on_line));

                                                    Line Line_comp = new Line(Lt.StartPoint, L1.EndPoint);

                                                    if (Line_comp.Length > L1.Length)
                                                    {
                                                        L1.StartPoint = Lt.StartPoint;

                                                        Line Line_perp = new Line(L1.StartPoint, L1.EndPoint);
                                                        Line_perp.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, Circ1.Center));

                                                        Point3dCollection Coll3 = new Point3dCollection();
                                                        Line_perp.IntersectWith(Circ1, Intersect.ExtendThis, Coll3, IntPtr.Zero, IntPtr.Zero);

                                                        if (Coll3.Count == 2)
                                                        {
                                                            Line_perp = new Line(Coll3[0], Coll3[1]);
                                                            Line_perp.TransformBy(Matrix3d.Scaling((Line_perp.Length + 2 * Extension_value) / Line_perp.Length, Circ1.Center));
                                                            Array.Resize(ref Lines, Lines.Length + 1);
                                                            Lines[Lines.Length - 1] = Line_perp;
                                                        }


                                                    }


                                                }
                                                if (Pc_start == PointContainment.Outside && Pc_end == PointContainment.Inside)
                                                {
                                                    Line Lt = new Line(L1.EndPoint, Int_on_line);
                                                    double SCfac = (Circ1.Diameter + Extension_value) / Lt.Length;
                                                    Lt.TransformBy(Matrix3d.Scaling(SCfac, Int_on_line));


                                                    Line Line_comp = new Line(Lt.StartPoint, L1.StartPoint);

                                                    if (Line_comp.Length > L1.Length)
                                                    {
                                                        L1.EndPoint = Lt.StartPoint;
                                                        Line Line_perp = new Line(L1.StartPoint, L1.EndPoint);
                                                        Line_perp.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, Circ1.Center));

                                                        Point3dCollection Coll3 = new Point3dCollection();
                                                        Line_perp.IntersectWith(Circ1, Intersect.ExtendThis, Coll3, IntPtr.Zero, IntPtr.Zero);

                                                        if (Coll3.Count == 2)
                                                        {
                                                            Line_perp = new Line(Coll3[0], Coll3[1]);
                                                            Line_perp.TransformBy(Matrix3d.Scaling((Line_perp.Length + 2 * Extension_value) / Line_perp.Length, Circ1.Center));
                                                            Array.Resize(ref Lines, Lines.Length + 1);
                                                            Lines[Lines.Length - 1] = Line_perp;
                                                        }
                                                    }
                                                }

                                            }

                                            else if (Coll1.Count == 2)
                                            {
                                                Point3d Int_on_line1 = L1.GetClosestPointTo(Coll1[0], Vector3d.ZAxis, false);
                                                Point3d Int_on_line2 = L1.GetClosestPointTo(Coll1[1], Vector3d.ZAxis, false);

                                                Line Line_perp = new Line(Int_on_line1, Int_on_line2);
                                                Line_perp.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, Circ1.Center));

                                                Line_perp.TransformBy(Matrix3d.Scaling((Line_perp.Length + 2 * Extension_value) / Line_perp.Length, Circ1.Center));
                                                Array.Resize(ref Lines, Lines.Length + 1);
                                                Lines[Lines.Length - 1] = Line_perp;
                                            }
                                        }
                                    }


                                    if (Colectie_elipse_online.Count > 0)
                                    {
                                        foreach (ObjectId ent_id1 in Colectie_elipse_online)
                                        {
                                            Ellipse Ellipsa1 = Trans1.GetObject(ent_id1, OpenMode.ForRead) as Ellipse;
                                            if (Ellipsa1 != null)
                                            {
                                                Point3dCollection Coll1 = new Point3dCollection();
                                                L1.IntersectWith(Ellipsa1, Intersect.OnBothOperands, Coll1, IntPtr.Zero, IntPtr.Zero);

                                                if (Coll1.Count == 1)
                                                {
                                                    Point3d Int_on_line = L1.GetClosestPointTo(Coll1[0], Vector3d.ZAxis, false);
                                                    DBObjectCollection ellcol = new DBObjectCollection();
                                                    ellcol.Add(Ellipsa1);
                                                    DBObjectCollection Region_ell = new DBObjectCollection();
                                                    Region_ell = Autodesk.AutoCAD.DatabaseServices.Region.CreateFromCurves(ellcol);
                                                    Autodesk.AutoCAD.DatabaseServices.Region Regioncerc = Region_ell[0] as Autodesk.AutoCAD.DatabaseServices.Region;

                                                    PointContainment Pc_start = Get_Point_Containment(Regioncerc, L1.StartPoint);
                                                    PointContainment Pc_end = Get_Point_Containment(Regioncerc, L1.EndPoint);
                                                    if (Pc_start == PointContainment.Inside && Pc_end == PointContainment.Outside)
                                                    {
                                                        Line Lt = new Line(L1.StartPoint, Int_on_line);

                                                        Point3dCollection Coll2 = new Point3dCollection();
                                                        Lt.IntersectWith(Ellipsa1, Intersect.ExtendThis, Coll2, IntPtr.Zero, IntPtr.Zero);

                                                        if (Coll2.Count == 2)
                                                        {
                                                            Lt = new Line(Coll2[0], Coll2[1]);
                                                            double SCfac = (Lt.Length + Extension_value) / Lt.Length;
                                                            Lt.TransformBy(Matrix3d.Scaling(SCfac, Int_on_line));

                                                            Line Line_comp1 = new Line(Lt.StartPoint, L1.EndPoint);
                                                            Line Line_comp2 = new Line(Lt.EndPoint, L1.EndPoint);
                                                            Line Line_comp = new Line();
                                                            if (Line_comp1.Length > Line_comp2.Length)
                                                            {
                                                                Line_comp = Line_comp1;
                                                            }
                                                            else
                                                            {
                                                                Line_comp = Line_comp2;
                                                            }

                                                            if (L1.Length < Line_comp.Length)
                                                            {
                                                                L1.StartPoint = Line_comp.StartPoint;
                                                                L1.EndPoint = Line_comp.EndPoint;

                                                                Line Line_perp = new Line(Coll2[0], Coll2[1]);
                                                                Line_perp.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, Ellipsa1.Center));

                                                                Point3dCollection Coll3 = new Point3dCollection();
                                                                Line_perp.IntersectWith(Ellipsa1, Intersect.ExtendThis, Coll3, IntPtr.Zero, IntPtr.Zero);

                                                                if (Coll3.Count == 2)
                                                                {
                                                                    Line_perp = new Line(Coll3[0], Coll3[1]);
                                                                    Line_perp.TransformBy(Matrix3d.Scaling((Line_perp.Length + 2 * Extension_value) / Line_perp.Length, Ellipsa1.Center));
                                                                    Array.Resize(ref Lines, Lines.Length + 1);
                                                                    Lines[Lines.Length - 1] = Line_perp;
                                                                }


                                                            }


                                                        }
                                                    }

                                                    if (Pc_start == PointContainment.Outside && Pc_end == PointContainment.Inside)
                                                    {
                                                        Line Lt = new Line(L1.EndPoint, Int_on_line);

                                                        Point3dCollection Coll2 = new Point3dCollection();
                                                        Lt.IntersectWith(Ellipsa1, Intersect.ExtendThis, Coll2, IntPtr.Zero, IntPtr.Zero);

                                                        if (Coll2.Count == 2)
                                                        {
                                                            Lt = new Line(Coll2[0], Coll2[1]);
                                                            double SCfac = (Lt.Length + Extension_value) / Lt.Length;
                                                            Lt.TransformBy(Matrix3d.Scaling(SCfac, Int_on_line));


                                                            Line Line_comp1 = new Line(Lt.StartPoint, L1.StartPoint);
                                                            Line Line_comp2 = new Line(Lt.EndPoint, L1.StartPoint);
                                                            Line Line_comp = new Line();
                                                            if (Line_comp1.Length > Line_comp2.Length)
                                                            {
                                                                Line_comp = Line_comp1;
                                                            }
                                                            else
                                                            {
                                                                Line_comp = Line_comp2;
                                                            }

                                                            if (L1.Length < Line_comp.Length)
                                                            {
                                                                L1.StartPoint = Line_comp.StartPoint;
                                                                L1.EndPoint = Line_comp.EndPoint;

                                                                Line Line_perp = new Line(Coll2[0], Coll2[1]);
                                                                Line_perp.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, Ellipsa1.Center));

                                                                Point3dCollection Coll3 = new Point3dCollection();
                                                                Line_perp.IntersectWith(Ellipsa1, Intersect.ExtendThis, Coll3, IntPtr.Zero, IntPtr.Zero);

                                                                if (Coll3.Count == 2)
                                                                {
                                                                    Line_perp = new Line(Coll3[0], Coll3[1]);
                                                                    Line_perp.TransformBy(Matrix3d.Scaling((Line_perp.Length + 2*Extension_value) / Line_perp.Length, Ellipsa1.Center));
                                                                    Array.Resize(ref Lines, Lines.Length + 1);
                                                                    Lines[Lines.Length - 1] = Line_perp;
                                                                }

                                                            }

                                                        }
                                                    }
                                                }

                                                else if (Coll1.Count == 2)
                                                {
                                                    Point3d Int_on_line1 = L1.GetClosestPointTo(Coll1[0], Vector3d.ZAxis, false);
                                                    Point3d Int_on_line2 = L1.GetClosestPointTo(Coll1[1], Vector3d.ZAxis, false);

                                                    Line Line_perp = new Line(Int_on_line1, Int_on_line2);
                                                    Line_perp.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, Ellipsa1.Center));

                                                    Point3dCollection Coll3 = new Point3dCollection();
                                                    Line_perp.IntersectWith(Ellipsa1, Intersect.ExtendThis, Coll3, IntPtr.Zero, IntPtr.Zero);

                                                    if (Coll3.Count == 2)
                                                    {
                                                        Line_perp = new Line(Coll3[0], Coll3[1]);
                                                        Line_perp.TransformBy(Matrix3d.Scaling((Line_perp.Length + 2 * Extension_value) / Line_perp.Length, Ellipsa1.Center));
                                                        Array.Resize(ref Lines, Lines.Length + 1);
                                                        Lines[Lines.Length - 1] = Line_perp;
                                                    }

                                                }


                                            }
                                        }


                                    }
                                }

                                index1 = 0;
                                do
                                {
                                    Line l1 = Lines[index1];
                                    double b1 = l1.Angle;
                                    Point3d p11 = l1.StartPoint;
                                    Point3d p12 = l1.EndPoint;

                                    for (int index2 = 0; index2 < Lines.Length; ++index2)
                                    {
                                        if (index1 != index2)
                                        {
                                            Line l2 = Lines[index2];

                                            if (Math.Round(l2.Length, 4) == 0)
                                            {
                                                Lines[index2] = Lines[Lines.Length - 1];
                                                Lines[index1] = l1;
                                                index2 = Lines.Length;
                                                Array.Resize(ref Lines, Lines.Length - 1);
                                                index1 = index1 - 1;
                                            }

                                            double b2 = l2.Angle;
                                            Point3d p21 = l2.StartPoint;
                                            Point3d p22 = l2.EndPoint;

                                            if (Math.Round(b1, 3) == Math.Round(b2, 3) ||
                                                Math.Round(b1, 3) == Math.Round(b2 + Math.PI, 3) ||
                                                Math.Round(b1 + Math.PI, 3) == Math.Round(b2, 3) ||
                                                Math.Round(b1, 3) == Math.Round(b2 - Math.PI, 3) ||
                                                Math.Round(b1 - Math.PI, 3) == Math.Round(b2, 3) ||
                                                Math.Round(b1 - Math.PI, 3) == Math.Round(b2 + Math.PI, 3) ||
                                                Math.Round(b1 + Math.PI, 3) == Math.Round(b2 - Math.PI, 3))
                                            {

                                                bool Changed = false;

                                                if (pointpelinie(l1, p21) == true && pointpelinie(l1, p22) == true)
                                                {
                                                    Lines[index2] = Lines[Lines.Length - 1];
                                                    Array.Resize(ref Lines, Lines.Length - 1);
                                                    index2 = Lines.Length;
                                                }
                                                else
                                                {

                                                    if (pointpelinie(l2, p11) == false)
                                                    {
                                                        if (pointpelinie(l1, p21) == true && pointpelinie(l1, p22) == false)
                                                        {
                                                            l1.EndPoint = p22;
                                                            Changed = true;
                                                        }

                                                        if (pointpelinie(l1, p22) == true && pointpelinie(l1, p21) == false)
                                                        {
                                                            l1.EndPoint = p21;
                                                            Changed = true;
                                                        }
                                                    }


                                                    if (pointpelinie(l2, p11) == true)
                                                    {
                                                        if (pointpelinie(l1, p21) == true && pointpelinie(l1, p22) == false)
                                                        {
                                                            l1.StartPoint = p22;
                                                            Changed = true;
                                                        }

                                                        if (pointpelinie(l1, p22) == true && pointpelinie(l1, p21) == false)
                                                        {
                                                            l1.StartPoint = p21;
                                                            Changed = true;
                                                        }

                                                        if (pointpelinie(l2, p12) == true)
                                                        {
                                                            l1.StartPoint = p21;
                                                            l1.EndPoint = p22;
                                                            Changed = true;
                                                        }
                                                    }

                                                }
                                                if (Changed == true)
                                                {
                                                    Lines[index1] = l1;
                                                    index1 = index1 - 1;

                                                    Lines[index2] = Lines[Lines.Length - 1];
                                                    Array.Resize(ref Lines, Lines.Length - 1);
                                                    index2 = Lines.Length;
                                                }

                                            }
                                        }
                                    }

                                    index1 = index1 + 1;

                                } while (index1 < Lines.Length);

                            }
                        }



                        for (int i = 0; i < Lines.Length; ++i)
                        {
                            Line Line_new = new Line();
                            Line_new = Lines[i].Clone() as Line;
                            Line_new.Layer = layer_name;
                            Line_new.ColorIndex = 256;
                            Line_new.Linetype = "ByLayer";
                            Line_new.LinetypeScale = Ltscale;

                            BTrecord.AppendEntity(Line_new);
                            Trans1.AddNewlyCreatedDBObject(Line_new, true);

                        }


                    }
                    if (Col_cercuriMS != null)
                    {
                        if (Col_cercuriMS.Count > 0)
                        {
                            foreach (ObjectId obj1 in Col_cercuriMS)
                            {
                                Entity ent1 = Trans1.GetObject(obj1, OpenMode.ForWrite) as Entity;
                                if (ent1 != null)
                                {
                                    ent1.Erase();
                                }
                            }
                        }
                    }

                    if (Col_cercuriPS != null)
                    {
                        if (Col_cercuriPS.Count > 0)
                        {
                            foreach (ObjectId obj1 in Col_cercuriPS)
                            {
                                Entity ent1 = Trans1.GetObject(obj1, OpenMode.ForWrite) as Entity;
                                if (ent1 != null)
                                {
                                    ent1.Erase();
                                }
                            }
                        }
                    }

                    Trans1.Commit();
                }
                else
                {
                    MessageBox.Show("No CaDWorx_Plant_Component or Line on " + "\u0022" + "cl" + "\u0022" + " layer found ");
                    Freeze_operations = false;
                    return;
                }


            }



        }


        private bool pointpelinie(Line l1, Point3d pt1)
        {
            bool este1 = true;

            Point3d p2 = l1.GetClosestPointTo(pt1, Vector3d.ZAxis, false);
            double d = pt1.DistanceTo(p2);

            if (d > 0.0001)
            {
                este1 = false;
            }


            return este1;

        }

        private Matrix3d GetArcMatrix(Arc Arc1)
        {
            Point3d origin = Arc1.StartPoint;
            Vector3d zAxis = Arc1.Normal;
            Vector3d xAxis = origin.GetVectorTo(Arc1.EndPoint);
            Vector3d yAxis = zAxis.CrossProduct(xAxis);
            return Matrix3d.AlignCoordinateSystem(
            Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis,
            origin, xAxis, yAxis, zAxis);
        }


        public static PointContainment Get_Point_Containment(Autodesk.AutoCAD.DatabaseServices.Region Region1, Point3d Pt1)
        {
            PointContainment result = PointContainment.Outside;

            /// Get a Brep object representing the region:
            using (Brep brep = new Brep(Region1))
            {
                if (brep != null)
                {

                    // Get the PointContainment and the BrepEntity at the given point:

                    using (BrepEntity ent = brep.GetPointContainment(Pt1, out result))
                    {

                        /// GetPointContainment() returns PointContainment.OnBoundary
                        /// when the picked point is either inside the region's area
                        /// or exactly on an edge.
                        /// 
                        /// So, to distinguish between a point on an edge and a point
                        /// inside the region, we must check the type of the returned
                        /// BrepEntity:
                        /// 
                        /// If the picked point was on an edge, the returned BrepEntity 
                        /// will be an Edge object. If the point was inside the boundary,
                        /// the returned BrepEntity will be a Face object.
                        //
                        /// So if the returned BrepEntity's type is a Face, we return
                        /// PointContainment.Inside:



                        if (ent is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                        {
                            result = PointContainment.Inside;
                        }
                    }
                }
            }
            return result;
        }

        private void Load_clients_to_combobox()
        {
            string fisier_config = "G:\\_HMM\\Design\\Facilities\\F-Gen\\CLTool\\CL_CLients.xlsx";
            if (System.IO.File.Exists(fisier_config) == true)
            {
                Microsoft.Office.Interop.Excel.Application Excel1 = new Microsoft.Office.Interop.Excel.Application();

                Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;
                Microsoft.Office.Interop.Excel.Worksheet W1 = null;
                try
                {
                    if (Excel1 == null)
                    {
                        MessageBox.Show("PROBLEM WITH EXCEL!");
                        return;
                    }
                    Excel1.Visible = false;

                    Workbook1 = Excel1.Workbooks.Open(fisier_config);
                    W1 = Workbook1.Worksheets[1];

                    dtc = Functions.Build_Data_table_CONFIG_from_excel(W1, 2);
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
                if (dtc.Rows.Count > 0)
                {
                    for (int i = 0; i < dtc.Rows.Count; ++i)
                    {
                        if (dtc.Rows[i][0] != DBNull.Value)
                        {
                            String client1 = dtc.Rows[i][0].ToString();
                            comboBox_Client.Items.Add(client1);

                        }
                    }
                    comboBox_Client.SelectedIndex = 0;
                }

            }
            else
            {
                MessageBox.Show("CL_CLients.xlsx not found!");
                Freeze_operations = false;
                return;
            }
        }

        private void creaza_line_types_file()
        {

            if (dtc != null)
            {
                if (dtc.Rows.Count > 0)
                {

                    for (int i = 0; i < dtc.Rows.Count; ++i)
                    {
                        if (dtc.Rows[i][0] != DBNull.Value && dtc.Rows[i][4] != DBNull.Value && dtc.Rows[i][5] != DBNull.Value)
                        {
                            String client1 = dtc.Rows[i][0].ToString();
                            if (String.IsNullOrEmpty(client1) == false)
                            {
                                string Fisier = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + client1 + ".lin";
                                if (System.IO.File.Exists(Fisier) == true)
                                {
                                    System.IO.File.Delete(Fisier);
                                }
                                System.IO.FileStream fs = System.IO.File.Create(Fisier);
                                fs.Close();

                                using (System.IO.StreamWriter sw1 = new System.IO.StreamWriter(Fisier))
                                {
                                    sw1.Write(dtc.Rows[i][4].ToString() + "\r\n" + dtc.Rows[i][5].ToString());
                                    sw1.Close();
                                }
                            }
                        }

                    }
                }
            }
        }


        private void Incarca_line_type(String LineType_name, String Linetype_file, Boolean overwrite_from_disk)
        {


            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("FILEDIA", 0);
            if (System.IO.File.Exists(Linetype_file) == true)
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                using (DocumentLock l1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        LinetypeTable Linetype_table = trans1.GetObject(ThisDrawing.Database.LinetypeTableId, OpenMode.ForRead) as LinetypeTable;
                        if (Linetype_table.Has(LineType_name) == false)
                        {
                            Linetype_table.UpgradeOpen();
                            ThisDrawing.Database.LoadLineTypeFile(LineType_name, Linetype_file);
                        }
                        else
                        {
                            if (overwrite_from_disk == true)
                            {
                                ThisDrawing.SendStringToExecute("-linetype Load " + LineType_name + "\r" + Linetype_file + "\r" + "Yes" + "\r\n", false, false, true);
                            }


                        }
                        trans1.Commit();
                    }
                }

            }



            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("FILEDIA" + "\r" + "1" + "\r", false, false, true);



        }

        private bool Verifica_daca_trebuie_procesat(ObjectId vp_id, System.Data.DataTable dt_layers_obj, ObjectId Objid_1)
        {


            if (dt_layers_obj != null && dt_frozen_xref_layers != null)
            {
                if (dt_layers_obj.Rows.Count > 0 && dt_frozen_xref_layers.Rows.Count > 0)
                {
                    for (int i = 0; i < dt_layers_obj.Rows.Count; ++i)
                    {
                        if (dt_layers_obj.Rows[i]["id"] != DBNull.Value)
                        {
                            string id1 = dt_layers_obj.Rows[i]["id"].ToString();
                            if (id1 == Objid_1.Handle.Value.ToString())
                            {
                                if (dt_layers_obj.Rows[i]["ln"] != DBNull.Value)
                                {
                                    string ln1 = dt_layers_obj.Rows[i]["ln"].ToString();
                                    string vpid1 = vp_id.Handle.Value.ToString();
                                    for (int j = 0; j < dt_frozen_xref_layers.Rows.Count; ++j)
                                    {
                                        if (dt_frozen_xref_layers.Rows[j]["id"] != DBNull.Value && dt_frozen_xref_layers.Rows[j]["ln"] != DBNull.Value)
                                        {
                                            string ln2 = dt_frozen_xref_layers.Rows[j]["ln"].ToString();
                                            string vpid2 = dt_frozen_xref_layers.Rows[j]["id"].ToString();

                                            if (vpid2 == vpid1 && ln2 == ln1)
                                            {
                                                return false;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return true;
        }

        private bool Verifica_layerul_obj(string ln)
        {


            if (dt_frozen_xref_layers != null)
            {
                if (dt_frozen_xref_layers.Rows.Count > 0)
                {
                    for (int i = 0; i < dt_frozen_xref_layers.Rows.Count; ++i)
                    {
                        if (dt_frozen_xref_layers.Rows[i]["ln"] != DBNull.Value)
                        {
                            string layer_name1 = dt_frozen_xref_layers.Rows[i]["ln"].ToString();
                            if (layer_name1 == ln)
                            {
                                return false;
                            }
                        }
                    }
                }
            }
            return true;
        }



    }
}
