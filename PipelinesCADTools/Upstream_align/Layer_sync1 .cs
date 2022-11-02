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
    public partial class layer_sync1 : Form
    {

        Layer_sync_mainform LS = null;
        string layer_state1 = "";
        List<string> lista_fisiere_dest;




        private void set_enable_false(object sender)
        {
            List<System.Windows.Forms.Button> lista_butoane = new List<Button>();
            lista_butoane.Add(Button_apply_layer_state);
            lista_butoane.Add(button_select_dwg_source);
            lista_butoane.Add(button_select_dwg_destination);

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
            lista_butoane.Add(Button_apply_layer_state);
            lista_butoane.Add(button_select_dwg_source);
            lista_butoane.Add(button_select_dwg_destination);
            foreach (System.Windows.Forms.Button bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }

        public layer_sync1()
        {
            InitializeComponent();
        }

        private void layer_sync_page_load(object sender, EventArgs e)
        {
            LS = this.MdiParent as Layer_sync_mainform;
        }



        private void Button_apply_layer_state_Click(object sender, EventArgs e)
        {
            set_enable_false(sender);
            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        if (lista_fisiere_dest != null && lista_fisiere_dest.Count > 0)
                        {
                            for (int i = 0; i < lista_fisiere_dest.Count; ++i)
                            {
                                string fisier = Convert.ToString(lista_fisiere_dest[i]);
                                if (System.IO.File.Exists(fisier) == true)
                                {

                                    using (Database Database2 = new Database(false, true))
                                    {
                                        Database2.ReadDwgFile(fisier, FileOpenMode.OpenForReadAndWriteNoShare, true, null);
                                        Database2.CloseInput(true);
                                        HostApplicationServices.WorkingDatabase = Database2;
                                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans2 = Database2.TransactionManager.StartTransaction())
                                        {
                                            Database2.Audit(true, true);
                                            LayerStateManager layer_state_manager;
                                            layer_state_manager = Database2.LayerStateManager;

                                            string path1 = "C:\\Users\\pop70694\\Documents\\Work Files\\2018-04-20 Layer sync\\Alignments\\" + layer_state1 + ".las";
                                            if (System.IO.File.Exists(path1) == true)
                                            {


                                                if (layer_state_manager.HasLayerState(layer_state1) == true)
                                                {
                                                    layer_state_manager.DeleteLayerState(layer_state1);
                                                }



                                                try
                                                {
                                                    layer_state_manager.ImportLayerState(path1);
                                                }
                                                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                                                {
                                                    Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(ex.Message);
                                                }
                                                

                                                // layer_state_manager.ImportLayerStateFromDb(layer_state1, hostbase_path);

                                                layer_state_manager.RestoreLayerState(layer_state1, ObjectId.Null, 0, LayerStateMasks.Color | LayerStateMasks.CurrentViewport
                                                                            | LayerStateMasks.Frozen | LayerStateMasks.LineType | LayerStateMasks.LineWeight | LayerStateMasks.Locked
                                                                            | LayerStateMasks.NewViewport | LayerStateMasks.On | LayerStateMasks.Plot | LayerStateMasks.PlotStyle);
                                                Database2.Visretain = true;
                                                Trans2.Commit();

                                            }





                                            HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                                            if (System.IO.File.Exists(path1) == true) Database2.SaveAs(fisier, true, DwgVersion.Current, ThisDrawing.Database.SecurityParameters);
                                        }
                                    }
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

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();
        }

        private void button_select_dwg_source_Click(object sender, EventArgs e)
        {

            using (OpenFileDialog fbd = new OpenFileDialog())
            {
                fbd.Multiselect = false;
                fbd.Filter = "drawings (*.dwg)|*.dwg";
                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {

                    string hostbase_path = fbd.FileName;
                    set_enable_false(sender);

                    ObjectId[] Empty_array = null;
                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;

                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    try
                    {
                        using (DocumentLock lock1 = ThisDrawing.LockDocument())
                        {
                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                            {
                                BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                                BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                                using (Database Database2 = new Database(false, true))
                                {
                                    Database2.ReadDwgFile(hostbase_path, FileOpenMode.OpenForReadAndAllShare, false, null);
                                    Database2.CloseInput(true);
                                    HostApplicationServices.WorkingDatabase = Database2;
                                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans2 = Database2.TransactionManager.StartTransaction())
                                    {
                                        LayerStateManager layer_state_manager;
                                        layer_state_manager = Database2.LayerStateManager;



                                        DBDictionary dict1 = Trans2.GetObject(layer_state_manager.LayerStatesDictionaryId(true), OpenMode.ForRead) as DBDictionary;



                                        int i = 0;
                                        foreach (DBDictionaryEntry acDbDictEnt in dict1)
                                        {
                                            if (i == 0) layer_state1 = acDbDictEnt.Key;
                                            ++i;
                                        }
                                        string path1 = "C:\\Users\\pop70694\\Documents\\Work Files\\2018-04-20 Layer sync\\Alignments\\" + layer_state1 + ".las";
                                        if (System.IO.File.Exists(path1) == true)
                                        {
                                            System.IO.File.Delete(path1);
                                        }

                                        layer_state_manager.ExportLayerState(layer_state1, path1);
                                        Trans2.Commit();
                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                                    }
                                }
                                Trans1.Commit();

                                label_base.Text = "Host Base = " + System.IO.Path.GetFileNameWithoutExtension(hostbase_path);

                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                    Editor1.SetImpliedSelection(Empty_array);
                    Editor1.WriteMessage("\nCommand:");
                    set_enable_true();
                }

            }



        }


        private void button_select_dwg_destination_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog fbd = new OpenFileDialog())
            {
                fbd.Multiselect = true;
                fbd.Filter = "drawings (*.dwg)|*.dwg";
                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    lista_fisiere_dest = new List<string>();
                    for (int i = 0; i < fbd.FileNames.Count(); ++i)
                    {
                        lista_fisiere_dest.Add(fbd.FileNames[i]);
                    }

                    for (int j = 0; j < lista_fisiere_dest.Count; ++j)
                    {
                        if (j == 0)
                        {
                            label_dwg.Text = System.IO.Path.GetFileNameWithoutExtension(Convert.ToString(lista_fisiere_dest[j]));
                        }
                        else
                        {
                            Label label1 = new Label();
                            label1.Location = new Point(label_dwg.Location.X, label_dwg.Location.Y + j * (label_dwg.Height + 8));
                            label1.BackColor = label_dwg.BackColor;
                            label1.ForeColor = label_dwg.ForeColor;
                            label1.Font = label_dwg.Font;
                            label1.Size = label_dwg.Size;
                            label1.FlatStyle = label_dwg.FlatStyle;
                            label1.Text = System.IO.Path.GetFileNameWithoutExtension(Convert.ToString(lista_fisiere_dest[j]));
                            panel_destination.Controls.Add(label1);
                            label1.Click += delegate (object s, EventArgs e1)
                            {
                                label_Click(label1, e1);
                            };

                        }


                    }
                }
                else
                {
                    lista_fisiere_dest = null;
                }
            }
        }

        private void label_Click(object sender, EventArgs e)
        {

        }

    }
}
