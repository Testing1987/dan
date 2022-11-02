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
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace SuperLayerManager
{
    public partial class Layer_manager_Form : Form
    {
        public Layer_manager_Form()
        {
            InitializeComponent();
        }
        
        Autodesk.AutoCAD.DatabaseServices.LayerTable LayerTable1;
        String[] Layer_list;
        System.Windows.Forms.TextBox[] TextBox_layer;
        System.Windows.Forms.Button[] Button_on_off;
        System.Windows.Forms.Button[] Button_th_fr;
        System.Windows.Forms.Button[] Button_change;
        System.Windows.Forms.CheckBox[] Check_box_select;
        System.Windows.Forms.Button[] Button_NO_PLOT;
        String[] Lista_layere_selectate;


        private void Layer_manager_Form_Load(object sender, EventArgs e)
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
                            Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                            LayerTable1 = (Autodesk.AutoCAD.DatabaseServices.LayerTable) Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                            int Idx1 = 0;
                            string  Clayer = ((LayerTableRecord) Trans1.GetObject(ThisDrawing.Database.Clayer,OpenMode.ForRead)).Name;
                            foreach (ObjectId ID1 in LayerTable1)
                            {
                                LayerTableRecord Layer1 = (LayerTableRecord) Trans1.GetObject(ID1,OpenMode.ForRead);                             
                                Array.Resize(ref Layer_list, Idx1 + 1);
                                Layer_list[Idx1] = Layer1.Name;
                               
                                Idx1 = Idx1 + 1;
                            }

                            Array.Sort(Layer_list);

                            Array.Resize(ref TextBox_layer, Idx1);
                            Array.Resize(ref Button_on_off, Idx1);
                            Array.Resize(ref Button_th_fr, Idx1);
                            Array.Resize(ref Button_change, Idx1);
                            Array.Resize(ref Check_box_select, Idx1);
                            Array.Resize(ref Button_NO_PLOT, Idx1);

                            for (int i = 0; i < Layer_list.Length; ++i)
                            {
                                LayerTableRecord Layer1 = (LayerTableRecord)LayerTable1[Layer_list[i]].GetObject(OpenMode.ForRead);    

                                TextBox_layer[i] = new TextBox();
                                TextBox_layer[i].Location = new System.Drawing.Point(6, 6 + i * 21);
                                TextBox_layer[i].Text = Layer_list[i];
                                TextBox_layer[i].Size = new System.Drawing.Size(panel1.Width - 202, 20);
                                TextBox_layer[i].Font = new System.Drawing.Font("Arial", 9, FontStyle.Bold); // new System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold);

                                


                                if (Layer_list[i] == Clayer)
                                {
                                    TextBox_layer[i].BackColor = System.Drawing.Color.PowderBlue;
                                }
                                else
                                {
                                    if (Layer1.IsOff == true | Layer1.IsFrozen == true)
                                    {
                                        TextBox_layer[i].BackColor = System.Drawing.Color.Red;
                                    }
                                    else
                                    {
                                        TextBox_layer[i].BackColor = System.Drawing.Color.Yellow;
                                    }
                                }

                                panel1.Controls.Add(TextBox_layer[i]);
                                TextBox_layer[i].Click += new EventHandler(Textbox_layer_click);

                                TextBox_layer[i].MouseClick += new MouseEventHandler(Textbox_layer_MOUSE_rename);


                                Button_on_off[i] = new Button();
                                Button_on_off[i].Location =  new System.Drawing.Point(TextBox_layer[i].Width +6, 6 + i * 21+1);
                                Button_on_off[i].Size = new System.Drawing.Size(35,19);
                                Button_on_off[i].ForeColor = System.Drawing.Color.Black;
                                Button_on_off[i].FlatStyle =System.Windows.Forms.FlatStyle.Popup;
                                Button_on_off[i].BackColor =System.Drawing.Color.LightGreen;
                                Button_on_off[i].Font = new System.Drawing.Font("Arial", 8, System.Drawing.FontStyle.Bold);
                               
                                if (Layer1.IsOff == false)
                                {
                                    Button_on_off[i].Text = "ON";
                                    Button_on_off[i].BackColor = System.Drawing.Color.LightGreen;
                                }
                                else
                                {
                                    Button_on_off[i].Text = "OFF";
                                    Button_on_off[i].BackColor = System.Drawing.Color.Red;
                                }

                                panel1.Controls.Add(Button_on_off[i]);

                                Button_on_off[i].Click += delegate(object s, EventArgs e1)
                                
                                {
                                    Button_on_off_click(s, e1, Layer1.Name);
                                };
                               
                                Button_change[i] = new Button();
                                Button_change[i].Location = new System.Drawing.Point(TextBox_layer[i].Width + Button_on_off[i].Width + 9, 6 + i * 21 + 1);
                                Button_change[i].Size = new System.Drawing.Size(30, 19);
                                Button_change[i].BackColor = System.Drawing.Color.LightYellow;
                                Button_change[i].FlatStyle = System.Windows.Forms.FlatStyle.Flat;
                                Button_change[i].FlatAppearance.BorderSize = 0;
                                panel1.Controls.Add(Button_change[i]);

                                Button_change[i].Click += delegate(object s, EventArgs e1)
                                {
                                    Button_change_layer_click(s, e1, Layer1.Name); 
                                };


                                Button_th_fr[i] = new Button();
                                Button_th_fr[i].Location = new System.Drawing.Point(TextBox_layer[i].Width + Button_on_off[i].Width + Button_change[i].Width + 12, 6 + i * 21 + 1);
                                Button_th_fr[i].Size = new System.Drawing.Size(35, 19);
                                Button_th_fr[i].ForeColor = System.Drawing.Color.Black;
                                Button_th_fr[i].FlatStyle = System.Windows.Forms.FlatStyle.Popup;
                                Button_th_fr[i].BackColor = System.Drawing.Color.LightGreen;
                                Button_th_fr[i].Font = new System.Drawing.Font("Arial", 8, System.Drawing.FontStyle.Bold);

                                if (Layer1.IsFrozen == false)
                                {
                                    Button_th_fr[i].Text = "TH";
                                    Button_th_fr[i].BackColor = System.Drawing.Color.LightGreen;
                                }
                                else
                                {
                                    Button_th_fr[i].Text = "FR";
                                    Button_th_fr[i].BackColor = System.Drawing.Color.Red;
                                }


                                panel1.Controls.Add(Button_th_fr[i]);


                                Button_th_fr[i].Click += delegate(object s, EventArgs e1)
                                {
                                    Button_TH_FR_click(s, e1, Layer1.Name);
                                };


                                Button_NO_PLOT[i] = new Button();
                                Button_NO_PLOT[i].Location = new System.Drawing.Point(TextBox_layer[i].Width + Button_on_off[i].Width + Button_change[i].Width + Button_th_fr[i].Width + 15, 6 + i * 21 + 1);
                                Button_NO_PLOT[i].Size = new System.Drawing.Size(35, 19);
                                Button_NO_PLOT[i].ForeColor = System.Drawing.Color.Black;
                                Button_NO_PLOT[i].FlatStyle = System.Windows.Forms.FlatStyle.Popup;
                                Button_NO_PLOT[i].BackColor = System.Drawing.Color.LightGreen;
                                Button_NO_PLOT[i].Font = new System.Drawing.Font("Arial", 8, System.Drawing.FontStyle.Bold);

                                if (Layer1.IsPlottable == true)
                                {
                                    Button_NO_PLOT[i].Text = "P";
                                    Button_NO_PLOT[i].BackColor = System.Drawing.Color.LightGreen;
                                }
                                else
                                {
                                    Button_NO_PLOT[i].Text = "NP";
                                    Button_NO_PLOT[i].BackColor = System.Drawing.Color.Red;
                                }


                                panel1.Controls.Add(Button_NO_PLOT[i]);


                                Button_NO_PLOT[i].Click += delegate(object s, EventArgs e1)
                                {
                                    Button_NO_PLOT_click(s, e1, Layer1.Name);
                                };


                                Check_box_select[i]= new CheckBox();
                                Check_box_select[i].Location = new System.Drawing.Point(TextBox_layer[i].Width + Button_on_off[i].Width + Button_change[i].Width + Button_th_fr[i].Width + Button_NO_PLOT[i].Width + 17, 6 + i * 21 + 4);
                                Check_box_select[i].Size = new System.Drawing.Size(15, 15);



                                panel1.Controls.Add(Check_box_select[i]);


                                Check_box_select[i].CheckedChanged += delegate(object s, EventArgs e1)
                                {
                                    checkBox1_CheckedChanged(s, e1, Layer1.Name);
                                };

                            }



                            Trans1.Commit();
                        }
                    
                    }

                    Button but_NoPlot_on_off = new Button();
                    but_NoPlot_on_off.Location = new System.Drawing.Point(6, 12 + Layer_list.Length * 21);
                    but_NoPlot_on_off.Size = new System.Drawing.Size(125, 30);
                    but_NoPlot_on_off.ForeColor = System.Drawing.Color.Black;
                    but_NoPlot_on_off.Font = new System.Drawing.Font("Arial", 9, System.Drawing.FontStyle.Bold);
                    but_NoPlot_on_off.Text = "NO PLOT not here";                 
                    if (LayerTable1.Has("NO PLOT") == true)
                    {
                        using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                        {
                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                            {
                                LayerTableRecord Layer_NO_PLOT = (LayerTableRecord)Trans1.GetObject(LayerTable1["NO PLOT"], OpenMode.ForRead);
                                if (Layer_NO_PLOT.IsOff == false)
                                {
                                    but_NoPlot_on_off.Text = "NO PLOT IS ON";
                                    but_NoPlot_on_off.BackColor = System.Drawing.Color.SandyBrown;
                                }
                                else
                                {
                                    but_NoPlot_on_off.Text = "NO PLOT IS OFF";
                                    but_NoPlot_on_off.BackColor = System.Drawing.Color.Red;
                                }
                            }
                        }
                    }
                    but_NoPlot_on_off.FlatStyle = FlatStyle.Standard;
                    panel1.Controls.Add(but_NoPlot_on_off);
                    but_NoPlot_on_off.Click += delegate(object s, EventArgs e1)
                    {
                        but_NoPlot_on_off_click(s, e1);
                    };


                    Button but_obj_to_NoPlot = new Button();
                    but_obj_to_NoPlot.Location = new System.Drawing.Point(6, 14 + Layer_list.Length * 21 + but_NoPlot_on_off.Height);
                    but_obj_to_NoPlot.Size = new System.Drawing.Size(125, 30);
                    but_obj_to_NoPlot.ForeColor = System.Drawing.Color.Black;
                    but_obj_to_NoPlot.Font = new System.Drawing.Font("Arial", 9, System.Drawing.FontStyle.Bold);
                    but_obj_to_NoPlot.Text = "OBJ -> NO PLOT";
                    but_obj_to_NoPlot.BackColor = System.Drawing.Color.SandyBrown;
                    but_obj_to_NoPlot.FlatStyle = FlatStyle.Standard;

                    panel1.Controls.Add(but_obj_to_NoPlot);

                    but_obj_to_NoPlot.Click += delegate(object s, EventArgs e1)
                    {
                        but_obj_to_NoPlot_click(s, e1);
                    };

                    Button but_No_plot_current = new Button();
                    but_No_plot_current.Location = new System.Drawing.Point(6, 16 + Layer_list.Length * 21 + but_NoPlot_on_off.Height + but_obj_to_NoPlot.Height);
                    but_No_plot_current.Size = new System.Drawing.Size(125, 30);
                    but_No_plot_current.ForeColor = System.Drawing.Color.Black;
                    but_No_plot_current.Font = new System.Drawing.Font("Arial", 8, System.Drawing.FontStyle.Bold);
                    but_No_plot_current.Text = "NO PLOT CURRENT";
                    but_No_plot_current.BackColor = System.Drawing.Color.Gold;
                    but_No_plot_current.FlatStyle = FlatStyle.Standard;

                    panel1.Controls.Add(but_No_plot_current);


                    but_No_plot_current.Click += delegate(object s, EventArgs e1)
                    {
                        but_No_plot_current_click(s, e1);
                    };

                    Button but_Refresh = new Button();
                    but_Refresh.Location = new System.Drawing.Point(6, 20 + Layer_list.Length * 21 + but_NoPlot_on_off.Height + but_obj_to_NoPlot.Height + but_No_plot_current.Height);
                    but_Refresh.Size = new System.Drawing.Size(125, 30);
                    but_Refresh.ForeColor = System.Drawing.Color.Black;
                    but_Refresh.Font = new System.Drawing.Font("Arial", 8, System.Drawing.FontStyle.Bold);
                    but_Refresh.Text = "REFRESH";
                    but_Refresh.BackColor = System.Drawing.Color.Aqua;
                    but_Refresh.FlatStyle = FlatStyle.Standard;

                    panel1.Controls.Add(but_Refresh);


                    but_Refresh.Click += delegate(object s, EventArgs e1)
                    {
                        but_Refresh_click(s, e1);
                    };

                    Button but_ALL_on_off = new Button();
                    but_ALL_on_off.Location = new System.Drawing.Point(6 + but_NoPlot_on_off.Width + 25, 12 + Layer_list.Length * 21);
                    but_ALL_on_off.Size = new System.Drawing.Size(125, 30);
                    but_ALL_on_off.ForeColor = System.Drawing.Color.Black;
                    but_ALL_on_off.Font = new System.Drawing.Font("Arial", 9, System.Drawing.FontStyle.Bold);
                    but_ALL_on_off.Text = "ALL OFF";
                    but_ALL_on_off.BackColor = System.Drawing.Color.LightGreen;

                    but_ALL_on_off.FlatStyle = FlatStyle.Standard;
                    panel1.Controls.Add(but_ALL_on_off);
                    but_ALL_on_off.Click += delegate(object s, EventArgs e1)
                    {
                        but_ALL_on_off_click(s, e1);
                    };

                    Button but_ALL_th_fr = new Button();
                    but_ALL_th_fr.Location = new System.Drawing.Point(6 + but_NoPlot_on_off.Width + 25, 14 + Layer_list.Length * 21 + but_NoPlot_on_off.Height);
                    but_ALL_th_fr.Size = new System.Drawing.Size(125, 30);
                    but_ALL_th_fr.ForeColor = System.Drawing.Color.Black;
                    but_ALL_th_fr.Font = new System.Drawing.Font("Arial", 9, System.Drawing.FontStyle.Bold);
                    but_ALL_th_fr.Text = "ALL FREEZE";
                    but_ALL_th_fr.BackColor = System.Drawing.Color.LightGreen;

                    but_ALL_th_fr.FlatStyle = FlatStyle.Standard;
                    panel1.Controls.Add(but_ALL_th_fr);
                    but_ALL_th_fr.Click += delegate(object s, EventArgs e1)
                    {
                        but_ALL_th_fr_click(s, e1);
                    };



                    Button but_Delete = new Button();
                    but_Delete.Location = new System.Drawing.Point(6 + but_NoPlot_on_off.Width + 25, 16 + Layer_list.Length * 21 + but_NoPlot_on_off.Height + but_obj_to_NoPlot.Height);
                    but_Delete.Size = new System.Drawing.Size(200, 30);
                    but_Delete.ForeColor = System.Drawing.Color.Black;
                    but_Delete.Font = new System.Drawing.Font("Arial", 9, System.Drawing.FontStyle.Bold);
                    but_Delete.BackColor = System.Drawing.Color.Red;
                    but_Delete.Text = "DELETE SELECTED LAYERS";

                    but_Delete.FlatStyle = FlatStyle.Standard;

                    panel1.Controls.Add(but_Delete);

                    but_Delete.Click += delegate(object s, EventArgs e1)
                    {
                        but_DELETE_click(s, e1);
                    };



                    panel1.Width = this.Width - 20;
                    panel1.Height = this.Height - 42;
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            
        

        }
        private void Textbox_layer_click(object sender, EventArgs e)
        {
            TextBox TextBox_layer1 = (TextBox)sender;

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        LayerTable1 = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        string Clayer = ((LayerTableRecord)Trans1.GetObject(ThisDrawing.Database.Clayer, OpenMode.ForRead)).Name;
                        
                        LayerTableRecord Layer_selectat = (LayerTableRecord)Trans1.GetObject(LayerTable1[TextBox_layer1.Text], OpenMode.ForRead);
                        
                        string NumeL = TextBox_layer1.Text;

                        if (NumeL.Contains("|") == false & NumeL.Contains("Defpoints") == false)
                        {
                            for (int i = 0; i < Layer_list.Length; ++i)
                            {
                                LayerTableRecord Layer1 = (LayerTableRecord)Trans1.GetObject(ThisDrawing.Database.Clayer, OpenMode.ForRead);
                                if (Layer_list[i] == Clayer)
                                {
                                    TextBox_layer[i].BackColor = System.Drawing.Color.Yellow;

                                }
                                if (Layer1.IsOff == true | Layer1.IsFrozen == true)
                                {
                                    TextBox_layer[i].BackColor = System.Drawing.Color.Red;
                                }

                                if ((Layer1.IsOff == true | Layer1.IsFrozen == true) & NumeL.Contains("|") == true)
                                {
                                    TextBox_layer[i].BackColor = System.Drawing.Color.Firebrick;
                                }

                            }

                            ThisDrawing.Database.Clayer = LayerTable1[NumeL];
                            TextBox_layer1.BackColor = System.Drawing.Color.PowderBlue;

                        }



                        
                       
                        

                        Trans1.Commit();
                    }

                }

                panel1.Width = this.Width - 20;
                panel1.Height = this.Height - 42;
                this.Refresh();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Textbox_layer_MOUSE_rename(object sender, MouseEventArgs e)
        {

            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                TextBox TextBox_layer1 = (TextBox)sender;

                try
                {
                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                            LayerTable1 = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                            string NumeL = TextBox_layer1.Text;
                            if (LayerTable1.Has(NumeL) == true)
                            {
                                LayerTableRecord Layer_selectat = (LayerTableRecord)Trans1.GetObject(LayerTable1[NumeL], OpenMode.ForRead);
                                
                                if (NumeL.Contains("|") == false & NumeL.Contains("Defpoints") == false)
                                {
                                    //string New_name = inpu


                                }


                            }

                            

                            









                            Trans1.Commit();
                        }

                    }

                    panel1.Width = this.Width - 20;
                    panel1.Height = this.Height - 42;
                    this.Refresh();
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            
           
        }
        
        private void Button_on_off_click(object sender, EventArgs e, string Layer_name)
        {
            Button Button_on_off = (Button)sender;

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        LayerTable1 = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        string Clayer = ((LayerTableRecord)Trans1.GetObject(ThisDrawing.Database.Clayer, OpenMode.ForRead)).Name;

                        LayerTableRecord Layer_selectat = (LayerTableRecord)Trans1.GetObject(LayerTable1[Layer_name], OpenMode.ForWrite);
                        int index_TB = -1;
                        for (int i = 0; i < TextBox_layer.Length; ++i)
                        {
                            if (TextBox_layer[i].Text == Layer_selectat.Name)
                            {
                                index_TB = i;
                                i = TextBox_layer.Length;
                            }
                        }


                        if (index_TB != -1)
                        {
                            if (Button_on_off.Text == "ON")
                            {
                                Layer_selectat.IsOff = true;
                                Button_on_off.Text = "OFF";
                                if (Layer_name == Clayer)
                                {
                                    TextBox_layer[index_TB ].BackColor = System.Drawing.Color.PowderBlue;
                                    

                                }
                                else
                                {
                                    TextBox_layer[index_TB ].BackColor = System.Drawing.Color.Red;
                                    
                                }

                                Button_on_off.BackColor = System.Drawing.Color.Red;

                            }

                            else
                            {
                                Layer_selectat.IsOff = false;
                                Button_on_off.Text = "ON";
                                if (Layer_name == Clayer)
                                {
                                    TextBox_layer[index_TB ].BackColor = System.Drawing.Color.PowderBlue;
                                }
                                else
                                {
                                    TextBox_layer[index_TB ].BackColor = System.Drawing.Color.Yellow;
                                }

                                Button_on_off.BackColor = System.Drawing.Color.LightGreen;

                            }
                        }
                        Trans1.Commit();
                    }

                }

                panel1.Width = this.Width - 20;
                panel1.Height = this.Height - 42;
                this.Refresh();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void Button_change_layer_click(object sender, EventArgs e, string Layer_name)
        {
            Button Button_change = (Button)sender;

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Button_change.BackColor = System.Drawing.Color.Brown;
                this.Refresh();

                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        LayerTableRecord Layer1 = (LayerTableRecord)LayerTable1[Layer_name].GetObject(OpenMode.ForRead);
                        if (Layer1 != null)
                        {

                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                            Rezultat1 = Editor1.SelectImplied();

                            if (Rezultat1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {
                                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                Prompt_rez.MessageForAdding = "\nSelect objects:";
                                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                                Prompt_rez.SingleOnly = false;
                                Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                            }



                            if (Rezultat1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {

                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                            
                            
                            for (int i = 0; i < Rezultat1.Value.Count; ++i)
                            {
                                Autodesk.AutoCAD.EditorInput.SelectedObject Obj1 = Rezultat1.Value[i];
                                Entity Ent1 = (Entity)Obj1.ObjectId.GetObject(OpenMode.ForWrite);
                                Ent1.Layer = Layer_name;

                            }

                            foreach (Control CTRL in panel1.Controls)
                            {
                                if (CTRL is Panel)
                                {
                                    CTRL.BackColor = System.Drawing.Color.LightYellow;
                                }                                  
                            }
                            Button_change.BackColor = System.Drawing.Color.Brown;
                            ObjectId[] null1 = null;

                            Editor1.SetImpliedSelection(null1);
                            Trans1.Commit();
                            
                        }



                    }

                }



        


                panel1.Width = this.Width - 20;
                panel1.Height = this.Height - 42;
                this.Refresh();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Button_TH_FR_click(object sender, EventArgs e, string Layer_name)
        {
            Button Button_TH_fr = (Button)sender;

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        LayerTable1 = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        string Clayer = ((LayerTableRecord)Trans1.GetObject(ThisDrawing.Database.Clayer, OpenMode.ForRead)).Name;

                        LayerTableRecord Layer_selectat = (LayerTableRecord)Trans1.GetObject(LayerTable1[Layer_name], OpenMode.ForWrite);
                        int index_TB = -1;
                        for (int i = 0; i < TextBox_layer.Length; ++i)
                        {
                            if (TextBox_layer[i].Text == Layer_selectat.Name)
                            {
                                index_TB = i;
                                i = TextBox_layer.Length;
                            }
                        }


                        if (index_TB != -1 & Layer_name != Clayer)
                        {
                            if (Button_TH_fr.Text == "TH")
                            {
                                Layer_selectat.IsFrozen = true;
                                Button_TH_fr.Text = "FR";

                                TextBox_layer[index_TB].BackColor = System.Drawing.Color.Red;



                                Button_TH_fr.BackColor = System.Drawing.Color.Red;

                            }

                            else
                            {
                                Layer_selectat.IsFrozen = false;
                                Button_TH_fr.Text = "TH";

                                TextBox_layer[index_TB].BackColor = System.Drawing.Color.Yellow;

                                Button_TH_fr.BackColor = System.Drawing.Color.LightGreen;

                            }
                        }
                        Trans1.Commit();
                    }

                }

                panel1.Width = this.Width - 20;
                panel1.Height = this.Height - 42;
                this.Refresh();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Button_NO_PLOT_click(object sender, EventArgs e, string Layer_name)
        {
            Button Button_PLOTTABLE = (Button)sender;

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        LayerTable1 = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                       

                        LayerTableRecord Layer_selectat = (LayerTableRecord)Trans1.GetObject(LayerTable1[Layer_name], OpenMode.ForWrite);
                        int index_TB = -1;
                        for (int i = 0; i < TextBox_layer.Length; ++i)
                        {
                            if (TextBox_layer[i].Text == Layer_selectat.Name)
                            {
                                index_TB = i;
                                i = TextBox_layer.Length;
                            }
                        }


                        if (index_TB != -1 )
                        {
                            if (Button_PLOTTABLE.Text == "P")
                            {
                                Layer_selectat.IsPlottable = false;
                                Button_PLOTTABLE.Text = "NP";
                               Button_PLOTTABLE.BackColor = System.Drawing.Color.Red;
                            }

                            else
                            {
                                Layer_selectat.IsPlottable = true;
                                Button_PLOTTABLE.Text = "P";
                                Button_PLOTTABLE.BackColor = System.Drawing.Color.LightGreen;
                            }
                        }
                        Trans1.Commit();
                    }

                }

                panel1.Width = this.Width - 20;
                panel1.Height = this.Height - 42;
                this.Refresh();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void but_No_plot_current_click(object sender, EventArgs e)
        {
            Button Button_TH_fr = (Button)sender;

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        LayerTable1 = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                       
                        
                        string Clayer = ((LayerTableRecord)Trans1.GetObject(ThisDrawing.Database.Clayer, OpenMode.ForRead)).Name;

                        
                        if (LayerTable1.Has("NO PLOT") == true)
                        {
                            if (Clayer != "NO PLOT")
                            {
                                LayerTableRecord NO_PLOT_LAYER = (LayerTableRecord)Trans1.GetObject(LayerTable1["NO PLOT"], OpenMode.ForRead);

                                ThisDrawing.Database.Clayer = NO_PLOT_LAYER.ObjectId;
                            }
                        }
                        else
                        {
                            Creaza_layer("NO PLOT", 40, false);
                            LayerTableRecord NO_PLOT_LAYER = (LayerTableRecord)Trans1.GetObject(LayerTable1["NO PLOT"], OpenMode.ForRead);

                            ThisDrawing.Database.Clayer = NO_PLOT_LAYER.ObjectId;
                        }
                        
                        
                        Trans1.Commit();
                    }

                }

                panel1.Controls.Clear();
                Layer_manager_Form_Load(sender, e);
                panel1.Width = this.Width - 20;
                panel1.Height = this.Height - 42;
                this.Refresh();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void but_obj_to_NoPlot_click(object sender, EventArgs e)
        {
            

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
               
                this.Refresh();

                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        if (LayerTable1.Has("NO PLOT") == false)
                        {
                            Creaza_layer("NO PLOT", 40, false);
                        }

                        LayerTableRecord Layer1 = (LayerTableRecord)LayerTable1["NO PLOT"].GetObject(OpenMode.ForRead);
                        if (Layer1 != null)
                        {

                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                            Rezultat1 = Editor1.SelectImplied();

                            if (Rezultat1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {
                                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                Prompt_rez.MessageForAdding = "\nSelect objects:";
                                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                                Prompt_rez.SingleOnly = false;
                                Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                            }



                            if (Rezultat1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {

                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);


                            for (int i = 0; i < Rezultat1.Value.Count; ++i)
                            {
                                Autodesk.AutoCAD.EditorInput.SelectedObject Obj1 = Rezultat1.Value[i];
                                Entity Ent1 = (Entity)Obj1.ObjectId.GetObject(OpenMode.ForWrite);
                                Ent1.Layer = "NO PLOT";

                            }

                            foreach (Control CTRL in panel1.Controls)
                            {
                                if (CTRL is Panel)
                                {
                                    CTRL.BackColor = System.Drawing.Color.LightYellow;
                                }
                            }
                           
                            ObjectId[] null1 = null;

                            Editor1.SetImpliedSelection(null1);
                            Trans1.Commit();

                        }



                    }

                }



                panel1.Controls.Clear();
                Layer_manager_Form_Load(sender, e);


                panel1.Width = this.Width - 20;
                panel1.Height = this.Height - 42;
                this.Refresh();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void but_NoPlot_on_off_click(object sender, EventArgs e)
        {
            Button Button_on_off = (Button)sender;

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        LayerTable1 = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

                        if (LayerTable1.Has("NO PLOT") == true)
                        {
                            LayerTableRecord Layer_selectat = (LayerTableRecord)Trans1.GetObject(LayerTable1["NO PLOT"], OpenMode.ForWrite);
                            if (Layer_selectat.IsOff == false)
                            {
                                Layer_selectat.IsOff = true;                           
                            }
                            else
                            {
                                Layer_selectat.IsOff = false;                         
                            }
                        }
                        Trans1.Commit();
                    }
                }

                panel1.Controls.Clear();
                Layer_manager_Form_Load(sender, e);
                panel1.Width = this.Width - 20;
                panel1.Height = this.Height - 42;
                this.Refresh();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e, string Layer_name)
        {
            CheckBox CheckBox1 = (CheckBox)sender;


            if (Lista_layere_selectate != null)
            {
                if (Layer_name.Contains("|") == false)
                {
                    try
                    {
                        int nr_val = Lista_layere_selectate.Length;

                        if (CheckBox1.Checked == true)
                        {
                            if (Lista_layere_selectate.Contains(Layer_name) == false)
                            {
                                Array.Resize(ref Lista_layere_selectate, nr_val + 1);
                                Lista_layere_selectate[nr_val] = Layer_name;
                            }
                        }

                        else
                        {
                            if (Lista_layere_selectate.Contains(Layer_name) == true)
                            {
                                Lista_layere_selectate = Lista_layere_selectate.Except(new string[] { Layer_name }).ToArray();
                            }
                        }


                        panel1.Width = this.Width - 20;
                        panel1.Height = this.Height - 42;
                        this.Refresh();
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
            else
            {
                if (Layer_name.Contains("|") == false)
                {
                    try
                    {
                        int nr_val = 0;

                        if (CheckBox1.Checked == true)
                        {
                                Array.Resize(ref Lista_layere_selectate, nr_val + 1);
                                Lista_layere_selectate[nr_val] = Layer_name;
                            
                        }

                        
                        panel1.Width = this.Width - 20;
                        panel1.Height = this.Height - 42;
                        this.Refresh();
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
        }

        private void but_DELETE_click(object sender, EventArgs e)
        {
           

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                DialogResult result = MessageBox.Show("Do you want to delete the selected layers?", "Delete", MessageBoxButtons.YesNo);


                if (result == DialogResult.Yes)
                {


                    if (Lista_layere_selectate.Length > 0)
                    {
                        using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                        {
                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                            {
                                Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                                Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                                LayerTable1 = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                                ThisDrawing.Database.Clayer = LayerTable1["0"];
                                for (int i = 0; i < Lista_layere_selectate.Length; ++i)
                                {
                                    if (LayerTable1.Has(Lista_layere_selectate[i]) == true)
                                        if (Lista_layere_selectate[i].Contains("|") == false )
                                        {

                                            {

                                                Autodesk.AutoCAD.DatabaseServices.TypedValue[] Filtru1 = new TypedValue[1];
                                                Filtru1[0] = new TypedValue((int)Autodesk.AutoCAD.DatabaseServices.DxfCode.LayerName, Lista_layere_selectate[i]);

                                                Autodesk.AutoCAD.EditorInput.SelectionFilter Selection_Filter1 = new Autodesk.AutoCAD.EditorInput.SelectionFilter(Filtru1);
                                                Autodesk.AutoCAD.EditorInput.PromptSelectionResult Selection_result1;

                                                Selection_result1 = Editor1.SelectAll(Selection_Filter1);
                                                Autodesk.AutoCAD.EditorInput.SelectionSet Selset1;

                                                Selset1 = Selection_result1.Value;
                                                if (Selset1 != null)
                                                {
                                                    if (Selset1.Count > 0)
                                                    {
                                                        for (int j = 0; j < Selset1.Count; ++j)
                                                        {
                                                            Autodesk.AutoCAD.EditorInput.SelectedObject Obj1 = Selset1[j];
                                                            Autodesk.AutoCAD.DatabaseServices.Entity Ent1 = (Entity)Trans1.GetObject(Obj1.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                                                            Ent1.Erase();
                                                        }

                                                    }
                                                }

                                                if (Lista_layere_selectate[i] != "0")
                                                {
                                                    LayerTableRecord Layer1 = (LayerTableRecord)LayerTable1[Lista_layere_selectate[i]].GetObject(OpenMode.ForWrite);
                                                    Layer1.Erase();
                                                }
                                            }
                                        }
                                }

                                Trans1.Commit();
                            }
                        }

                        panel1.Controls.Clear();
                        Layer_manager_Form_Load(sender, e);
                        panel1.Width = this.Width - 20;
                        panel1.Height = this.Height - 42;
                        this.Refresh();
                    }

                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void but_ALL_on_off_click(object sender, EventArgs e)
        {
            Button Button_on_off = (Button)sender;

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        LayerTable1 = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        string Clayer = ((LayerTableRecord)Trans1.GetObject(ThisDrawing.Database.Clayer, OpenMode.ForRead)).Name;
                       
                            if (Button_on_off.Text ==  "ALL OFF")
                            {                              
                                Button_on_off.Text = "ALL ON";
                                foreach (ObjectId id1 in LayerTable1)
                                {
                                    LayerTableRecord Layer1 = (LayerTableRecord)Trans1.GetObject(id1, OpenMode.ForWrite);
                                    Layer1.IsOff = true;
                                }
                                Button_on_off.BackColor = System.Drawing.Color.Red;                         
                            }

                            else
                            {
                                Button_on_off.Text = "ALL OFF";
                                foreach (ObjectId id1 in LayerTable1)
                                {
                                    LayerTableRecord Layer1 = (LayerTableRecord)Trans1.GetObject(id1, OpenMode.ForWrite);
                                    Layer1.IsOff = false;
                                }
                                Button_on_off.BackColor = System.Drawing.Color.LightGreen;                          
                            }
                        
                        Trans1.Commit();
                    }

                }

                panel1.Width = this.Width - 20;
                panel1.Height = this.Height - 42;
                this.Refresh();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void but_ALL_th_fr_click(object sender, EventArgs e)
        {
            Button Button_THAW_FREEZE = (Button)sender;

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        LayerTable1 = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        string Clayer = ((LayerTableRecord)Trans1.GetObject(ThisDrawing.Database.Clayer, OpenMode.ForRead)).Name;

                        if (Button_THAW_FREEZE.Text == "ALL FREEZE")
                        {
                            Button_THAW_FREEZE.Text = "ALL THAW";
                            foreach (ObjectId id1 in LayerTable1)
                            {
                                LayerTableRecord Layer1 = (LayerTableRecord)Trans1.GetObject(id1, OpenMode.ForWrite);
                                if (Layer1.Name != Clayer)
                                {
                                    Layer1.IsFrozen = true;
                                }
                            }
                            Button_THAW_FREEZE.BackColor = System.Drawing.Color.Red;
                        }

                        else
                        {
                            Button_THAW_FREEZE.Text = "ALL FREEZE";
                            foreach (ObjectId id1 in LayerTable1)
                            {
                                LayerTableRecord Layer1 = (LayerTableRecord)Trans1.GetObject(id1, OpenMode.ForWrite);
                                if (Layer1.Name != Clayer)
                                {
                                    Layer1.IsFrozen = false;
                                }
                            }
                            Button_THAW_FREEZE.BackColor = System.Drawing.Color.LightGreen;
                        }

                        Trans1.Commit();
                    }

                }

                panel1.Width = this.Width - 20;
                panel1.Height = this.Height - 42;
                this.Refresh();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void but_Refresh_click(object sender, EventArgs e)
        {
            panel1.Controls.Clear();
            Layer_manager_Form_Load(sender, e);
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
           Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
           Editor1.Regen();
        }

        public void Creaza_layer(string Layername, short Culoare, bool Plot)
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




    }
}
