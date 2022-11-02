using System;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System.Runtime.InteropServices;

namespace Alignment_mdi
{
    public class zLAUNCH
    {
        [CommandMethod("sgen")]
        public void Show_Sgen_mainForm()
        {
            if (Functions.isSECURE() == true)
            {
                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is Alignment_mdi._SGEN_mainform)
                    {
                        Forma1.Focus();
                        Forma1.WindowState = System.Windows.Forms.FormWindowState.Normal;
                        Forma1.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - Forma1.Width) / 2, (Screen.PrimaryScreen.WorkingArea.Height - Forma1.Height) / 2);
                        return;
                    }
                }
                try
                {
                    Alignment_mdi._SGEN_mainform forma2 = new Alignment_mdi._SGEN_mainform();
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma2);
                    forma2.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - forma2.Width) / 2,
                         (Screen.PrimaryScreen.WorkingArea.Height - forma2.Height) / 2);
                }
                catch (System.Exception EX)
                {
                    MessageBox.Show(EX.Message);
                }
            }
            else
            {
                return;
            }
        }


        [CommandMethod("DynablockVisibilityStates")]
        public void DynablockVisibilityStates()
        {
            Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            PromptResult pr = ThisDrawing.Editor.GetString("\nEnter block name: ");

            if (pr.Status != PromptStatus.OK) return;
            using (Transaction Trans1 = ThisDrawing.Database.TransactionManager.StartTransaction())
            {
                BlockTable BlockTable1 = Trans1.GetObject(ThisDrawing.Database.BlockTableId, OpenMode.ForRead) as BlockTable;
                if (!BlockTable1.Has(pr.StringResult))
                {
                    ThisDrawing.Editor.WriteMessage("\nBlock doesn't exist :(");
                    return;
                }
                BlockTableRecord BTrecord = Trans1.GetObject(BlockTable1[pr.StringResult], OpenMode.ForRead) as BlockTableRecord;
                if (!BTrecord.IsDynamicBlock)
                {
                    ThisDrawing.Editor.WriteMessage("\nNot a dynamic block :(");
                    //return;
                }

                if (BTrecord.ExtensionDictionary == null)
                {
                    ThisDrawing.Editor.WriteMessage("\nNo ExtensionDictionary :(");
                    //return;
                }

                DBDictionary dico = Trans1.GetObject(BTrecord.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;

                if (!dico.Contains("ACAD_ENHANCEDBLOCK"))
                {
                    ThisDrawing.Editor.WriteMessage("\nACAD_ENHANCEDBLOCK Entry not found :(");
                    //return;
                }

                ObjectId graphId = dico.GetAt("ACAD_ENHANCEDBLOCK");

                System.Collections.Generic.List<object> parameterIds = acdbEntGetObjects(graphId, 360);

                foreach (object parameterId in parameterIds)
                {
                    ObjectId id = (ObjectId)parameterId;

                    if (id.ObjectClass.Name == "AcDbBlockVisibilityParameter")
                    {
                        System.Collections.Generic.List<TypedValue> visibilityParam = acdbEntGetTypedVals(id);
                        System.Collections.Generic.List<TypedValue>.Enumerator enumerator = visibilityParam.GetEnumerator();

                        while (enumerator.MoveNext())
                        {
                            if (enumerator.Current.TypeCode == 303)
                            {
                                string group = (string)enumerator.Current.Value;
                                enumerator.MoveNext();
                                int nbEntitiesInGroup = (int)enumerator.Current.Value;
                                ThisDrawing.Editor.WriteMessage("\n . Visibility Group: " + group + " Nb Entities in group: " + nbEntitiesInGroup);

                                for (int i = 0; i < nbEntitiesInGroup; ++i)
                                {
                                    enumerator.MoveNext();
                                    ObjectId entityId = (ObjectId)enumerator.Current.Value;
                                    Entity entity = Trans1.GetObject(entityId, OpenMode.ForRead) as Entity;
                                    ThisDrawing.Editor.WriteMessage("\n    - " + entity.ToString() + " " + entityId.ToString());
                                }
                            }
                        }
                        break;
                    }
                }
                Trans1.Commit();
            }
        }



        public struct ads_name
        {
            IntPtr a;
            IntPtr b;
        };



        [DllImport("acdb18.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "?acdbGetAdsName@@YA?AW4ErrorStatus@Acad@@AAY01JVAcDbObjectId@@@Z")]

        public static extern int acdbGetAdsName(ref ads_name name, ObjectId objId);

        [DllImport("acad.exe", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl, EntryPoint = "acdbEntGet")]
        public static extern System.IntPtr acdbEntGet(ref ads_name ename);



        private System.Collections.Generic.List<object> acdbEntGetObjects(ObjectId id, short dxfcode)
        {
            System.Collections.Generic.List<object> result = new System.Collections.Generic.List<object>();
            ads_name name = new ads_name();
            int res = acdbGetAdsName(ref name, id);
            ResultBuffer rb = new ResultBuffer();
            Autodesk.AutoCAD.Runtime.Interop.AttachUnmanagedObject(rb, acdbEntGet(ref name), true);

            ResultBufferEnumerator iter = rb.GetEnumerator();
            while (iter.MoveNext())
            {
                TypedValue typedValue = (TypedValue)iter.Current;
                if (typedValue.TypeCode == dxfcode)
                {
                    result.Add(typedValue.Value);
                }
            }
            return result;
        }



        private System.Collections.Generic.List<TypedValue> acdbEntGetTypedVals(ObjectId id)
        {

            System.Collections.Generic.List<TypedValue> result = new System.Collections.Generic.List<TypedValue>();
            ads_name name = new ads_name();
            int res = acdbGetAdsName(ref name, id);
            ResultBuffer rb = new ResultBuffer();
            Autodesk.AutoCAD.Runtime.Interop.AttachUnmanagedObject(rb, acdbEntGet(ref name), true);
            ResultBufferEnumerator iter = rb.GetEnumerator();
            while (iter.MoveNext())
            {
                result.Add((TypedValue)iter.Current);
            }
            return result;

        }

    }
}
