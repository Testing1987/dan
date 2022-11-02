using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectSpecificPrograms.PennEast.NJDEP
{
    public class BlockAligner
    {
        #region FHA

        public static Point3d Last_Placed_FHA;
        public static Point3d FHA_Header_Position;
        public static Point3d NewUtil_Position;
        public static Point3d AccessTo_Position;
        public static Point3d ActiveDisturb_Position;
        public static Point3d ZeroFHA = new Point3d(0, 0, 0);

        public static string tag_FHA;
        public static string value_FHA;
        public static bool has_newutil;
        public static bool has_accessto;

        /// <summary>
        /// Searches the drawing for the block titled "FHA_HEADER" and then, if found, will run the methods to place the panel blocks with matching attributes on the "WATERBODY_NAME" tag under it.
        /// </summary>
        [CommandMethod("FHA_IMPACT_ALIGN")]
        public void AlignFHABlocksFHA()
        {
            Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor Editor1 = ThisDrawing.Editor;

            using (DocumentLock Lock1 = ThisDrawing.LockDocument())
            {
                using (Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    BlockTable BlockTable = Trans1.GetObject(ThisDrawing.Database.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTR = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                    foreach (ObjectId id1 in BTR)
                    {
                        has_newutil = false;
                        BlockReference FHA_HEADER = Trans1.GetObject(id1, OpenMode.ForRead) as BlockReference;
                        if (FHA_HEADER != null && FHA_HEADER.Name == BlockNameConstants.FHA_Header)
                        {
                            AttributeCollection AtrCol1 = FHA_HEADER.AttributeCollection;
                            foreach (ObjectId AtrId1 in AtrCol1)
                            {
                                AttributeReference AtrRef1 = Trans1.GetObject(AtrId1, OpenMode.ForRead) as AttributeReference;
                                if (AtrRef1 != null && AtrRef1.Tag == "WATERBODY_NAME")
                                {
                                    tag_FHA = AtrRef1.Tag;
                                    value_FHA = AtrRef1.TextString;
                                    FHA_Header_Position = FHA_HEADER.Position;
                                    Last_Placed_FHA = FHA_HEADER.Position;
                                }
                            }
                            Find_and_Place_NewUtil_Block();
                            Find_and_Place_AccessTo_Block();
                            Find_and_Place_ActiveDisturb_Block();
                            has_newutil = false;
                            has_accessto = false;
                        }
                    }
                    Trans1.Commit();
                }
            }
        }

        /// <summary>
        /// Searches the drawing for the New Utility Impact block and, if found, places it underneath the FHA_HEADER block with the matching attribute data.
        /// </summary>
        public static void Find_and_Place_NewUtil_Block()
        {
            Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            using (DocumentLock Lock1 = ThisDrawing.LockDocument())
            {
                using (Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    BlockTable BlockTable = Trans1.GetObject(ThisDrawing.Database.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTR = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                    foreach (ObjectId id2 in BTR)
                    {
                        BlockReference FHA_New_Utility = Trans1.GetObject(id2, OpenMode.ForWrite) as BlockReference;
                        if (FHA_New_Utility != null && FHA_New_Utility.Name == BlockNameConstants.FHA_NewUtility)
                        {
                            AttributeCollection AtrCol1 = FHA_New_Utility.AttributeCollection;
                            foreach (ObjectId AtrId1 in AtrCol1)
                            {
                                AttributeReference AtrRef1 = Trans1.GetObject(AtrId1, OpenMode.ForRead) as AttributeReference;
                                if (AtrRef1 != null && AtrRef1.TextString == value_FHA)
                                {
                                    has_newutil = true;
                                    FHA_New_Utility.Position = FHA_Header_Position;
                                    NewUtil_Position = FHA_New_Utility.Position;
                                    Last_Placed_FHA = FHA_New_Utility.Position;
                                }
                            }
                        }
                    }
                    Trans1.Commit();
                }
            }
        }

        /// <summary>
        /// Searches the drawing for the Access To the Project Impact Block with matching attribute values and, if found, places it under the New Utility block. If the New Utility block is not present, it places it underneath the header block.
        /// </summary>
        public static void Find_and_Place_AccessTo_Block()
        {
            Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            using (DocumentLock Lock1 = ThisDrawing.LockDocument())
            {
                using (Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    BlockTable BlockTable = Trans1.GetObject(ThisDrawing.Database.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTR = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                    foreach (ObjectId id3 in BTR)
                    {
                        BlockReference FHA_Access_To = Trans1.GetObject(id3, OpenMode.ForWrite) as BlockReference;
                        if (FHA_Access_To != null && FHA_Access_To.Name == BlockNameConstants.FHA_AccessTo)
                        {
                            AttributeCollection AtrCol1 = FHA_Access_To.AttributeCollection;
                            foreach (ObjectId AtrId1 in AtrCol1)
                            {
                                AttributeReference AtrRef1 = Trans1.GetObject(AtrId1, OpenMode.ForRead) as AttributeReference;
                                if (AtrRef1 != null && AtrRef1.TextString == value_FHA)
                                {
                                    has_accessto = true;
                                    if (has_newutil == true)
                                    {
                                        FHA_Access_To.Position = NewUtil_Position;
                                        Vector3d vector1 = new Vector3d(0, -0.5, 0);
                                        FHA_Access_To.TransformBy(Matrix3d.Displacement(vector1));
                                        AccessTo_Position = FHA_Access_To.Position;
                                        Last_Placed_FHA = FHA_Access_To.Position;
                                        break;
                                    }
                                    //FIND OUT WHY THIS ELSE STATEMENT IS BEING HIT. ANSWER: NEEDED TO ADD A BREAK AT LINE 133 TO STOP THE FOR-EACH LOOP FROM RESETTING THE BOOL.
                                    else
                                    {
                                        FHA_Access_To.Position = FHA_Header_Position;
                                        AccessTo_Position = FHA_Access_To.Position;
                                        Last_Placed_FHA = FHA_Access_To.Position;
                                    }
                                }
                            }
                        }
                    }
                    Trans1.Commit();
                }
            }
        }

        /// <summary>
        /// Searches the drawing for the Actively Disturbed area block and, if found, places it under the Impact Block, the New Utility block, or the FHA_HEADER block when present, in that order.
        /// </summary>
        public static void Find_and_Place_ActiveDisturb_Block()
        {
            Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            using (DocumentLock Lock1 = ThisDrawing.LockDocument())
            {
                using (Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    BlockTable BlockTable = Trans1.GetObject(ThisDrawing.Database.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTR = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                    foreach (ObjectId id4 in BTR)
                    {
                        BlockReference FHA_Active_Disturb = Trans1.GetObject(id4, OpenMode.ForWrite) as BlockReference;
                        if (FHA_Active_Disturb != null && FHA_Active_Disturb.Name == BlockNameConstants.FHA_ActivelyDisturbed)
                        {
                            AttributeCollection AtrCol1 = FHA_Active_Disturb.AttributeCollection;
                            foreach (ObjectId AtrId1 in AtrCol1)
                            {
                                AttributeReference AtrRef1 = Trans1.GetObject(AtrId1, OpenMode.ForRead) as AttributeReference;
                                if (AtrRef1 != null && AtrRef1.TextString == value_FHA)
                                {
                                    if (has_accessto == true)
                                    {
                                        FHA_Active_Disturb.Position = AccessTo_Position;
                                        Vector3d vector1 = new Vector3d(0, -0.5, 0);
                                        FHA_Active_Disturb.TransformBy(Matrix3d.Displacement(vector1));

                                        ActiveDisturb_Position = FHA_Active_Disturb.Position;
                                        has_newutil = false;
                                        has_accessto = false;
                                        break;
                                    }
                                    else
                                    {
                                        if (has_newutil == true)
                                        {
                                            FHA_Active_Disturb.Position = NewUtil_Position;
                                            Vector3d vector1 = new Vector3d(0, -0.5, 0);
                                            FHA_Active_Disturb.TransformBy(Matrix3d.Displacement(vector1));
                                            AccessTo_Position = FHA_Active_Disturb.Position;
                                            break;
                                        }
                                        else
                                        {
                                            FHA_Active_Disturb.Position = FHA_Header_Position;
                                            Last_Placed_FHA = FHA_Active_Disturb.Position;
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    Trans1.Commit();
                }
            }
        }

        #endregion

        #region FWW

        public static Point3d Last_Placed_FWW;
        public static Point3d FWW_Header_Position;
        public static Point3d PFO_Position;
        public static Point3d PEM_Position;
        public static Point3d PSS_Position;
        public static Point3d OWTR_Position;
        public static Point3d TransNW_Position;
        public static Point3d TransWD_Position;
        public static Point3d TransMROW_Position;
        public static Point3d ZeroFWW = new Point3d(0, 0, 0);
        public static string tag_FWW;
        public static string value_FWW;
        public static bool has_PFO;
        public static bool has_PEM;
        public static bool has_PSS;
        public static bool has_OWTR;
        public static bool has_TransNW;
        public static bool has_TransWD;
        public static bool has_TransMROW;


        [CommandMethod("FWW_IMPACT_ALIGN")]
        public void AlignBlocksFWW()
        {
            Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor Editor1 = ThisDrawing.Editor;

            using (DocumentLock Lock1 = ThisDrawing.LockDocument())
            {
                using (Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    BlockTable BlockTable = Trans1.GetObject(ThisDrawing.Database.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTR = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                    foreach (ObjectId id1 in BTR)
                    {
                        has_newutil = false;
                        BlockReference FWW_HEADER = Trans1.GetObject(id1, OpenMode.ForRead) as BlockReference;
                        if (FWW_HEADER != null && FWW_HEADER.Name == "FWW_HEADER")
                        {
                            AttributeCollection AtrCol1 = FWW_HEADER.AttributeCollection;
                            foreach (ObjectId AtrId1 in AtrCol1)
                            {
                                AttributeReference AtrRef1 = Trans1.GetObject(AtrId1, OpenMode.ForRead) as AttributeReference;
                                if (AtrRef1 != null && AtrRef1.Tag == "FWW_ID")
                                {
                                    tag_FWW = AtrRef1.Tag;
                                    value_FWW = AtrRef1.TextString;
                                    FWW_Header_Position = FWW_HEADER.Position;
                                    Last_Placed_FWW = FWW_HEADER.Position;
                                }
                            }
                            Find_and_Place_PFO();
                            Find_and_Place_PEM();
                            Find_and_Place_PSS();
                            Find_and_Place_Open_Water();
                            Find_and_Place_TransNW();
                            Find_and_Place_TransWD();
                            Find_and_Place_TransMROW();
                            has_PFO = false;
                            has_PEM = false;
                            has_PSS = false;
                            has_OWTR = false;
                            has_TransNW = false;
                            has_TransWD = false;
                            has_TransMROW = false;
                        }
                    }
                    Trans1.Commit();
                }
            }
        }

        public static void Find_and_Place_PFO()
        {
            Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            using (DocumentLock Lock1 = ThisDrawing.LockDocument())
            {
                using (Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    BlockTable BlockTable = Trans1.GetObject(ThisDrawing.Database.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTR = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                    foreach (ObjectId id2 in BTR)
                    {
                        BlockReference FWW_PFO_Block = Trans1.GetObject(id2, OpenMode.ForWrite) as BlockReference;
                        if (FWW_PFO_Block != null && FWW_PFO_Block.Name == "PFO_IMPACT_BLOCK")
                        {

                            AttributeCollection AtrCol1 = FWW_PFO_Block.AttributeCollection;
                            foreach (ObjectId AtrId1 in AtrCol1)
                            {
                                AttributeReference AtrRef1 = Trans1.GetObject(AtrId1, OpenMode.ForRead) as AttributeReference;
                                if (AtrRef1 != null && AtrRef1.TextString == value_FWW)
                                {
                                    has_PFO = true;
                                    FWW_PFO_Block.Position = FWW_Header_Position;
                                    PFO_Position = FWW_PFO_Block.Position;
                                    Last_Placed_FWW = FWW_PFO_Block.Position;
                                }
                            }
                        }
                    }
                    Trans1.Commit();
                }
            }
        }
        public static void Find_and_Place_PEM()
        {
            Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            using (DocumentLock Lock1 = ThisDrawing.LockDocument())
            {
                using (Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    BlockTable BlockTable = Trans1.GetObject(ThisDrawing.Database.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTR = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                    foreach (ObjectId id3 in BTR)
                    {
                        BlockReference FWW_PEM_Impact = Trans1.GetObject(id3, OpenMode.ForWrite) as BlockReference;
                        if (FWW_PEM_Impact != null && FWW_PEM_Impact.Name == "PEM_IMPACT_BLOCK")
                        {
                            AttributeCollection AtrCol1 = FWW_PEM_Impact.AttributeCollection;
                            foreach (ObjectId AtrId1 in AtrCol1)
                            {
                                AttributeReference AtrRef1 = Trans1.GetObject(AtrId1, OpenMode.ForRead) as AttributeReference;
                                if (AtrRef1 != null && AtrRef1.TextString == value_FWW)
                                {
                                    has_PEM = true;
                                    if (has_PFO == true)
                                    {
                                        FWW_PEM_Impact.Position = PFO_Position;
                                        Vector3d vector1 = new Vector3d(0, -0.23, 0);
                                        FWW_PEM_Impact.TransformBy(Matrix3d.Displacement(vector1));
                                        PEM_Position = FWW_PEM_Impact.Position;
                                        Last_Placed_FWW = FWW_PEM_Impact.Position;
                                        break;
                                    }
                                    else
                                    {
                                        FWW_PEM_Impact.Position = FWW_Header_Position;
                                        PEM_Position = FWW_PEM_Impact.Position;
                                        Last_Placed_FWW = FWW_PEM_Impact.Position;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    Trans1.Commit();
                }
            }
        }
        public static void Find_and_Place_PSS()
        {
            Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            using (DocumentLock Lock1 = ThisDrawing.LockDocument())
            {
                using (Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    BlockTable BlockTable = Trans1.GetObject(ThisDrawing.Database.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTR = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                    foreach (ObjectId id3 in BTR)
                    {
                        BlockReference FWW_PSS_Impact = Trans1.GetObject(id3, OpenMode.ForWrite) as BlockReference;
                        if (FWW_PSS_Impact != null && FWW_PSS_Impact.Name == "PSS_IMPACT_BLOCK")
                        {
                            AttributeCollection AtrCol1 = FWW_PSS_Impact.AttributeCollection;
                            foreach (ObjectId AtrId1 in AtrCol1)
                            {
                                AttributeReference AtrRef1 = Trans1.GetObject(AtrId1, OpenMode.ForRead) as AttributeReference;
                                if (AtrRef1 != null && AtrRef1.TextString == value_FWW)
                                {
                                    has_PSS = true;
                                    if (has_PEM == true)
                                    {
                                        FWW_PSS_Impact.Position = PEM_Position;
                                        Vector3d vector1 = new Vector3d(0, -0.23, 0);
                                        FWW_PSS_Impact.TransformBy(Matrix3d.Displacement(vector1));
                                        PSS_Position = FWW_PSS_Impact.Position;
                                        Last_Placed_FWW = FWW_PSS_Impact.Position;
                                        break;
                                    }
                                    else
                                    {
                                        if (has_PFO == true)
                                        {
                                            FWW_PSS_Impact.Position = PFO_Position;
                                            Vector3d vector1 = new Vector3d(0, -0.23, 0);
                                            FWW_PSS_Impact.TransformBy(Matrix3d.Displacement(vector1));
                                            PSS_Position = FWW_PSS_Impact.Position;
                                            Last_Placed_FWW = FWW_PSS_Impact.Position;
                                            break;
                                        }
                                        else
                                        {
                                            FWW_PSS_Impact.Position = FWW_Header_Position;
                                            PSS_Position = FWW_PSS_Impact.Position;
                                            Last_Placed_FWW = FWW_PSS_Impact.Position;
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    Trans1.Commit();
                }
            }
        }
        public static void Find_and_Place_Open_Water()
        {
            Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            using (DocumentLock Lock1 = ThisDrawing.LockDocument())
            {
                using (Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    BlockTable BlockTable = Trans1.GetObject(ThisDrawing.Database.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTR = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                    foreach (ObjectId id3 in BTR)
                    {
                        BlockReference FWW_OWTR_Impact = Trans1.GetObject(id3, OpenMode.ForWrite) as BlockReference;
                        if (FWW_OWTR_Impact != null && FWW_OWTR_Impact.Name == "OPEN_WATER_IMPACT_BLOCK")
                        {
                            AttributeCollection AtrCol1 = FWW_OWTR_Impact.AttributeCollection;
                            foreach (ObjectId AtrId1 in AtrCol1)
                            {
                                AttributeReference AtrRef1 = Trans1.GetObject(AtrId1, OpenMode.ForRead) as AttributeReference;
                                if (AtrRef1 != null && AtrRef1.TextString == value_FWW)
                                {
                                    has_OWTR = true;
                                    if (has_PSS == true)
                                    {
                                        FWW_OWTR_Impact.Position = PSS_Position;
                                        Vector3d vector1 = new Vector3d(0, -0.23, 0);
                                        FWW_OWTR_Impact.TransformBy(Matrix3d.Displacement(vector1));
                                        OWTR_Position = FWW_OWTR_Impact.Position;
                                        Last_Placed_FWW = FWW_OWTR_Impact.Position;
                                        break;
                                    }
                                    else
                                    {
                                        if (has_PEM == true)
                                        {
                                            FWW_OWTR_Impact.Position = PEM_Position;
                                            Vector3d vector1 = new Vector3d(0, -0.23, 0);
                                            FWW_OWTR_Impact.TransformBy(Matrix3d.Displacement(vector1));
                                            OWTR_Position = FWW_OWTR_Impact.Position;
                                            Last_Placed_FWW = FWW_OWTR_Impact.Position;
                                            break;
                                        }
                                        else
                                        {
                                            if (has_PFO == true)
                                            {
                                                FWW_OWTR_Impact.Position = PFO_Position;
                                                Vector3d vector1 = new Vector3d(0, -0.23, 0);
                                                FWW_OWTR_Impact.TransformBy(Matrix3d.Displacement(vector1));
                                                OWTR_Position = FWW_OWTR_Impact.Position;
                                                Last_Placed_FWW = FWW_OWTR_Impact.Position;
                                                break;
                                            }
                                            else
                                            {
                                                FWW_OWTR_Impact.Position = FWW_Header_Position;
                                                OWTR_Position = FWW_OWTR_Impact.Position;
                                                Last_Placed_FWW = FWW_OWTR_Impact.Position;
                                                break;
                                            }
                                        }
                                    }

                                }
                            }
                        }
                    }
                    Trans1.Commit();
                }
            }
        }
        public static void Find_and_Place_TransNW()
        {
            Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            using (DocumentLock Lock1 = ThisDrawing.LockDocument())
            {
                using (Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    BlockTable BlockTable = Trans1.GetObject(ThisDrawing.Database.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTR = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                    foreach (ObjectId id3 in BTR)
                    {
                        BlockReference FWW_TransNW = Trans1.GetObject(id3, OpenMode.ForWrite) as BlockReference;
                        if (FWW_TransNW != null && FWW_TransNW.Name == "TRANS_AREA_NONWOODED_TEMP_IMPACT_BLOCK")
                        {
                            AttributeCollection AtrCol1 = FWW_TransNW.AttributeCollection;
                            foreach (ObjectId AtrId1 in AtrCol1)
                            {
                                AttributeReference AtrRef1 = Trans1.GetObject(AtrId1, OpenMode.ForRead) as AttributeReference;
                                if (AtrRef1 != null && AtrRef1.TextString == value_FWW)
                                {
                                    has_TransNW = true;
                                    if (has_OWTR == true)
                                    {
                                        FWW_TransNW.Position = OWTR_Position;
                                        Vector3d vector1 = new Vector3d(0, -0.23, 0);
                                        FWW_TransNW.TransformBy(Matrix3d.Displacement(vector1));
                                        TransNW_Position = FWW_TransNW.Position;
                                        Last_Placed_FWW = FWW_TransNW.Position;
                                        break;
                                    }
                                    else
                                    {
                                        if (has_PSS == true)
                                        {
                                            FWW_TransNW.Position = PSS_Position;
                                            Vector3d vector1 = new Vector3d(0, -0.23, 0);
                                            FWW_TransNW.TransformBy(Matrix3d.Displacement(vector1));
                                            TransNW_Position = FWW_TransNW.Position;
                                            Last_Placed_FWW = FWW_TransNW.Position;
                                            break;
                                        }
                                        else
                                        {
                                            if (has_PEM == true)
                                            {
                                                FWW_TransNW.Position = PEM_Position;
                                                Vector3d vector1 = new Vector3d(0, -0.23, 0);
                                                FWW_TransNW.TransformBy(Matrix3d.Displacement(vector1));
                                                TransNW_Position = FWW_TransNW.Position;
                                                Last_Placed_FWW = FWW_TransNW.Position;
                                                break;
                                            }
                                            else
                                            {
                                                if (has_PFO == true)
                                                {
                                                    FWW_TransNW.Position = PFO_Position;
                                                    Vector3d vector1 = new Vector3d(0, -0.23, 0);
                                                    FWW_TransNW.TransformBy(Matrix3d.Displacement(vector1));
                                                    TransNW_Position = FWW_TransNW.Position;
                                                    Last_Placed_FWW = FWW_TransNW.Position;
                                                    break;
                                                }
                                                else
                                                {
                                                    FWW_TransNW.Position = FWW_Header_Position;
                                                    TransNW_Position = FWW_TransNW.Position;
                                                    Last_Placed_FWW = FWW_TransNW.Position;
                                                    break;
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
            }
        }
        public static void Find_and_Place_TransWD()
        {
            Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            using (DocumentLock Lock1 = ThisDrawing.LockDocument())
            {
                using (Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    BlockTable BlockTable = Trans1.GetObject(ThisDrawing.Database.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTR = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                    foreach (ObjectId id3 in BTR)
                    {
                        BlockReference FWW_TransWD = Trans1.GetObject(id3, OpenMode.ForWrite) as BlockReference;
                        if (FWW_TransWD != null && FWW_TransWD.Name == "TRANS_AREA_WOODED_TEMP_IMPACT_BLOCK")
                        {
                            AttributeCollection AtrCol1 = FWW_TransWD.AttributeCollection;
                            foreach (ObjectId AtrId1 in AtrCol1)
                            {
                                AttributeReference AtrRef1 = Trans1.GetObject(AtrId1, OpenMode.ForRead) as AttributeReference;
                                if (AtrRef1 != null && AtrRef1.TextString == value_FWW)
                                {
                                    has_TransWD = true;

                                    if (has_TransNW == true)
                                    {
                                        FWW_TransWD.Position = TransNW_Position;
                                        Vector3d vector1 = new Vector3d(0, -0.23, 0);
                                        FWW_TransWD.TransformBy(Matrix3d.Displacement(vector1));
                                        TransWD_Position = FWW_TransWD.Position;
                                        Last_Placed_FWW = FWW_TransWD.Position;
                                        break;
                                    }
                                    else
                                    {
                                        if (has_OWTR == true)
                                        {
                                            FWW_TransWD.Position = OWTR_Position;
                                            Vector3d vector1 = new Vector3d(0, -0.23, 0);
                                            FWW_TransWD.TransformBy(Matrix3d.Displacement(vector1));
                                            TransWD_Position = FWW_TransWD.Position;
                                            Last_Placed_FWW = FWW_TransWD.Position;
                                            break;
                                        }
                                        else
                                        {
                                            if (has_PSS == true)
                                            {
                                                FWW_TransWD.Position = PSS_Position;
                                                Vector3d vector1 = new Vector3d(0, -0.23, 0);
                                                FWW_TransWD.TransformBy(Matrix3d.Displacement(vector1));
                                                TransWD_Position = FWW_TransWD.Position;
                                                Last_Placed_FWW = FWW_TransWD.Position;
                                                break;
                                            }
                                            else
                                            {
                                                if (has_PEM == true)
                                                {
                                                    FWW_TransWD.Position = PEM_Position;
                                                    Vector3d vector1 = new Vector3d(0, -0.23, 0);
                                                    FWW_TransWD.TransformBy(Matrix3d.Displacement(vector1));
                                                    TransWD_Position = FWW_TransWD.Position;
                                                    Last_Placed_FWW = FWW_TransWD.Position;
                                                    break;
                                                }
                                                else
                                                {
                                                    if (has_PFO == true)
                                                    {
                                                        FWW_TransWD.Position = PFO_Position;
                                                        Vector3d vector1 = new Vector3d(0, -0.23, 0);
                                                        FWW_TransWD.TransformBy(Matrix3d.Displacement(vector1));
                                                        TransWD_Position = FWW_TransWD.Position;
                                                        Last_Placed_FWW = FWW_TransWD.Position;
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        FWW_TransWD.Position = FWW_Header_Position;
                                                        TransWD_Position = FWW_TransWD.Position;
                                                        Last_Placed_FWW = FWW_TransWD.Position;
                                                        break;
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
            }
        }
        public static void Find_and_Place_TransMROW()
        {
            Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            using (DocumentLock Lock1 = ThisDrawing.LockDocument())
            {
                using (Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    BlockTable BlockTable = Trans1.GetObject(ThisDrawing.Database.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTR = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                    foreach (ObjectId id3 in BTR)
                    {
                        BlockReference FWW_TransMROW = Trans1.GetObject(id3, OpenMode.ForWrite) as BlockReference;
                        if (FWW_TransMROW != null && FWW_TransMROW.Name == "TRANS_INSIDE_MROW_IMPACT_BLOCK")
                        {
                            AttributeCollection AtrCol1 = FWW_TransMROW.AttributeCollection;
                            foreach (ObjectId AtrId1 in AtrCol1)
                            {
                                AttributeReference AtrRef1 = Trans1.GetObject(AtrId1, OpenMode.ForRead) as AttributeReference;
                                if (AtrRef1 != null && AtrRef1.TextString == value_FWW)
                                {
                                    has_TransMROW = true;
                                    if (has_TransWD == true)
                                    {
                                        FWW_TransMROW.Position = TransWD_Position;
                                        Vector3d vector1 = new Vector3d(0, -0.23, 0);
                                        FWW_TransMROW.TransformBy(Matrix3d.Displacement(vector1));
                                        TransMROW_Position = FWW_TransMROW.Position;
                                        Last_Placed_FWW = FWW_TransMROW.Position;
                                        break;
                                    }

                                    else
                                    {
                                        if (has_TransNW == true)
                                        {
                                            FWW_TransMROW.Position = TransNW_Position;
                                            Vector3d vector1 = new Vector3d(0, -0.23, 0);
                                            FWW_TransMROW.TransformBy(Matrix3d.Displacement(vector1));
                                            TransMROW_Position = FWW_TransMROW.Position;
                                            Last_Placed_FWW = FWW_TransMROW.Position;
                                            break;
                                        }
                                        else
                                        {
                                            if (has_OWTR == true)
                                            {
                                                FWW_TransMROW.Position = OWTR_Position;
                                                Vector3d vector1 = new Vector3d(0, -0.23, 0);
                                                FWW_TransMROW.TransformBy(Matrix3d.Displacement(vector1));
                                                TransMROW_Position = FWW_TransMROW.Position;
                                                Last_Placed_FWW = FWW_TransMROW.Position;
                                                break;
                                            }
                                            else
                                            {
                                                if (has_PSS == true)
                                                {
                                                    FWW_TransMROW.Position = PSS_Position;
                                                    Vector3d vector1 = new Vector3d(0, -0.23, 0);
                                                    FWW_TransMROW.TransformBy(Matrix3d.Displacement(vector1));
                                                    TransMROW_Position = FWW_TransMROW.Position;
                                                    Last_Placed_FWW = FWW_TransMROW.Position;
                                                    break;
                                                }
                                                else
                                                {
                                                    if (has_PEM == true)
                                                    {
                                                        FWW_TransMROW.Position = PEM_Position;
                                                        Vector3d vector1 = new Vector3d(0, -0.23, 0);
                                                        FWW_TransMROW.TransformBy(Matrix3d.Displacement(vector1));
                                                        TransMROW_Position = FWW_TransMROW.Position;
                                                        Last_Placed_FWW = FWW_TransMROW.Position;
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        if (has_PFO == true)
                                                        {
                                                            FWW_TransMROW.Position = PFO_Position;
                                                            Vector3d vector1 = new Vector3d(0, -0.23, 0);
                                                            FWW_TransMROW.TransformBy(Matrix3d.Displacement(vector1));
                                                            TransMROW_Position = FWW_TransMROW.Position;
                                                            Last_Placed_FWW = FWW_TransMROW.Position;
                                                            break;
                                                        }
                                                        else
                                                        {
                                                            FWW_TransMROW.Position = FWW_Header_Position;
                                                            TransMROW_Position = FWW_TransMROW.Position;
                                                            Last_Placed_FWW = FWW_TransMROW.Position;
                                                            break;
                                                        }
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
            }
        }

        #endregion
    }
}
