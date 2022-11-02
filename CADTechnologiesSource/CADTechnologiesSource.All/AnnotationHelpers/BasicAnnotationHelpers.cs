using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;

namespace CADTechnologiesSource.All.AnnotationHelpers
{
    public class BasicAnnotationHelpers
    {
        /// <summary>
        /// Creates a standalone MTEXT object and allows you to set various properties to it.
        /// </summary>
        /// <param name="location">Where to insert the MText.</param>
        /// <param name="contents">String contents of the MText.</param>
        /// <param name="height">Text Height of the MText.</param>
        /// <param name="color">Color applied to the MText.</param>
        /// <param name="usebackgroundmask">Determines if a background mask will be used. Must be true or false.</param>
        /// <param name="usebackgroundcolor">Determines if the background mask color will use the background color of the drawing. Must be true or false.</param>
        /// <param name="backgroundscale">Sets the scale of the background mask. 1.0 will be the same size as the text.</param>
        public static void createMtext(Point3d location, AttachmentPoint attachmentpoint, string contents, double height, short color, bool usebackgroundmask, bool usebackgroundcolor, double backgroundscale)
        {
            TransactionHelpers.TransactionMethods.ModifyDrawingWithinTransaction((transaction, database, blocktable, blocktablerecord) =>
            {
                MText mt = new MText();
                mt.SetDatabaseDefaults();
                mt.Location = location;
                mt.Attachment = attachmentpoint;
                mt.Contents = contents;
                mt.Height = height;
                mt.Color = Color.FromColorIndex(ColorMethod.ByAci, color);
                mt.BackgroundFill = usebackgroundmask;
                mt.UseBackgroundColor = usebackgroundcolor;
                mt.BackgroundScaleFactor = backgroundscale;
                blocktablerecord.AppendEntity(mt);
                transaction.AddNewlyCreatedDBObject(mt, true);
            });
        }

        /// <summary>
        /// Creates a basic Multileader with MText content and a single leader line. Allows you to set various properties during creation.
        /// </summary>
        /// <param name="createleader">Adds a new leader cluster to this MLeader object. A leader cluster is made up by a dog-leg and some leader lines.</param>
        /// <param name="createleaderline">Adds a leaderline to the leader cluster with specified index.</param>
        /// <param name="firstvertex">Inserts a vertex to the leader line with specified index as new leader head.</param>
        /// <param name="lastvertex">Appends a vertex to the specified leader line as the new leader tail.</param>
        /// <param name="DeltaX">Distance between the first and second points on the mleader along the X axis.</param>
        /// <param name="DeltaY">Distance between the first and second points on the mleader along the Y axis.</param>
        /// <param name="insertionpoint">The point the arrow is pointing at.</param>
        /// <param name="leadertype">The style of the leader line (straight line, spline, etc).</param>
        /// <param name="contenttype">The content hosted by the Mleader (MText, for example).</param>
        /// <param name="mtext">the MText being inserted for the MLeader to host.</param>
        /// <param name="textheight">The height of the MText being hosted by the MLeader.</param>
        /// <param name="arrowheadsize">The size of the arrowhead on the MLeader.</param>
        /// <param name="landinggap">The distance between the end of the dogleg and the beginning of the hosted content.</param>
        /// <param name="dogleglength">The length of the straight line that leads towards the hosted content.</param>
        /// <param name="annotativestate">Determines if the MLeader should be annotative or not.</param>
        /// <param name="layer">The layer that the MLeader will be created on.</param>
        /// <param name="transformrotation">Rotates the MLeader by this amount.</param>
        public static void createMleader(int createleader, int createleaderline, int firstvertex, int lastvertex, int DeltaX, int DeltaY, Point3d insertionpoint, LeaderType leadertype, ContentType contenttype, MText mtext, TextAttachmentType textattachmenttype, LeaderDirectionType leaderdirectiontype, double textheight, double arrowheadsize, double landinggap, double dogleglength, AnnotativeStates annotativestate, double transformrotation)
        {
            TransactionHelpers.TransactionMethods.ModifyDrawingWithinTransaction((transaction, database, blocktable, blocktablerecord) =>
            {
                MLeader ml = new MLeader();
                ml.SetDatabaseDefaults();
                createleader = ml.AddLeader();
                createleaderline = ml.AddLeaderLine(createleader);
                ml.AddFirstVertex(createleaderline, new Point3d(insertionpoint.X, insertionpoint.Y, 0));
                ml.AddLastVertex(createleaderline, new Point3d(insertionpoint.X + DeltaX, insertionpoint.Y + DeltaY, 0));
                ml.LeaderLineType = leadertype;
                ml.ContentType = contenttype;
                ml.MText = mtext;
                ml.TextHeight = textheight;
                ml.ArrowSize = arrowheadsize;
                ml.DoglegLength = dogleglength;
                ml.LandingGap = landinggap;
                ml.Annotative = annotativestate;
                //ml.Layer = layer;
                ml.TransformBy(Matrix3d.Rotation(transformrotation, Vector3d.ZAxis, insertionpoint));
                blocktablerecord.AppendEntity(ml);
                transaction.AddNewlyCreatedDBObject(ml, true);
                ml.MoveMLeader(new Vector3d(), MoveType.MoveAllPoints);
                ml.SetTextAttachmentType(textattachmenttype, leaderdirectiontype);
                ml.SetTextAttachmentType(textattachmenttype, leaderdirectiontype);
            });
        }

        /// <summary>
        /// Creates a reference MLeader with three points and placeholder MText Contents. Does not accept any parameters. Use for reference only.
        /// </summary>
        public static void CreateReferenceThreePointVerticalMLeader()
        {
            var mLeader = new MLeader();
            var leaderline_index = new int();
            var leaderpoint = new int();
            Point3d insertPoint = new Point3d(0, 0, 0);

            Document thisDocument = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            Database thisDatabase = thisDocument.Database;
            Editor ThisEditor = thisDocument.Editor;

            using (DocumentLock docLock = thisDocument.LockDocument())
            {
                using (Transaction trans1 = thisDatabase.TransactionManager.StartTransaction())
                {

                    BlockTable blockTable;
                    blockTable = trans1.GetObject(thisDatabase.BlockTableId, OpenMode.ForRead) as BlockTable;

                    BlockTableRecord blockTableRecord = trans1.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    MText mText = new MText();
                    mText.Contents = "Placeholder Text";

                    mText.Attachment = AttachmentPoint.MiddleLeft;
                    mText.BackgroundFill = true;
                    mText.UseBackgroundColor = true;
                    mText.BackgroundScaleFactor = 1.2;
                    mLeader.SetDatabaseDefaults();
                    leaderline_index = mLeader.AddLeader();
                    leaderpoint = mLeader.AddLeaderLine(leaderline_index);
                    mLeader.AddFirstVertex(leaderpoint, insertPoint);
                    mLeader.AddLastVertex(leaderpoint, new Point3d(4, 0, 0));

                    // calling AddLastVertex again will add more vertices to the leader line. The last vertex you add is where the text will appear.
                    mLeader.AddLastVertex(leaderline_index, new Point3d(8, 4, 0));

                    mLeader.LeaderLineType = LeaderType.StraightLeader;
                    mLeader.ContentType = ContentType.MTextContent;
                    mLeader.MText = mText;
                    mLeader.TextHeight = 4;
                    mLeader.LandingGap = 3;
                    mLeader.ArrowSize = 0;
                    mLeader.DoglegLength = 2;

                    //Append the mLeader to the block table and add it to the database.
                    blockTableRecord.AppendEntity(mLeader);
                    trans1.AddNewlyCreatedDBObject(mLeader, true);

                    // rotate leader to make it vertical.
                    mLeader.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, insertPoint));

                    //MoveMLeader will reset the MLeader and fix a bug with the text showing up in a wrong location
                    mLeader.MoveMLeader(new Vector3d(), MoveType.MoveAllPoints);

                    //you must call set text attachment for right and left leader after the entitiy is appended before you commit.
                    mLeader.SetTextAttachmentType(TextAttachmentType.AttachmentBottomLine, LeaderDirectionType.RightLeader);
                    mLeader.SetTextAttachmentType(TextAttachmentType.AttachmentBottomLine, LeaderDirectionType.LeftLeader);
                    trans1.Commit();
                }
            }
        }
    }
}
