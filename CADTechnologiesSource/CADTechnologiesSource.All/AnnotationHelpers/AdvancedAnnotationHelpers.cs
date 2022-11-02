using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;

namespace CADTechnologiesSource.All.AnnotationHelpers
{
    public class AdvancedAnnotationHelpers
    {
        /// <summary>
        /// Returns an Mtext entity that can be passed into another method, for example. Can be used to create an MText to be passed into a multileader.
        /// </summary>
        /// <param name="location">Where to insert the MText.</param>
        /// <param name="attachmentpoint"></param>
        /// <param name="contents">Text contents of the MText.</param>
        /// <param name="height">Text Height of the MText.</param>
        /// <param name="color">Color applied to the MText.</param>
        /// <param name="usebackgroundmask">Determines if a background mask will be used. Must be true or false.</param>
        /// <param name="usebackgroundcolor">Determines if the background mask color will use the background color of the drawing. Must be true or false.</param>
        /// <param name="backgroundscale">Sets the scale of the background mask. 1.0 will be the same size as the text.</param>
        /// <returns></returns>
        public static MText ReturnMText(AttachmentPoint attachmentpoint, string contents, double height, short color, bool usebackgroundmask, bool usebackgroundcolor, double backgroundscale)
        {
            return AnnotationDelegates.CreateMTextDelegate((transaction, database, blocktable, blocktablerecord) =>
            {
                MText mt = new MText();
                mt.SetDatabaseDefaults();
                //mt.Location = location;
                mt.Attachment = attachmentpoint;
                mt.Contents = contents;
                mt.Height = height;
                mt.Color = Color.FromColorIndex(ColorMethod.ByAci, color);
                mt.BackgroundFill = usebackgroundmask;
                mt.UseBackgroundColor = usebackgroundcolor;
                mt.BackgroundScaleFactor = backgroundscale;
                blocktablerecord.AppendEntity(mt);
                transaction.AddNewlyCreatedDBObject(mt, true);
                //make sure to get ObjectId only AFTER adding to db.
                return mt;
            });
        }
    }
}
