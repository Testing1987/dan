using Autodesk.AutoCAD.ApplicationServices;

using Autodesk.AutoCAD.DatabaseServices;

using Autodesk.AutoCAD.EditorInput;

using Autodesk.AutoCAD.Geometry;

using Autodesk.AutoCAD.Runtime;

using System;



public class Commands

{

    [CommandMethod("QT")]

    static public void QuickText()

    {

        Document doc =

          Application.DocumentManager.MdiActiveDocument;

        Database db = doc.Database;

        Editor ed = doc.Editor;



        PromptStringOptions pso =

          new PromptStringOptions("\nEnter text string");

        pso.AllowSpaces = true;

        PromptResult pr = ed.GetString(pso);



        if (pr.Status != PromptStatus.OK)

            return;



        Transaction tr =

          doc.TransactionManager.StartTransaction();

        using (tr)

        {

            BlockTableRecord btr =

              (BlockTableRecord)tr.GetObject(

                db.CurrentSpaceId, OpenMode.ForWrite

              );



            // Create the text object, set its normal and contents



            DBText txt = new DBText();

            txt.Normal =

              ed.CurrentUserCoordinateSystem.

                CoordinateSystem3d.Zaxis;

            txt.TextString = pr.StringResult;



            // We'll add the text to the database before jigging

            // it - this allows alignment adjustments to be

            // reflected



            btr.AppendEntity(txt);

            tr.AddNewlyCreatedDBObject(txt, true);



            // Create our jig



            TextPlacementJig pj = new TextPlacementJig(tr, db, txt);



            // Loop as we run our jig, as we may have keywords



            PromptStatus stat = PromptStatus.Keyword;

            while (stat == PromptStatus.Keyword)

            {

                PromptResult res = ed.Drag(pj);

                stat = res.Status;

                if (

                  stat != PromptStatus.OK &&

                  stat != PromptStatus.Keyword

                )

                    return;

            }



            tr.Commit();

        }

    }



    class TextPlacementJig : EntityJig

    {

        // Declare some internal state



        Database _db;
        Transaction _tr;
        Point3d _position;
        double _angle, _txtSize;



        // Constructor



        public TextPlacementJig(          Transaction tr, Database db, Entity ent        ) : base(ent)

        {
            _db = db;
           _tr = tr;
            _angle = 0;
            _txtSize = 1;

        }



        protected override SamplerStatus Sampler(

          JigPrompts jp

        )

        {

            // We acquire a point but with keywords



            JigPromptPointOptions po =              new JigPromptPointOptions(                "\nPosition of text"              );



            po.UserInputControls =

              (UserInputControls.Accept3dCoordinates |

                UserInputControls.NullResponseAccepted |

                UserInputControls.NoNegativeResponseAccepted |

                UserInputControls.GovernedByOrthoMode);



            po.SetMessageAndKeywords(

              "\nSpecify position of text or " +

              "[Bold/Italic/LArger/Smaller/" +

               "ROtate90/LEft/Middle/RIght]: ",

              "Bold Italic LArger Smaller " +

              "ROtate90 LEft Middle RIght"

            );



            PromptPointResult ppr = jp.AcquirePoint(po);



            if (ppr.Status == PromptStatus.Keyword)

            {

                switch (ppr.StringResult)

                {

                    case "Bold":

                        {

                            // TODO



                            break;

                        }

                    case "Italic":

                        {

                            // TODO



                            break;

                        }

                    case "LArger":

                        {

                            // Multiple the text size by two



                            _txtSize *= 2;

                            break;

                        }

                    case "Smaller":

                        {

                            // Divide the text size by two



                            _txtSize /= 2;

                            break;

                        }

                    case "ROtate90":

                        {

                            // To rotate clockwise we subtract 90 degrees and

                            // then normalise the angle between 0 and 360



                            _angle -= Math.PI / 2;

                            while (_angle < Math.PI * 2)

                            {

                                _angle += Math.PI * 2;

                            }

                            break;

                        }

                    case "LEft":

                        {

                            // TODO



                            break;

                        }

                    case "RIght":

                        {

                            // TODO



                            break;

                        }

                    case "Middle":

                        {

                            // TODO



                            break;

                        }

                }



                return SamplerStatus.OK;

            }

            else if (ppr.Status == PromptStatus.OK)

            {

                // Check if it has changed or not (reduces flicker)



                if (

                  _position.DistanceTo(ppr.Value) <

                    Tolerance.Global.EqualPoint

                )

                    return SamplerStatus.NoChange;



                _position = ppr.Value;

                return SamplerStatus.OK;

            }



            return SamplerStatus.Cancel;

        }



        protected override bool Update()

        {

            // Set properties on our text object



            DBText txt = (DBText)Entity;



            txt.Position = _position;

            txt.Height = _txtSize;

            txt.Rotation = _angle;



            return true;

        }

    }

}
