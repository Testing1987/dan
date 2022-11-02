Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Public Class _3d_Ubolt_form

    Private Sub _3d_Ubolt_form_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        ComboBox_NPS.Items.Add("1")
        ComboBox_NPS.Items.Add("2")
        ComboBox_NPS.Items.Add("3")
        ComboBox_NPS.Items.Add("4")
        ComboBox_NPS.Items.Add("6")
        ComboBox_NPS.Items.Add("8")
        ComboBox_NPS.Items.Add("10")
        ComboBox_NPS.Items.Add("12")
        ComboBox_NPS.Items.Add("16")
        ComboBox_NPS.Items.Add("20")
        ComboBox_NPS.Items.Add("24")
        ComboBox_NPS.Items.Add("30")
        ComboBox_NPS.Items.Add("36")
        ComboBox_NPS.Items.Add("42")
        ComboBox_NPS.Items.Add("48")
        ComboBox_NPS.SelectedIndex = 0
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Try

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor

            Button1.Visible = False

            If IsNumeric(TextBox_stud_diam_mm.Text) = False Then
                MsgBox("Not numeric bar diameter")
                Button1.Visible = True
                Exit Sub
            End If

            If IsNumeric(TextBox_extend_mm.Text) = False Then
                MsgBox("Not numeric length of straight portion")
                Button1.Visible = True
                Exit Sub
            End If

            If IsNumeric(TextBox_plate_thickness.Text) = False Then
                MsgBox("Not numeric plate thickness")
                Button1.Visible = True
                Exit Sub
            End If

            If IsNumeric(TextBox_GAP_mm.Text) = False Then
                MsgBox("Not numeric gap")
                Button1.Visible = True
                Exit Sub
            End If

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Using Lock As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim Curent_UCS As Matrix3d

                    Curent_UCS = Editor1.CurrentUserCoordinateSystem

                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)

                    Dim GAP As Double = CDbl(TextBox_GAP_mm.Text)
                    Dim Diam_bara As Double = CDbl(TextBox_stud_diam_mm.Text)
                    Dim Extra_lungime As Double = CDbl(TextBox_extend_mm.Text)
                    Dim Grosime_placa As Double = CDbl(TextBox_plate_thickness.Text)
                    Dim OD As Double = 2 * Get_from_NPS_radius_for_pipes_from_inches_to_milimeters(CDbl(ComboBox_NPS.Text))

                    Editor1.CurrentUserCoordinateSystem = WCS_align()

                    Dim PP_start As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select point")
                    PP_start.AllowNone = True

                    Dim Point_start_result As Autodesk.AutoCAD.EditorInput.PromptPointResult
                    Point_start_result = Editor1.GetPoint(PP_start)
                    If Not Point_start_result.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Button1.Visible = True
                        Exit Sub
                    End If
                    Dim Punct_start As Point3d = Point_start_result.Value

                    Dim PP_UNGHI As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select rotation")
                    PP_UNGHI.AllowNone = True
                    PP_UNGHI.UseBasePoint = True
                    PP_UNGHI.BasePoint = Punct_start

                    Dim Point_rot_result As Autodesk.AutoCAD.EditorInput.PromptPointResult
                    Point_rot_result = Editor1.GetPoint(PP_UNGHI)
                    If Not Point_rot_result.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Button1.Visible = True
                        Exit Sub
                    End If

                    Dim Solid1 As New Solid3d
                    Solid1 = Creaza_Hexagon(New Point3d(Punct_start.X - OD / 2 - GAP - Diam_bara / 2, Punct_start.Y, Punct_start.Z - OD / 2 - Grosime_placa), 1.5 * Diam_bara / 2 / Cos(30 * PI / 180), -0.8 * Diam_bara, 0)

                    Dim Solid2 As New Solid3d
                    Solid2 = Creaza_Hexagon(New Point3d(Punct_start.X + OD / 2 + GAP + Diam_bara / 2, Punct_start.Y, Punct_start.Z - OD / 2 - Grosime_placa), 1.5 * Diam_bara / 2 / Cos(30 * PI / 180), -0.8 * Diam_bara, 0)

                    Dim Solid3 As New Solid3d
                    Solid3 = Creaza_Hexagon(New Point3d(Punct_start.X - OD / 2 - GAP - Diam_bara / 2, Punct_start.Y, Punct_start.Z - OD / 2 - Grosime_placa - 0.8 * Diam_bara), 1.5 * Diam_bara / 2 / Cos(30 * PI / 180), -0.8 * Diam_bara, PI / 2)

                    Dim Solid4 As New Solid3d
                    Solid4 = Creaza_Hexagon(New Point3d(Punct_start.X + OD / 2 + GAP + Diam_bara / 2, Punct_start.Y, Punct_start.Z - OD / 2 - Grosime_placa - 0.8 * Diam_bara), 1.5 * Diam_bara / 2 / Cos(30 * PI / 180), -0.8 * Diam_bara, PI / 2)

                    Editor1.CurrentUserCoordinateSystem = UCS_align_FRONT()

                    Dim Solid5 As New Solid3d
                    Solid5 = Creaza_uBOLT(Punct_start, Diam_bara, OD, GAP, Extra_lungime + OD / 2)

                    Editor1.CurrentUserCoordinateSystem = WCS_align()

                    Dim Planul_curent As Plane = New Plane(New Point3d(0, 0, 0), Editor1.CurrentUserCoordinateSystem.CoordinateSystem3d.Zaxis)

                    Dim Rotatie_in_xyPlane As Double = Punct_start.GetVectorTo(Point_rot_result.Value).AngleOnPlane(Planul_curent) + PI / 2
                    Solid1.TransformBy(Matrix3d.Rotation(Rotatie_in_xyPlane, Vector3d.ZAxis, Punct_start))
                    Solid2.TransformBy(Matrix3d.Rotation(Rotatie_in_xyPlane, Vector3d.ZAxis, Punct_start))
                    Solid3.TransformBy(Matrix3d.Rotation(Rotatie_in_xyPlane, Vector3d.ZAxis, Punct_start))
                    Solid4.TransformBy(Matrix3d.Rotation(Rotatie_in_xyPlane, Vector3d.ZAxis, Punct_start))
                    Solid5.TransformBy(Matrix3d.Rotation(Rotatie_in_xyPlane, Vector3d.ZAxis, Punct_start))

                    Dim vZ As New Vector3d
                    vZ = Solid1.GeometricExtents.MinPoint.GetVectorTo(Solid2.GeometricExtents.MinPoint)

                    Planul_curent = New Plane(New Point3d(0, 0, 0), vZ)
                    Rotatie_in_xyPlane = Punct_start.GetVectorTo(Point_rot_result.Value).AngleOnPlane(Planul_curent) - PI
                    Solid1.TransformBy(Matrix3d.Rotation(Rotatie_in_xyPlane, vZ, Punct_start))
                    Solid2.TransformBy(Matrix3d.Rotation(Rotatie_in_xyPlane, vZ, Punct_start))
                    Solid3.TransformBy(Matrix3d.Rotation(Rotatie_in_xyPlane, vZ, Punct_start))
                    Solid4.TransformBy(Matrix3d.Rotation(Rotatie_in_xyPlane, vZ, Punct_start))
                    Solid5.TransformBy(Matrix3d.Rotation(Rotatie_in_xyPlane, vZ, Punct_start))

                    Trans1.Commit()

                    ThisDrawing.Editor.WriteMessage(vbLf & "Command:")

                End Using

            End Using ' asta e de la trans1

            Button1.Visible = True

        Catch ex As Exception
            Button1.Visible = True
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Function Creaza_Hexagon(ByVal Centre_point As Point3d, ByVal radius_hexagon As Double, ByVal height_hexagon As Double, ByVal valoare_rotatie_along_z_axis As Double) As Solid3d
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim Polygon3d As Solid3d = Nothing

        Try

            Using Lock1 As DocumentLock = ThisDrawing.LockDocument


                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)



                    Dim Poly1 As New Autodesk.AutoCAD.DatabaseServices.Polyline

                    Dim Point00 As New Point2d(Centre_point.X + radius_hexagon, Centre_point.Y)
                    Dim Point01 As New Point2d(Centre_point.X + radius_hexagon * Cos(PI / 3), Centre_point.Y + radius_hexagon * Sin(PI / 3))
                    Dim Point02 As New Point2d(Centre_point.X - radius_hexagon * Cos(PI / 3), Centre_point.Y + radius_hexagon * Sin(PI / 3))
                    Dim Point03 As New Point2d(Centre_point.X - radius_hexagon, Centre_point.Y)
                    Dim Point04 As New Point2d(Centre_point.X - radius_hexagon * Cos(PI / 3), Centre_point.Y - radius_hexagon * Sin(PI / 3))
                    Dim Point05 As New Point2d(Centre_point.X + radius_hexagon * Cos(PI / 3), Centre_point.Y - radius_hexagon * Sin(PI / 3))

                    Poly1.AddVertexAt(0, Point00, 0, 0, 0)
                    Poly1.AddVertexAt(1, Point01, 0, 0, 0)
                    Poly1.AddVertexAt(2, Point02, 0, 0, 0)
                    Poly1.AddVertexAt(3, Point03, 0, 0, 0)
                    Poly1.AddVertexAt(4, Point04, 0, 0, 0)
                    Poly1.AddVertexAt(5, Point05, 0, 0, 0)
                    Poly1.Closed = True

                    Poly1.TransformBy(Matrix3d.Displacement(Poly1.GetPointAtParameter(0).GetVectorTo(New Point3d(Centre_point.X + radius_hexagon, Centre_point.Y, Centre_point.Z))))

                    Dim Segments_collection As New DBObjectCollection
                    Segments_collection.Add(Poly1)


                    Dim Colectie_Regiune As New DBObjectCollection
                    Colectie_Regiune = Autodesk.AutoCAD.DatabaseServices.Region.CreateFromCurves(Segments_collection)

                    Dim Regiunea As New Region
                    Regiunea = Colectie_Regiune(0)
                    Polygon3d = New Solid3d
                    Polygon3d.RecordHistory = True
                    Polygon3d.ShowHistory = True

                    Polygon3d.Extrude(Regiunea, height_hexagon, 0)

                    Polygon3d.TransformBy(Matrix3d.Rotation(valoare_rotatie_along_z_axis, Vector3d.ZAxis, Centre_point))

                    BTrecord.AppendEntity(Polygon3d)
                    Trans1.AddNewlyCreatedDBObject(Polygon3d, True)



                    Trans1.Commit()


                End Using ' asta e de la trans1




            End Using

        Catch ex As Exception
            Polygon3d = Nothing
            MsgBox(ex.Message)
        End Try
        Return Polygon3d
    End Function
    Public Function Creaza_uBOLT(ByVal punct_start As Point3d, ByVal OD_bara As Double, ByVal OD_pipe As Double, ByVal Gap As Double, ByVal Lungime_de_la_centru As Double) As Solid3d
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim Ubolt As Solid3d = Nothing

        Try

            Using Lock1 As DocumentLock = ThisDrawing.LockDocument


                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)



                    Dim Poly1 As New Autodesk.AutoCAD.DatabaseServices.Polyline

                    Dim Point00 As New Point2d(-punct_start.X, punct_start.Z)
                    Dim Point01 As New Point2d(-punct_start.X, punct_start.Z + Lungime_de_la_centru)
                    Dim Point02 As New Point2d(-punct_start.X + OD_bara + 2 * Gap + OD_pipe, punct_start.Z + Lungime_de_la_centru)
                    Dim Point03 As New Point2d(-punct_start.X + OD_bara + 2 * Gap + OD_pipe, punct_start.Z)


                    Poly1.AddVertexAt(0, Point00, 0, 0, 0)
                    Poly1.AddVertexAt(1, Point01, -1, 0, 0)
                    Poly1.AddVertexAt(2, Point02, 0, 0, 0)
                    Poly1.AddVertexAt(3, Point03, 0, 0, 0)
                    Poly1.Normal = Vector3d.YAxis

                    Poly1.TransformBy(Matrix3d.Displacement(Poly1.GetPointAtParameter(0).GetVectorTo(New Point3d(punct_start.X + OD_bara / 2 + OD_pipe / 2 + Gap, punct_start.Y, punct_start.Z - Lungime_de_la_centru))))
                    BTrecord.AppendEntity(Poly1)
                    Trans1.AddNewlyCreatedDBObject(Poly1, True)

                    Dim Circle1 As New Circle
                    Circle1.Center = Poly1.StartPoint
                    Circle1.Radius = OD_bara / 2
                    Circle1.Normal = Vector3d.ZAxis


                    Dim Ent1 As Entity = Circle1
                    Dim Path1 As Curve = Poly1

                    Dim Sweep_builder As New SweepOptionsBuilder
                    Sweep_builder.Align = SweepOptionsAlignOption.AlignSweepEntityToPath
                    Sweep_builder.BasePoint = Poly1.StartPoint
                    Sweep_builder.Bank = True

                    Ubolt = New Solid3d
                    Ubolt.CreateSweptSolid(Ent1, Path1, Sweep_builder.ToSweepOptions)

                    BTrecord.AppendEntity(Ubolt)
                    Trans1.AddNewlyCreatedDBObject(Ubolt, True)

                    Poly1.Erase()

                    Trans1.Commit()


                End Using ' asta e de la trans1




            End Using

        Catch ex As Exception
            Ubolt = Nothing
            MsgBox(ex.Message)
        End Try
        Return Ubolt
    End Function

End Class