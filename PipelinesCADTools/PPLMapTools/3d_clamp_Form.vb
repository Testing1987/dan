Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class _3d_clamp_Form
    Private Sub _3d_clamp_Form_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        With ComboBox_bolt_nps
            .Items.Add("1/4")
            .Items.Add("5/16")
            .Items.Add("3/8")
            .Items.Add("7/16")
            .Items.Add("1/2")
            .Items.Add("9/16")
            .Items.Add("5/8")
            .Items.Add("3/4")
            .Items.Add("7/8")
            .Items.Add("1")
            .Items.Add("1 + 1/8")
            .Items.Add("1 + 1/4")
            .Items.Add("1 + 3/8")
            .Items.Add("1 + 1/2")
            .Items.Add("1 + 5/8")
            .Items.Add("1 + 3/4")
            .Items.Add("1 + 7/8")
            .Items.Add("2")
            .Items.Add("2 + 1/4")
            .Items.Add("2 + 1/2")
            .Items.Add("2 + 3/4")

        End With
        ComboBox_bolt_nps.SelectedIndex = 7
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Try

            If IsNumeric((TextBox_clamp_width_inches.Text)) = False Then
                MsgBox("Non numeric clamp width")
                Exit Sub
            End If
            If IsNumeric((TextBox_half_of_length_in.Text)) = False Then
                MsgBox("Non numeric length")
                Exit Sub
            End If
            If IsNumeric((TextBox_od_mm.Text)) = False Then
                MsgBox("Non numeric outside pipe diameter")
                Exit Sub
            End If
            If IsNumeric((TextBox_plate_separation_inches.Text)) = False Then
                MsgBox("Non numeric plate separation")
                Exit Sub
            End If
            If IsNumeric((TextBox_plate_thickness_inches.Text)) = False Then
                MsgBox("Non numeric plate thickness")
                Exit Sub
            End If
            If IsNumeric((TextBox_Bolt_dist_fromCL_inches.Text)) = False Then
                MsgBox("Non numeric Bolt separation")
                Exit Sub
            End If


            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Using Lock As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction



                    Dim Curent_UCS As Matrix3d
                    Curent_UCS = Editor1.CurrentUserCoordinateSystem


                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)


                    Dim Latime_clamp As Double = CDbl(TextBox_clamp_width_inches.Text) * 25.4
                    Dim Lungime_pe_jumatate As Double = CDbl(TextBox_half_of_length_in.Text) * 25.4
                    Dim Separatie_plate As Double = CDbl(TextBox_plate_separation_inches.Text) * 25.4
                    Dim Grosime_plate As Double = CDbl(TextBox_plate_thickness_inches.Text) * 25.4
                    Dim OD_pipe As Double = CDbl(TextBox_od_mm.Text)
                    Dim Dist_bolt_CL As Double = CDbl(TextBox_Bolt_dist_fromCL_inches.Text) * 25.4


                    Editor1.CurrentUserCoordinateSystem = WCS_align()

                    Dim Colectie_Obj_ID As New ObjectIdCollection

                    Dim PP_start As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select point")
                    PP_start.AllowNone = True

                    Dim Point_start_result As Autodesk.AutoCAD.EditorInput.PromptPointResult
                    Point_start_result = Editor1.GetPoint(PP_start)
                    If Not Point_start_result.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Exit Sub
                    End If

                    Dim PP_directie As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select rotation on XY plane")
                    PP_directie.AllowNone = True
                    PP_directie.UseBasePoint = True
                    PP_directie.BasePoint = Point_start_result.Value

                    Dim Point_dir_result As Autodesk.AutoCAD.EditorInput.PromptPointResult

                    Point_dir_result = Editor1.GetPoint(PP_directie)
                    If Not Point_dir_result.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Exit Sub
                    End If


                    Dim Punct_start As Point3d = Point_start_result.Value
                    Dim Punct_dir As Point3d = Point_dir_result.Value
                    Dim Rotatie_xy As Double = GET_Bearing_rad(Punct_start.X, Punct_start.Y, Punct_dir.X, Punct_dir.Y)
                    Dim Rotatie_dir_matrix As Matrix3d = Matrix3d.Rotation(Rotatie_xy, Vector3d.ZAxis, Punct_start)


                    Dim Solid1 As Solid3d = Creaza_rounded_plate_jos(Punct_start, Grosime_plate, Separatie_plate, Lungime_pe_jumatate, OD_pipe, Latime_clamp)
                    Solid1.TransformBy(Rotatie_dir_matrix)
                    Colectie_Obj_ID.Add(Solid1.ObjectId)

                    Dim Solid2 As Solid3d = Creaza_rounded_plate_SUS(New Point3d(Punct_start.X, Punct_start.Y, Punct_start.Z + Grosime_plate + Separatie_plate), Grosime_plate, Separatie_plate, Lungime_pe_jumatate, OD_pipe, Latime_clamp)
                    Solid2.TransformBy(Rotatie_dir_matrix)
                    Colectie_Obj_ID.Add(Solid2.ObjectId)
                    Dim Inaltime_cilindru As Double = Grosime_plate * 2 + Separatie_plate + get_nut_thickness_inches(ComboBox_bolt_nps.Text) * 25.4 + 2 * get_WASHER_thickness_inches(ComboBox_bolt_nps.Text) * 25.4

                    Dim Increment_l As Double = 25.4 / 8
                    Inaltime_cilindru = Ceiling(Inaltime_cilindru / Increment_l) * Increment_l

                    Dim Punct_Bolt_Stanga As New Point3d(Punct_start.X + Lungime_pe_jumatate - Dist_bolt_CL, Punct_start.Y + Latime_clamp / 2, Punct_start.Z + 2 * Grosime_plate + Separatie_plate)
                    Dim Solid3 As Solid3d = Creaza_Cilindru(Punct_Bolt_Stanga, 25.4 * get_nut_thickness_inches(ComboBox_bolt_nps.Text) / 2, Inaltime_cilindru)
                    Solid3.TransformBy(Rotatie_dir_matrix)
                    Colectie_Obj_ID.Add(Solid3.ObjectId)

                    Dim Punct_Bolt_dreapta As New Point3d(Punct_start.X + Lungime_pe_jumatate + Dist_bolt_CL, Punct_start.Y + Latime_clamp / 2, Punct_start.Z + 2 * Grosime_plate + Separatie_plate)
                    Dim Solid4 As Solid3d = Creaza_Cilindru(Punct_Bolt_dreapta, 25.4 * get_nut_thickness_inches(ComboBox_bolt_nps.Text) / 2, Inaltime_cilindru)
                    Solid4.TransformBy(Rotatie_dir_matrix)
                    Colectie_Obj_ID.Add(Solid4.ObjectId)

                    Dim Solid5 As Solid3d = Creaza_Hexagon(Punct_Bolt_Stanga, 0.5 * get_hexagon_diam_inches(ComboBox_bolt_nps.Text) * 25.4, get_bolt_head_thickness_inches(ComboBox_bolt_nps.Text) * 25.4)
                    Solid5.TransformBy(Rotatie_dir_matrix)
                    Colectie_Obj_ID.Add(Solid5.ObjectId)

                    Dim Solid6 As Solid3d = Creaza_Hexagon(Punct_Bolt_dreapta, 0.5 * get_hexagon_diam_inches(ComboBox_bolt_nps.Text) * 25.4, get_bolt_head_thickness_inches(ComboBox_bolt_nps.Text) * 25.4)
                    Solid6.TransformBy(Rotatie_dir_matrix)
                    Colectie_Obj_ID.Add(Solid6.ObjectId)

                    Dim Point_nut_stanga As New Point3d(Punct_Bolt_Stanga.X, Punct_Bolt_Stanga.Y, Punct_Bolt_Stanga.Z - Grosime_plate * 2 - Separatie_plate - get_nut_thickness_inches(ComboBox_bolt_nps.Text) * 25.4)
                    Dim Solid7 As Solid3d = Creaza_Hexagon(Point_nut_stanga, 0.5 * get_hexagon_diam_inches(ComboBox_bolt_nps.Text) * 25.4, get_nut_thickness_inches(ComboBox_bolt_nps.Text) * 25.4)
                    Solid7.TransformBy(Rotatie_dir_matrix)
                    Colectie_Obj_ID.Add(Solid7.ObjectId)

                    Dim Point_nut_dreapta As New Point3d(Punct_Bolt_dreapta.X, Punct_Bolt_dreapta.Y, Punct_Bolt_dreapta.Z - Grosime_plate * 2 - Separatie_plate - get_nut_thickness_inches(ComboBox_bolt_nps.Text) * 25.4)
                    Dim Solid8 As Solid3d = Creaza_Hexagon(Point_nut_dreapta, 0.5 * get_hexagon_diam_inches(ComboBox_bolt_nps.Text) * 25.4, get_nut_thickness_inches(ComboBox_bolt_nps.Text) * 25.4)
                    Solid8.TransformBy(Rotatie_dir_matrix)
                    Colectie_Obj_ID.Add(Solid8.ObjectId)


                    Creaza_group_of_dbobjects(Colectie_Obj_ID, "DAN POPESCU GROUP")

                    Editor1.CurrentUserCoordinateSystem = Curent_UCS
                    Trans1.Commit()
                End Using
            End Using
            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Function Creaza_rounded_plate_jos(ByVal Punct_prim_vertex As Point3d, ByVal Grosime_plate As Double, ByVal separatie_plate As Double, ByVal Lungime_plate_per_2 As Double, ByVal OD As Double, ByVal Latime_plate As Double) As Solid3d
        Try
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Using Lock1 As DocumentLock = ThisDrawing.LockDocument


                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                    Dim Poly1 As New Autodesk.AutoCAD.DatabaseServices.Polyline

                    Dim Punct0 As New Point2d(Punct_prim_vertex.X, Punct_prim_vertex.Z)
                    Dim Punct1 As New Point2d(Punct_prim_vertex.X, Punct_prim_vertex.Z + Grosime_plate)

                    Dim Centru_cerc As New Point3d(Punct_prim_vertex.X + Lungime_plate_per_2, Punct_prim_vertex.Y, Punct_prim_vertex.Z + Grosime_plate + separatie_plate / 2)
                    Dim Cerc_interior As New Circle(Centru_cerc, Vector3d.YAxis, OD / 2)
                    Dim Cerc_exterior As New Circle(Centru_cerc, Vector3d.YAxis, OD / 2 + Grosime_plate)
                    Dim Linie_sus As New Line(New Point3d(Punct_prim_vertex.X, Punct_prim_vertex.Y, Punct_prim_vertex.Z + Grosime_plate), New Point3d(Punct_prim_vertex.X + 2 * Lungime_plate_per_2, Punct_prim_vertex.Y, Punct_prim_vertex.Z + Grosime_plate))
                    Dim Linie_jos As New Line(New Point3d(Punct_prim_vertex.X, Punct_prim_vertex.Y, Punct_prim_vertex.Z), New Point3d(Punct_prim_vertex.X + 2 * Lungime_plate_per_2, Punct_prim_vertex.Y, Punct_prim_vertex.Z))
                    Dim Colectie_sus As New Point3dCollection
                    Cerc_interior.IntersectWith(Linie_sus, Intersect.OnBothOperands, Colectie_sus, IntPtr.Zero, IntPtr.Zero)

                    Dim Colectie_jos As New Point3dCollection
                    Cerc_exterior.IntersectWith(Linie_jos, Intersect.OnBothOperands, Colectie_jos, IntPtr.Zero, IntPtr.Zero)

                    Dim Punct2 As New Point2d(Colectie_sus(0).X, Colectie_sus(0).Z)
                    Dim Punct3 As New Point2d(Colectie_sus(1).X, Colectie_sus(1).Z)
                    Dim Punct4 As New Point2d(Punct_prim_vertex.X + 2 * Lungime_plate_per_2, Punct_prim_vertex.Z + Grosime_plate)
                    Dim Punct5 As New Point2d(Punct_prim_vertex.X + 2 * Lungime_plate_per_2, Punct_prim_vertex.Z)
                    Dim Punct6 As New Point2d(Colectie_jos(1).X, Colectie_jos(1).Z)
                    Dim Punct7 As New Point2d(Colectie_jos(0).X, Colectie_jos(0).Z)

                    Dim Linie1 As New Line(Colectie_sus(0), Colectie_sus(1))
                    Dim Linie2 As New Line(Linie1.GetClosestPointTo(Centru_cerc, True), Centru_cerc)
                    Dim H1 As Double = Linie2.Length
                    Dim L1 As Double = Linie1.Length
                    Dim Alpha1 As Double = 2 * Atan(0.5 * L1 / H1)
                    Dim Bulge1 As Double = Tan(Alpha1 / 4)

                    Dim Linie11 As New Line(Colectie_jos(0), Colectie_jos(1))
                    Dim Linie21 As New Line(Linie11.GetClosestPointTo(Centru_cerc, True), Centru_cerc)
                    Dim H2 As Double = Linie21.Length
                    Dim L2 As Double = Linie11.Length
                    Dim Alpha2 As Double = 2 * Atan(0.5 * L2 / H2)
                    Dim Bulge2 As Double = Tan(Alpha2 / 4)


                    Poly1.AddVertexAt(0, Punct0, 0, 0, 0)
                    Poly1.AddVertexAt(1, Punct1, 0, 0, 0)
                    Poly1.AddVertexAt(2, Punct2, Bulge1, 0, 0)
                    Poly1.AddVertexAt(3, Punct3, 0, 0, 0)
                    Poly1.AddVertexAt(4, Punct4, 0, 0, 0)
                    Poly1.AddVertexAt(5, Punct5, 0, 0, 0)
                    Poly1.AddVertexAt(6, Punct6, -Bulge2, 0, 0)
                    Poly1.AddVertexAt(7, Punct7, 0, 0, 0)

                    Poly1.Closed = True
                    Poly1.Normal = Vector3d.YAxis

                    Dim Point3d_move As New Point3d
                    Point3d_move = Poly1.GetPoint3dAt(0)

                    Dim Vector1 As Vector3d = (Point3d_move.GetVectorTo(Punct_prim_vertex))

                    Poly1.TransformBy(Matrix3d.Displacement(Vector1))
                    Poly1.TransformBy(Matrix3d.Rotation(PI, Vector3d.ZAxis, Punct_prim_vertex))


                    Dim Colectie_segmente As New DBObjectCollection
                    Colectie_segmente.Add(Poly1)

                    Dim Colectie_Regiune As New DBObjectCollection
                    Colectie_Regiune = Autodesk.AutoCAD.DatabaseServices.Region.CreateFromCurves(Colectie_segmente)

                    Dim Regiunea As New Region
                    Regiunea = Colectie_Regiune(0)
                    Dim Cub3d As New Solid3d
                    Cub3d.RecordHistory = True
                    Cub3d.Extrude(Regiunea, -Latime_plate, 0)



                    BTrecord.AppendEntity(Cub3d)
                    Trans1.AddNewlyCreatedDBObject(Cub3d, True)


                    Trans1.Commit()
                    Return Cub3d
                End Using ' asta e de la trans1




            End Using

        Catch ex As Exception

            MsgBox(ex.Message)
        End Try

    End Function
    Public Function Creaza_rounded_plate_SUS(ByVal Punct_prim_vertex As Point3d, ByVal Grosime_plate As Double, ByVal separatie_plate As Double, ByVal Lungime_plate_per_2 As Double, ByVal OD As Double, ByVal Latime_plate As Double) As Solid3d
        Try
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Using Lock1 As DocumentLock = ThisDrawing.LockDocument


                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                    Dim Poly1 As New Autodesk.AutoCAD.DatabaseServices.Polyline

                    Dim Punct0 As New Point2d(Punct_prim_vertex.X, Punct_prim_vertex.Z)
                    Dim Punct1 As New Point2d(Punct_prim_vertex.X, Punct_prim_vertex.Z + Grosime_plate)

                    Dim Centru_cerc As New Point3d(Punct_prim_vertex.X + Lungime_plate_per_2, Punct_prim_vertex.Y, Punct_prim_vertex.Z - separatie_plate / 2)


                    Dim Cerc_exterior As New Circle(Centru_cerc, Vector3d.YAxis, OD / 2 + Grosime_plate)
                    Dim Linie_sus As New Line(New Point3d(Punct_prim_vertex.X, Punct_prim_vertex.Y, Punct_prim_vertex.Z + Grosime_plate), New Point3d(Punct_prim_vertex.X + 2 * Lungime_plate_per_2, Punct_prim_vertex.Y, Punct_prim_vertex.Z + Grosime_plate))
                    Dim Colectie_sus As New Point3dCollection
                    Cerc_exterior.IntersectWith(Linie_sus, Intersect.OnBothOperands, Colectie_sus, IntPtr.Zero, IntPtr.Zero)

                    Dim Cerc_interior As New Circle(Centru_cerc, Vector3d.YAxis, OD / 2)
                    Dim Linie_jos As New Line(Punct_prim_vertex, New Point3d(Punct_prim_vertex.X + 2 * Lungime_plate_per_2, Punct_prim_vertex.Y, Punct_prim_vertex.Z))
                    Dim Colectie_jos As New Point3dCollection
                    Cerc_interior.IntersectWith(Linie_jos, Intersect.OnBothOperands, Colectie_jos, IntPtr.Zero, IntPtr.Zero)

                    Dim Punct2 As New Point2d(Colectie_sus(1).X, Colectie_sus(1).Z)
                    Dim Punct3 As New Point2d(Colectie_sus(0).X, Colectie_sus(0).Z)
                    Dim Punct4 As New Point2d(Punct_prim_vertex.X + 2 * Lungime_plate_per_2, Punct_prim_vertex.Z + Grosime_plate)
                    Dim Punct5 As New Point2d(Punct_prim_vertex.X + 2 * Lungime_plate_per_2, Punct_prim_vertex.Z)
                    Dim Punct6 As New Point2d(Colectie_jos(0).X, Colectie_jos(0).Z)
                    Dim Punct7 As New Point2d(Colectie_jos(1).X, Colectie_jos(1).Z)

                    Dim Linie1 As New Line(Colectie_sus(0), Colectie_sus(1))
                    Dim Linie2 As New Line(Linie1.GetClosestPointTo(Centru_cerc, True), Centru_cerc)
                    Dim H1 As Double = Linie2.Length
                    Dim L1 As Double = Linie1.Length
                    Dim Alpha1 As Double = 2 * Atan(0.5 * L1 / H1)
                    Dim Bulge1 As Double = Tan(Alpha1 / 4)

                    Dim Linie11 As New Line(Colectie_jos(0), Colectie_jos(1))
                    Dim Linie21 As New Line(Linie11.GetClosestPointTo(Centru_cerc, True), Centru_cerc)
                    Dim H2 As Double = Linie21.Length
                    Dim L2 As Double = Linie11.Length
                    Dim Alpha2 As Double = 2 * Atan(0.5 * L2 / H2)
                    Dim Bulge2 As Double = Tan(Alpha2 / 4)


                    Poly1.AddVertexAt(0, Punct0, 0, 0, 0)
                    Poly1.AddVertexAt(1, Punct1, 0, 0, 0)
                    Poly1.AddVertexAt(2, Punct2, -Bulge1, 0, 0)
                    Poly1.AddVertexAt(3, Punct3, 0, 0, 0)
                    Poly1.AddVertexAt(4, Punct4, 0, 0, 0)
                    Poly1.AddVertexAt(5, Punct5, 0, 0, 0)
                    Poly1.AddVertexAt(6, Punct6, Bulge2, 0, 0)
                    Poly1.AddVertexAt(7, Punct7, 0, 0, 0)

                    Poly1.Closed = True
                    Poly1.Normal = Vector3d.YAxis

                    Dim Point3d_move As New Point3d
                    Point3d_move = Poly1.GetPoint3dAt(0)

                    Dim Vector1 As Vector3d = Point3d_move.GetVectorTo(Punct_prim_vertex)

                    Poly1.TransformBy(Matrix3d.Displacement(Vector1))
                    Poly1.TransformBy(Matrix3d.Rotation(PI, Vector3d.ZAxis, Punct_prim_vertex))


                    Dim Colectie_segmente As New DBObjectCollection
                    Colectie_segmente.Add(Poly1)

                    Dim Colectie_Regiune As New DBObjectCollection
                    Colectie_Regiune = Autodesk.AutoCAD.DatabaseServices.Region.CreateFromCurves(Colectie_segmente)

                    Dim Regiunea As New Region
                    Regiunea = Colectie_Regiune(0)
                    Dim Cub3d As New Solid3d
                    Cub3d.RecordHistory = True
                    Cub3d.Extrude(Regiunea, -Latime_plate, 0)

                    BTrecord.AppendEntity(Cub3d)
                    Trans1.AddNewlyCreatedDBObject(Cub3d, True)


                    Trans1.Commit()
                    Return Cub3d
                End Using ' asta e de la trans1




            End Using

        Catch ex As Exception

            MsgBox(ex.Message)
        End Try

    End Function

    Public Function Creaza_Cilindru(ByVal Punct_Centru As Point3d, ByVal Raza As Double, ByVal inaltime_cilindru As Double) As Solid3d
        Try
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Using Lock1 As DocumentLock = ThisDrawing.LockDocument


                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)

                    Dim Cilindru1 As New Solid3d
                    Cilindru1.RecordHistory = True
                    Cilindru1.CreateFrustum(inaltime_cilindru, Raza, Raza, Raza)
                    BTrecord.AppendEntity(Cilindru1)
                    Trans1.AddNewlyCreatedDBObject(Cilindru1, True)
                    'Cilindru1.TransformBy(Matrix3d.Displacement(Punct_Centru - Point3d.Origin))
                    Cilindru1.TransformBy(Matrix3d.Displacement(Point3d.Origin.GetVectorTo(New Point3d(Punct_Centru.X, Punct_Centru.Y, Punct_Centru.Z - inaltime_cilindru / 2))))

                    Trans1.Commit()
                    Return Cilindru1
                End Using ' asta e de la trans1




            End Using

        Catch ex As Exception

            MsgBox(ex.Message)
        End Try

    End Function

    Public Function Creaza_Hexagon(ByVal Centre_point As Point3d, ByVal radius_hexagon As Double, ByVal height_hexagon As Double) As Solid3d
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




    Public Function get_hexagon_diam_inches(ByVal combo_value As String) As Double
        Select Case combo_value
            Case "1/4"
                Return 0.578
            Case "5/16"
                Return 0.686
            Case "3/8"
                Return 0.794
            Case "7/16"
                Return 0.902
            Case "1/2"
                Return 1.011
            Case "9/16"
                Return 1.119
            Case "5/8"
                Return 1.227
            Case "3/4"
                Return 1.444
            Case "7/8"
                Return 1.66
            Case "1"
                Return 1.877
            Case "1 + 1/8"
                Return 2.093
            Case "1 + 1/4"
                Return 2.31
            Case "1 + 3/8"
                Return 2.527
            Case "1 + 1/2"
                Return 2.743
            Case "1 + 5/8"
                Return 2.96
            Case "1 + 3/4"
                Return 3.176
            Case "1 + 7/8"
                Return 3.393
            Case "2"
                Return 3.609
            Case "2 + 1/4"
                Return 4.043
            Case "2 + 1/2"
                Return 4.476
            Case "2 + 3/4"
                Return 4.909

            Case Else
                Return 0

        End Select


    End Function




    Public Function get_nut_thickness_inches(ByVal combo_value As String) As Double
        Select Case combo_value
            Case "1/4"
                Return 0.25
            Case "5/16"
                Return 0.3125
            Case "3/8"
                Return 0.375
            Case "7/16"
                Return 0.4375
            Case "1/2"
                Return 0.5
            Case "9/16"
                Return 0.5625
            Case "5/8"
                Return 0.625
            Case "3/4"
                Return 0.75
            Case "7/8"
                Return 0.875
            Case "1"
                Return 1
            Case "1 + 1/8"
                Return 1.125
            Case "1 + 1/4"
                Return 1.25
            Case "1 + 3/8"
                Return 1.375
            Case "1 + 1/2"
                Return 1.5
            Case "1 + 5/8"
                Return 1.625
            Case "1 + 3/4"
                Return 1.75
            Case "1 + 7/8"
                Return 1.875
            Case "2"
                Return 2
            Case "2 + 1/4"
                Return 2.25
            Case "2 + 1/2"
                Return 2.5
            Case "2 + 3/4"
                Return 2.75




            Case Else
                Return 0

        End Select


    End Function


    Public Function get_bolt_head_thickness_inches(ByVal combo_value As String) As Double
        Select Case combo_value
            Case "1/4"
                Return 0.25
            Case "5/16"
                Return 0.296875
            Case "3/8"
                Return 0.34375
            Case "7/16"
                Return 0.390625
            Case "1/2"
                Return 0.4375
            Case "9/16"
                Return 0.484375
            Case "5/8"
                Return 0.53125
            Case "3/4"
                Return 0.625
            Case "7/8"
                Return 0.71875
            Case "1"
                Return 0.8125
            Case "1 + 1/8"
                Return 0.90625
            Case "1 + 1/4"
                Return 1
            Case "1 + 3/8"
                Return 1.09375
            Case "1 + 1/2"
                Return 1.1875
            Case "1 + 5/8"
                Return 1.28125
            Case "1 + 3/4"
                Return 1.375
            Case "1 + 7/8"
                Return 1.46875
            Case "2"
                Return 1.5625
            Case "2 + 1/4"
                Return 1.75
            Case "2 + 1/2"
                Return 1.9375
            Case "2 + 3/4"
                Return 2.125
           


            Case Else
                Return 0

        End Select


    End Function
    Public Function get_WASHER_thickness_inches(ByVal combo_value As String) As Double
        Select Case combo_value
            Case "1/4"
                Return 0.065
            Case "5/16"
                Return 0.065
            Case "3/8"
                Return 0.065
            Case "7/16"
                Return 0.083
            Case "1/2"
                Return 0.095
            Case "9/16"
                Return 0.5625
            Case "5/8"
                Return 0.109
            Case "3/4"
                Return 0.18
            Case "7/8"
                Return 0.18
            Case "1"
                Return 0.18
            Case "1 + 1/8"
                Return 0.18
            Case "1 + 1/4"
                Return 0.18
            Case "1 + 3/8"
                Return 0.18
            Case "1 + 1/2"
                Return 0.18
            Case "1 + 5/8"
                Return 0.18
            Case "1 + 3/4"
                Return 0.18
            Case "1 + 7/8"
                Return 0.18
            Case "2"
                Return 0.18
            Case "2 + 1/4"
                Return 0.18
            Case "2 + 1/2"
                Return 0.22
            Case "2 + 3/4"
                Return 0.22


            Case Else
                Return 0

        End Select


    End Function

End Class