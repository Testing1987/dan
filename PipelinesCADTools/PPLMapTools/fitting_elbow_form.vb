Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Public Class fitting_elbow_form
    Private Sub fitting_elbow_form_Load(sender As Object, e As EventArgs) Handles Me.Load
        ComboBox_nps.Items.Add("NPS 1/2")
        ComboBox_nps.Items.Add("NPS 1")
        ComboBox_nps.Items.Add("NPS 1.5")
        ComboBox_nps.Items.Add("NPS 2")
        ComboBox_nps.Items.Add("NPS 3")
        ComboBox_nps.Items.Add("NPS 4")
        ComboBox_nps.Items.Add("NPS 6")
        ComboBox_nps.Items.Add("NPS 8")
        ComboBox_nps.Items.Add("NPS 10")
        ComboBox_nps.Items.Add("NPS 12")
        ComboBox_nps.Items.Add("NPS 16")
        ComboBox_nps.Items.Add("NPS 20")
        ComboBox_nps.Items.Add("NPS 24")
        ComboBox_nps.Items.Add("NPS 30")
        ComboBox_nps.Items.Add("NPS 36")
        ComboBox_nps.Items.Add("NPS 42")
        ComboBox_nps.Items.Add("NPS 48")
        ComboBox_nps.SelectedIndex = 13
        TextBox_radius1.Text = 3 * 2 * Get_from_NPS_radius_for_pipes_from_inches_to_milimeters(30) / 1000
        Incarca_existing_layers_to_combobox(ComboBox_layer_CL)
        Incarca_existing_layers_to_combobox(ComboBox_layer_OD)
        If ComboBox_layer_CL.Items.Contains("PCENTRE") = True Then
            ComboBox_layer_CL.SelectedIndex = ComboBox_layer_CL.Items.IndexOf("PCENTRE")
        End If
        If ComboBox_layer_OD.Items.Contains("PNEW") = True Then
            ComboBox_layer_OD.SelectedIndex = ComboBox_layer_OD.Items.IndexOf("PNEW")
        End If
    End Sub
    Private Sub Button_fitting_Click(sender As Object, e As EventArgs) Handles Button_fitting.Click
        Try
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Using Lock1 As Autodesk.AutoCAD.ApplicationServices.DocumentLock = ThisDrawing.LockDocument
                Dim R1 As Double
                If IsNumeric(TextBox_radius1.Text) = True Then
                    R1 = CDbl(TextBox_radius1.Text)
                Else
                    MsgBox("Verify the radius")
                    Exit Sub
                End If

                If IsNumeric(TextBox_tangent_length.Text) = False Then
                    MsgBox("Specify the tangent length")
                    Exit Sub
                End If

                If IsNumeric(TextBox_elbow_angle.Text) = False Then
                    MsgBox("Specify the elbow angle")
                    Exit Sub
                End If


                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                Editor1 = ThisDrawing.Editor

                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)

                    Dim Point_start As Autodesk.AutoCAD.EditorInput.PromptPointResult

                    Dim PP_start As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Pick starting point:")
                    PP_start.AllowNone = True
                    Point_start = Editor1.GetPoint(PP_start)
                    If Not Point_start.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Editor1.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If

                    Dim Point_end As Autodesk.AutoCAD.EditorInput.PromptPointResult

                    Dim PP_end As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Pick the point used to calculate elevation diference:")
                    PP_end.AllowNone = True
                    PP_end.UseBasePoint = True
                    PP_end.BasePoint = Point_start.Value
                    Point_end = Editor1.GetPoint(PP_end)
                    If Not Point_end.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Editor1.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If

                    Dim x0, y0, y1, x1 As Double
                    x0 = Point_start.Value.X
                    y0 = Point_start.Value.Y
                    y1 = Point_end.Value.Y

                    x1 = Point_end.Value.X
                    Dim Pt1 As New Point3d
                    Dim Pt2 As New Point3d
                    Dim Pt3 As New Point3d
                    Dim Pt4 As New Point3d
                    Dim Pt5 As New Point3d
                    Dim Pt6 As New Point3d
                    Dim Pt7 As New Point3d
                    Dim Pt8 As New Point3d

                    Pt1 = New Point3d(x0, y0, 0)
                    Dim Tangent As Double = CDbl(TextBox_tangent_length.Text)
                    Dim Angle1 As Double = CDbl(TextBox_elbow_angle.Text)
                    Dim Factor_stanga_dreapta As Double = 1

                    If y1 > y0 Then
                        If x1 >= x0 Then
                            Pt2 = New Point3d(x0 + Tangent, y0, 0)
                            Dim Centru_cerc1 As New Point3d(Pt2.X, Pt2.Y + R1, 0)
                            Dim Cerc1 As New Circle
                            Cerc1.Radius = R1
                            Cerc1.Center = Centru_cerc1

                            Dim Linie1 As New Line(Centru_cerc1, New Point3d(Pt2.X, Pt2.Y - 10, 0))
                            Linie1.TransformBy(Matrix3d.Rotation(Angle1 * PI / 180, Vector3d.ZAxis, Centru_cerc1))
                            Dim Col1 As New Point3dCollection
                            Cerc1.IntersectWith(Linie1, Intersect.OnBothOperands, Col1, IntPtr.Zero, IntPtr.Zero)
                            Pt3 = Col1(0)
                            Linie1.TransformBy(Matrix3d.Rotation(-PI / 2, Vector3d.ZAxis, Pt3))
                            Dim Linie_1 As New Line(Pt3, Linie1.StartPoint)
                            Pt4 = Linie_1.GetPointAtDist(Tangent)

                            Dim Linie2 As New Line(New Point3d(-10, y1, 0), New Point3d(10, y1, 0))
                            Dim Col2 As New Point3dCollection
                            Linie1.IntersectWith(Linie2, Intersect.ExtendBoth, Col2, IntPtr.Zero, IntPtr.Zero)
                            Dim PointI As New Point3d
                            PointI = Col2(0)
                            Dim Linie3 As New Line(PointI, Pt3)
                            Pt6 = Linie3.GetPointAtDist(R1 * Tan(0.5 * Angle1 * PI / 180))
                            Pt5 = Linie3.GetPointAtDist(Tangent + R1 * Tan(0.5 * Angle1 * PI / 180))
                            Linie3 = New Line(Pt6, Pt3)
                            Linie3.TransformBy(Matrix3d.Rotation(PI / 2, Vector3d.ZAxis, Pt6))
                            Linie3.TransformBy(Matrix3d.Scaling(1000, Pt6))
                            Dim Cerc3 As New Circle
                            Cerc3.Radius = R1
                            Cerc3.Center = Pt6
                            Dim Col3 As New Point3dCollection
                            Linie3.IntersectWith(Cerc3, Intersect.OnBothOperands, Col3, IntPtr.Zero, IntPtr.Zero)
                            Dim Centru_cerc3 As New Point3d
                            Centru_cerc3 = Col3(0)
                            Pt7 = New Point3d(Centru_cerc3.X, Centru_cerc3.Y + R1, 0)
                            Pt8 = New Point3d(Pt7.X + Tangent, Pt7.Y, 0)
                        Else
                            Factor_stanga_dreapta = -1
                            Pt2 = New Point3d(x0 - Tangent, y0, 0)
                            Dim Centru_cerc1 As New Point3d(Pt2.X, Pt2.Y + R1, 0)
                            Dim Cerc1 As New Circle
                            Cerc1.Radius = R1
                            Cerc1.Center = Centru_cerc1

                            Dim Linie1 As New Line(Centru_cerc1, New Point3d(Pt2.X, Pt2.Y - 10, 0))
                            Linie1.TransformBy(Matrix3d.Rotation(-Angle1 * PI / 180, Vector3d.ZAxis, Centru_cerc1))
                            Dim Col1 As New Point3dCollection
                            Cerc1.IntersectWith(Linie1, Intersect.OnBothOperands, Col1, IntPtr.Zero, IntPtr.Zero)
                            Pt3 = Col1(0)
                            Linie1.TransformBy(Matrix3d.Rotation(PI / 2, Vector3d.ZAxis, Pt3))
                            Dim Linie_1 As New Line(Pt3, Linie1.StartPoint)
                            Pt4 = Linie_1.GetPointAtDist(Tangent)

                            Dim Linie2 As New Line(New Point3d(-10, y1, 0), New Point3d(10, y1, 0))
                            Dim Col2 As New Point3dCollection
                            Linie1.IntersectWith(Linie2, Intersect.ExtendBoth, Col2, IntPtr.Zero, IntPtr.Zero)
                            Dim PointI As New Point3d
                            PointI = Col2(0)
                            Dim Linie3 As New Line(PointI, Pt3)
                            Pt6 = Linie3.GetPointAtDist(R1 * Tan(0.5 * Angle1 * PI / 180))
                            Pt5 = Linie3.GetPointAtDist(Tangent + R1 * Tan(0.5 * Angle1 * PI / 180))
                            Linie3 = New Line(Pt6, Pt3)
                            Linie3.TransformBy(Matrix3d.Rotation(-PI / 2, Vector3d.ZAxis, Pt6))
                            Linie3.TransformBy(Matrix3d.Scaling(1000, Pt6))
                            Dim Cerc3 As New Circle
                            Cerc3.Radius = R1
                            Cerc3.Center = Pt6
                            Dim Col3 As New Point3dCollection
                            Linie3.IntersectWith(Cerc3, Intersect.OnBothOperands, Col3, IntPtr.Zero, IntPtr.Zero)
                            Dim Centru_cerc3 As New Point3d
                            Centru_cerc3 = Col3(0)
                            Pt7 = New Point3d(Centru_cerc3.X, Centru_cerc3.Y + R1, 0)
                            Pt8 = New Point3d(Pt7.X - Tangent, Pt7.Y, 0)

                        End If
                    Else
                        If x1 >= x0 Then
                            Factor_stanga_dreapta = -1
                            Pt2 = New Point3d(x0 + Tangent, y0, 0)
                            Dim Centru_cerc1 As New Point3d(Pt2.X, Pt2.Y - R1, 0)
                            Dim Cerc1 As New Circle
                            Cerc1.Radius = R1
                            Cerc1.Center = Centru_cerc1

                            Dim Linie1 As New Line(Centru_cerc1, New Point3d(Pt2.X, Pt2.Y, 0))
                            Linie1.TransformBy(Matrix3d.Rotation(-Angle1 * PI / 180, Vector3d.ZAxis, Centru_cerc1))
                            Linie1.TransformBy(Matrix3d.Scaling(2, Centru_cerc1))
                            Dim Col1 As New Point3dCollection
                            Cerc1.IntersectWith(Linie1, Intersect.OnBothOperands, Col1, IntPtr.Zero, IntPtr.Zero)
                            Pt3 = Col1(0)
                            Linie1.TransformBy(Matrix3d.Rotation(PI / 2, Vector3d.ZAxis, Pt3))
                            Dim Linie_1 As New Line(Pt3, Linie1.StartPoint)
                            Pt4 = Linie_1.GetPointAtDist(Tangent)

                            Dim Linie2 As New Line(New Point3d(-10, y1, 0), New Point3d(10, y1, 0))
                            Dim Col2 As New Point3dCollection
                            Linie1.IntersectWith(Linie2, Intersect.ExtendBoth, Col2, IntPtr.Zero, IntPtr.Zero)
                            Dim PointI As New Point3d
                            PointI = Col2(0)
                            Dim Linie3 As New Line(PointI, Pt3)
                            Pt6 = Linie3.GetPointAtDist(R1 * Tan(0.5 * Angle1 * PI / 180))
                            Pt5 = Linie3.GetPointAtDist(Tangent + R1 * Tan(0.5 * Angle1 * PI / 180))
                            Linie3 = New Line(Pt6, Pt3)
                            Linie3.TransformBy(Matrix3d.Rotation(-PI / 2, Vector3d.ZAxis, Pt6))
                            Linie3.TransformBy(Matrix3d.Scaling(1000, Pt6))
                            Dim Cerc3 As New Circle
                            Cerc3.Radius = R1
                            Cerc3.Center = Pt6
                            Dim Col3 As New Point3dCollection
                            Linie3.IntersectWith(Cerc3, Intersect.OnBothOperands, Col3, IntPtr.Zero, IntPtr.Zero)
                            Dim Centru_cerc3 As New Point3d
                            Centru_cerc3 = Col3(0)
                            Pt7 = New Point3d(Centru_cerc3.X, Centru_cerc3.Y - R1, 0)
                            Pt8 = New Point3d(Pt7.X + Tangent, Pt7.Y, 0)
                        Else

                            Pt2 = New Point3d(x0 - Tangent, y0, 0)
                            Dim Centru_cerc1 As New Point3d(Pt2.X, Pt2.Y - R1, 0)
                            Dim Cerc1 As New Circle
                            Cerc1.Radius = R1
                            Cerc1.Center = Centru_cerc1

                            Dim Linie1 As New Line(Centru_cerc1, New Point3d(Pt2.X, Pt2.Y, 0))
                            Linie1.TransformBy(Matrix3d.Rotation(Angle1 * PI / 180, Vector3d.ZAxis, Centru_cerc1))
                            Linie1.TransformBy(Matrix3d.Scaling(2, Centru_cerc1))
                            Dim Col1 As New Point3dCollection
                            Cerc1.IntersectWith(Linie1, Intersect.OnBothOperands, Col1, IntPtr.Zero, IntPtr.Zero)
                            Pt3 = Col1(0)
                            Linie1.TransformBy(Matrix3d.Rotation(-PI / 2, Vector3d.ZAxis, Pt3))
                            Dim Linie_1 As New Line(Pt3, Linie1.StartPoint)
                            Pt4 = Linie_1.GetPointAtDist(Tangent)

                            Dim Linie2 As New Line(New Point3d(-10, y1, 0), New Point3d(10, y1, 0))
                            Dim Col2 As New Point3dCollection
                            Linie1.IntersectWith(Linie2, Intersect.ExtendBoth, Col2, IntPtr.Zero, IntPtr.Zero)
                            Dim PointI As New Point3d
                            PointI = Col2(0)
                            Dim Linie3 As New Line(PointI, Pt3)
                            Pt6 = Linie3.GetPointAtDist(R1 * Tan(0.5 * Angle1 * PI / 180))
                            Pt5 = Linie3.GetPointAtDist(Tangent + R1 * Tan(0.5 * Angle1 * PI / 180))
                            Linie3 = New Line(Pt6, Pt3)
                            Linie3.TransformBy(Matrix3d.Rotation(PI / 2, Vector3d.ZAxis, Pt6))
                            Linie3.TransformBy(Matrix3d.Scaling(1000, Pt6))
                            Dim Cerc3 As New Circle
                            Cerc3.Radius = R1
                            Cerc3.Center = Pt6
                            Dim Col3 As New Point3dCollection
                            Linie3.IntersectWith(Cerc3, Intersect.OnBothOperands, Col3, IntPtr.Zero, IntPtr.Zero)
                            Dim Centru_cerc3 As New Point3d
                            Centru_cerc3 = Col3(0)
                            Pt7 = New Point3d(Centru_cerc3.X, Centru_cerc3.Y - R1, 0)
                            Pt8 = New Point3d(Pt7.X - Tangent, Pt7.Y, 0)

                        End If

                    End If

                    Dim Bulge1 As Double = Tan(0.25 * Angle1 * PI / 180)

                    Dim Layer_pipe As String = ComboBox_layer_CL.Text




                    Dim Poly_fillet As New Autodesk.AutoCAD.DatabaseServices.Polyline
                    Poly_fillet.Layer = Layer_pipe


                    Poly_fillet.AddVertexAt(0, New Autodesk.AutoCAD.Geometry.Point2d(Pt1.X, Pt1.Y), 0, 0, 0)
                    Poly_fillet.AddVertexAt(1, New Autodesk.AutoCAD.Geometry.Point2d(Pt2.X, Pt2.Y), Factor_stanga_dreapta * Bulge1, 0, 0)
                    Poly_fillet.AddVertexAt(2, New Autodesk.AutoCAD.Geometry.Point2d(Pt3.X, Pt3.Y), 0, 0, 0)
                    Poly_fillet.AddVertexAt(3, New Autodesk.AutoCAD.Geometry.Point2d(Pt4.X, Pt4.Y), 0, 0, 0)
                    Poly_fillet.AddVertexAt(4, New Autodesk.AutoCAD.Geometry.Point2d(Pt5.X, Pt5.Y), 0, 0, 0)
                    Poly_fillet.AddVertexAt(5, New Autodesk.AutoCAD.Geometry.Point2d(Pt6.X, Pt6.Y), -Factor_stanga_dreapta * Bulge1, 0, 0)
                    Poly_fillet.AddVertexAt(6, New Autodesk.AutoCAD.Geometry.Point2d(Pt7.X, Pt7.Y), 0, 0, 0)
                    Poly_fillet.AddVertexAt(7, New Autodesk.AutoCAD.Geometry.Point2d(Pt8.X, Pt8.Y), 0, 0, 0)

                    BTrecord.AppendEntity(Poly_fillet)
                    Trans1.AddNewlyCreatedDBObject(Poly_fillet, True)




                    If CheckBox_draw_od.Checked = True Then
                        Dim Layer_od As String = ComboBox_layer_OD.Text


                        Dim Number_string As String
                        Number_string = Replace(ComboBox_nps.Text, "NPS ", "")

                        If IsNumeric(Number_string) = True And IsNumeric(TextBox_diameter_X.Text) = True Then
                            Dim R_pipe As Double
                            Dim NR As Integer = CInt(Number_string)
                            R_pipe = Get_from_NPS_radius_for_pipes_from_inches_to_milimeters(NR) / 1000

                            Dim Dx As Double = CDbl(TextBox_diameter_X.Text)

                            Dim Obj_coll1 As DBObjectCollection = Poly_fillet.GetOffsetCurves(R_pipe)
                            Dim Obj_coll2 As DBObjectCollection = Poly_fillet.GetOffsetCurves(-R_pipe)
                            For Each acEnt As Entity In Obj_coll1
                                acEnt.Layer = Layer_od
                                '' Add each offset object
                                BTrecord.AppendEntity(acEnt)

                                Trans1.AddNewlyCreatedDBObject(acEnt, True)
                            Next
                            For Each acEnt As Entity In Obj_coll2
                                acEnt.Layer = Layer_od
                                '' Add each offset object
                                BTrecord.AppendEntity(acEnt)
                                Trans1.AddNewlyCreatedDBObject(acEnt, True)
                            Next

                            Dim Line_od_temp As New Autodesk.AutoCAD.DatabaseServices.Line(Pt1, Pt2)

                            Line_od_temp.TransformBy(Matrix3d.Rotation(PI / 2, Vector3d.ZAxis, Pt1))
                            Line_od_temp.TransformBy(Matrix3d.Scaling(1000, Pt1))


                            Dim Point_od1 As New Point3d
                            Point_od1 = Line_od_temp.GetPointAtDist(R_pipe)
                            Line_od_temp.TransformBy(Matrix3d.Rotation(PI, Vector3d.ZAxis, Pt1))

                            Dim Point_od2 As New Point3d
                            Point_od2 = Line_od_temp.GetPointAtDist(R_pipe)

                            Dim Line_od As New Line(Point_od1, Point_od2)
                            Line_od.Layer = Layer_od
                            Line_od.ColorIndex = 256
                            Line_od.Linetype = "BYLAYER"
                            Line_od.LineWeight = LineWeight.ByLayer
                            BTrecord.AppendEntity(Line_od)
                            Trans1.AddNewlyCreatedDBObject(Line_od, True)

                            Dim Line_od_temp1 As New Autodesk.AutoCAD.DatabaseServices.Line(Pt8, Pt7)

                            Line_od_temp1.TransformBy(Matrix3d.Rotation(PI / 2, Vector3d.ZAxis, Pt8))
                            Line_od_temp1.TransformBy(Matrix3d.Scaling(1000, Pt8))


                            Dim Point_od11 As New Point3d
                            Point_od11 = Line_od_temp1.GetPointAtDist(R_pipe)
                            Line_od_temp1.TransformBy(Matrix3d.Rotation(PI, Vector3d.ZAxis, Pt8))

                            Dim Point_od12 As New Point3d
                            Point_od12 = Line_od_temp1.GetPointAtDist(R_pipe)

                            Dim Line_od1 As New Line(Point_od11, Point_od12)
                            Line_od1.Layer = Layer_od
                            Line_od1.ColorIndex = 256
                            Line_od1.Linetype = "BYLAYER"
                            Line_od1.LineWeight = LineWeight.ByLayer
                            BTrecord.AppendEntity(Line_od1)
                            Trans1.AddNewlyCreatedDBObject(Line_od1, True)

                            Dim Line_od_temp2 As New Autodesk.AutoCAD.DatabaseServices.Line(Pt4, Pt5)
                            Line_od_temp2.TransformBy(Matrix3d.Rotation(PI / 2, Vector3d.ZAxis, Pt4))
                            Line_od_temp2.TransformBy(Matrix3d.Scaling(1000, Pt4))


                            Dim Point_od3 As New Point3d
                            Point_od3 = Line_od_temp2.GetPointAtDist(R_pipe)
                            Line_od_temp2.TransformBy(Matrix3d.Rotation(PI, Vector3d.ZAxis, Pt4))

                            Dim Point_od31 As New Point3d
                            Point_od31 = Line_od_temp2.GetPointAtDist(R_pipe)

                            Dim Line_od3 As New Line(Point_od3, Point_od31)
                            Line_od3.Layer = Layer_od
                            Line_od3.ColorIndex = 256
                            Line_od3.Linetype = "BYLAYER"
                            Line_od3.LineWeight = LineWeight.ByLayer
                            BTrecord.AppendEntity(Line_od3)
                            Trans1.AddNewlyCreatedDBObject(Line_od3, True)

                            Dim Line_od_temp3 As New Autodesk.AutoCAD.DatabaseServices.Line(Pt5, Pt4)
                            Line_od_temp3.TransformBy(Matrix3d.Rotation(PI / 2, Vector3d.ZAxis, Pt5))
                            Line_od_temp3.TransformBy(Matrix3d.Scaling(1000, Pt5))


                            Dim Point_od4 As New Point3d
                            Point_od4 = Line_od_temp3.GetPointAtDist(R_pipe)
                            Line_od_temp3.TransformBy(Matrix3d.Rotation(PI, Vector3d.ZAxis, Pt5))

                            Dim Point_od41 As New Point3d
                            Point_od41 = Line_od_temp3.GetPointAtDist(R_pipe)

                            Dim Line_od4 As New Line(Point_od4, Point_od41)
                            Line_od4.Layer = Layer_od
                            Line_od4.ColorIndex = 256
                            Line_od4.Linetype = "BYLAYER"
                            Line_od4.LineWeight = LineWeight.ByLayer
                            BTrecord.AppendEntity(Line_od4)
                            Trans1.AddNewlyCreatedDBObject(Line_od4, True)

                        End If
                    End If






                    Trans1.Commit()



                End Using

                Dim Empty_array() As ObjectId
                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
            End Using 'asta e de la lock1

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub ComboBox_nps_TabIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_nps.TextChanged, TextBox_diameter_X.TextChanged
        Try
            Dim Number_string As String
            Number_string = Replace(ComboBox_nps.Text, "NPS ", "")
            If IsNumeric(Number_string) = True And IsNumeric(TextBox_diameter_X.Text) = True Then
                Dim NR As Integer = CInt(Number_string)
                Dim Dx As Double = CDbl(TextBox_diameter_X.Text)
                TextBox_radius1.Text = Dx * 2 * Get_from_NPS_radius_for_pipes_from_inches_to_milimeters(NR) / 1000


                TextBox_report.Text = ComboBox_nps.Text & vbCrLf & "Diameter = " & 2 * Get_from_NPS_radius_for_pipes_from_inches_to_milimeters(NR) / 1000 & " m" _
                              & vbCrLf & "Radius = " & Get_from_NPS_radius_for_pipes_from_inches_to_milimeters(NR) / 1000 & " m"

            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Panel2_Click(sender As Object, e As EventArgs) Handles Panel2.Click
        Incarca_existing_layers_to_combobox(ComboBox_layer_CL)
        Incarca_existing_layers_to_combobox(ComboBox_layer_OD)
    End Sub
End Class