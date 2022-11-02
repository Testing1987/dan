Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class HDD_Form

    Dim Index1 As Integer = 666
    Dim OldCL_layer As String = ""
    Dim Old_layer_text As String = ""
    Private Sub HDD_Form_Load(sender As Object, e As System.EventArgs) Handles Me.Load, Me.Click
        TextBox_depth.Select()
        Incarca_existing_textstyles_to_combobox(ComboBox_text_styles)
        If ComboBox_text_styles.Items.Count > 0 Then ComboBox_text_styles.SelectedIndex = 0
        If ComboBox_text_styles.Items.Contains("ROMANS") = True Then
            ComboBox_text_styles.SelectedIndex = ComboBox_text_styles.Items.IndexOf("ROMANS")
        End If
        If ComboBox_text_styles.Items.Contains("Romans") = True Then
            ComboBox_text_styles.SelectedIndex = ComboBox_text_styles.Items.IndexOf("Romans")
        End If
        If ComboBox_text_styles.Items.Contains("romans") = True Then
            ComboBox_text_styles.SelectedIndex = ComboBox_text_styles.Items.IndexOf("romans")
        End If
        Incarca_existing_layers_to_combobox(ComboBox_layer_hdd_cl)
        Incarca_existing_layers_to_combobox(ComboBox_layer_text)
        If ComboBox_layer_hdd_cl.Items.Contains("PCENTRE") = True Then
            ComboBox_layer_hdd_cl.SelectedIndex = ComboBox_layer_hdd_cl.Items.IndexOf("PCENTRE")
        End If
        If ComboBox_layer_text.Items.Contains("TEXT") = True Then
            ComboBox_layer_text.SelectedIndex = ComboBox_layer_text.Items.IndexOf("TEXT")
        End If
    End Sub

    Private Sub Button_create_3d_curve_Click(sender As System.Object, e As System.EventArgs) Handles Button_3d_curve.Click
        Try

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Using Lock1 As Autodesk.AutoCAD.ApplicationServices.DocumentLock = ThisDrawing.LockDocument
                Dim Angle_left_3d, Angle_right_3d As Double


                If IsNumeric(TextBox_3d_angle_left.Text) = True Then
                    Angle_left_3d = CDbl(TextBox_3d_angle_left.Text)
                Else
                    MsgBox("Verify the 3d angle on left")
                    TextBox_3d_angle_left.Focus()
                    TextBox_3d_angle_left.Select()
                    Exit Sub
                End If

                If IsNumeric(TextBox_3d_angle_right.Text) = True Then
                    Angle_right_3d = CDbl(TextBox_3d_angle_right.Text)
                Else
                    MsgBox("Verify the 3d angle on right")
                    TextBox_3d_angle_right.Focus()
                    TextBox_3d_angle_right.Select()
                    Exit Sub
                End If

                If Angle_left_3d = 0 And Angle_right_3d = 0 Then
                    MsgBox("The 3d angles are 0 no compound 3D curve will be drafted")
                    Exit Sub
                End If


                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                Editor1 = ThisDrawing.Editor





                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction




                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord

                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)


                    Dim Rezultat_graph As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_graph As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_graph.MessageForAdding = vbLf & "Select the graph HDD polyline:"
                    Object_Prompt_graph.SingleOnly = True
                    Rezultat_graph = Editor1.GetSelection(Object_Prompt_graph)

                    If Rezultat_graph.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        MsgBox("Nothing has been drafted, please try again")
                        Exit Sub
                    End If


                    Dim Ent_POLY_poly_graph As Entity
                    Ent_POLY_poly_graph = Rezultat_graph.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim Poly_fillet As Polyline
                    If TypeOf Ent_POLY_poly_graph Is Polyline Then
                        Poly_fillet = Ent_POLY_poly_graph
                    Else
                        MsgBox("Your selection is not a HDD polyline, please try again")
                        Exit Sub
                    End If





                    Dim Poly_clone As Polyline

                    If Not TextBox_3d_angle_left.Text = 0 Then
                        If IsNumeric(TextBox_3d_angle_left.Text) = True Then
                            ThisDrawing.Editor.CurrentUserCoordinateSystem = WCS_align()
                            Poly_clone = New Polyline
                            Poly_clone = Poly_fillet.Clone
                            Poly_clone.TransformBy(Matrix3d.Rotation(Angle_left_3d * PI / 180, Vector3d.XAxis, Poly_fillet.GetPoint3dAt(2)))
                            Poly_clone.RemoveVertexAt(5)
                            Poly_clone.RemoveVertexAt(4)
                            Poly_clone.RemoveVertexAt(3)

                            Poly_clone.ColorIndex = 1
                            BTrecord.AppendEntity(Poly_clone)
                            Trans1.AddNewlyCreatedDBObject(Poly_clone, True)



                        End If
                    End If

                    If Not TextBox_3d_angle_right.Text = 0 Then
                        If IsNumeric(TextBox_3d_angle_right.Text) = True Then
                            ThisDrawing.Editor.CurrentUserCoordinateSystem = WCS_align()
                            Poly_clone = New Polyline
                            Poly_clone = Poly_fillet.Clone
                            Poly_clone.TransformBy(Matrix3d.Rotation(Angle_right_3d * PI / 180, Vector3d.XAxis, Poly_fillet.GetPoint3dAt(2)))
                            Poly_clone.RemoveVertexAt(0)
                            Poly_clone.RemoveVertexAt(0)
                            Poly_clone.RemoveVertexAt(0)

                            Poly_clone.ColorIndex = 1
                            BTrecord.AppendEntity(Poly_clone)
                            Trans1.AddNewlyCreatedDBObject(Poly_clone, True)
                        End If
                    End If




                    If Not TextBox_3d_angle_left.Text = 0 Then
                        If IsNumeric(TextBox_3d_angle_left.Text) = True Then
                            Dim Rezultat_PLAN As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                            Dim Object_Prompt_PLAN As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                            Object_Prompt_PLAN.MessageForAdding = vbLf & "Select the plan view polyline:"
                            Object_Prompt_PLAN.SingleOnly = True
                            Rezultat_PLAN = Editor1.GetSelection(Object_Prompt_PLAN)

                            If Rezultat_PLAN.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                MsgBox("Nothing has been drafted, please try again")
                                Exit Sub
                            End If


                            Dim Ent_POLY_PLAN As Entity
                            Ent_POLY_PLAN = Rezultat_PLAN.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                            Dim Curva As Curve
                            If TypeOf Ent_POLY_PLAN Is Curve Then
                                Curva = Ent_POLY_PLAN
                            Else
                                MsgBox("Your selection is not a polyline 2D/3D or line, please try again")
                                Exit Sub
                            End If


                            Dim Point_plan_start As Autodesk.AutoCAD.EditorInput.PromptPointResult
                            Dim PP_start As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify HDD right point position on the plan view:")
                            PP_start.AllowNone = False
                            PP_start.UseBasePoint = True
                            PP_start.BasePoint = Poly_fillet.EndPoint
                            Point_plan_start = Editor1.GetPoint(PP_start)
                            If Point_plan_start.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                MsgBox("Nothing has been drafted, please try again")
                                Exit Sub
                            End If




                            Dim Punct_pe_poly_end As New Point3d
                            Punct_pe_poly_end = Curva.GetClosestPointTo(Point_plan_start.Value, Vector3d.ZAxis, False)


                            Dim Point_plan_direction As Autodesk.AutoCAD.EditorInput.PromptPointResult
                            Dim PP_direction As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify HDD direction on the plan view:")
                            PP_direction.AllowNone = False
                            PP_direction.UseBasePoint = True
                            PP_direction.BasePoint = Point_plan_start.Value
                            Point_plan_direction = Editor1.GetPoint(PP_direction)
                            If Point_plan_direction.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                MsgBox("Nothing has been drafted, please try again")
                                Exit Sub
                            End If

                            Dim Punct_pe_poly_directie As New Point3d
                            Punct_pe_poly_directie = Curva.GetClosestPointTo(Point_plan_direction.Value, Vector3d.ZAxis, False)



                            Dim Planul_curent As Plane = New Plane(New Point3d(0, 0, 0), Editor1.CurrentUserCoordinateSystem.CoordinateSystem3d.Zaxis)

                            Dim Rotatie_in_xyPlane As Double = Punct_pe_poly_end.GetVectorTo(Punct_pe_poly_directie).AngleOnPlane(Planul_curent) + PI
                            Dim Vector_rotatie_90 As New Vector3d
                            Vector_rotatie_90 = Punct_pe_poly_end.GetVectorTo(Punct_pe_poly_directie)

                            Dim Poly_copy_2 As New Polyline
                            Poly_copy_2 = Poly_fillet.Clone
                            Poly_copy_2.TransformBy(Matrix3d.Displacement(Poly_fillet.EndPoint.GetVectorTo(Punct_pe_poly_end)))
                            Poly_copy_2.TransformBy(Matrix3d.Rotation(Rotatie_in_xyPlane, Vector3d.ZAxis, Punct_pe_poly_end))
                            Poly_copy_2.TransformBy(Matrix3d.Rotation(-PI / 2, Vector_rotatie_90, Punct_pe_poly_end))
                            BTrecord.AppendEntity(Poly_copy_2)
                            Trans1.AddNewlyCreatedDBObject(Poly_copy_2, True)

                            Dim Poly_copy_1 As New Polyline
                            Poly_copy_1 = Poly_clone.Clone
                            Poly_copy_1.TransformBy(Matrix3d.Displacement(Poly_copy_1.EndPoint.GetVectorTo(Poly_copy_2.GetPoint3dAt(2))))
                            Poly_copy_1.TransformBy(Matrix3d.Rotation(Rotatie_in_xyPlane, Vector3d.ZAxis, Poly_copy_2.GetPoint3dAt(2)))
                            Poly_copy_1.TransformBy(Matrix3d.Rotation(-PI / 2, Vector_rotatie_90, Poly_copy_2.GetPoint3dAt(2)))

                            BTrecord.AppendEntity(Poly_copy_1)
                            Trans1.AddNewlyCreatedDBObject(Poly_copy_1, True)





                        End If
                    End If

                    If Not TextBox_3d_angle_right.Text = 0 Then
                        If IsNumeric(TextBox_3d_angle_right.Text) = True Then



                            Dim Rezultat_PLAN As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                            Dim Object_Prompt_PLAN As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                            Object_Prompt_PLAN.MessageForAdding = vbLf & "Select the plan view polyline:"
                            Object_Prompt_PLAN.SingleOnly = True
                            Rezultat_PLAN = Editor1.GetSelection(Object_Prompt_PLAN)

                            If Rezultat_PLAN.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                MsgBox("Nothing has been drafted, please try again")
                                Exit Sub
                            End If


                            Dim Ent_POLY_PLAN As Entity
                            Ent_POLY_PLAN = Rezultat_PLAN.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                            Dim Curva As Curve
                            If TypeOf Ent_POLY_PLAN Is Curve Then
                                Curva = Ent_POLY_PLAN
                            Else
                                MsgBox("Your selection is not a polyline 2D/3D or line, please try again")
                                Exit Sub
                            End If


                            Dim Point_plan_start As Autodesk.AutoCAD.EditorInput.PromptPointResult
                            Dim PP_start As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify HDD left point position on the plan view:")
                            PP_start.AllowNone = False
                            PP_start.UseBasePoint = True
                            PP_start.BasePoint = Poly_fillet.StartPoint
                            Point_plan_start = Editor1.GetPoint(PP_start)
                            If Point_plan_start.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                MsgBox("Nothing has been drafted, please try again")
                                Exit Sub
                            End If




                            Dim Punct_pe_poly_start As New Point3d
                            Punct_pe_poly_start = Curva.GetClosestPointTo(Point_plan_start.Value, Vector3d.ZAxis, False)


                            Dim Point_plan_direction As Autodesk.AutoCAD.EditorInput.PromptPointResult
                            Dim PP_direction As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify HDD direction:")
                            PP_direction.AllowNone = False
                            PP_direction.UseBasePoint = True
                            PP_direction.BasePoint = Point_plan_start.Value
                            Point_plan_direction = Editor1.GetPoint(PP_direction)
                            If Point_plan_direction.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                MsgBox("Nothing has been drafted, please try again")
                                Exit Sub
                            End If

                            Dim Punct_pe_poly_directie As New Point3d
                            Punct_pe_poly_directie = Curva.GetClosestPointTo(Point_plan_direction.Value, Vector3d.ZAxis, False)



                            Dim Planul_curent As Plane = New Plane(New Point3d(0, 0, 0), Editor1.CurrentUserCoordinateSystem.CoordinateSystem3d.Zaxis)

                            Dim Rotatie_in_xyPlane As Double = Punct_pe_poly_start.GetVectorTo(Punct_pe_poly_directie).AngleOnPlane(Planul_curent)
                            Dim Vector_rotatie_90 As New Vector3d
                            Vector_rotatie_90 = Punct_pe_poly_start.GetVectorTo(Punct_pe_poly_directie)


                            Dim Poly_copy_1 As New Polyline
                            Poly_copy_1 = Poly_fillet.Clone
                            Poly_copy_1.TransformBy(Matrix3d.Displacement(Poly_copy_1.StartPoint.GetVectorTo(Punct_pe_poly_start)))
                            Poly_copy_1.TransformBy(Matrix3d.Rotation(Rotatie_in_xyPlane, Vector3d.ZAxis, Punct_pe_poly_start))

                            Poly_copy_1.TransformBy(Matrix3d.Rotation(PI / 2, Vector_rotatie_90, Punct_pe_poly_start))

                            BTrecord.AppendEntity(Poly_copy_1)
                            Trans1.AddNewlyCreatedDBObject(Poly_copy_1, True)

                            Dim Poly_copy_2 As New Polyline
                            Poly_copy_2 = Poly_clone.Clone
                            Poly_copy_2.TransformBy(Matrix3d.Displacement(Poly_copy_2.StartPoint.GetVectorTo(Poly_copy_1.GetPoint3dAt(3))))
                            Poly_copy_2.TransformBy(Matrix3d.Rotation(Rotatie_in_xyPlane, Vector3d.ZAxis, Poly_copy_1.GetPoint3dAt(3)))
                            Poly_copy_2.TransformBy(Matrix3d.Rotation(PI / 2, Vector_rotatie_90, Poly_copy_1.GetPoint3dAt(3)))
                            BTrecord.AppendEntity(Poly_copy_2)
                            Trans1.AddNewlyCreatedDBObject(Poly_copy_2, True)

                        End If
                    End If










                    Editor1.Regen()
                    Trans1.Commit()




                End Using







                Dim Empty_array() As ObjectId
                Editor1.SetImpliedSelection(Empty_array)
            End Using 'asta e de la lock1
        Catch ex As Exception

            'Exit Sub
            MsgBox(ex.Message)
        End Try
    End Sub



    Private Sub Button_create_hdd_Click(sender As System.Object, e As System.EventArgs) Handles Button_create_hdd.Click
        Try

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Using Lock1 As Autodesk.AutoCAD.ApplicationServices.DocumentLock = ThisDrawing.LockDocument
                Dim Depth, Angle_left, Angle_right, Radius_left, Radius_right, Straight_portion As Double
                If IsNumeric(TextBox_depth.Text) = True Then
                    Depth = CDbl(TextBox_depth.Text)
                Else
                    MsgBox("Non Numeric Depth")
                    TextBox_depth.Focus()
                    TextBox_depth.Select()
                    Exit Sub
                End If

                If IsNumeric(TextBox_angle_left.Text) = True Then
                    Angle_left = CDbl(TextBox_angle_left.Text)
                Else
                    MsgBox("Verify the Left angle")
                    TextBox_angle_left.Focus()
                    TextBox_angle_left.Select()
                    Exit Sub
                End If

                If IsNumeric(TextBox_angle_right.Text) = True Then
                    Angle_right = CDbl(TextBox_angle_right.Text)
                Else
                    MsgBox("Verify the Right angle")
                    TextBox_angle_right.Focus()
                    TextBox_angle_right.Select()
                    Exit Sub
                End If

                If IsNumeric(TextBox_radius_Left.Text) = True Then
                    Radius_left = CDbl(TextBox_radius_Left.Text)
                Else
                    MsgBox("Verify the Left Radius")
                    TextBox_radius_Left.Focus()
                    TextBox_radius_Left.Select()
                    Exit Sub
                End If

                If IsNumeric(TextBox_Radius_right.Text) = True Then
                    Radius_right = CDbl(TextBox_Radius_right.Text)
                Else
                    MsgBox("Verify the right Radius")
                    TextBox_Radius_right.Focus()
                    TextBox_Radius_right.Select()
                    Exit Sub
                End If

                If IsNumeric(TextBox_straight_length.Text) = True Then
                    Straight_portion = CDbl(TextBox_straight_length.Text)
                Else
                    MsgBox("Verify the Straight segment")
                    TextBox_straight_length.Focus()
                    TextBox_straight_length.Select()
                    Exit Sub
                End If


                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                Editor1 = ThisDrawing.Editor

                If CheckBox_pick_start_end.Checked = False Then
                    Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult

                    Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions("Specify 0+000 position:" & vbCrLf)
                    PP1.AllowNone = False
                    Point1 = Editor1.GetPoint(PP1)
                    If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                        Exit Sub
                    End If

                    Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt2.MessageForAdding = "Select the ground polyline:"
                    Object_Prompt2.SingleOnly = True
                    Rezultat2 = Editor1.GetSelection(Object_Prompt2)

                    If Rezultat2.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Exit Sub
                    End If

                    If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK And Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK And IsNothing(Rezultat2) = False Then

                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat2.Value.Item(0)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then

                                Dim Poly1 As Polyline = Ent1
                                Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord

                                BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                                BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)


                                Dim x0, y0, x01, x02, x03, x04, y01, y02, y03, y04, X11, X12, X13, X14, Y11, Y12, Y13, Y14 As Double
                                Dim x0T, y0T, x1T, y1T As Double


                                x0 = Point1.Value.X
                                y0 = Point1.Value.Y - Depth
                                x01 = x0 - Straight_portion / 2
                                y01 = y0
                                y02 = y01 + Radius_left
                                x02 = x01
                                y03 = y02 + Radius_left * Sin(3 * PI / 2 - Angle_left * PI / 180)
                                x03 = x02 + Radius_left * Cos(3 * PI / 2 - Angle_left * PI / 180)
                                y0T = y03 + 10 * Sin(PI - Angle_left * PI / 180)
                                x0T = x03 + 10 * Cos(PI - Angle_left * PI / 180)

                                X11 = x0 + Straight_portion / 2
                                Y11 = y0
                                Y12 = Y11 + Radius_right
                                X12 = X11
                                Y13 = Y12 - Radius_right * Cos(Angle_right * PI / 180)
                                X13 = X12 + Radius_right * Sin(Angle_right * PI / 180)
                                y1T = Y13 + 10 * Sin(Angle_right * PI / 180)
                                x1T = X13 + 10 * Cos(Angle_right * PI / 180)

                                Dim Line0 As New Autodesk.AutoCAD.DatabaseServices.Line
                                Line0.StartPoint = New Point3d(x03, y03, 0)
                                Line0.EndPoint = New Point3d(x0T, y0T, 0)

                                Dim Point01 As New Point3dCollection
                                Line0.IntersectWith(Poly1, Intersect.ExtendBoth, Point01, IntPtr.Zero, IntPtr.Zero)
                                If Point01.Count > 0 Then
                                    x04 = Point01(0).X
                                    y04 = Point01(0).Y
                                Else
                                    MsgBox("The HDD profile doesn't intersect the ground profile on left" & vbCrLf & "Please verify your parameters")
                                    Exit Sub
                                End If

                                Dim Line1 As New Autodesk.AutoCAD.DatabaseServices.Line
                                Line1.StartPoint = New Point3d(X13, Y13, 0)
                                Line1.EndPoint = New Point3d(x1T, y1T, 0)

                                Dim Point11 As New Point3dCollection
                                Line1.IntersectWith(Poly1, Intersect.ExtendBoth, Point11, IntPtr.Zero, IntPtr.Zero)
                                If Point11.Count > 0 Then
                                    X14 = Point11(0).X
                                    Y14 = Point11(0).Y
                                Else
                                    MsgBox("The HDD profile doesn't intersect the ground profile on right" & vbCrLf & "Please verify your parameters")
                                    Exit Sub
                                End If







                                'BULGE CALCS
                                ' BULGE = TAN(Unghi/4)


                                Dim Bulge0 As Double = Tan((Angle_left * PI / 180) / 4)
                                Dim Bulge1 As Double = Tan((Angle_right * PI / 180) / 4)




                                Dim Poly_fillet As New Autodesk.AutoCAD.DatabaseServices.Polyline
                                Poly_fillet.Layer = ComboBox_layer_hdd_cl.Text


                                Poly_fillet.AddVertexAt(0, New Autodesk.AutoCAD.Geometry.Point2d(x04, y04), 0, 0, 0)
                                Poly_fillet.AddVertexAt(1, New Autodesk.AutoCAD.Geometry.Point2d(x03, y03), Bulge0, 0, 0)
                                Poly_fillet.AddVertexAt(2, New Autodesk.AutoCAD.Geometry.Point2d(x01, y01), 0, 0, 0)
                                Poly_fillet.AddVertexAt(3, New Autodesk.AutoCAD.Geometry.Point2d(X11, Y11), Bulge1, 0, 0)
                                Poly_fillet.AddVertexAt(4, New Autodesk.AutoCAD.Geometry.Point2d(X13, Y13), 0, 0, 0)
                                Poly_fillet.AddVertexAt(5, New Autodesk.AutoCAD.Geometry.Point2d(X14, Y14), 0, 0, 0)

                                BTrecord.AppendEntity(Poly_fillet)
                                Trans1.AddNewlyCreatedDBObject(Poly_fillet, True)

                                Dim Text_style_table As Autodesk.AutoCAD.DatabaseServices.TextStyleTable = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                If Text_style_table.Has(ComboBox_text_styles.Text) = True Then
                                    Dim Text25id As Autodesk.AutoCAD.DatabaseServices.ObjectId = Text_style_table.Item(ComboBox_text_styles.Text)




                                    Dim Nr_zecimale As Integer
                                    If CheckBox_0DEC.Checked = True Then Nr_zecimale = 0
                                    If CheckBox_1DEC.Checked = True Then Nr_zecimale = 1

                                    Dim Tang_length_stanga As Double = Round(GET_distanta_Double_XY(x04, y04, x03, y03), Nr_zecimale)
                                    Dim Mtext_tang_length_stanga As New MText
                                    Mtext_tang_length_stanga.TextHeight = 2.5
                                    Mtext_tang_length_stanga.Contents = "TANGENT LENGTH " & Get_String_Rounded(Tang_length_stanga, Nr_zecimale) & "m"
                                    Mtext_tang_length_stanga.TextStyleId = Text25id
                                    Mtext_tang_length_stanga.Layer = ComboBox_layer_text.Text
                                    Mtext_tang_length_stanga.Location = New Point3d((x04 + x03) / 2, (y03 + y04) / 2, 0)
                                    BTrecord.AppendEntity(Mtext_tang_length_stanga)
                                    Trans1.AddNewlyCreatedDBObject(Mtext_tang_length_stanga, True)





                                    Dim Mtext1 As New MText
                                    Mtext1.Layer = ComboBox_layer_text.Text
                                    Mtext1.TextStyleId = Text25id
                                    Mtext1.TextHeight = 2.5
                                    Dim Arc_length1 As Double = Radius_left * Angle_left * PI / 180
                                    Mtext1.Contents = "ARC LENGTH " & Get_String_Rounded(Arc_length1, Nr_zecimale) & "m" & vbCrLf & "ARC RADIUS " & Radius_left & "m"
                                    Mtext1.Attachment = AttachmentPoint.TopRight
                                    Mtext1.Location = New Point3d((x03 + x01) / 2, (y03 + y01) / 2, 0)
                                    BTrecord.AppendEntity(Mtext1)
                                    Trans1.AddNewlyCreatedDBObject(Mtext1, True)




                                    Dim Tang_length_dreapta As Double = Round(GET_distanta_Double_XY(X14, Y14, X13, Y13), Nr_zecimale)

                                    Dim Mtext_tang_length_dreapta As New MText
                                    Mtext_tang_length_dreapta.TextHeight = 2.5
                                    Mtext_tang_length_dreapta.Contents = "TANGENT LENGTH " & Get_String_Rounded(Tang_length_dreapta, Nr_zecimale) & "m"
                                    Mtext_tang_length_dreapta.TextStyleId = Text25id
                                    Mtext_tang_length_dreapta.Location = New Point3d((X14 + X13) / 2, (Y13 + Y14) / 2, 0)
                                    Mtext_tang_length_dreapta.Layer = ComboBox_layer_text.Text
                                    BTrecord.AppendEntity(Mtext_tang_length_dreapta)
                                    Trans1.AddNewlyCreatedDBObject(Mtext_tang_length_dreapta, True)

                                    Dim Mtext2 As New MText
                                    Mtext2.Layer = ComboBox_layer_text.Text
                                    Mtext2.TextStyleId = Text25id
                                    Mtext2.TextHeight = 2.5
                                    Dim Arc_length2 As Double = Radius_right * Angle_right * PI / 180
                                    Mtext2.Contents = "ARC LENGTH " & Get_String_Rounded(Arc_length2, Nr_zecimale) & "m" & vbCrLf & "ARC RADIUS " & Radius_right & "m"
                                    Mtext2.Location = New Point3d((X13 + X11) / 2, (Y13 + Y11) / 2, 0)
                                    Mtext2.Attachment = AttachmentPoint.TopLeft
                                    BTrecord.AppendEntity(Mtext2)
                                    Trans1.AddNewlyCreatedDBObject(Mtext2, True)


                                    Dim Mtext3 As New MText
                                    Mtext3.Layer = ComboBox_layer_text.Text
                                    Mtext3.TextStyleId = Text25id
                                    Mtext3.TextHeight = 2.5
                                    Dim Straight_length As Double = GET_distanta_Double_XY(x01, y01, X11, Y11)
                                    Mtext3.Contents = "TANGENT LENGTH " & Get_String_Rounded(Straight_length, Nr_zecimale) & "m"
                                    Mtext3.Attachment = AttachmentPoint.TopCenter
                                    Mtext3.Location = New Point3d((x01 + X11) / 2, (y01 + Y11) / 2, 0)
                                    BTrecord.AppendEntity(Mtext3)
                                    Trans1.AddNewlyCreatedDBObject(Mtext3, True)

                                    Dim Entry_angle_string_left As String
                                    Dim Entry_angle_string_dreapta As String

                                    If y04 < Y14 Then
                                        Entry_angle_string_left = "ENTRY ANGLE "
                                        Entry_angle_string_dreapta = "EXIT ANGLE "
                                    Else
                                        Entry_angle_string_left = "EXIT ANGLE "
                                        Entry_angle_string_dreapta = "ENTRY ANGLE "
                                    End If


                                    Dim Mtext_left As New MText
                                    Mtext_left.Layer = ComboBox_layer_text.Text
                                    Mtext_left.TextStyleId = Text25id
                                    Mtext_left.TextHeight = 2.5
                                    Mtext_left.Contents = Entry_angle_string_left & Angle_left & Chr(176)
                                    Mtext_left.Attachment = AttachmentPoint.BottomRight
                                    Mtext_left.Location = New Point3d(x04 - 5, y04 + 5, 0)
                                    BTrecord.AppendEntity(Mtext_left)
                                    Trans1.AddNewlyCreatedDBObject(Mtext_left, True)

                                    Dim Mtext_right As New MText
                                    Mtext_right.Layer = ComboBox_layer_text.Text
                                    Mtext_right.TextStyleId = Text25id
                                    Mtext_right.TextHeight = 2.5
                                    Mtext_right.Contents = Entry_angle_string_dreapta & Angle_right & Chr(176)
                                    Mtext_right.Attachment = AttachmentPoint.BottomLeft
                                    Mtext_right.Location = New Point3d(X14 + 5, Y14 + 5, 0)
                                    BTrecord.AppendEntity(Mtext_right)
                                    Trans1.AddNewlyCreatedDBObject(Mtext_right, True)



                                    Dim Mtext_Total_length As New MText
                                    Mtext_Total_length.Layer = ComboBox_layer_text.Text
                                    Mtext_Total_length.TextStyleId = Text25id
                                    Mtext_Total_length.TextHeight = 2.5
                                    Mtext_Total_length.Contents = "TOTAL HDD PROFILE LENGTH " & Get_String_Rounded(Tang_length_stanga + Round(Arc_length1, Nr_zecimale) + Tang_length_dreapta + Round(Arc_length2, Nr_zecimale) + Round(Straight_length, Nr_zecimale), Nr_zecimale) & "m"
                                    Mtext_Total_length.Attachment = AttachmentPoint.BottomLeft
                                    Mtext_Total_length.Location = New Point3d(x01, y01 - 10, 0)
                                    BTrecord.AppendEntity(Mtext_Total_length)
                                    Trans1.AddNewlyCreatedDBObject(Mtext_Total_length, True)



                                Else
                                    MsgBox("no such a text style in your drawing")

                                    Exit Sub
                                End If ' ASTA E DE LA If Text_style_table.Has(ComboBox_text_styles.text) = True Then




                            End If ' ASTA E DE LA If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Line And TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Line



                            Editor1.Regen()
                            Trans1.Commit()




                        End Using


                    End If ' asta e de la Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK
                End If 'ASTA E DE LA If CheckBox_pick_start_end.Checked = FALSE

                If CheckBox_pick_start_end.Checked = True Then
                    Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult

                    Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions("Specify entry point:" & vbCrLf)
                    PP1.AllowNone = False
                    Point1 = Editor1.GetPoint(PP1)
                    If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                        Exit Sub
                    End If


                    Dim Point2 As Autodesk.AutoCAD.EditorInput.PromptPointResult

                    Dim PP2 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions("Specify exit point:" & vbCrLf)
                    PP2.AllowNone = False
                    Point2 = Editor1.GetPoint(PP2)
                    If Point2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                        Exit Sub
                    End If


                    Dim Point3 As Autodesk.AutoCAD.EditorInput.PromptPointResult

                    Dim PP3 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions("Specify the point for applying the depth:" & vbCrLf)
                    PP3.AllowNone = False
                    Point3 = Editor1.GetPoint(PP3)
                    If Point3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                        Exit Sub
                    End If



                    If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK And Point2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK And Point3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord

                            BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)


                            Dim x1, x2, x4, y1, y2, y4 As Double



                            x1 = Point1.Value.X
                            y1 = Point1.Value.Y
                            x2 = Point2.Value.X
                            y2 = Point2.Value.Y

                            Dim x1T, y1T As Double
                            x1T = x1 + 10 * Cos(Angle_left * PI / 180)
                            y1T = y1 - 10 * Sin(Angle_left * PI / 180)

                            Dim x2T, y2T As Double
                            x2T = x2 - 10 * Cos(Angle_right * PI / 180)
                            y2T = y2 - 10 * Sin(Angle_right * PI / 180)

                            Dim Line1 As New Line
                            Line1.StartPoint = New Point3d(x1, y1, 0)
                            Line1.EndPoint = New Point3d(x1T, y1T, 0)

                            Dim Line2 As New Line
                            Line2.StartPoint = New Point3d(x2, y2, 0)
                            Line2.EndPoint = New Point3d(x2T, y2T, 0)


                            Dim Point01 As New Point3dCollection
                            Line1.IntersectWith(Line2, Intersect.ExtendBoth, Point01, IntPtr.Zero, IntPtr.Zero)
                            x4 = Point01(0).X
                            y4 = Point01(0).Y


                            Dim x11, y11, x22, y22, x5, y5 As Double
                            x5 = Point3.Value.X
                            y5 = Point3.Value.Y - Depth

                            Dim x3T, y3T As Double
                            x3T = x5 - 10
                            y3T = y5

                            Dim Linie_dreapta As New Line
                            Linie_dreapta.StartPoint = New Point3d(x5, y5, 0)
                            Linie_dreapta.EndPoint = New Point3d(x3T, y3T, 0)

                            Dim Line14 As New Line
                            Line14.StartPoint = New Point3d(x1, y1, 0)
                            Line14.EndPoint = New Point3d(x4, y4, 0)




                            Dim Point02 As New Point3dCollection
                            Linie_dreapta.IntersectWith(Line14, Intersect.ExtendBoth, Point02, IntPtr.Zero, IntPtr.Zero)
                            x11 = Point02(0).X
                            y11 = Point02(0).Y

                            Dim Line24 As New Line
                            Line24.StartPoint = New Point3d(x2, y2, 0)
                            Line24.EndPoint = New Point3d(x4, y4, 0)



                            Dim Point03 As New Point3dCollection
                            Linie_dreapta.IntersectWith(Line24, Intersect.ExtendBoth, Point03, IntPtr.Zero, IntPtr.Zero)
                            x22 = Point03(0).X
                            y22 = Point03(0).Y


                            Dim D1 As Double = Radius_left * Tan(0.5 * Angle_left * PI / 180)

                            Dim Xc11, Yc11 As Double
                            Xc11 = x11 + D1
                            Yc11 = y11 + Radius_left



                            Dim XO1, YO1 As Double
                            XO1 = x11 - D1 * Cos(Angle_left * PI / 180)
                            YO1 = y11 + D1 * Sin(Angle_left * PI / 180)


                            Dim Xs1, Ys1 As Double
                            Xs1 = Xc11
                            Ys1 = y11

                            Dim D2 As Double = Radius_right * Tan(0.5 * Angle_right * PI / 180)

                            Dim Xc22, Yc22 As Double
                            Xc22 = x22 - D2
                            Yc22 = y22 + Radius_right



                            Dim XO2, YO2 As Double
                            XO2 = x22 + D2 * Cos(Angle_right * PI / 180)
                            YO2 = y22 + D2 * Sin(Angle_right * PI / 180)

                            Dim Xs2, Ys2 As Double
                            Xs2 = Xc22
                            Ys2 = y22


                            'BULGE CALCS
                            ' BULGE = TAN(Unghi/4)


                            Dim Bulge0 As Double = Tan((Angle_left * PI / 180) / 4)
                            Dim Bulge1 As Double = Tan((Angle_right * PI / 180) / 4)





                            Dim Poly_fillet As New Autodesk.AutoCAD.DatabaseServices.Polyline
                            Poly_fillet.Layer = ComboBox_layer_hdd_cl.Text


                            Poly_fillet.AddVertexAt(0, New Autodesk.AutoCAD.Geometry.Point2d(x1, y1), 0, 0, 0)
                            Poly_fillet.AddVertexAt(1, New Autodesk.AutoCAD.Geometry.Point2d(XO1, YO1), Bulge0, 0, 0)
                            Poly_fillet.AddVertexAt(2, New Autodesk.AutoCAD.Geometry.Point2d(Xs1, Ys1), 0, 0, 0)
                            Poly_fillet.AddVertexAt(3, New Autodesk.AutoCAD.Geometry.Point2d(Xs2, Ys2), Bulge1, 0, 0)
                            Poly_fillet.AddVertexAt(4, New Autodesk.AutoCAD.Geometry.Point2d(XO2, YO2), 0, 0, 0)
                            Poly_fillet.AddVertexAt(5, New Autodesk.AutoCAD.Geometry.Point2d(x2, y2), 0, 0, 0)
                            BTrecord.AppendEntity(Poly_fillet)
                            Trans1.AddNewlyCreatedDBObject(Poly_fillet, True)


                            Dim Text_style_table As Autodesk.AutoCAD.DatabaseServices.TextStyleTable = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                            If Text_style_table.Has(ComboBox_text_styles.Text) = True Then
                                Dim Text25id As Autodesk.AutoCAD.DatabaseServices.ObjectId = Text_style_table.Item(ComboBox_text_styles.Text)

                                Dim Nr_zecimale As Integer
                                If CheckBox_0DEC.Checked = True Then Nr_zecimale = 0
                                If CheckBox_1DEC.Checked = True Then Nr_zecimale = 1
                                Dim Tang_length_stanga As Double = Round(GET_distanta_Double_XY(x1, y1, XO1, YO1), Nr_zecimale)
                                Dim Mtext_tang_length_stanga As New MText
                                Mtext_tang_length_stanga.TextHeight = 2.5
                                Mtext_tang_length_stanga.Contents = "TANGENT LENGTH " & Get_String_Rounded(Tang_length_stanga, Nr_zecimale) & "m"
                                Mtext_tang_length_stanga.TextStyleId = Text25id
                                Mtext_tang_length_stanga.Layer = ComboBox_layer_text.Text
                                Mtext_tang_length_stanga.Location = New Point3d((x1 + XO1) / 2, (YO1 + y1) / 2, 0)
                                BTrecord.AppendEntity(Mtext_tang_length_stanga)
                                Trans1.AddNewlyCreatedDBObject(Mtext_tang_length_stanga, True)

                                Dim Mtext1 As New MText
                                Mtext1.Layer = ComboBox_layer_text.Text
                                Mtext1.TextStyleId = Text25id
                                Mtext1.TextHeight = 2.5
                                Dim Arc_length1 As Double = Radius_left * Angle_left * PI / 180
                                Mtext1.Contents = "ARC LENGTH " & Get_String_Rounded(Arc_length1, Nr_zecimale) & "m" & vbCrLf & "ARC RADIUS " & Radius_left & "m"
                                Mtext1.Location = New Point3d((XO1 + Xs1) / 2, (YO1 + Ys1) / 2, 0)
                                Mtext1.Attachment = AttachmentPoint.TopRight
                                BTrecord.AppendEntity(Mtext1)
                                Trans1.AddNewlyCreatedDBObject(Mtext1, True)


                                Dim Tang_length_dreapta As Double = Round(GET_distanta_Double_XY(XO2, YO2, x2, y2), Nr_zecimale)
                                Dim Mtext_tang_length_dreapta As New MText
                                Mtext_tang_length_dreapta.TextHeight = 2.5
                                Mtext_tang_length_dreapta.Contents = "TANGENT LENGTH " & Get_String_Rounded(Tang_length_dreapta, Nr_zecimale) & "m"
                                Mtext_tang_length_dreapta.TextStyleId = Text25id
                                Mtext_tang_length_dreapta.Location = New Point3d((XO2 + x2) / 2, (y2 + YO2) / 2, 0)
                                Mtext_tang_length_dreapta.Layer = ComboBox_layer_text.Text
                                BTrecord.AppendEntity(Mtext_tang_length_dreapta)
                                Trans1.AddNewlyCreatedDBObject(Mtext_tang_length_dreapta, True)

                                Dim Mtext2 As New MText
                                Mtext2.Layer = ComboBox_layer_text.Text
                                Mtext2.TextStyleId = Text25id
                                Mtext2.TextHeight = 2.5
                                Dim Arc_length2 As Double = Radius_right * Angle_right * PI / 180
                                Mtext2.Contents = "ARC LENGTH " & Get_String_Rounded(Arc_length2, Nr_zecimale) & "m" & vbCrLf & "ARC RADIUS " & Radius_right & "m"
                                Mtext2.Location = New Point3d((Xs2 + XO2) / 2, (Ys2 + YO2) / 2, 0)
                                Mtext2.Attachment = AttachmentPoint.TopLeft
                                BTrecord.AppendEntity(Mtext2)
                                Trans1.AddNewlyCreatedDBObject(Mtext2, True)

                                Dim Mtext3 As New MText
                                Mtext3.Layer = ComboBox_layer_text.Text
                                Mtext3.TextStyleId = Text25id
                                Mtext3.TextHeight = 2.5
                                Dim Straight_length As Double = GET_distanta_Double_XY(Xs1, Ys1, Xs2, Ys2)
                                Mtext3.Contents = "TANGENT LENGTH " & Get_String_Rounded(Straight_length, Nr_zecimale) & "m"
                                Mtext3.Attachment = AttachmentPoint.TopCenter
                                Mtext3.Location = New Point3d((Xs1 + Xs2) / 2, (Ys1 + Ys2) / 2, 0)
                                BTrecord.AppendEntity(Mtext3)
                                Trans1.AddNewlyCreatedDBObject(Mtext3, True)
                                Dim Entry_angle_string_left As String
                                Dim Entry_angle_string_dreapta As String

                                If y1 < y2 Then
                                    Entry_angle_string_left = "ENTRY ANGLE "
                                    Entry_angle_string_dreapta = "EXIT ANGLE "
                                Else
                                    Entry_angle_string_left = "EXIT ANGLE "
                                    Entry_angle_string_dreapta = "ENTRY ANGLE "
                                End If

                                Dim Mtext_left As New MText
                                Mtext_left.Layer = ComboBox_layer_text.Text
                                Mtext_left.TextStyleId = Text25id
                                Mtext_left.TextHeight = 2.5
                                Mtext_left.Contents = Entry_angle_string_left & Angle_left & Chr(176)
                                Mtext_left.Attachment = AttachmentPoint.BottomRight
                                Mtext_left.Location = New Point3d(x1 - 5, y1 + 5, 0)
                                BTrecord.AppendEntity(Mtext_left)
                                Trans1.AddNewlyCreatedDBObject(Mtext_left, True)

                                Dim Mtext_right As New MText
                                Mtext_right.Layer = ComboBox_layer_text.Text
                                Mtext_right.TextStyleId = Text25id
                                Mtext_right.TextHeight = 2.5
                                Mtext_right.Contents = Entry_angle_string_dreapta & Angle_right & Chr(176)
                                Mtext_right.Attachment = AttachmentPoint.BottomLeft
                                Mtext_right.Location = New Point3d(x2 + 5, y2 + 5, 0)
                                BTrecord.AppendEntity(Mtext_right)
                                Trans1.AddNewlyCreatedDBObject(Mtext_right, True)

                                Dim Mtext_Total_length As New MText
                                Mtext_Total_length.Layer = ComboBox_layer_text.Text
                                Mtext_Total_length.TextStyleId = Text25id
                                Mtext_Total_length.TextHeight = 2.5
                                Mtext_Total_length.Contents = "TOTAL HDD PROFILE LENGTH " & Get_String_Rounded(Tang_length_stanga + Round(Arc_length1, Nr_zecimale) + Tang_length_dreapta + Round(Arc_length2, Nr_zecimale) + Round(Straight_length, Nr_zecimale), Nr_zecimale) & "m"
                                Mtext_Total_length.Attachment = AttachmentPoint.BottomLeft
                                Mtext_Total_length.Location = New Point3d(Xs1, Ys1 - 10, 0)
                                BTrecord.AppendEntity(Mtext_Total_length)
                                Trans1.AddNewlyCreatedDBObject(Mtext_Total_length, True)


                            End If ' asta e de la If Text_style_table.Has(ComboBox_text_styles.text) = True 
                            Editor1.Regen()
                            Trans1.Commit()




                        End Using


                    End If 'asta e de la  If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK And Point2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK And Point3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK
                End If 'ASTA E DE LA If CheckBox_pick_start_end.Checked = TRUE


                Dim Empty_array() As ObjectId
                Editor1.SetImpliedSelection(Empty_array)
            End Using 'asta e de la lock1
        Catch ex As Exception

            'Exit Sub
            MsgBox(ex.Message)
        End Try
    End Sub



    Private Sub CheckBox_pick_start_end_CheckedChanged(sender As Object, e As System.EventArgs) Handles CheckBox_pick_start_end.CheckedChanged
        If CheckBox_pick_start_end.Checked = True Then
            Panel_straight_portion.Visible = False
            Label_3d_angle_left.Visible = False
            Label_3d_angle_right.Visible = False
            TextBox_3d_angle_left.Visible = False
            TextBox_3d_angle_right.Visible = False
        Else
            Panel_straight_portion.Visible = True
            Label_3d_angle_left.Visible = True
            Label_3d_angle_right.Visible = True
            TextBox_3d_angle_left.Visible = True
            TextBox_3d_angle_right.Visible = True
        End If
    End Sub



    Private Sub TextBox_depth_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox_depth.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            TextBox_angle_left.SelectAll()
            TextBox_angle_left.Focus()
        End If
    End Sub


    Private Sub TextBox_angle_left_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox_angle_left.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            TextBox_angle_right.SelectAll()
            TextBox_angle_right.Focus()
        End If
    End Sub

    Private Sub TextBox_angle_right_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox_angle_right.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            TextBox_radius_Left.SelectAll()
            TextBox_radius_Left.Focus()
        End If
    End Sub
    Private Sub TextBox_radius_Left_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox_radius_Left.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            TextBox_Radius_right.SelectAll()
            TextBox_Radius_right.Focus()
        End If
    End Sub

    Private Sub TextBox_radius_RIGHT_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox_Radius_right.KeyDown
        If CheckBox_pick_start_end.Checked = False Then
            If e.KeyCode = Windows.Forms.Keys.Enter Then
                TextBox_straight_length.SelectAll()
                TextBox_straight_length.Focus()
            End If
        End If
    End Sub

    Private Sub TextBox_straight_left_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox_straight_length.KeyDown
        If CheckBox_pick_start_end.Checked = False Then
            If e.KeyCode = Windows.Forms.Keys.Enter Then

            End If
        End If
    End Sub

    Private Sub CheckBox_0DEC_CheckedChanged(sender As Object, e As System.EventArgs) Handles CheckBox_0DEC.CheckedChanged
        If CheckBox_0DEC.Checked = True Then
            CheckBox_1DEC.Checked = False
        Else
            CheckBox_1DEC.Checked = True
        End If
    End Sub
    Private Sub CheckBox_1DEC_CheckedChanged(sender As Object, e As System.EventArgs) Handles CheckBox_1DEC.CheckedChanged
        If CheckBox_1DEC.Checked = True Then
            CheckBox_0DEC.Checked = False
        Else
            CheckBox_0DEC.Checked = True
        End If
    End Sub

    Private Sub TextBox_3d_angle_left_TextChanged(sender As Object, e As System.EventArgs) Handles TextBox_3d_angle_left.TextChanged
        TextBox_3d_angle_right.Text = 0
    End Sub
    Private Sub TextBox_3d_angle_right_TextChanged(sender As Object, e As System.EventArgs) Handles TextBox_3d_angle_right.TextChanged
        TextBox_3d_angle_left.Text = 0
    End Sub

    Private Sub Panel5_Click(sender As Object, e As EventArgs) Handles Panel_formating.Click
        Incarca_existing_textstyles_to_combobox(ComboBox_text_styles)
        If ComboBox_text_styles.Items.Count > 0 Then ComboBox_text_styles.SelectedIndex = 0
        Incarca_existing_layers_to_combobox(ComboBox_layer_hdd_cl)
        Incarca_existing_layers_to_combobox(ComboBox_layer_text)
    End Sub

End Class

