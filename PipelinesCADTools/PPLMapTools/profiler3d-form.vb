Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.ApplicationServices
Imports System.Math

Public Class profiler3d_form

    Private Sub Button_draw_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_draw.Click
        If issecure() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
        ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try
            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument

            Using Lock1
                Dim Grid_Z_Max, Grid_Z_Min As Double
                If IsNumeric(TextBox_L_elevation.Text) = True Then
                    Grid_Z_Min = CDbl(TextBox_L_elevation.Text)
                Else
                    MsgBox("Non numeric lowest elevation")
                    Commands_class_plain_autocad.Profile_table.Clear()
                    Me.Close()
                    Exit Sub
                End If

                If IsNumeric(TextBox_H_Elevation.Text) = True Then
                    Grid_Z_Max = CDbl(TextBox_H_Elevation.Text)
                Else
                    MsgBox("Non numeric highest elevation")
                    Commands_class_plain_autocad.Profile_table.Clear()
                    Me.Close()
                    Exit Sub
                End If

                Dim CSF As Double = 1




                Dim Color_index_grid As Integer

                If IsNumeric(TextBox_color_index_grid_lines.Text) = True Then
                    Color_index_grid = CInt(TextBox_color_index_grid_lines.Text)
                Else
                    MsgBox("Non numeric viewport scale")
                    Commands_class_plain_autocad.Profile_table.Clear()
                    Me.Close()
                    Exit Sub
                End If
                If ComboBox_LAYER_TEXT.Text = "" Or ComboBox_LAYER_GRIDLINES.Text = "" Or ComboBox_LAYER_POLYLINE.Text = "" Then
                    MsgBox("Layer not specified")
                    Commands_class_plain_autocad.Profile_table.Clear()
                    Me.Close()
                    Exit Sub
                End If


                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim Index_minus As Double
                    Dim Index_plus As Double



                    Dim Text_style_table As Autodesk.AutoCAD.DatabaseServices.TextStyleTable = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                    Dim Text_style_ID As Autodesk.AutoCAD.DatabaseServices.ObjectId = Text_style_table.Item(ComboBox_text_styles.Text)


                    Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                    Dim Index_On_the_poly_for_zero As Double = -1

                    If CheckBox_PICK_ZERO.Checked = True Then
1234:
                        Dim Prompt_zero As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Please Pick the 0+000 position on 3Dpolyline")
                        Dim Point111 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                        Prompt_zero.AllowNone = False
                        Point111 = Editor1.GetPoint(Prompt_zero)
                        If Point111.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Exit Sub
                        End If
                        Dim X001 As Double = Point111.Value.X
                        Dim Y001 As Double = Point111.Value.Y

                        Dim Pline3d As New Autodesk.AutoCAD.DatabaseServices.Polyline3d
                        Pline3d = Trans1.GetObject(Commands_class_plain_autocad.Poly3Did, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                        Dim p3d As Point3d = Pline3d.GetClosestPointTo(New Point3d(X001, Y001, 0), Vector3d.ZAxis, False)

                        Index_On_the_poly_for_zero = Pline3d.GetParameterAtPoint(p3d)
                        Index_minus = Math.Floor(Index_On_the_poly_for_zero)
                        Index_plus = Math.Ceiling(Index_On_the_poly_for_zero)
                        Dim Chain_minus, Chain_plus, Dist1 As Double
                        Dist1 = Commands_class_plain_autocad.Profile_table.Rows.Item(Index_plus).Item("Total_Chainage_grid") - Commands_class_plain_autocad.Profile_table.Rows.Item(Index_minus).Item("Total_Chainage_grid")
                        Chain_minus = (Index_On_the_poly_for_zero - Index_minus) * Dist1
                        Chain_plus = (Index_plus - Index_On_the_poly_for_zero) * Dist1

                        Commands_class_plain_autocad.Profile_table.Rows.Item(Index_minus).Item("Modified_Chainage_grid") = Chain_minus

                        For i = Index_minus - 1 To 0 Step -1
                            Commands_class_plain_autocad.Profile_table.Rows.Item(i).Item("Modified_Chainage_grid") = Commands_class_plain_autocad.Profile_table.Rows.Item(i + 1).Item("Modified_Chainage_grid") + Commands_class_plain_autocad.Profile_table.Rows.Item(i + 1).Item("Partial_Chainage_grid")
                        Next
                        Commands_class_plain_autocad.Profile_table.Rows.Item(Index_plus).Item("Modified_Chainage_grid") = Chain_plus
                        For i = Index_plus + 1 To Commands_class_plain_autocad.Profile_table.Rows.Count - 1
                            Commands_class_plain_autocad.Profile_table.Rows.Item(i).Item("Modified_Chainage_grid") = Commands_class_plain_autocad.Profile_table.Rows.Item(i - 1).Item("Modified_Chainage_grid") + Commands_class_plain_autocad.Profile_table.Rows.Item(i).Item("Partial_Chainage_grid")
                        Next

                    End If

                    Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                    Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Please pick the grid location:")

                    PP1.AllowNone = False
                    Point1 = Editor1.GetPoint(PP1)
                    If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                        Commands_class_plain_autocad.Profile_table.Clear()
                        Me.Close()
                        Exit Sub
                    End If



                    Dim Xg0, Yg0 As Double
                    Xg0 = Point1.Value.X
                    Yg0 = Point1.Value.Y


                    Dim Pt1(0 To 1) As Double
                    Dim Pt2(0 To 1) As Double
                    Dim Pt3(0 To 1) As Double
                    Dim Pt4(0 To 1) As Double
                    Dim Pt5(0 To 1) As Double
                    Dim Pt6(0 To 1) As Double
                    Dim Pt7(0 To 1) As Double
                    Dim Pt8(0 To 1) As Double
                    Dim Pt9(0 To 1) As Double
                    Dim Pt10(0 To 1) As Double
                    Dim Pt11(0 To 1) As Double
                    Dim Pt12(0 To 1) As Double
                    Dim Pt13(0 To 1) As Double
                    Dim Pt14(0 To 1) As Double
                    Dim Pt15(0 To 1) As Double
                    Dim Pt16(0 To 1) As Double
                    Dim Pt17(0 To 1) As Double
                    Dim Pt18(0 To 1) As Double
                    Dim Pt19(0 To 1) As Double
                    Dim Pt20(0 To 1) As Double



                    Dim Nr_vert_Lines As Double
                    Dim Nr_hor_Lines As Double
                    Dim Horiz_increment, Vert_increment As Double
                    Dim hSCALE, vSCALE, Vertical_exag As Double


                    If IsNumeric(TextBox_Hincr.Text) = True Then
                        Horiz_increment = CDbl(TextBox_Hincr.Text)
                    Else
                        MsgBox("Non numeric horizontal increment")
                        Commands_class_plain_autocad.Profile_table.Clear()
                        Me.Close()
                        Exit Sub
                    End If

                    If IsNumeric(TextBox_Vincr.Text) = True Then
                        Vert_increment = CDbl(TextBox_Vincr.Text)
                    Else
                        MsgBox("Non numeric vertical increment")
                        Commands_class_plain_autocad.Profile_table.Clear()
                        Me.Close()
                        Exit Sub
                    End If
                    Dim Nr_vert_Lines_stanga, Nr_vert_Lines_dreapta, Chainage_start, Chainage_end As Double


                    If CheckBox_PICK_ZERO.Checked = True Then
                        Chainage_start = Commands_class_plain_autocad.Profile_table.Rows.Item(0).Item("Modified_Chainage_grid") / CSF
                        Chainage_end = Commands_class_plain_autocad.Profile_table.Rows.Item(Commands_class_plain_autocad.Profile_table.Rows.Count - 1).Item("Modified_Chainage_grid") / CSF


                        If Horiz_increment < Chainage_end Then
                            Nr_vert_Lines_dreapta = Ceiling(Chainage_end / Horiz_increment)
                        Else
                            Nr_vert_Lines_dreapta = 1
                        End If


                        If Horiz_increment < Chainage_start Then
                            Nr_vert_Lines_stanga = Floor(Chainage_start / Horiz_increment) + 1
                        Else
                            Nr_vert_Lines_stanga = 1
                        End If
                    Else
                        Nr_vert_Lines = Fix((Commands_class_plain_autocad.Total_chainage_grid / CSF) / Horiz_increment) + 1

                    End If

                    Nr_hor_Lines = Fix((Grid_Z_Max - Grid_Z_Min) / Vert_increment) + 1


                    Pt1(0) = Xg0
                    Pt1(1) = Yg0



                    Dim Factor_print_viewport As Double
                    Factor_print_viewport = 1000 / 1 'Printing_scale / Viewport_scale

                    If IsNumeric(TextBox_text_height.Text) = False Then
                        MsgBox("Non numeric text height")
                        Exit Sub
                    End If
                    Dim Text_height As Double = CDbl(TextBox_text_height.Text)

                    If IsNumeric(TextBox_Hscale.Text) = True Then
                        hSCALE = CDbl(TextBox_Hscale.Text)
                    Else
                        MsgBox("Non numeric horizontal scale")
                        Commands_class_plain_autocad.Profile_table.Clear()
                        Me.Close()
                        Exit Sub
                    End If

                    If IsNumeric(TextBox_Vscale.Text) = True Then
                        vSCALE = CDbl(TextBox_Vscale.Text)
                    Else
                        MsgBox("Non numeric horizontal scale")
                        Commands_class_plain_autocad.Profile_table.Clear()
                        Me.Close()
                        Exit Sub
                    End If
                    Vertical_exag = hSCALE / vSCALE

                    Pt2(0) = Xg0
                    Pt2(1) = Yg0 + Nr_hor_Lines * Vert_increment * (Factor_print_viewport / hSCALE) * Vertical_exag



                    Dim Vert_Zero_line As New Autodesk.AutoCAD.DatabaseServices.Line
                    Vert_Zero_line.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Pt1(0), Pt1(1), 0)
                    Vert_Zero_line.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Pt2(0), Pt2(1), 0)
                    Vert_Zero_line.Layer = ComboBox_LAYER_GRIDLINES.Text
                    If CheckBox_PICK_ZERO.Checked = True Then
                        Vert_Zero_line.ColorIndex = Color_index_grid
                        Vert_Zero_line.Linetype = ComboBox_LINETYPE.Text
                        Vert_Zero_line.LineWeight = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight000
                    End If
                    BTrecord.AppendEntity(Vert_Zero_line)
                    Trans1.AddNewlyCreatedDBObject(Vert_Zero_line, True)

                    If CheckBox_PICK_ZERO.Checked = True Then
                        For i = 1 To Nr_vert_Lines_dreapta
                            Dim Xv1, Yv1, Xv2, Yv2 As Double
                            Xv1 = Pt1(0) + i * Horiz_increment * (Factor_print_viewport / hSCALE)
                            Xv2 = Pt1(0) + i * Horiz_increment * (Factor_print_viewport / hSCALE)

                            Yv1 = Pt1(1)
                            Yv2 = Pt2(1)
                            Dim V_off_line_dreapta As New Autodesk.AutoCAD.DatabaseServices.Line
                            V_off_line_dreapta.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xv1, Yv1, 0)
                            V_off_line_dreapta.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xv2, Yv2, 0)
                            V_off_line_dreapta.Layer = ComboBox_LAYER_GRIDLINES.Text
                            If i < Nr_vert_Lines_dreapta Then
                                V_off_line_dreapta.ColorIndex = Color_index_grid
                                V_off_line_dreapta.Linetype = ComboBox_LINETYPE.Text
                                V_off_line_dreapta.LineWeight = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight000
                            Else
                                V_off_line_dreapta.Linetype = "CONTINUOUS"
                            End If


                            BTrecord.AppendEntity(V_off_line_dreapta)
                            Trans1.AddNewlyCreatedDBObject(V_off_line_dreapta, True)

                        Next

                        For i = 1 To Nr_vert_Lines_stanga
                            Dim Xv1, Yv1, Xv2, Yv2 As Double
                            Xv1 = Pt1(0) - i * Horiz_increment * (Factor_print_viewport / hSCALE)
                            Xv2 = Pt1(0) - i * Horiz_increment * (Factor_print_viewport / hSCALE)

                            Yv1 = Pt1(1)
                            Yv2 = Pt2(1)
                            Dim V_off_line_stanga As New Autodesk.AutoCAD.DatabaseServices.Line
                            V_off_line_stanga.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xv1, Yv1, 0)
                            V_off_line_stanga.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xv2, Yv2, 0)
                            V_off_line_stanga.Layer = ComboBox_LAYER_GRIDLINES.Text

                            If i < Nr_vert_Lines_stanga Then
                                V_off_line_stanga.ColorIndex = Color_index_grid
                                V_off_line_stanga.Linetype = ComboBox_LINETYPE.Text
                                V_off_line_stanga.LineWeight = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight000
                            Else
                                V_off_line_stanga.Linetype = "CONTINUOUS"
                            End If

                            BTrecord.AppendEntity(V_off_line_stanga)
                            Trans1.AddNewlyCreatedDBObject(V_off_line_stanga, True)

                        Next

                    Else

                        For i = 1 To Nr_vert_Lines
                            Dim Xv1, Yv1, Xv2, Yv2 As Double
                            Xv1 = Pt1(0) + i * Horiz_increment * (Factor_print_viewport / hSCALE)
                            Xv2 = Pt1(0) + i * Horiz_increment * (Factor_print_viewport / hSCALE)

                            Yv1 = Pt1(1)
                            Yv2 = Pt2(1)
                            Dim V_off_line As New Autodesk.AutoCAD.DatabaseServices.Line
                            V_off_line.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xv1, Yv1, 0)
                            V_off_line.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xv2, Yv2, 0)
                            V_off_line.Layer = ComboBox_LAYER_GRIDLINES.Text

                            If i < Nr_vert_Lines Then
                                V_off_line.ColorIndex = Color_index_grid
                                V_off_line.Linetype = ComboBox_LINETYPE.Text
                                If i = 1 Then V_off_line.Linetype = "CONTINUOUS"
                                V_off_line.LineWeight = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight000
                            Else
                                V_off_line.Linetype = "CONTINUOUS"
                            End If

                            BTrecord.AppendEntity(V_off_line)
                            Trans1.AddNewlyCreatedDBObject(V_off_line, True)

                        Next
                    End If 'CheckBox_PICK_ZERO.Checked = True 



                    If CheckBox_PICK_ZERO.Checked = True Then
                        Pt3(0) = Pt1(0) - Nr_vert_Lines_stanga * Horiz_increment * (Factor_print_viewport / hSCALE)
                        Pt3(1) = Pt1(1)

                        Pt4(0) = Pt1(0) + Nr_vert_Lines_dreapta * Horiz_increment * (Factor_print_viewport / hSCALE)
                        Pt4(1) = Pt1(1)

                    Else
                        Pt3(0) = Pt1(0)
                        Pt3(1) = Pt1(1)

                        Pt4(0) = Pt1(0) + Nr_vert_Lines * Horiz_increment * (Factor_print_viewport / hSCALE)
                        Pt4(1) = Pt1(1)
                    End If

                    Dim Horiz_Zero_line As New Autodesk.AutoCAD.DatabaseServices.Line
                    Horiz_Zero_line.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Pt3(0), Pt3(1), 0)
                    Horiz_Zero_line.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Pt4(0), Pt4(1), 0)
                    Horiz_Zero_line.Layer = ComboBox_LAYER_GRIDLINES.Text
                    Horiz_Zero_line.Linetype = "CONTINUOUS"
                    BTrecord.AppendEntity(Horiz_Zero_line)
                    Trans1.AddNewlyCreatedDBObject(Horiz_Zero_line, True)


                    For i = 1 To Nr_hor_Lines
                        Dim Xh1, Yh1, Xh2, Yh2 As Double
                        Xh1 = Pt3(0)
                        Xh2 = Pt4(0)

                        Yh1 = Pt3(1) + i * (Factor_print_viewport / hSCALE) * Vertical_exag * Vert_increment
                        Yh2 = Yh1
                        Dim h_off_line As New Autodesk.AutoCAD.DatabaseServices.Line
                        h_off_line.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xh1, Yh1, 0)
                        h_off_line.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xh2, Yh2, 0)

                        h_off_line.ColorIndex = Color_index_grid
                        h_off_line.Linetype = ComboBox_LINETYPE.Text
                        h_off_line.LineWeight = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight000
                        h_off_line.Layer = ComboBox_LAYER_GRIDLINES.Text
                        BTrecord.AppendEntity(h_off_line)
                        Trans1.AddNewlyCreatedDBObject(h_off_line, True)

                    Next





                    Pt6(0) = Pt3(0) ' - 10 * (Factor_print_viewport / 1000)
                    Pt6(1) = Pt1(1) '

                    Pt7(0) = Pt4(0) ' + 10 * (Factor_print_viewport / 1000)
                    Pt7(1) = Pt1(1) '


                    Pt8(0) = Pt1(0)
                    Pt8(1) = Pt1(1) - 1.2 * Text_height

                    Dim MText_ch_Hor As New Autodesk.AutoCAD.DatabaseServices.MText
                    MText_ch_Hor.Contents = "0+000"
                    MText_ch_Hor.Layer = ComboBox_LAYER_TEXT.Text
                    MText_ch_Hor.TextStyleId = Text_style_ID
                    MText_ch_Hor.TextHeight = Text_height
                    MText_ch_Hor.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.TopCenter
                    MText_ch_Hor.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt8(0), Pt8(1), 0)
                    BTrecord.AppendEntity(MText_ch_Hor)
                    Trans1.AddNewlyCreatedDBObject(MText_ch_Hor, True)

                    If CheckBox_PICK_ZERO.Checked = True Then
                        For i = 1 To Nr_vert_Lines_dreapta - 1

                            Pt9(0) = Pt8(0) + i * Horiz_increment * (Factor_print_viewport / hSCALE)
                            Pt9(1) = Pt8(1)

                            Dim Text_ch_Hor_h_dreapta As String = ""

                            If i * Horiz_increment < 10 Then
                                Text_ch_Hor_h_dreapta = "0+00" & i * Horiz_increment
                            End If

                            If i * Horiz_increment < 100 And i * Horiz_increment >= 10 Then
                                Text_ch_Hor_h_dreapta = "0+0" & i * Horiz_increment
                            End If

                            If i * Horiz_increment < 1000 And i * Horiz_increment >= 100 Then
                                Text_ch_Hor_h_dreapta = "0+" & i * Horiz_increment
                            End If

                            Dim String22 As String
                            String22 = Round(i * Horiz_increment, 0).ToString

                            If i * Horiz_increment >= 1000 Then
                                Text_ch_Hor_h_dreapta = Strings.Left(String22, Len(String22) - 3) & "+" & Strings.Right(String22, 3)
                            End If



                            Dim MText_ch_Hor_h_dreapta As New Autodesk.AutoCAD.DatabaseServices.MText
                            MText_ch_Hor_h_dreapta.Contents = Text_ch_Hor_h_dreapta
                            MText_ch_Hor_h_dreapta.Layer = ComboBox_LAYER_TEXT.Text
                            MText_ch_Hor_h_dreapta.TextStyleId = Text_style_ID
                            MText_ch_Hor_h_dreapta.TextHeight = Text_height
                            MText_ch_Hor_h_dreapta.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.TopCenter
                            MText_ch_Hor_h_dreapta.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt9(0), Pt9(1), 0)
                            BTrecord.AppendEntity(MText_ch_Hor_h_dreapta)
                            Trans1.AddNewlyCreatedDBObject(MText_ch_Hor_h_dreapta, True)


                        Next

                        For i = 1 To Nr_vert_Lines_stanga - 1

                            Pt9(0) = Pt8(0) - i * Horiz_increment * (Factor_print_viewport / hSCALE)
                            Pt9(1) = Pt8(1)

                            Dim Text_ch_Hor_h_stanga As String

                            If i * Horiz_increment < 10 Then
                                Text_ch_Hor_h_stanga = "-0+00" & i * Horiz_increment
                            End If

                            If i * Horiz_increment < 100 And i * Horiz_increment >= 10 Then
                                Text_ch_Hor_h_stanga = "-0+0" & i * Horiz_increment
                            End If

                            If i * Horiz_increment < 1000 And i * Horiz_increment >= 100 Then
                                Text_ch_Hor_h_stanga = "-0+" & i * Horiz_increment
                            End If

                            Dim String22 As String
                            String22 = Round(i * Horiz_increment, 0).ToString

                            If i * Horiz_increment >= 1000 Then
                                Text_ch_Hor_h_stanga = "-" & Strings.Left(String22, Len(String22) - 3) & "+" & Strings.Right(String22, 3)
                            End If


                            Dim MText_ch_Hor_h_stanga As New Autodesk.AutoCAD.DatabaseServices.MText
                            MText_ch_Hor_h_stanga.Contents = Text_ch_Hor_h_stanga
                            MText_ch_Hor_h_stanga.Layer = ComboBox_LAYER_TEXT.Text
                            MText_ch_Hor_h_stanga.TextStyleId = Text_style_ID
                            MText_ch_Hor_h_stanga.TextHeight = Text_height
                            MText_ch_Hor_h_stanga.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.TopCenter
                            MText_ch_Hor_h_stanga.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt9(0), Pt9(1), 0)
                            BTrecord.AppendEntity(MText_ch_Hor_h_stanga)
                            Trans1.AddNewlyCreatedDBObject(MText_ch_Hor_h_stanga, True)

                        Next

                    Else
                        For i = 1 To Nr_vert_Lines

                            Pt9(0) = Pt8(0) + i * Horiz_increment * (Factor_print_viewport / hSCALE)
                            Pt9(1) = Pt8(1)

                            Dim Text_ch_Hor_h As String

                            If i * Horiz_increment < 10 Then
                                Text_ch_Hor_h = "0+00" & i * Horiz_increment
                            End If

                            If i * Horiz_increment < 100 And i * Horiz_increment >= 10 Then
                                Text_ch_Hor_h = "0+0" & i * Horiz_increment
                            End If

                            If i * Horiz_increment < 1000 And i * Horiz_increment >= 100 Then
                                Text_ch_Hor_h = "0+" & i * Horiz_increment
                            End If

                            Dim String22 As String
                            String22 = Round(i * Horiz_increment, 0).ToString

                            If i * Horiz_increment >= 1000 Then
                                Text_ch_Hor_h = Strings.Left(String22, Len(String22) - 3) & "+" & Strings.Right(String22, 3)
                            End If

                            Dim MMtext_ch_Hor_h_stanga As New Autodesk.AutoCAD.DatabaseServices.MText
                            MMtext_ch_Hor_h_stanga.Contents = Text_ch_Hor_h
                            MMtext_ch_Hor_h_stanga.Layer = ComboBox_LAYER_TEXT.Text
                            MMtext_ch_Hor_h_stanga.TextStyleId = Text_style_ID
                            MMtext_ch_Hor_h_stanga.TextHeight = Text_height
                            MMtext_ch_Hor_h_stanga.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.TopCenter
                            MMtext_ch_Hor_h_stanga.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt9(0), Pt9(1), 0)
                            BTrecord.AppendEntity(MMtext_ch_Hor_h_stanga)
                            Trans1.AddNewlyCreatedDBObject(MMtext_ch_Hor_h_stanga, True)


                        Next
                    End If 'asta e de la If CheckBox_PICK_ZERO.Checked = True



                    Pt10(0) = Pt6(0) - 0.5 * Text_height
                    Pt10(1) = Pt6(1)

                    Pt12(0) = Pt7(0) + 0.5 * Text_height
                    Pt12(1) = Pt7(1)

                    Dim Vstring As String
                    Vstring = Round(Grid_Z_Min, 0).ToString

                    Dim Mtext_ch_Ver_left As New Autodesk.AutoCAD.DatabaseServices.MText
                    Mtext_ch_Ver_left.Contents = Vstring
                    Mtext_ch_Ver_left.Layer = ComboBox_LAYER_TEXT.Text
                    Mtext_ch_Ver_left.TextStyleId = Text_style_ID
                    Mtext_ch_Ver_left.TextHeight = Text_height
                    Mtext_ch_Ver_left.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.MiddleRight
                    Mtext_ch_Ver_left.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt10(0), Pt10(1), 0)
                    BTrecord.AppendEntity(Mtext_ch_Ver_left)
                    Trans1.AddNewlyCreatedDBObject(Mtext_ch_Ver_left, True)

                    Dim Mtext_ch_Ver_right As New Autodesk.AutoCAD.DatabaseServices.MText
                    Mtext_ch_Ver_right.Contents = Vstring
                    Mtext_ch_Ver_right.Layer = ComboBox_LAYER_TEXT.Text
                    Mtext_ch_Ver_right.TextStyleId = Text_style_ID
                    Mtext_ch_Ver_right.TextHeight = Text_height
                    Mtext_ch_Ver_right.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.MiddleLeft
                    Mtext_ch_Ver_right.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt12(0), Pt12(1), 0)
                    BTrecord.AppendEntity(Mtext_ch_Ver_right)
                    Trans1.AddNewlyCreatedDBObject(Mtext_ch_Ver_right, True)

                    For i = 1 To Nr_hor_Lines
                        Pt11(0) = Pt10(0)
                        Pt11(1) = Pt10(1) + i * Vert_increment * (Factor_print_viewport / hSCALE) * Vertical_exag
                        Pt13(0) = Pt12(0)
                        Pt13(1) = Pt12(1) + i * Vert_increment * (Factor_print_viewport / hSCALE) * Vertical_exag



                        Vstring = Round((i * Vert_increment) + Grid_Z_Min, 0).ToString

                        Dim Mtext_ch_Ver_left1 As New Autodesk.AutoCAD.DatabaseServices.MText
                        Mtext_ch_Ver_left1.Contents = Vstring
                        Mtext_ch_Ver_left1.Layer = ComboBox_LAYER_TEXT.Text
                        Mtext_ch_Ver_left1.TextStyleId = Text_style_ID
                        Mtext_ch_Ver_left1.TextHeight = Text_height
                        Mtext_ch_Ver_left1.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.MiddleRight
                        Mtext_ch_Ver_left1.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt11(0), Pt11(1), 0)
                        BTrecord.AppendEntity(Mtext_ch_Ver_left1)
                        Trans1.AddNewlyCreatedDBObject(Mtext_ch_Ver_left1, True)

                        Dim Mtext_ch_Ver_right1 As New Autodesk.AutoCAD.DatabaseServices.MText
                        Mtext_ch_Ver_right1.Contents = Vstring
                        Mtext_ch_Ver_right1.Layer = ComboBox_LAYER_TEXT.Text
                        Mtext_ch_Ver_right1.TextStyleId = Text_style_ID
                        Mtext_ch_Ver_right1.TextHeight = Text_height
                        Mtext_ch_Ver_right1.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.MiddleLeft
                        Mtext_ch_Ver_right1.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt13(0), Pt13(1), 0)
                        BTrecord.AppendEntity(Mtext_ch_Ver_right1)
                        Trans1.AddNewlyCreatedDBObject(Mtext_ch_Ver_right1, True)


                    Next

                    Pt14(0) = (Pt3(0) + Pt4(0)) / 2
                    Pt14(1) = (Pt3(1) + Pt4(1)) / 2 - 6.5 * Text_height
                    Pt15(0) = (Pt3(0) + Pt4(0)) / 2
                    Pt15(1) = (Pt3(1) + Pt4(1)) / 2 - 10.5 * Text_height


                    Dim Titlu1 As New Autodesk.AutoCAD.DatabaseServices.MText
                    Titlu1.Contents = "PROFILE ALONG "
                    Titlu1.Layer = ComboBox_LAYER_TEXT.Text
                    Titlu1.TextStyleId = Text_style_ID
                    Titlu1.TextHeight = 3 * Text_height
                    Titlu1.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.MiddleCenter
                    Titlu1.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt14(0), Pt14(1), 0)
                    BTrecord.AppendEntity(Titlu1)
                    Trans1.AddNewlyCreatedDBObject(Titlu1, True)

                    Dim Titlu2 As New Autodesk.AutoCAD.DatabaseServices.MText
                    Titlu2.Contents = "SCALE - HORIZ=1:" & hSCALE & "  VERT=1:" & vSCALE
                    Titlu2.Layer = ComboBox_LAYER_TEXT.Text
                    Titlu2.TextStyleId = Text_style_ID
                    Titlu2.TextHeight = 2 * Text_height
                    Titlu2.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.MiddleCenter
                    Titlu2.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt15(0), Pt15(1), 0)
                    BTrecord.AppendEntity(Titlu2)
                    Trans1.AddNewlyCreatedDBObject(Titlu2, True)




                    If CheckBox_PICK_ZERO.Checked = True Then
                        Pt19(0) = Pt1(0) - Nr_vert_Lines_stanga * Horiz_increment * (Factor_print_viewport / hSCALE)
                        Pt19(1) = Pt9(1) '- 1.5 * (Factor_print_viewport / 500)
                        Pt20(0) = Pt1(0) + Nr_vert_Lines_dreapta * Horiz_increment * (Factor_print_viewport / hSCALE)
                        Pt20(1) = Pt9(1) '- 1.5 * (Factor_print_viewport / 500)
                    Else
                        Pt19(0) = Pt1(0)
                        Pt19(1) = Pt9(1) - 1.5 * Text_height
                        Pt20(0) = Pt1(0) + Nr_vert_Lines * Horiz_increment * (Factor_print_viewport / hSCALE)
                        Pt20(1) = Pt9(1) - 1.5 * Text_height

                    End If 'asta e de la If CheckBox_PICK_ZERO.Checked = True

                    Dim UCS_CURENT As Matrix3d = Editor1.CurrentUserCoordinateSystem

                    Dim Punct1 As New Point3d(Commands_class_plain_autocad.Profile_table.Rows.Item(0).Item("X"), Commands_class_plain_autocad.Profile_table.Rows.Item(0).Item("Y"), 0)
                    Dim Punct2 As New Point3d(Commands_class_plain_autocad.Profile_table.Rows.Item(Commands_class_plain_autocad.Profile_table.Rows.Count - 1).Item("X"), Commands_class_plain_autocad.Profile_table.Rows.Item(Commands_class_plain_autocad.Profile_table.Rows.Count - 1).Item("Y"), 0)

                    Dim punct1_wcs As Point3d = Punct1.TransformBy(UCS_CURENT)
                    Dim punct2_wcs As Point3d = Punct2.TransformBy(UCS_CURENT)
                    Dim X1_wcs, Y1_wcs, X2_wcs, Y2_wcs As Double
                    X1_wcs = punct1_wcs.X
                    Y1_wcs = punct1_wcs.Y
                    X2_wcs = punct2_wcs.X
                    Y2_wcs = punct2_wcs.Y



                    Dim Bearing_for_title As Double = GET_Bearing_rad(X1_wcs, Y1_wcs, X2_wcs, Y2_wcs)


                    Dim String_left, String_right As String

                    If Bearing_for_title > 0 And Bearing_for_title < 15 * PI / 180 Then
                        String_left = "WEST"
                        String_right = "EAST"
                    End If
                    If Bearing_for_title <= 75 * PI / 180 And Bearing_for_title > 15 * PI / 180 Then
                        String_left = "SOUTHWEST"
                        String_right = "NORTHEAST"
                    End If
                    If Bearing_for_title <= 115 * PI / 180 And Bearing_for_title > 75 * PI / 180 Then
                        String_left = "SOUTH"
                        String_right = "NORTH"
                    End If
                    If Bearing_for_title <= 165 * PI / 180 And Bearing_for_title > 115 * PI / 180 Then
                        String_left = "SOUTHEAST"
                        String_right = "NORTHWEST"
                    End If
                    If Bearing_for_title <= 195 * PI / 180 And Bearing_for_title > 165 * PI / 180 Then
                        String_left = "EAST"
                        String_right = "WEST"
                    End If
                    If Bearing_for_title <= 255 * PI / 180 And Bearing_for_title > 195 * PI / 180 Then
                        String_left = "NORTHEAST"
                        String_right = "SOUTHWEST"
                    End If
                    If Bearing_for_title <= 285 * PI / 180 And Bearing_for_title > 255 * PI / 180 Then
                        String_left = "NORTH"
                        String_right = "SOUTH"
                    End If
                    If Bearing_for_title <= 345 * PI / 180 And Bearing_for_title > 285 * PI / 180 Then
                        String_left = "NORTHWEST"
                        String_right = "SOUTHEAST"
                    End If
                    If Bearing_for_title > 345 * PI / 180 Then
                        String_left = "WEST"
                        String_right = "EAST"
                    End If






                    Dim North_South1 As New Autodesk.AutoCAD.DatabaseServices.MText
                    North_South1.Contents = String_left
                    North_South1.Layer = ComboBox_LAYER_TEXT.Text
                    North_South1.TextStyleId = Text_style_ID
                    North_South1.TextHeight = Text_height
                    North_South1.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.TopCenter
                    North_South1.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt19(0), Pt19(1), 0)
                    BTrecord.AppendEntity(North_South1)
                    Trans1.AddNewlyCreatedDBObject(North_South1, True)

                    Dim North_South2 As New Autodesk.AutoCAD.DatabaseServices.MText
                    North_South2.Contents = String_right
                    North_South2.Layer = ComboBox_LAYER_TEXT.Text
                    North_South2.TextStyleId = Text_style_ID
                    North_South2.TextHeight = Text_height
                    North_South2.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.TopCenter
                    North_South2.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt20(0), Pt20(1), 0)
                    BTrecord.AppendEntity(North_South2)
                    Trans1.AddNewlyCreatedDBObject(North_South2, True)



                    Dim Profile_x_point(0 To Commands_class_plain_autocad.Profile_table.Rows.Count - 1) As Double
                    Dim Profile_y_point(0 To Commands_class_plain_autocad.Profile_table.Rows.Count - 1) As Double
                    Dim Ground_3d_length As Double
                    Dim Poly1 As New Autodesk.AutoCAD.DatabaseServices.Polyline()


                    Profile_x_point(0) = Pt1(0) - Chainage_start * (Factor_print_viewport / hSCALE)
                    Profile_y_point(0) = Yg0 + (Commands_class_plain_autocad.Profile_table.Rows.Item(0).Item("Z") - Grid_Z_Min) * (Factor_print_viewport / hSCALE) * Vertical_exag

                    Poly1.AddVertexAt(0, New Autodesk.AutoCAD.Geometry.Point2d(Profile_x_point(0), Profile_y_point(0)), 0, 0, 0)
                    'here You can put CSF!
                    For i = 1 To Commands_class_plain_autocad.Profile_table.Rows.Count - 1

                        Dim Grid_Chainage As Double = Commands_class_plain_autocad.Profile_table.Rows.Item(i).Item("Partial_Chainage_grid")
                        Profile_x_point(i) = Profile_x_point(i - 1) + Grid_Chainage * (Factor_print_viewport / hSCALE) / CSF
                        Profile_y_point(i) = Yg0 + (Commands_class_plain_autocad.Profile_table.Rows.Item(i).Item("Z") - Grid_Z_Min) * (Factor_print_viewport / hSCALE) * Vertical_exag
                        Poly1.AddVertexAt(i, New Autodesk.AutoCAD.Geometry.Point2d(Profile_x_point(i), Profile_y_point(i)), 0, 0, 0)
                        Ground_3d_length = Ground_3d_length + GET_distanta3d_Double_with_CSF(Commands_class_plain_autocad.Profile_table.Rows.Item(i - 1).Item("X"), Commands_class_plain_autocad.Profile_table.Rows.Item(i - 1).Item("Y"), Commands_class_plain_autocad.Profile_table.Rows.Item(i - 1).Item("Z"), Commands_class_plain_autocad.Profile_table.Rows.Item(i).Item("X"), Commands_class_plain_autocad.Profile_table.Rows.Item(i).Item("Y"), Commands_class_plain_autocad.Profile_table.Rows.Item(i).Item("Z"), CSF)

                    Next

                    Poly1.Layer = ComboBox_LAYER_POLYLINE.Text
                    BTrecord.AppendEntity(Poly1)
                    Trans1.AddNewlyCreatedDBObject(Poly1, True)

                    Dim Pt21(1) As Double

                    Pt21(0) = Pt15(0)
                    Pt21(1) = Pt15(1) - 5 * Text_height

                    Dim Report_Mtext As New Autodesk.AutoCAD.DatabaseServices.MText()
                    Report_Mtext.SetDatabaseDefaults()
                    Report_Mtext.LineSpacingFactor = 1
                    Report_Mtext.Contents = "COMBINED SCALE FACTOR APPLIED = " & CSF & vbCrLf & _
                                        "TOTAL 2D LENGTH OF THE PROFILE LINE = " & Round(Commands_class_plain_autocad.Total_chainage_grid / CSF, 4) & vbCrLf & _
                                        "TOTAL 3D LENGTH OF THE PROFILE LINE = " & Round(Ground_3d_length, 4)


                    Report_Mtext.TextHeight = Text_height
                    Report_Mtext.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt21(0), Pt21(1), 0)
                    Report_Mtext.Rotation = 0
                    Report_Mtext.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.TopLeft
                    Report_Mtext.Layer = "0"
                    BTrecord.AppendEntity(Report_Mtext)
                    Trans1.AddNewlyCreatedDBObject(Report_Mtext, True)

                    Editor1.Regen()
                    Trans1.Commit()
                End Using ' asta e de la transaction
                Commands_class_plain_autocad.Profile_table.Clear()
                Commands_class_plain_autocad.Total_chainage_grid = 0
            End Using ' ASTA E DE LA LOCK1
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
            Commands_class_plain_autocad.Profile_table.Clear()
            Me.Close()
        End Try
    End Sub



    Private Sub TextBox_csf_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyValue = Windows.Forms.Keys.Enter Then
            With TextBox_Hscale
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub


    Private Sub TextBox_Hscale_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox_Hscale.KeyDown, TextBox_text_height.KeyDown, TextBox_color_index_grid_lines.KeyDown
        If e.KeyValue = Windows.Forms.Keys.Enter Then
            With TextBox_Vscale
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub


    Private Sub TextBox_Vscale_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox_Vscale.KeyDown
        If e.KeyValue = Windows.Forms.Keys.Enter Then
            With TextBox_Hincr
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub

    Private Sub TextBox_Hincr_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox_Hincr.KeyDown
        If e.KeyValue = Windows.Forms.Keys.Enter Then
            With TextBox_Vincr
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub

    Private Sub TextBox_Vincr_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox_Vincr.KeyDown
        If e.KeyValue = Windows.Forms.Keys.Enter Then
            With TextBox_L_elevation
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub



    Private Sub TextBox_L_elevation_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox_L_elevation.KeyDown
        If e.KeyValue = Windows.Forms.Keys.Enter Then
            With TextBox_H_Elevation
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub

    Private Sub TextBox_H_Elevation_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox_H_Elevation.KeyDown
        If e.KeyValue = Windows.Forms.Keys.Enter Then
            Button_draw_Click(sender, e)

        End If
    End Sub

    Private Sub Profiler_form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim Inaltimea1 As Double = 1000
        'AM ADAUGAT COD IN LANSARE COMENZI DESPRE COMBO BOX
    End Sub

   
End Class