Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Public Class Viewports_along_graph_form
    Dim Colectie1 As New Specialized.StringCollection
    Dim Elevatia_cunoscuta As Double = -100000
    Dim Chainage_cunoscuta As Double = -100000
    Dim Chainage_Y As Double
    Dim Elevation_X1 As Double
    Dim Elevation_X2 As Double
    Dim Point_cunoscut As New Point3d
    Dim Data_table_poly As System.Data.DataTable
    Dim Old1, Old2, Old3, Old4, Old5, Old6, Old8, Old9, Old10, Old11, Old12, Old13, Old14, Old15 As String
    Dim t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13, t14, t15 As String

    Dim Freeze_operations As Boolean = False
    Dim Data_table_Matchlines As System.Data.DataTable
    Dim Data_table_station_equation As System.Data.DataTable
    Dim PolyCL As Polyline
    Dim PolyCL3D As Polyline3d
    Dim Poly_length As Double

    Dim Viewport_height As Double = 0
    Dim Viewport_width As Double = 0
    Dim Viewport_scale As Double = 0


    Dim Empty_array() As ObjectId

    Private Sub Viewports_along_graph_form_Load(sender As Object, e As EventArgs) Handles Me.Load, Panel_blocks.Click
        Button_draw.Visible = False
        Old1 = TextBox_W2.Text
        Old2 = TextBox_W3.Text
        Old3 = TextBox_deltaY_Vw2.Text
        Old4 = TextBox_H1.Text
        Old5 = TextBox_delta_x1.Text
        Old6 = TextBox_delta_x2.Text
        Old8 = TextBox_H_SCALE.Text
        Old9 = TextBox_V_SCALE.Text
        Old10 = TextBox_x.Text
        Old11 = TextBox_y.Text
        Old12 = TextBox_H2.Text
        Old13 = TextBox_shiftY_viewport.Text
        Panel_matchline_label.Visible = False
        Old14 = TextBox_Viewport_spacing1.Text
        Old15 = TextBox_Viewport_spacing2.Text
        Incarca_existing_Blocks_to_combobox(ComboBox_blocks_left)
        Incarca_existing_Blocks_to_combobox(ComboBox_blocks_right)
        Incarca_existing_layers_to_combobox(ComboBox_layer_blocks)

    End Sub

    Private Sub Button_load_graph_Click(sender As Object, e As EventArgs) Handles Button_load_graph.Click

        Try

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
            ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            ascunde_butoanele_pentru_forms(Me, Colectie1)

            Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Rezultat_hor As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_prompt_hor As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_prompt_hor.MessageForAdding = vbLf & "Select a known vertical line (STATION) and the label for it:"

                    Object_prompt_hor.SingleOnly = False
                    Rezultat_hor = Editor1.GetSelection(Object_prompt_hor)


                    Dim Rezultat_vert As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_prompt_vert As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_prompt_vert.MessageForAdding = vbLf & "Select a known horizontal line (ELEVATION) and the label for it:"

                    Object_prompt_vert.SingleOnly = False
                    Rezultat_vert = Editor1.GetSelection(Object_prompt_vert)

                    Dim Rezultat_vert2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_prompt_vert2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_prompt_vert2.MessageForAdding = vbLf & "Select an elevation label from the other side:"
                    Object_prompt_vert2.SingleOnly = True
                    Rezultat_vert2 = Editor1.GetSelection(Object_prompt_vert2)

                    Dim x01, y01, x02, y02 As Double
                    Dim x03, y03, x04, y04 As Double

                    If Rezultat_hor.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK And Rezultat_vert.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK And Rezultat_vert2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Dim Empty_array() As ObjectId


                        Dim mText_cunoscut As Autodesk.AutoCAD.DatabaseServices.MText
                        Dim Text_cunoscut As Autodesk.AutoCAD.DatabaseServices.DBText
                        Dim Linia_cunoscuta As Autodesk.AutoCAD.DatabaseServices.Line
                        Dim PolyLinia_cunoscuta As Autodesk.AutoCAD.DatabaseServices.Polyline



                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_vert.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                        Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj2 = Rezultat_vert.Value.Item(1)
                        Dim Ent2 As Entity
                        Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)

                        Dim Obj3 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj3 = Rezultat_hor.Value.Item(0)
                        Dim Ent3 As Entity
                        Ent3 = Obj3.ObjectId.GetObject(OpenMode.ForRead)

                        Dim Obj4 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Dim Ent4 As Entity
                        Obj4 = Rezultat_hor.Value.Item(1)
                        Ent4 = Obj4.ObjectId.GetObject(OpenMode.ForRead)

                        Dim Obj5 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj5 = Rezultat_vert2.Value.Item(0)
                        Dim Ent5 As Entity
                        Ent5 = Obj5.ObjectId.GetObject(OpenMode.ForRead)



                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                            mText_cunoscut = Ent1
                            Dim String_Mtext As String = Replace(mText_cunoscut.Text, "'", "")

                            If IsNumeric(String_Mtext) = True Then
                                Elevatia_cunoscuta = CDbl(String_Mtext)
                                Elevation_X2 = mText_cunoscut.Location.X
                            End If

                        End If

                        If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                            mText_cunoscut = Ent2
                            Dim String_Mtext As String = Replace(mText_cunoscut.Text, "'", "")
                            If IsNumeric(String_Mtext) = True Then
                                Elevatia_cunoscuta = CDbl(String_Mtext)
                                Elevation_X2 = mText_cunoscut.Location.X
                            End If

                        End If

                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                            Text_cunoscut = Ent1
                            Dim String_Text As String = Replace(Text_cunoscut.TextString, "'", "")
                            If IsNumeric(String_Text) = True Then
                                Elevatia_cunoscuta = CDbl(String_Text)
                                Elevation_X2 = Text_cunoscut.Position.X
                            End If

                        End If

                        If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                            Text_cunoscut = Ent2
                            Dim String_Text As String = Replace(Text_cunoscut.TextString, "'", "")
                            If IsNumeric(String_Text) = True Then
                                Elevatia_cunoscuta = CDbl(String_Text)
                                Elevation_X2 = Text_cunoscut.Position.X
                            End If

                        End If

                        If TypeOf Ent5 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                            mText_cunoscut = Ent5
                            Dim String_Mtext As String = Replace(mText_cunoscut.Text, "'", "")

                            If IsNumeric(String_Mtext) = True Then
                                Elevation_X1 = mText_cunoscut.Location.X
                            End If

                        End If

                        If TypeOf Ent5 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                            Text_cunoscut = Ent5
                            Dim String_Text As String = Replace(Text_cunoscut.TextString, "'", "")
                            If IsNumeric(String_Text) = True Then
                                Elevation_X1 = Text_cunoscut.Position.X
                            End If

                        End If
                        If Elevatia_cunoscuta = -100000 Then
                            afiseaza_butoanele_pentru_forms(Me, Colectie1)
                            Editor1.SetImpliedSelection(Empty_array)
                            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If



                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                            Linia_cunoscuta = Ent1
                            x01 = Linia_cunoscuta.StartPoint.X
                            y01 = Linia_cunoscuta.StartPoint.Y
                            x02 = Linia_cunoscuta.EndPoint.X
                            y02 = Linia_cunoscuta.EndPoint.Y
                            If Abs(y01 - y02) > 0.001 Then
                                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                Editor1.SetImpliedSelection(Empty_array)
                                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                Exit Sub
                            End If
                        End If


                        If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                            Linia_cunoscuta = Ent2
                            x01 = Linia_cunoscuta.StartPoint.X
                            y01 = Linia_cunoscuta.StartPoint.Y
                            x02 = Linia_cunoscuta.EndPoint.X
                            y02 = Linia_cunoscuta.EndPoint.Y
                            If Abs(y01 - y02) > 0.001 Then
                                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                Editor1.SetImpliedSelection(Empty_array)
                                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                Exit Sub
                            End If
                        End If

                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                            PolyLinia_cunoscuta = Ent1

                            x01 = PolyLinia_cunoscuta.StartPoint.X
                            y01 = PolyLinia_cunoscuta.StartPoint.Y
                            x02 = PolyLinia_cunoscuta.EndPoint.X
                            y02 = PolyLinia_cunoscuta.EndPoint.Y
                            If Abs(y01 - y02) > 0.001 Then
                                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                Editor1.SetImpliedSelection(Empty_array)
                                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                Exit Sub
                            End If
                        End If
                        If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                            PolyLinia_cunoscuta = Ent2

                            x01 = PolyLinia_cunoscuta.StartPoint.X
                            y01 = PolyLinia_cunoscuta.StartPoint.Y
                            x02 = PolyLinia_cunoscuta.EndPoint.X
                            y02 = PolyLinia_cunoscuta.EndPoint.Y
                            If Abs(y01 - y02) > 0.001 Then
                                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                Editor1.SetImpliedSelection(Empty_array)
                                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                Exit Sub
                            End If

                        End If




                        If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                            mText_cunoscut = Ent3
                            Dim numar_fara_plus As String = Replace(mText_cunoscut.Text, "+", "")
                            If IsNumeric(numar_fara_plus) = True Then
                                Chainage_cunoscuta = CDbl(numar_fara_plus)
                                Chainage_Y = mText_cunoscut.Location.Y
                            End If

                        End If

                        If TypeOf Ent4 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                            mText_cunoscut = Ent4
                            Dim numar_fara_plus As String = Replace(mText_cunoscut.Text, "+", "")
                            If IsNumeric(numar_fara_plus) = True Then
                                Chainage_cunoscuta = CDbl(numar_fara_plus)
                                Chainage_Y = mText_cunoscut.Location.Y
                            End If

                        End If

                        If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                            Text_cunoscut = Ent3
                            Dim numar_fara_plus As String = Replace(Text_cunoscut.TextString, "+", "")
                            If IsNumeric(numar_fara_plus) = True Then
                                Chainage_cunoscuta = CDbl(numar_fara_plus)
                                Chainage_Y = Text_cunoscut.Position.Y
                            End If

                        End If

                        If TypeOf Ent4 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                            Text_cunoscut = Ent4
                            Dim numar_fara_plus As String = Replace(Text_cunoscut.TextString, "+", "")
                            If IsNumeric(numar_fara_plus) = True Then
                                Chainage_cunoscuta = CDbl(numar_fara_plus)
                                Chainage_Y = Text_cunoscut.Position.Y
                            End If

                        End If



                        If Chainage_cunoscuta = -100000 Then
                            afiseaza_butoanele_pentru_forms(Me, Colectie1)
                            Editor1.SetImpliedSelection(Empty_array)
                            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If



                        If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                            Linia_cunoscuta = Ent3
                            x03 = Linia_cunoscuta.StartPoint.X
                            y03 = Linia_cunoscuta.StartPoint.Y
                            x04 = Linia_cunoscuta.EndPoint.X
                            y04 = Linia_cunoscuta.EndPoint.Y
                            If Abs(x03 - x04) > 0.001 Then
                                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                Editor1.SetImpliedSelection(Empty_array)
                                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                Exit Sub
                            End If
                        End If


                        If TypeOf Ent4 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                            Linia_cunoscuta = Ent4
                            x03 = Linia_cunoscuta.StartPoint.X
                            y03 = Linia_cunoscuta.StartPoint.Y
                            x04 = Linia_cunoscuta.EndPoint.X
                            y04 = Linia_cunoscuta.EndPoint.Y
                            If Abs(x03 - x04) > 0.001 Then
                                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                Editor1.SetImpliedSelection(Empty_array)
                                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                Exit Sub
                            End If

                        End If

                        If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                            PolyLinia_cunoscuta = Ent3
                            x03 = PolyLinia_cunoscuta.StartPoint.X
                            y03 = PolyLinia_cunoscuta.StartPoint.Y
                            x04 = PolyLinia_cunoscuta.EndPoint.X
                            y04 = PolyLinia_cunoscuta.EndPoint.Y
                            If Abs(x03 - x04) > 0.001 Then
                                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                Editor1.SetImpliedSelection(Empty_array)
                                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                Exit Sub
                            End If
                        End If

                        If TypeOf Ent4 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                            PolyLinia_cunoscuta = Ent4
                            x03 = PolyLinia_cunoscuta.StartPoint.X
                            y03 = PolyLinia_cunoscuta.StartPoint.Y
                            x04 = PolyLinia_cunoscuta.EndPoint.X
                            y04 = PolyLinia_cunoscuta.EndPoint.Y
                            If Abs(x03 - x04) > 0.001 Then
                                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                Editor1.SetImpliedSelection(Empty_array)
                                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                Exit Sub
                            End If
                        End If
                    End If

                    Dim Linie1 As New Line(New Point3d(x01, y01, 0), New Point3d(x02, y02, 0))
                    Dim Linie2 As New Line(New Point3d(x03, y03, 0), New Point3d(x04, y04, 0))

                    If Linie1.Length > 0.01 And Linie2.Length > 0.01 Then
                        Dim Colint1 As New Point3dCollection
                        Linie1.IntersectWith(Linie2, Intersect.ExtendBoth, Colint1, IntPtr.Zero, IntPtr.Zero)
                        If Colint1.Count > 0 Then
                            Dim Rezultat_Poly As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                            Dim Object_prompt_Poly As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                            Object_prompt_Poly.MessageForAdding = vbLf & "Select the ground Polyline"

                            Object_prompt_Poly.SingleOnly = True
                            Rezultat_Poly = Editor1.GetSelection(Object_prompt_Poly)
                            If Rezultat_Poly.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                If TypeOf Rezultat_Poly.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead) Is Polyline Then
                                    Dim Poly1 As Polyline
                                    Poly1 = Rezultat_Poly.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)

                                    Data_table_poly = New System.Data.DataTable
                                    Data_table_poly.Columns.Add("X", GetType(Double))
                                    Data_table_poly.Columns.Add("Y", GetType(Double))

                                    For i = 0 To Poly1.NumberOfVertices - 1
                                        Data_table_poly.Rows.Add()
                                        Data_table_poly.Rows(i).Item("X") = Poly1.GetPoint3dAt(i).X
                                        Data_table_poly.Rows(i).Item("Y") = Poly1.GetPoint3dAt(i).Y
                                    Next

                                    Point_cunoscut = Colint1(0)
                                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                    Button_draw.Visible = True
                                Else
                                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                End If
                            Else
                                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                            End If


                        Else
                            afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        End If
                    Else
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                    End If



                End Using
            End Using

        Catch ex As Exception
            MsgBox(ex.Message)
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

    End Sub


    Private Sub Button_pick_corner_Click(sender As Object, e As EventArgs) Handles Button_pick_corner.Click
        Try

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
            ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            ascunde_butoanele_pentru_forms(Me, Colectie1)

            Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Pt_rezult As Autodesk.AutoCAD.EditorInput.PromptPointResult
                    Dim Prompt_pt As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify THE VIEWPORT lower left point")

                    Prompt_pt.AllowNone = True
                    Pt_rezult = Editor1.GetPoint(Prompt_pt)

                    If Pt_rezult.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        TextBox_x.Text = Get_String_Rounded(Pt_rezult.Value.X, 2)
                        TextBox_y.Text = Get_String_Rounded(Pt_rezult.Value.Y, 2)
                    End If



                    afiseaza_butoanele_pentru_forms(Me, Colectie1)

                End Using
            End Using

        Catch ex As Exception
            MsgBox(ex.Message)
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
    End Sub

    Private Sub CheckBox_USA_style_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_USA_style.CheckedChanged

        If CheckBox_USA_style.Checked = True Then

            TextBox_W2.Text = "80"
            TextBox_W3.Text = "80"
            TextBox_deltaY_Vw2.Text = "9"
            TextBox_H1.Text = "30"
            TextBox_delta_x1.Text = "33.24"
            TextBox_delta_x2.Text = "21.56"
            RadioButton_left_right.Checked = True
            TextBox_H_SCALE.Text = "1"
            TextBox_V_SCALE.Text = "1"
            TextBox_start.Text = "0"
            TextBox_end.Text = "5300"

            TextBox_x.Text = "1190"
            TextBox_y.Text = "1482.44"
            TextBox_H2.Text = "400"
            TextBox_shiftY_viewport.Text = "10"
            CheckBox_label_match.Checked = True
            TextBox_Viewport_spacing1.Text = "25"
            TextBox_Viewport_spacing2.Text = "25"

            t1 = TextBox_W2.Text
            t2 = TextBox_Viewport_spacing1.Text
            t3 = TextBox_delta_x2.Text
            t4 = TextBox_deltaY_Vw2.Text
            t5 = TextBox_shiftY_viewport.Text
            t6 = TextBox_Viewport_spacing2.Text
            t7 = TextBox_H1.Text
            t8 = TextBox_W3.Text
            t9 = TextBox_H2.Text
            t10 = TextBox_delta_x1.Text
            t11 = TextBox_match_textstyle.Text
            t12 = TextBox_match_text_height.Text
            t13 = TextBox_match_deltaX.Text
            t14 = TextBox_x.Text
            t15 = TextBox_y.Text
        Else

            TextBox_W2.Text = Old1
            TextBox_W3.Text = Old2
            TextBox_deltaY_Vw2.Text = Old3
            TextBox_H1.Text = Old4
            TextBox_delta_x1.Text = Old5
            TextBox_delta_x2.Text = Old6
            RadioButton_Right_left.Checked = True
            TextBox_H_SCALE.Text = Old8
            TextBox_V_SCALE.Text = Old9
            TextBox_x.Text = Old10
            TextBox_y.Text = Old11
            TextBox_H2.Text = Old12
            TextBox_shiftY_viewport.Text = Old13
            CheckBox_label_match.Checked = False
            TextBox_Viewport_spacing1.Text = Old14
            TextBox_Viewport_spacing2.Text = Old15

        End If
    End Sub

    Private Sub CheckBox_label_match_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_label_match.CheckedChanged
        If CheckBox_label_match.Checked = True Then
            Panel_matchline_label.Visible = True
        Else
            Panel_matchline_label.Visible = False
        End If
    End Sub



    Private Sub CheckBox_pick_middle_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_pick_middle.CheckedChanged, CheckBox_label_stationing.CheckedChanged
        If CheckBox_pick_middle.Checked = True Then
            Label_viewport_point.Text = "Middle point"
            Button_pick_corner.Text = "Pick Middle"
        Else
            Label_viewport_point.Text = "Lower Left Corner"
            Button_pick_corner.Text = "Pick Corner"
        End If
    End Sub

    Private Sub RadioButton_pen_east_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton_pen_east.CheckedChanged, RadioButton_default.CheckedChanged, RadioButton_etc_as_built.CheckedChanged, RadioButton_spire.CheckedChanged
        If RadioButton_pen_east.Checked = True Then
            TextBox_W2.Text = "30"
            TextBox_Viewport_spacing1.Text = "16"
            TextBox_delta_x2.Text = "-12"
            TextBox_deltaY_Vw2.Text = "-7"
            TextBox_shiftY_viewport.Text = "0"
            TextBox_Viewport_spacing2.Text = "16"
            TextBox_H1.Text = "32"
            TextBox_W3.Text = "30"
            TextBox_H2.Text = "490"
            TextBox_delta_x1.Text = "12"
            TextBox_match_textstyle.Text = "Arial"
            TextBox_match_text_height.Text = "8"
            TextBox_match_deltaX.Text = "8"
            TextBox_x.Text = "1981.7608"
            TextBox_y.Text = "260.00"
            CheckBox_pick_middle.Checked = True
            CheckBox_USA_style.Checked = True
            CheckBox_label_match.Checked = True
            TextBox_Match_layer.Text = "PROFILE-TEXT"
            CheckBox_label_stationing.Checked = True

        ElseIf RadioButton_default.Checked = True Then
            TextBox_W2.Text = t1
            TextBox_Viewport_spacing1.Text = t2
            TextBox_delta_x2.Text = t3
            TextBox_deltaY_Vw2.Text = t4
            TextBox_shiftY_viewport.Text = t5
            TextBox_Viewport_spacing2.Text = t6
            TextBox_H1.Text = t7
            TextBox_W3.Text = t8
            TextBox_H2.Text = t9
            TextBox_delta_x1.Text = t10
            TextBox_match_textstyle.Text = t11
            TextBox_match_text_height.Text = t12
            TextBox_match_deltaX.Text = t13
            TextBox_x.Text = t14
            TextBox_y.Text = t15

            CheckBox_pick_middle.Checked = False
            CheckBox_label_match.Checked = False
            CheckBox_USA_style.Checked = False
            CheckBox_label_stationing.Checked = False

        ElseIf RadioButton_etc_as_built.Checked = True Then



            TextBox_match_deltaX.Text = "8"
            TextBox_x.Text = "4090.84"
            TextBox_y.Text = "485.00"
            TextBox_shiftY_viewport.Text = "0"
            TextBox_deltaY_Vw2.Text = "-5"
            TextBox_H2.Text = "500"
            TextBox_W2.Text = "80"
            TextBox_W3.Text = "80"
            TextBox_delta_x1.Text = "33"
            TextBox_delta_x2.Text = "-33"
            TextBox_Viewport_spacing1.Text = "25"
            TextBox_Viewport_spacing2.Text = "25"
            TextBox_H1.Text = "30"
            CheckBox_pick_middle.Checked = True
            CheckBox_label_match.Checked = True
            CheckBox_USA_style.Checked = True
            TextBox_H_SCALE.Text = 1
            TextBox_V_SCALE.Text = 0.25
            RadioButton_left_right.Checked = True
            TextBox_match_text_height.Text = "16"
            TextBox_match_deltaX.Text = "12.5"
            TextBox_match_textstyle.Text = "ALIGNSTATEXT"
            TextBox_Match_layer.Text = "TEXT"
            CheckBox_label_stationing.Checked = False

        ElseIf RadioButton_spire.Checked = True Then
            TextBox_match_deltaX.Text = "8"
            TextBox_x.Text = "3639.493"
            TextBox_y.Text = "800"
            TextBox_shiftY_viewport.Text = "0"
            TextBox_deltaY_Vw2.Text = "10"
            TextBox_H2.Text = "600"
            TextBox_W2.Text = "80"
            TextBox_W3.Text = "80"
            TextBox_delta_x1.Text = "20"
            TextBox_delta_x2.Text = "0"
            TextBox_Viewport_spacing1.Text = "25"
            TextBox_Viewport_spacing2.Text = "25"
            TextBox_H1.Text = "30"
            CheckBox_pick_middle.Checked = True
            CheckBox_label_match.Checked = True
            CheckBox_USA_style.Checked = True
            TextBox_H_SCALE.Text = 1
            TextBox_V_SCALE.Text = 1
            RadioButton_left_right.Checked = True
            TextBox_match_text_height.Text = "16"
            TextBox_match_deltaX.Text = "12.5"
            TextBox_match_textstyle.Text = "Arial"
            TextBox_Match_layer.Text = "PROF_TEXT"
            CheckBox_label_stationing.Checked = False
        End If
    End Sub

    Private Sub Button_remove_items_list_Click(sender As Object, e As EventArgs) Handles Button_remove_items_list.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            If ListBox_DWG.Items.Count > 0 Then
                If ListBox_DWG.SelectedIndex >= 0 Then
                    ListBox_DWG.Items.RemoveAt((ListBox_DWG.SelectedIndex))
                    If ListBox_sheet_numbers.Items.Count > 0 Then
                        If ListBox_sheet_numbers.Items.Count >= ListBox_DWG.SelectedIndex Then
                            ListBox_sheet_numbers.Items.RemoveAt((ListBox_DWG.SelectedIndex))
                        End If
                    End If
                End If
            End If
            Freeze_operations = False
        End If
    End Sub
    Private Sub Button_clear_lists_Click(sender As Object, e As EventArgs) Handles Button_clear_lists.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            If ListBox_DWG.Items.Count > 0 Then
                ListBox_DWG.Items.Clear()
            End If
            If ListBox_sheet_numbers.Items.Count > 0 Then
                ListBox_sheet_numbers.Items.Clear()
            End If
            Freeze_operations = False
        End If
    End Sub
    Private Sub Button_load_DWG_Click(sender As Object, e As EventArgs) Handles Button_load_DWG.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
            W1 = get_new_worksheet_from_Excel()

            Dim column_file_name As String = TextBox_column_drawing_number.Text.ToUpper


            Try
                Dim FileBrowserDialog1 As New Windows.Forms.OpenFileDialog
                FileBrowserDialog1.Filter = "Drawing Files (*.dwg)|*.dwg|All Files (*.*)|*.*"
                FileBrowserDialog1.Multiselect = True

                If FileBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Dim i As Integer = 2
                    For Each file1 In FileBrowserDialog1.FileNames
                        ListBox_DWG.Items.Add(file1)
                        If IsNothing(W1) = False Then
                            Try
                                Dim Name1 As String = IO.Path.GetFileNameWithoutExtension(file1)
                                W1.Range(column_file_name & i).Value2 = Name1
                                i = i + 1
                            Catch ex As System.SystemException

                            End Try
                        End If


                    Next

                End If

            Catch ex As System.Exception
                Freeze_operations = False
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub
    Private Sub Button_load_layouts_Click(sender As Object, e As EventArgs) Handles Button_load_layouts.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
            W1 = Get_active_worksheet_from_Excel()
            Dim Start1 As Integer = 0
            Dim End1 As Integer = 0
            If IsNumeric(TextBox_SHEET_ROW_START.Text) = True Then
                Start1 = CInt(TextBox_SHEET_ROW_START.Text)
            End If
            If IsNumeric(TextBox_SHEET_ROW_END.Text) = True Then
                End1 = CInt(TextBox_SHEET_ROW_END.Text)
            End If

            If End1 = 0 Or Start1 = 0 Then
                Freeze_operations = False
                Exit Sub
            End If

            If End1 < Start1 Then
                Freeze_operations = False
                Exit Sub
            End If

            ListBox_DWG.Items.Clear()

            Dim column_file_name As String = TextBox_column_drawing_number.Text.ToUpper


            Try
                Using lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument

                    Using Trans1 As Transaction = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.TransactionManager.StartTransaction
                        Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                        Dim Layoutdict As DBDictionary = Trans1.GetObject(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database.LayoutDictionaryId, OpenMode.ForRead)
                        Dim BlockTable1 As BlockTable = Trans1.GetObject(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database.BlockTableId, OpenMode.ForRead)
                        For i = Start1 To End1
                            For Each entry As DBDictionaryEntry In Layoutdict
                                Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForRead)
                                If Layout1.TabOrder > 0 Then
                                    If Layout1.LayoutName = W1.Range(column_file_name & i).Value2 Then
                                        ListBox_DWG.Items.Add(Layout1.LayoutName)
                                    End If

                                End If
                            Next
                        Next
                    End Using
                End Using
            Catch ex As System.Exception
                Freeze_operations = False
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_load_sheet_from_excel_Click(sender As Object, e As EventArgs) Handles Button_load_band_from_excel.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Dim Start1 As Integer = 0
                Dim End1 As Integer = 0
                If IsNumeric(TextBox_SHEET_ROW_START.Text) = True Then
                    Start1 = CInt(TextBox_SHEET_ROW_START.Text)
                End If
                If IsNumeric(TextBox_SHEET_ROW_END.Text) = True Then
                    End1 = CInt(TextBox_SHEET_ROW_END.Text)
                End If

                If End1 = 0 Or Start1 = 0 Then
                    Freeze_operations = False
                    Exit Sub
                End If

                If End1 < Start1 Then
                    Freeze_operations = False
                    Exit Sub
                End If

                Dim Column_sheet As String = ""
                Column_sheet = TextBox_column_sheet_number.Text.ToUpper

                Dim Column_dwg As String = ""
                Column_dwg = TextBox_column_drawing_number.Text.ToUpper

                ListBox_sheet_numbers.Items.Clear()

                If ListBox_DWG.Items.Count > 0 Then
                    For i = 0 To ListBox_DWG.Items.Count - 1
                        ListBox_sheet_numbers.Items.Add(0)
                    Next
                    If End1 - Start1 + 1 = ListBox_DWG.Items.Count Then
                        For i = Start1 To End1
                            Dim Sheet_string As String = W1.Range(Column_sheet & i).Value2
                            Dim DWG_no As String = W1.Range(Column_dwg & i).Value2
                            If Not Sheet_string = "" Then

                                For j = 0 To ListBox_DWG.Items.Count - 1
                                    Dim DWG_no1 As String = ListBox_DWG.Items(j)
                                    Dim Nume1 As String = System.IO.Path.GetFileNameWithoutExtension(DWG_no1)
                                    If DWG_no = Nume1 Then
                                        ListBox_sheet_numbers.Items(j) = Sheet_string
                                        Exit For
                                    End If

                                    If DWG_no = DWG_no1 Then
                                        ListBox_sheet_numbers.Items(j) = Sheet_string
                                        Exit For
                                    End If
                                Next

                            End If
                        Next
                    End If
                End If


            Catch ex As System.Exception
                Freeze_operations = False
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_read_centerline_and_viewports_Click(sender As Object, e As EventArgs) Handles Button_read_centerline_and_viewports.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Dim Round1 As Integer = 0
            If IsNumeric(TextBox_dec1.Text) = True Then
                Round1 = CInt(TextBox_dec1.Text)
            End If

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor


            Editor1.SetImpliedSelection(Empty_array)
            Try
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                Colectie1 = New Specialized.StringCollection


                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptEntityResult

                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select centerline:")

                Object_Prompt.SetRejectMessage(vbLf & "Please select a lightweight polyline or a 3d polyline")
                Object_Prompt.AddAllowedClass(GetType(Polyline), True)
                Object_Prompt.AddAllowedClass(GetType(Polyline3d), True)


                Rezultat1 = Editor1.GetEntity(Object_Prompt)


                If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.WriteMessage(vbLf & "Command:")
                    Freeze_operations = False
                    Editor1.SetImpliedSelection(Empty_array)
                    Exit Sub
                End If

                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                    If IsNothing(Rezultat1) = False Then

                        Using Lock_dwg As DocumentLock = ThisDrawing.LockDocument


                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction


                                Dim PolyCL_for_viewports As Polyline = TryCast(Trans1.GetObject(Rezultat1.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Polyline)
                                If IsNothing(PolyCL_for_viewports) = False Then
                                    Poly_length = PolyCL_for_viewports.Length
                                    PolyCL = PolyCL_for_viewports
                                End If
                                Dim PolyCL3D_for_viewports As Polyline3d
                                PolyCL3D_for_viewports = TryCast(Trans1.GetObject(Rezultat1.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Polyline3d)

                                If IsNothing(PolyCL3D_for_viewports) = False Then
                                    Poly_length = PolyCL3D_for_viewports.Length
                                    PolyCL3D = PolyCL3D_for_viewports
                                    Dim Index_Poly As Integer = 0
                                    PolyCL_for_viewports = New Polyline
                                    For Each vId As Autodesk.AutoCAD.DatabaseServices.ObjectId In PolyCL3D_for_viewports
                                        Dim v3d As Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d = DirectCast(Trans1.GetObject _
                                                (vId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d)

                                        Dim x1 As Double = v3d.Position.X
                                        Dim y1 As Double = v3d.Position.Y
                                        Dim z1 As Double = v3d.Position.Z
                                        PolyCL_for_viewports.AddVertexAt(Index_Poly, New Point2d(x1, y1), 0, 0, 0)
                                        Index_Poly = Index_Poly + 1
                                    Next
                                    PolyCL_for_viewports.Elevation = 0
                                End If

                                If Not PolyCL_for_viewports.Elevation = 0 Then
                                    Freeze_operations = False
                                    MsgBox("CL Polyline is not at elevation 0")
                                    Exit Sub

                                End If

                                Data_table_Matchlines = New System.Data.DataTable
                                Data_table_Matchlines.Columns.Add("STATION1", GetType(Double))
                                Data_table_Matchlines.Columns.Add("STATION2", GetType(Double))
                                Data_table_Matchlines.Columns.Add("X1", GetType(Double))
                                Data_table_Matchlines.Columns.Add("Y1", GetType(Double))
                                Data_table_Matchlines.Columns.Add("X2", GetType(Double))
                                Data_table_Matchlines.Columns.Add("Y2", GetType(Double))

                                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                                Dim LayerTable1 As Autodesk.AutoCAD.DatabaseServices.LayerTable
                                LayerTable1 = Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                Dim Index_dataT As Double = 0


                                For Each objID As ObjectId In BTrecord
                                    Dim Rectangle_poly As Entity = Trans1.GetObject(objID, OpenMode.ForRead)

                                    Dim Executa As Boolean = False
                                    If TypeOf Rectangle_poly Is Polyline Then
                                        If Not Rectangle_poly.ObjectId = PolyCL_for_viewports.ObjectId Then
                                            If IsNothing(PolyCL3D_for_viewports) = False Then
                                                If Not Rectangle_poly.ObjectId = PolyCL3D_for_viewports.ObjectId Then
                                                    Executa = True
                                                End If
                                            Else
                                                Executa = True
                                            End If
                                        End If
                                    End If

                                    If Executa = True Then
                                        Dim Viewport_poly As Polyline = Rectangle_poly
                                        Dim LayerTableRecord1 As Autodesk.AutoCAD.DatabaseServices.LayerTableRecord
                                        LayerTableRecord1 = LayerTable1(Viewport_poly.Layer).GetObject(OpenMode.ForRead)


                                        If Viewport_poly.NumberOfVertices >= 4 And LayerTableRecord1.IsOff = False And LayerTableRecord1.IsFrozen = False Then
                                            Viewport_poly.UpgradeOpen()
                                            Viewport_poly.Elevation = 0
                                            Dim Col_int As New Point3dCollection
                                            Col_int = Intersect_on_both_operands(PolyCL_for_viewports, Viewport_poly)
                                            If Col_int.Count = 2 Then
                                                Dim Station1 As Double = 0
                                                Dim Station2 As Double = 0
                                                Dim Este_zero As Boolean = False
                                                Dim Nr_values As Integer = Col_int.Count
                                                If Nr_values > 2 Then Nr_values = 2


                                                Dim Point_on_poly1 As New Point3d()
                                                Point_on_poly1 = Col_int(0)
                                                Station1 = PolyCL_for_viewports.GetDistAtPoint(Point_on_poly1)

                                                If Round(Station1, 0) = 0 Then
                                                    Este_zero = True
                                                End If

                                                If IsNothing(PolyCL3D_for_viewports) = False Then
                                                    Dim Param1 As Double = PolyCL_for_viewports.GetParameterAtPoint(Point_on_poly1)
                                                    Station1 = PolyCL3D_for_viewports.GetDistanceAtParameter(Param1)
                                                End If

                                                Dim Point_on_poly2 As New Point3d()
                                                Point_on_poly2 = Col_int(1)
                                                Station2 = PolyCL_for_viewports.GetDistAtPoint(Point_on_poly2)

                                                If IsNothing(PolyCL3D_for_viewports) = False Then
                                                    Dim Param2 As Double = PolyCL_for_viewports.GetParameterAtPoint(Point_on_poly2)
                                                    Station2 = PolyCL3D_for_viewports.GetDistanceAtParameter(Param2)
                                                End If


                                                Dim Linie1 As New Line(Point_on_poly1, Point_on_poly2)

                                                Dim vpoly_exploded As New DBObjectCollection
                                                Viewport_poly.Explode(vpoly_exploded)



                                                If Station1 > Station2 Then
                                                    Dim T As Double = Station1
                                                    Station1 = Station2
                                                    Station2 = T
                                                End If

                                                If IsNothing(PolyCL3D_for_viewports) = False Then

                                                Else
                                                    If Station2 > PolyCL_for_viewports.Length Then
                                                        Station2 = PolyCL_for_viewports.Length
                                                    End If
                                                End If



                                                Dim X11, Y11, X21, Y22 As Double
                                                Dim LimitL As Double = 100
                                                For Each Ent4 As Entity In vpoly_exploded
                                                    If TypeOf (Ent4) Is Line Then
                                                        Dim Line2 As Line
                                                        Line2 = Ent4
                                                        If Line2.Length > LimitL Then
                                                            Dim Col_int1 As New Point3dCollection
                                                            Line2.IntersectWith(Linie1, Intersect.OnBothOperands, Col_int1, IntPtr.Zero, IntPtr.Zero)

                                                            If IsNothing(Col_int1) = True Then
                                                                Dim Pt1 As New Point3d
                                                                Dim Pt2 As New Point3d
                                                                Pt1 = Line2.GetClosestPointTo(Point_on_poly1, Vector3d.ZAxis, True)
                                                                Pt2 = Line2.GetClosestPointTo(Point_on_poly2, Vector3d.ZAxis, True)
                                                                If Pt1.GetVectorTo(Pt2).Length > LimitL Then
                                                                    X11 = Pt1.X
                                                                    Y11 = Pt1.Y
                                                                    X21 = Pt2.X
                                                                    Y22 = Pt2.Y
                                                                End If
                                                            Else
                                                                If Col_int1.Count = 0 Then
                                                                    Dim Pt1 As New Point3d
                                                                    Dim Pt2 As New Point3d
                                                                    Pt1 = Line2.GetClosestPointTo(Point_on_poly1, Vector3d.ZAxis, True)
                                                                    Pt2 = Line2.GetClosestPointTo(Point_on_poly2, Vector3d.ZAxis, True)
                                                                    If Pt1.GetVectorTo(Pt2).Length > LimitL Then
                                                                        X11 = Pt1.X
                                                                        Y11 = Pt1.Y
                                                                        X21 = Pt2.X
                                                                        Y22 = Pt2.Y
                                                                    End If

                                                                End If
                                                            End If



                                                        End If


                                                    End If

                                                Next


                                                Data_table_Matchlines.Rows.Add()
                                                Data_table_Matchlines.Rows(Index_dataT).Item("STATION1") = Round(Station1, Round1)
                                                Data_table_Matchlines.Rows(Index_dataT).Item("STATION2") = Round(Station2, Round1)
                                                Data_table_Matchlines.Rows(Index_dataT).Item("X1") = X11 'Viewport_poly.GetPointAtParameter(3).X
                                                Data_table_Matchlines.Rows(Index_dataT).Item("Y1") = Y11 'Viewport_poly.GetPointAtParameter(3).Y
                                                Data_table_Matchlines.Rows(Index_dataT).Item("X2") = X21 'Viewport_poly.GetPointAtParameter(2).X
                                                Data_table_Matchlines.Rows(Index_dataT).Item("Y2") = Y22 'Viewport_poly.GetPointAtParameter(2).Y
                                                Index_dataT = Index_dataT + 1


                                            End If


                                        End If

                                    End If


                                Next

                                Data_table_Matchlines = Sort_data_table(Data_table_Matchlines, "STATION1")

                                If Data_table_Matchlines.Rows.Count > 0 Then
                                    Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                                    W1 = get_new_worksheet_from_Excel()
                                    Dim Column_match As String = TextBox_column_sheet_number.Text.ToUpper
                                    Dim Idx_col As Integer = W1.Range(Column_match & "1").Column

                                    For i = 0 To Data_table_Matchlines.Rows.Count - 1
                                        W1.Range(Column_match & (i + 2).ToString).Value2 = Data_table_Matchlines.Rows(i).Item("STATION1") & " - " & Data_table_Matchlines.Rows(i).Item("STATION2")

                                        W1.Cells(i + 2, Idx_col + 2).Value2 = Data_table_Matchlines.Rows(i).Item("STATION1")
                                        W1.Cells(i + 2, Idx_col + 3).Value2 = Data_table_Matchlines.Rows(i).Item("STATION2")
                                        W1.Cells(i + 2, Idx_col + 4).FormulaR1C1 = "=RC[-2]-R[-1]C[-1]"
                                    Next
                                    W1.Cells(Data_table_Matchlines.Rows.Count + 2, Idx_col + 4).FormulaR1C1 = "=SUM(R[-" & (Data_table_Matchlines.Rows.Count).ToString & "]C:R[-1]C)"


                                End If


                                Trans1.Commit()


                            End Using
                        End Using

                    End If
                End If



                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")

            Catch ex As Exception
                Editor1.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try
            Freeze_operations = False
        End If
    End Sub


    Private Sub Button_draw_Click(sender As Object, e As EventArgs) Handles Button_draw.Click

        If Data_table_poly.Rows.Count > 0 Then
            Try
                If IsNumeric(TextBox_end.Text) = False Then
                    MsgBox("Please specify the END STATION!")
                    Exit Sub
                End If
                If IsNumeric(TextBox_H_SCALE.Text) = False Then
                    MsgBox("Please specify the HORIZONTAL SCALE!")
                    Exit Sub
                End If
                If IsNumeric(TextBox_H1.Text) = False Then
                    MsgBox("Please specify the HEIGHT!")
                    Exit Sub
                End If
                If IsNumeric(TextBox_H2.Text) = False Then
                    MsgBox("Please specify the HEIGHT!")
                    Exit Sub
                End If
                If IsNumeric(TextBox_start.Text) = False Then
                    MsgBox("Please specify the START STATION!")
                    Exit Sub
                End If
                If IsNumeric(TextBox_V_SCALE.Text) = False Then
                    MsgBox("Please specify the VERTICAL SCALE!")
                    Exit Sub
                End If
                If IsNumeric(TextBox_W3.Text) = False Then
                    MsgBox("Please specify the WIDTH!")
                    Exit Sub
                End If
                If IsNumeric(TextBox_W2.Text) = False Then
                    MsgBox("Please specify the WIDTH!")
                    Exit Sub
                End If
                If IsNumeric(TextBox_x.Text) = False Then
                    MsgBox("Please specify the X COORDINATE!")
                    Exit Sub
                End If
                If IsNumeric(TextBox_y.Text) = False Then
                    MsgBox("Please specify the Y COORDINATE!")
                    Exit Sub
                End If

                If IsNumeric(TextBox_viewport_scale.Text) = False Then
                    MsgBox("Please specify the viewport_scale")
                    Exit Sub
                End If


                Dim Sta1 As Double = CDbl(TextBox_start.Text)
                Dim Sta2 As Double = CDbl(TextBox_end.Text)
                Dim Hscale As Double = CDbl(TextBox_H_SCALE.Text)
                Dim Vscale As Double = CDbl(TextBox_V_SCALE.Text)
                Dim Viewport_scale As Double = CDbl(TextBox_viewport_scale.Text)


                If CheckBox_USA_style.Checked = True Then
                    Hscale = Hscale * 1000
                    Vscale = Vscale * 1000
                End If

                Dim H1 As Double = CDbl(TextBox_H1.Text)
                Dim H2 As Double = CDbl(TextBox_H2.Text)
                Dim W1 As Double = CDbl(TextBox_W3.Text)
                Dim W2 As Double = CDbl(TextBox_W2.Text)
                Dim x As Double = CDbl(TextBox_x.Text)
                Dim y As Double = CDbl(TextBox_y.Text)

                If Hscale <= 0 Or Vscale <= 0 Or H1 <= 0 Or H2 <= 0 Or W1 <= 0 Or W2 <= 0 Or Viewport_scale <= 0 Then
                    MsgBox("Negative values not allowed")
                    Exit Sub
                End If


                Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
                ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                ascunde_butoanele_pentru_forms(Me, Colectie1)

                Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Creaza_layer("NO PLOT", 41, "no plot", False)
                        Creaza_layer("GRID", 252, "grid", True)
                        If CheckBox_label_match.Checked = True Then
                            Creaza_layer(TextBox_Match_layer.Text, 7, "PROFILE TEXT", True)
                        End If
                        Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        Dim BTrecordMS As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BTrecordMS = Trans1.GetObject(BlockTable_data1(BlockTableRecord.ModelSpace), OpenMode.ForRead)
                        Dim BTrecordPS As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BTrecordPS = Trans1.GetObject(BlockTable_data1(BlockTableRecord.PaperSpace), OpenMode.ForWrite)
                        Dim Point0 As Point3d
                        Dim Point1 As Point3d
                        Dim Point2 As Point3d
                        Dim Point3 As Point3d

                        If RadioButton_left_right.Checked = True Then
                            Point0 = New Point3d(Point_cunoscut.X - Chainage_cunoscuta * 1000 / Hscale, Point_cunoscut.Y - Elevatia_cunoscuta * 1000 / Vscale, 0)
                            Point1 = New Point3d(Point0.X + Sta1 * 1000 / Hscale, Point0.Y + Elevatia_cunoscuta * 1000 / Vscale, 0)
                            Point2 = New Point3d(Point0.X + Sta2 * 1000 / Hscale, Point0.Y + Elevatia_cunoscuta * 1000 / Vscale, 0)
                        Else
                            Point0 = New Point3d(Point_cunoscut.X + Chainage_cunoscuta * 1000 / Hscale, Point_cunoscut.Y - Elevatia_cunoscuta * 1000 / Vscale, 0)
                            Point1 = New Point3d(Point0.X - Sta1 * 1000 / Hscale, Point0.Y + Elevatia_cunoscuta * 1000 / Vscale, 0)
                            Point2 = New Point3d(Point0.X - Sta2 * 1000 / Hscale, Point0.Y + Elevatia_cunoscuta * 1000 / Vscale, 0)
                        End If
                        Point3 = New Point3d((Point1.X + Point2.X) / 2, (Point1.Y + Point2.Y) / 2, 0)

                        Dim Poly1 As New Polyline
                        For i = 0 To Data_table_poly.Rows.Count - 1
                            Poly1.AddVertexAt(i, New Point2d(Data_table_poly.Rows(i).Item("X"), Data_table_poly.Rows(i).Item("Y")), 0, 0, 0)
                        Next

                        Dim Linie1 As New Line(New Point3d(Point1.X, Point1.Y - 100000, 0), New Point3d(Point1.X, Point1.Y + 100000, 0))
                        Dim Linie2 As New Line(New Point3d(Point2.X, Point2.Y - 100000, 0), New Point3d(Point2.X, Point2.Y + 100000, 0))

                        Dim ColInt1 As New Point3dCollection
                        Dim ColInt2 As New Point3dCollection
                        Poly1.IntersectWith(Linie1, Intersect.OnBothOperands, ColInt1, IntPtr.Zero, IntPtr.Zero)
                        Poly1.IntersectWith(Linie2, Intersect.OnBothOperands, ColInt2, IntPtr.Zero, IntPtr.Zero)

                        If ColInt1.Count > 0 And ColInt2.Count > 0 Then
                            Point3 = New Point3d((Point1.X + Point2.X) / 2, (ColInt1(0).Y + ColInt2(0).Y) / 2, 0)
                        End If

                        Dim DeltaY2 As Double

                        If IsNumeric(TextBox_deltaY_Vw2.Text) = True Then
                            DeltaY2 = CDbl(TextBox_shiftY_viewport.Text)
                        End If
                        If Not DeltaY2 = 0 Then
                            Point3 = New Point3d(Point3.X, Point3.Y + DeltaY2, 0)
                        End If

                        Dim Point4 As New Point3d(Point3.X - Abs(W1 - W2) * (1000 / Hscale) / 2, Chainage_Y, 0)

                        If Elevation_X1 < Elevation_X2 Then
                            Dim Temp1 As Double = Elevation_X1
                            Elevation_X1 = Elevation_X2
                            Elevation_X2 = Temp1
                        End If

                        Dim Point5 As New Point3d(Elevation_X2, Point3.Y, 0)
                        Dim Point6 As New Point3d(Elevation_X1, Point3.Y, 0)

                        Dim DeltaX1, DeltaX2, DeltaY As Double
                        If IsNumeric(TextBox_deltaY_Vw2.Text) = True Then
                            DeltaY = CDbl(TextBox_deltaY_Vw2.Text)
                        End If
                        If Not DeltaY = 0 Then
                            Point4 = New Point3d(Point4.X, Point4.Y + DeltaY, 0)
                        End If



                        If IsNumeric(TextBox_delta_x1.Text) = True Then
                            DeltaX1 = CDbl(TextBox_delta_x1.Text)
                        End If
                        If IsNumeric(TextBox_delta_x2.Text) = True Then
                            DeltaX2 = CDbl(TextBox_delta_x2.Text)
                        End If


                        If Not DeltaX2 = 0 Then
                            Point5 = New Point3d(Point5.X + DeltaX2, Point5.Y, 0)
                        End If
                        If Not DeltaX1 = 0 Then
                            Point6 = New Point3d(Point6.X + DeltaX1, Point6.Y, 0)
                        End If


                        Dim ExtraL As Double = 10

                        If CheckBox_USA_style.Checked = True Then ExtraL = 0

                        Dim L1 As Double = (Abs(Point1.X - Point2.X) + 2 * ExtraL) * Viewport_scale

                        Dim Spacing1 As Double = 0
                        If IsNumeric(TextBox_Viewport_spacing1.Text) = True Then
                            Spacing1 = CDbl(TextBox_Viewport_spacing1.Text)
                        End If

                        Dim Spacing2 As Double = 0
                        If IsNumeric(TextBox_Viewport_spacing2.Text) = True Then
                            Spacing2 = CDbl(TextBox_Viewport_spacing2.Text)
                        End If

                        If CheckBox_pick_middle.Checked = True Then
                            x = x - (L1 / 2 + (W2 + W1) / 2 + (Spacing1 + Spacing2) / 2)
                        End If

                        Dim Viewport1 As New Viewport
                        Viewport1.SetDatabaseDefaults()
                        Viewport1.CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(x + L1 / 2 + (W2 + W1) / 2 + (Spacing1 + Spacing2) / 2, y + H1 + H2 / 2, 0) ' asta e pozitia viewport in paper space
                        Viewport1.Height = H2
                        Viewport1.Width = L1
                        Viewport1.Layer = "GRID"
                        'Viewport1.ColorIndex = 1
                        Viewport1.ViewDirection = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis
                        Viewport1.ViewTarget = Point3 ' asta e pozitia viewport in MODEL space
                        Viewport1.ViewCenter = Autodesk.AutoCAD.Geometry.Point2d.Origin
                        Viewport1.TwistAngle = 0 ' asta e PT TWIST

                        BTrecordPS.AppendEntity(Viewport1)
                        Trans1.AddNewlyCreatedDBObject(Viewport1, True)

                        Viewport1.On = True
                        Viewport1.CustomScale = Viewport_scale
                        Viewport1.Locked = True

                        Dim Viewport2 As New Viewport
                        Viewport2.SetDatabaseDefaults()
                        Viewport2.CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(x + L1 / 2 + (W2 + W1) / 2 + (Spacing1 + Spacing2) / 2, y + H1 / 2, 0) ' asta e pozitia viewport in paper space
                        Viewport2.Height = H1
                        Viewport2.Width = L1
                        Viewport2.Layer = "NO PLOT"
                        'Viewport2.ColorIndex = 2
                        Viewport2.ViewDirection = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis
                        Viewport2.ViewTarget = Point4 ' asta e pozitia viewport in MODEL space
                        Viewport2.ViewCenter = Autodesk.AutoCAD.Geometry.Point2d.Origin
                        Viewport2.TwistAngle = 0 ' asta e PT TWIST

                        BTrecordPS.AppendEntity(Viewport2)
                        Trans1.AddNewlyCreatedDBObject(Viewport2, True)

                        Viewport2.On = True
                        Viewport2.CustomScale = Viewport_scale
                        Viewport2.Locked = True

                        Dim Viewport3 As New Viewport
                        Viewport3.SetDatabaseDefaults()
                        Viewport3.CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(x + W2 / 2, y + H1 + H2 / 2, 0) ' asta e pozitia viewport in paper space
                        Viewport3.Height = H2
                        Viewport3.Width = W2
                        Viewport3.Layer = "NO PLOT"
                        'Viewport3.ColorIndex = 3
                        Viewport3.ViewDirection = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis
                        Viewport3.ViewTarget = Point5 ' asta e pozitia viewport in MODEL space
                        Viewport3.ViewCenter = Autodesk.AutoCAD.Geometry.Point2d.Origin
                        Viewport3.TwistAngle = 0 ' asta e PT TWIST

                        BTrecordPS.AppendEntity(Viewport3)
                        Trans1.AddNewlyCreatedDBObject(Viewport3, True)

                        Viewport3.On = True
                        Viewport3.CustomScale = Viewport_scale
                        Viewport3.Locked = True

                        Dim Viewport4 As New Viewport
                        Viewport4.SetDatabaseDefaults()
                        Viewport4.CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(x + W2 + L1 + W1 / 2 + Spacing1 + Spacing2, y + H1 + H2 / 2, 0) ' asta e pozitia viewport in paper space
                        Viewport4.Height = H2
                        Viewport4.Width = W1
                        Viewport4.Layer = "NO PLOT"
                        'Viewport4.ColorIndex = 4
                        Viewport4.ViewDirection = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis
                        Viewport4.ViewTarget = Point6 ' asta e pozitia viewport in MODEL space
                        Viewport4.ViewCenter = Autodesk.AutoCAD.Geometry.Point2d.Origin
                        Viewport4.TwistAngle = 0 ' asta e PT TWIST

                        BTrecordPS.AppendEntity(Viewport4)
                        Trans1.AddNewlyCreatedDBObject(Viewport4, True)

                        Viewport4.On = True
                        Viewport4.CustomScale = Viewport_scale
                        Viewport4.Locked = True


                        If CheckBox_label_match.Checked = True Then

                            Dim Text_style_table As Autodesk.AutoCAD.DatabaseServices.TextStyleTable = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                            Dim TextStyleID As ObjectId = Nothing



                            For Each Text_id As ObjectId In Text_style_table
                                Dim TextStyle1 As TextStyleTableRecord = Trans1.GetObject(Text_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                If TextStyle1.Name.ToUpper = TextBox_match_textstyle.Text.ToUpper Then
                                    TextStyleID = TextStyle1.ObjectId
                                    Exit For
                                End If
                            Next

                            Dim Mtext_sta1 As New MText
                            Mtext_sta1.Contents = "STA " & Get_chainage_feet_from_double(Sta1, 0)
                            Mtext_sta1.Attachment = AttachmentPoint.MiddleCenter
                            Mtext_sta1.Rotation = PI / 2

                            Dim th As Double = 16
                            If IsNumeric(TextBox_match_text_height.Text) = True Then
                                th = CDbl(TextBox_match_text_height.Text)
                            End If

                            Mtext_sta1.TextHeight = th
                            If IsNothing(TextStyleID) = False Then
                                Mtext_sta1.TextStyleId = TextStyleID
                            End If

                            Mtext_sta1.Layer = TextBox_Match_layer.Text

                            Dim dM As Double = Spacing1 / 2
                            If IsNumeric(TextBox_match_deltaX.Text) = True Then
                                dM = CDbl(TextBox_match_deltaX.Text)
                            End If

                            Dim Pt1 As New Point3d(x + W2 + dM, y + H1 + H2 / 2, 0)
                            Mtext_sta1.Location = Pt1

                            BTrecordPS.AppendEntity(Mtext_sta1)
                            Trans1.AddNewlyCreatedDBObject(Mtext_sta1, True)

                            If CheckBox_label_stationing.Checked = True Then
                                Dim Mtext_sta3 As New MText
                                Mtext_sta3.Contents = "ELEVATION"
                                Mtext_sta3.Attachment = AttachmentPoint.MiddleCenter
                                Mtext_sta3.Rotation = PI / 2

                                Mtext_sta3.TextHeight = th
                                If IsNothing(TextStyleID) = False Then
                                    Mtext_sta3.TextStyleId = TextStyleID
                                End If

                                Mtext_sta3.Layer = TextBox_Match_layer.Text

                                Dim Pt3 As New Point3d(x - th, y + H1 + H2 / 2, 0)
                                Mtext_sta3.Location = Pt3

                                BTrecordPS.AppendEntity(Mtext_sta3)
                                Trans1.AddNewlyCreatedDBObject(Mtext_sta3, True)


                            End If


                            Dim Mtext_sta2 As New MText
                            Mtext_sta2.Contents = "STA " & Get_chainage_feet_from_double(Sta2, 0)
                            Mtext_sta2.Attachment = AttachmentPoint.MiddleCenter
                            Mtext_sta2.Rotation = PI / 2

                            Mtext_sta2.TextHeight = th
                            If IsNothing(TextStyleID) = False Then
                                Mtext_sta2.TextStyleId = TextStyleID
                            End If

                            Mtext_sta2.Layer = TextBox_Match_layer.Text

                            Dim dM2 As Double = Spacing2 / 2
                            If IsNumeric(TextBox_match_deltaX.Text) = True Then
                                dM2 = CDbl(TextBox_match_deltaX.Text)
                            End If

                            Dim Pt2 As New Point3d(x + W2 + Spacing1 + L1 + dM2, y + H1 + H2 / 2, 0)
                            Mtext_sta2.Location = Pt2

                            BTrecordPS.AppendEntity(Mtext_sta2)
                            Trans1.AddNewlyCreatedDBObject(Mtext_sta2, True)


                            If CheckBox_label_stationing.Checked = True Then
                                Dim Mtext_sta3 As New MText
                                Mtext_sta3.Contents = "ELEVATION"
                                Mtext_sta3.Attachment = AttachmentPoint.MiddleCenter
                                Mtext_sta3.Rotation = PI / 2

                                Mtext_sta3.TextHeight = th
                                If IsNothing(TextStyleID) = False Then
                                    Mtext_sta3.TextStyleId = TextStyleID
                                End If

                                Mtext_sta3.Layer = TextBox_Match_layer.Text

                                Dim Pt3 As New Point3d(x + W2 + Spacing1 + L1 + Spacing2 + W1 + th, y + H1 + H2 / 2, 0)
                                Mtext_sta3.Location = Pt3

                                BTrecordPS.AppendEntity(Mtext_sta3)
                                Trans1.AddNewlyCreatedDBObject(Mtext_sta3, True)

                                Dim Mtext_sta4 As New MText
                                Mtext_sta4.Contents = "STATIONING"
                                Mtext_sta4.Attachment = AttachmentPoint.MiddleCenter
                                Mtext_sta4.Rotation = 0

                                Mtext_sta4.TextHeight = th
                                If IsNothing(TextStyleID) = False Then
                                    Mtext_sta4.TextStyleId = TextStyleID
                                End If

                                Mtext_sta4.Layer = TextBox_Match_layer.Text


                                Dim Pt4 As New Point3d(x + W2 + Spacing1 + L1 / 2, y + th / 4, 0)
                                Mtext_sta4.Location = Pt4

                                BTrecordPS.AppendEntity(Mtext_sta4)
                                Trans1.AddNewlyCreatedDBObject(Mtext_sta4, True)


                            End If

                        End If


                        Trans1.Commit()

                    End Using
                End Using

            Catch ex As Exception
                MsgBox(ex.Message)
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
        End If
    End Sub


    Public Function Get_equation_value(ByVal Station_measured As Double) As Double
        Dim Valoare As Double = 0
        If IsNothing(Data_table_station_equation) = False Then
            If Data_table_station_equation.Rows.Count > 0 Then
                For i = 0 To Data_table_station_equation.Rows.Count - 1
                    If IsDBNull(Data_table_station_equation.Rows(i).Item("STATION_BACK")) = False And IsDBNull(Data_table_station_equation.Rows(i).Item("STATION_AHEAD")) = False Then
                        Dim Station_back As Double = Data_table_station_equation.Rows(i).Item("STATION_BACK")
                        Dim Station_ahead As Double = Data_table_station_equation.Rows(i).Item("STATION_AHEAD")

                        If Station_measured + Valoare < Station_back Then
                            Exit For
                        End If

                        Valoare = Valoare + Station_ahead - Station_back

                    End If
                Next
            End If


        End If


        Return Valoare
    End Function

    Private Sub Button_DWG_Insert_viewport_Click(sender As Object, e As EventArgs) Handles Button_DWG_Insert_viewport.Click


        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                If Not ListBox_DWG.Items.Count = ListBox_sheet_numbers.Items.Count Then
                    MsgBox("Please be sure numbers of DWG is equal to the number of bands")
                    Freeze_operations = False
                    Exit Sub
                End If

                If IsNothing(Data_table_poly) = True Then
                    MsgBox("Please LOAD THE graph")
                    Freeze_operations = False
                    Exit Sub
                End If

                If IsNumeric(TextBox_H_SCALE.Text) = False Then
                    MsgBox("Please specify the HORIZONTAL SCALE!")
                    Freeze_operations = False
                    Exit Sub
                End If
                If IsNumeric(TextBox_H1.Text) = False Then
                    MsgBox("Please specify the HEIGHT!")
                    Freeze_operations = False
                    Exit Sub
                End If
                If IsNumeric(TextBox_H2.Text) = False Then
                    MsgBox("Please specify the HEIGHT!")
                    Freeze_operations = False
                    Exit Sub
                End If

                If IsNumeric(TextBox_V_SCALE.Text) = False Then
                    MsgBox("Please specify the VERTICAL SCALE!")
                    Freeze_operations = False
                    Exit Sub
                End If
                If IsNumeric(TextBox_W3.Text) = False Then
                    MsgBox("Please specify the WIDTH!")
                    Freeze_operations = False
                    Exit Sub
                End If
                If IsNumeric(TextBox_W2.Text) = False Then
                    MsgBox("Please specify the WIDTH!")
                    Freeze_operations = False
                    Exit Sub
                End If
                If IsNumeric(TextBox_x.Text) = False Then
                    MsgBox("Please specify the X COORDINATE!")
                    Freeze_operations = False
                    Exit Sub
                End If
                If IsNumeric(TextBox_y.Text) = False Then
                    MsgBox("Please specify the Y COORDINATE!")
                    Freeze_operations = False
                    Exit Sub
                End If

                If IsNumeric(TextBox_viewport_scale.Text) = False Then
                    Freeze_operations = False
                    MsgBox("Please specify the viewport_scale")
                    Exit Sub
                End If


                If IsNumeric(TextBox_Layout_index_START.Text) = False Then
                    MsgBox("Please specify the layout for viewport insertion!")
                    Freeze_operations = False
                    Exit Sub
                End If
                If IsNumeric(TextBox_Layout_index_END.Text) = False Then
                    MsgBox("Please specify the layout for viewport insertion!")
                    Freeze_operations = False
                    Exit Sub
                End If


                Dim Sta1 As Double
                Dim Sta2 As Double
                Dim Hscale As Double = CDbl(TextBox_H_SCALE.Text)
                Dim Vscale As Double = CDbl(TextBox_V_SCALE.Text)
                Dim Viewport_scale As Double = CDbl(TextBox_viewport_scale.Text)


                If CheckBox_USA_style.Checked = True Then
                    Hscale = Hscale * 1000
                    Vscale = Vscale * 1000
                End If

                Dim H1 As Double = CDbl(TextBox_H1.Text)
                Dim H2 As Double = CDbl(TextBox_H2.Text)
                Dim W1 As Double = CDbl(TextBox_W3.Text)
                Dim W2 As Double = CDbl(TextBox_W2.Text)


                Dim Decimals As Integer = 0
                If IsNumeric(TextBox_dec1.Text) Then
                    Decimals = CInt(TextBox_dec1.Text)
                End If

                If Hscale <= 0 Or Vscale <= 0 Or H1 <= 0 Or H2 <= 0 Or W1 <= 0 Or W2 <= 0 Or Viewport_scale <= 0 Then
                    MsgBox("Negative values not allowed")
                    Freeze_operations = False
                    Exit Sub
                End If

                If ListBox_DWG.Items.Count > 0 Then
                    Using lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument

                        Using Trans11 As Transaction = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.TransactionManager.StartTransaction
                            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                            For i = 0 To ListBox_DWG.Items.Count - 1
                                Dim Match_string As String = ListBox_sheet_numbers.Items(i)

                                If Match_string.Contains(" - ") = True Then
                                    Dim Poz1 As Integer = InStr(Match_string, " - ")
                                    Dim Part1 As String = Strings.Left(Match_string, Poz1 - 1)
                                    Dim Part2 As String = Strings.Mid(Match_string, Poz1 + 3)
                                    If IsNumeric(Part1) = True And IsNumeric(Part2) = True Then
                                        Sta1 = CDbl(Part1)
                                        Sta2 = CDbl(Part2)

                                        Dim Drawing1 As String = ListBox_DWG.Items(i)



                                        If IO.File.Exists(Drawing1) = True Then
                                            Dim Database1 As New Database(False, True)
                                            Database1.ReadDwgFile(Drawing1, IO.FileShare.ReadWrite, True, Nothing)
                                            HostApplicationServices.WorkingDatabase = Database1

                                            Creaza_layer_with_database(Database1, "VP", 4, "VIEWPORT", False)
                                            Creaza_layer_with_database(Database1, "NO PLOT", 41, "no plot", False)
                                            Creaza_layer_with_database(Database1, "GRID", 252, "grid", True)
                                            Dim Layer_match_text As String = "0"
                                            If Not TextBox_Match_layer.Text = "" Then
                                                Layer_match_text = TextBox_Match_layer.Text
                                            End If

                                            If CheckBox_label_match.Checked = True And Not Layer_match_text = "0" Then
                                                Creaza_layer_with_database(Database1, TextBox_Match_layer.Text, 7, Layer_match_text, True)
                                            End If

                                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction

                                                Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                                                Dim Layoutdict As DBDictionary = Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead)

                                                Dim BlockTable1 As BlockTable = Trans1.GetObject(Database1.BlockTableId, OpenMode.ForRead)

                                                For Each entry As DBDictionaryEntry In Layoutdict
                                                    Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForRead)

                                                    If Not Layout1.TabOrder = 0 And Layout1.TabOrder >= CInt(TextBox_Layout_index_START.Text) And Layout1.TabOrder <= CInt(TextBox_Layout_index_END.Text) Then
                                                        Dim BTrecordPS As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(Layout1.BlockTableRecordId, OpenMode.ForWrite)

                                                        If Layout1.TabOrder > 1 Then
                                                            LayoutManager1.CurrentLayout = Layout1.LayoutName
                                                        End If

                                                        Dim x As Double = CDbl(TextBox_x.Text)
                                                        Dim y As Double = CDbl(TextBox_y.Text)

                                                        Dim Point0 As Point3d
                                                        Dim Point1 As Point3d
                                                        Dim Point2 As Point3d
                                                        Dim Point3 As Point3d

                                                        If RadioButton_left_right.Checked = True Then
                                                            Point0 = New Point3d(Point_cunoscut.X - Chainage_cunoscuta * 1000 / Hscale, Point_cunoscut.Y - Elevatia_cunoscuta * 1000 / Vscale, 0)
                                                            Point1 = New Point3d(Point0.X + Sta1 * 1000 / Hscale, Point0.Y + Elevatia_cunoscuta * 1000 / Vscale, 0)
                                                            Point2 = New Point3d(Point0.X + Sta2 * 1000 / Hscale, Point0.Y + Elevatia_cunoscuta * 1000 / Vscale, 0)
                                                        Else
                                                            Point0 = New Point3d(Point_cunoscut.X + Chainage_cunoscuta * 1000 / Hscale, Point_cunoscut.Y - Elevatia_cunoscuta * 1000 / Vscale, 0)
                                                            Point1 = New Point3d(Point0.X - Sta1 * 1000 / Hscale, Point0.Y + Elevatia_cunoscuta * 1000 / Vscale, 0)
                                                            Point2 = New Point3d(Point0.X - Sta2 * 1000 / Hscale, Point0.Y + Elevatia_cunoscuta * 1000 / Vscale, 0)
                                                        End If
                                                        Point3 = New Point3d((Point1.X + Point2.X) / 2, (Point1.Y + Point2.Y) / 2, 0)

                                                        Dim Poly1 As New Polyline
                                                        For j = 0 To Data_table_poly.Rows.Count - 1
                                                            Poly1.AddVertexAt(j, New Point2d(Data_table_poly.Rows(j).Item("X"), Data_table_poly.Rows(j).Item("Y")), 0, 0, 0)
                                                        Next

                                                        Dim Linie1 As New Line(New Point3d(Point1.X, Point1.Y - 100000, 0), New Point3d(Point1.X, Point1.Y + 100000, 0))
                                                        Dim Linie2 As New Line(New Point3d(Point2.X, Point2.Y - 100000, 0), New Point3d(Point2.X, Point2.Y + 100000, 0))

                                                        Dim ColInt1 As New Point3dCollection
                                                        Dim ColInt2 As New Point3dCollection
                                                        Poly1.IntersectWith(Linie1, Intersect.OnBothOperands, ColInt1, IntPtr.Zero, IntPtr.Zero)
                                                        Poly1.IntersectWith(Linie2, Intersect.OnBothOperands, ColInt2, IntPtr.Zero, IntPtr.Zero)

                                                        If ColInt1.Count > 0 And ColInt2.Count > 0 Then
                                                            Point3 = New Point3d((Point1.X + Point2.X) / 2, (ColInt1(0).Y + ColInt2(0).Y) / 2, 0)
                                                        End If

                                                        Dim DeltaY2 As Double

                                                        If IsNumeric(TextBox_deltaY_Vw2.Text) = True Then
                                                            DeltaY2 = CDbl(TextBox_shiftY_viewport.Text)
                                                        End If
                                                        If Not DeltaY2 = 0 Then
                                                            Point3 = New Point3d(Point3.X, Point3.Y + DeltaY2, 0)
                                                        End If

                                                        Dim Point4 As New Point3d(Point3.X - Abs(W1 - W2) * (1000 / Hscale) / 2, Chainage_Y, 0)

                                                        If Elevation_X1 < Elevation_X2 Then
                                                            Dim Temp1 As Double = Elevation_X1
                                                            Elevation_X1 = Elevation_X2
                                                            Elevation_X2 = Temp1
                                                        End If

                                                        Dim Point5 As New Point3d(Elevation_X2, Point3.Y, 0)
                                                        Dim Point6 As New Point3d(Elevation_X1, Point3.Y, 0)

                                                        Dim DeltaX1, DeltaX2, DeltaY As Double
                                                        If IsNumeric(TextBox_deltaY_Vw2.Text) = True Then
                                                            DeltaY = CDbl(TextBox_deltaY_Vw2.Text)
                                                        End If
                                                        If Not DeltaY = 0 Then
                                                            Point4 = New Point3d(Point4.X, Point4.Y + DeltaY, 0)
                                                        End If



                                                        If IsNumeric(TextBox_delta_x1.Text) = True Then
                                                            DeltaX1 = CDbl(TextBox_delta_x1.Text)
                                                        End If
                                                        If IsNumeric(TextBox_delta_x2.Text) = True Then
                                                            DeltaX2 = CDbl(TextBox_delta_x2.Text)
                                                        End If


                                                        If Not DeltaX2 = 0 Then
                                                            Point5 = New Point3d(Point5.X + DeltaX2, Point5.Y, 0)
                                                        End If
                                                        If Not DeltaX1 = 0 Then
                                                            Point6 = New Point3d(Point6.X + DeltaX1, Point6.Y, 0)
                                                        End If


                                                        Dim ExtraL As Double = 10

                                                        If CheckBox_USA_style.Checked = True Then ExtraL = 0

                                                        Dim L1 As Double = (Abs(Point1.X - Point2.X) + 2 * ExtraL) * Viewport_scale

                                                        Dim Spacing1 As Double = 0
                                                        If IsNumeric(TextBox_Viewport_spacing1.Text) = True Then
                                                            Spacing1 = CDbl(TextBox_Viewport_spacing1.Text)
                                                        End If

                                                        Dim Spacing2 As Double = 0
                                                        If IsNumeric(TextBox_Viewport_spacing2.Text) = True Then
                                                            Spacing2 = CDbl(TextBox_Viewport_spacing2.Text)
                                                        End If

                                                        If CheckBox_pick_middle.Checked = True Then
                                                            x = x - (L1 / 2 + (W2 + W1) / 2 + (Spacing1 + Spacing2) / 2)
                                                        End If

                                                        Dim Viewport1 As New Viewport
                                                        Viewport1.SetDatabaseDefaults()
                                                        Viewport1.CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(x + L1 / 2 + (W2 + W1) / 2 + (Spacing1 + Spacing2) / 2, y + H1 + H2 / 2, 0) ' asta e pozitia viewport in paper space
                                                        Viewport1.Height = H2
                                                        Viewport1.Width = L1
                                                        Viewport1.Layer = "GRID"
                                                        'Viewport1.ColorIndex = 1
                                                        Viewport1.ViewDirection = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis
                                                        Viewport1.ViewTarget = Point3 ' asta e pozitia viewport in MODEL space
                                                        Viewport1.ViewCenter = Autodesk.AutoCAD.Geometry.Point2d.Origin
                                                        Viewport1.TwistAngle = 0 ' asta e PT TWIST

                                                        BTrecordPS.AppendEntity(Viewport1)
                                                        Trans1.AddNewlyCreatedDBObject(Viewport1, True)

                                                        'Viewport1.On = True
                                                        Viewport1.CustomScale = Viewport_scale
                                                        Viewport1.Locked = True

                                                        Dim Viewport2 As New Viewport
                                                        Viewport2.SetDatabaseDefaults()
                                                        Viewport2.CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(x + L1 / 2 + (W2 + W1) / 2 + (Spacing1 + Spacing2) / 2, y + H1 / 2, 0) ' asta e pozitia viewport in paper space
                                                        Viewport2.Height = H1
                                                        Viewport2.Width = L1
                                                        Viewport2.Layer = "NO PLOT"
                                                        'Viewport2.ColorIndex = 2
                                                        Viewport2.ViewDirection = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis
                                                        Viewport2.ViewTarget = Point4 ' asta e pozitia viewport in MODEL space
                                                        Viewport2.ViewCenter = Autodesk.AutoCAD.Geometry.Point2d.Origin
                                                        Viewport2.TwistAngle = 0 ' asta e PT TWIST

                                                        BTrecordPS.AppendEntity(Viewport2)
                                                        Trans1.AddNewlyCreatedDBObject(Viewport2, True)

                                                        ' Viewport2.On = True
                                                        Viewport2.CustomScale = Viewport_scale
                                                        Viewport2.Locked = True

                                                        Dim Viewport3 As New Viewport
                                                        Viewport3.SetDatabaseDefaults()
                                                        Viewport3.CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(x + W2 / 2, y + H1 + H2 / 2, 0) ' asta e pozitia viewport in paper space
                                                        Viewport3.Height = H2
                                                        Viewport3.Width = W2
                                                        Viewport3.Layer = "NO PLOT"
                                                        'Viewport3.ColorIndex = 3
                                                        Viewport3.ViewDirection = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis
                                                        Viewport3.ViewTarget = Point5 ' asta e pozitia viewport in MODEL space
                                                        Viewport3.ViewCenter = Autodesk.AutoCAD.Geometry.Point2d.Origin
                                                        Viewport3.TwistAngle = 0 ' asta e PT TWIST

                                                        BTrecordPS.AppendEntity(Viewport3)
                                                        Trans1.AddNewlyCreatedDBObject(Viewport3, True)

                                                        'Viewport3.On = True
                                                        Viewport3.CustomScale = Viewport_scale
                                                        Viewport3.Locked = True

                                                        Dim Viewport4 As New Viewport
                                                        Viewport4.SetDatabaseDefaults()
                                                        Viewport4.CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(x + W2 + L1 + W1 / 2 + Spacing1 + Spacing2, y + H1 + H2 / 2, 0) ' asta e pozitia viewport in paper space
                                                        Viewport4.Height = H2
                                                        Viewport4.Width = W1
                                                        Viewport4.Layer = "NO PLOT"
                                                        'Viewport4.ColorIndex = 4
                                                        Viewport4.ViewDirection = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis
                                                        Viewport4.ViewTarget = Point6 ' asta e pozitia viewport in MODEL space
                                                        Viewport4.ViewCenter = Autodesk.AutoCAD.Geometry.Point2d.Origin
                                                        Viewport4.TwistAngle = 0 ' asta e PT TWIST

                                                        BTrecordPS.AppendEntity(Viewport4)
                                                        Trans1.AddNewlyCreatedDBObject(Viewport4, True)

                                                        'Viewport4.On = True
                                                        Viewport4.CustomScale = Viewport_scale
                                                        Viewport4.Locked = True


                                                        If CheckBox_label_match.Checked = True Then

                                                            Dim Text_style_table As Autodesk.AutoCAD.DatabaseServices.TextStyleTable = Trans1.GetObject(Database1.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                                            Dim TextStyleID As ObjectId = Nothing



                                                            For Each Text_id As ObjectId In Text_style_table
                                                                Dim TextStyle1 As TextStyleTableRecord = Trans1.GetObject(Text_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                                                If TextStyle1.Name.ToUpper = TextBox_match_textstyle.Text.ToUpper Then
                                                                    TextStyleID = TextStyle1.ObjectId
                                                                    Exit For
                                                                End If
                                                            Next

                                                            Dim Mtext_sta1 As New MText
                                                            Mtext_sta1.Contents = "STA " & Get_chainage_feet_from_double(Sta1 + Get_equation_value(Sta1), Decimals)
                                                            Mtext_sta1.Attachment = AttachmentPoint.MiddleCenter
                                                            Mtext_sta1.Rotation = PI / 2

                                                            Dim th As Double = 16
                                                            If IsNumeric(TextBox_match_text_height.Text) = True Then
                                                                th = CDbl(TextBox_match_text_height.Text)
                                                            End If

                                                            Mtext_sta1.TextHeight = th
                                                            If IsNothing(TextStyleID) = False Then
                                                                Mtext_sta1.TextStyleId = TextStyleID
                                                            End If

                                                            Mtext_sta1.Layer = Layer_match_text

                                                            Dim dM As Double = Spacing1 / 2
                                                            If IsNumeric(TextBox_match_deltaX.Text) = True Then
                                                                dM = CDbl(TextBox_match_deltaX.Text)
                                                            End If

                                                            Dim Pt1 As New Point3d(x + W2 + dM, y + H1 + H2 / 2, 0)
                                                            Mtext_sta1.Location = Pt1

                                                            BTrecordPS.AppendEntity(Mtext_sta1)
                                                            Trans1.AddNewlyCreatedDBObject(Mtext_sta1, True)

                                                            If CheckBox_label_stationing.Checked = True Then
                                                                Dim Mtext_sta3 As New MText
                                                                Mtext_sta3.Contents = "ELEVATION"
                                                                Mtext_sta3.Attachment = AttachmentPoint.MiddleCenter
                                                                Mtext_sta3.Rotation = PI / 2

                                                                Mtext_sta3.TextHeight = th
                                                                If IsNothing(TextStyleID) = False Then
                                                                    Mtext_sta3.TextStyleId = TextStyleID
                                                                End If

                                                                Mtext_sta3.Layer = TextBox_Match_layer.Text

                                                                Dim Pt3 As New Point3d(x - th, y + H1 + H2 / 2, 0)
                                                                Mtext_sta3.Location = Pt3

                                                                BTrecordPS.AppendEntity(Mtext_sta3)
                                                                Trans1.AddNewlyCreatedDBObject(Mtext_sta3, True)


                                                            End If


                                                            Dim Mtext_sta2 As New MText
                                                            Mtext_sta2.Contents = "STA " & Get_chainage_feet_from_double(Sta2 + Get_equation_value(Sta2), Decimals)
                                                            Mtext_sta2.Attachment = AttachmentPoint.MiddleCenter
                                                            Mtext_sta2.Rotation = PI / 2

                                                            Mtext_sta2.TextHeight = th
                                                            If IsNothing(TextStyleID) = False Then
                                                                Mtext_sta2.TextStyleId = TextStyleID
                                                            End If

                                                            Mtext_sta2.Layer = Layer_match_text

                                                            Dim dM2 As Double = Spacing2 / 2
                                                            If IsNumeric(TextBox_match_deltaX.Text) = True Then
                                                                dM2 = CDbl(TextBox_match_deltaX.Text)
                                                            End If

                                                            Dim Pt2 As New Point3d(x + W2 + Spacing1 + L1 + dM2, y + H1 + H2 / 2, 0)
                                                            Mtext_sta2.Location = Pt2

                                                            BTrecordPS.AppendEntity(Mtext_sta2)
                                                            Trans1.AddNewlyCreatedDBObject(Mtext_sta2, True)


                                                            If CheckBox_label_stationing.Checked = True Then
                                                                Dim Mtext_sta3 As New MText
                                                                Mtext_sta3.Contents = "ELEVATION"
                                                                Mtext_sta3.Attachment = AttachmentPoint.MiddleCenter
                                                                Mtext_sta3.Rotation = PI / 2

                                                                Mtext_sta3.TextHeight = th
                                                                If IsNothing(TextStyleID) = False Then
                                                                    Mtext_sta3.TextStyleId = TextStyleID
                                                                End If

                                                                Mtext_sta3.Layer = TextBox_Match_layer.Text
                                                                Dim Pt3 As New Point3d(x + W2 + Spacing1 + L1 + Spacing2 + W1 + th, y + H1 + H2 / 2, 0)
                                                                Mtext_sta3.Location = Pt3


                                                                BTrecordPS.AppendEntity(Mtext_sta3)
                                                                Trans1.AddNewlyCreatedDBObject(Mtext_sta3, True)

                                                                Dim Mtext_sta4 As New MText
                                                                Mtext_sta4.Contents = "STATIONING"
                                                                Mtext_sta4.Attachment = AttachmentPoint.MiddleCenter
                                                                Mtext_sta4.Rotation = 0

                                                                Mtext_sta4.TextHeight = th
                                                                If IsNothing(TextStyleID) = False Then
                                                                    Mtext_sta4.TextStyleId = TextStyleID
                                                                End If

                                                                Mtext_sta4.Layer = TextBox_Match_layer.Text


                                                                Dim Pt4 As New Point3d(x + W2 + Spacing1 + L1 / 2, y + th / 4, 0)
                                                                Mtext_sta4.Location = Pt4

                                                                BTrecordPS.AppendEntity(Mtext_sta4)
                                                                Trans1.AddNewlyCreatedDBObject(Mtext_sta4, True)


                                                            End If


                                                        End If
                                                    End If
                                                Next
                                                Trans1.Commit()
                                            End Using


                                            Database1.SaveAs(Drawing1, True, DwgVersion.Current, Database1.SecurityParameters)
                                            Database1.Dispose()
                                            HostApplicationServices.WorkingDatabase = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database
                                        End If

                                    End If
                                End If
                            Next

                            Trans11.Commit()
                        End Using
                    End Using
                End If





            Catch ex As Exception
                MsgBox(ex.Message)

            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: DONE!")
            Freeze_operations = False

        End If
    End Sub

    Private Sub Button_layout_Insert_viewport_Click(sender As Object, e As EventArgs) Handles Button_layout_Insert_viewport.Click


        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                If Not ListBox_DWG.Items.Count = ListBox_sheet_numbers.Items.Count Then
                    MsgBox("Please be sure numbers of DWG is equal to the number of bands")
                    Freeze_operations = False
                    Exit Sub
                End If

                If IsNothing(Data_table_poly) = True Then
                    MsgBox("Please LOAD THE graph")
                    Freeze_operations = False
                    Exit Sub
                End If

                If IsNumeric(TextBox_H_SCALE.Text) = False Then
                    MsgBox("Please specify the HORIZONTAL SCALE!")
                    Freeze_operations = False
                    Exit Sub
                End If
                If IsNumeric(TextBox_H1.Text) = False Then
                    MsgBox("Please specify the HEIGHT!")
                    Freeze_operations = False
                    Exit Sub
                End If
                If IsNumeric(TextBox_H2.Text) = False Then
                    MsgBox("Please specify the HEIGHT!")
                    Freeze_operations = False
                    Exit Sub
                End If

                If IsNumeric(TextBox_V_SCALE.Text) = False Then
                    MsgBox("Please specify the VERTICAL SCALE!")
                    Freeze_operations = False
                    Exit Sub
                End If
                If IsNumeric(TextBox_W3.Text) = False Then
                    MsgBox("Please specify the WIDTH!")
                    Freeze_operations = False
                    Exit Sub
                End If
                If IsNumeric(TextBox_W2.Text) = False Then
                    MsgBox("Please specify the WIDTH!")
                    Freeze_operations = False
                    Exit Sub
                End If
                If IsNumeric(TextBox_x.Text) = False Then
                    MsgBox("Please specify the X COORDINATE!")
                    Freeze_operations = False
                    Exit Sub
                End If
                If IsNumeric(TextBox_y.Text) = False Then
                    MsgBox("Please specify the Y COORDINATE!")
                    Freeze_operations = False
                    Exit Sub
                End If

                If IsNumeric(TextBox_viewport_scale.Text) = False Then
                    Freeze_operations = False
                    MsgBox("Please specify the viewport_scale")
                    Exit Sub
                End If



                Dim Sta1 As Double
                Dim Sta2 As Double
                Dim Hscale As Double = CDbl(TextBox_H_SCALE.Text)
                Dim Vscale As Double = CDbl(TextBox_V_SCALE.Text)
                Dim Viewport_scale As Double = CDbl(TextBox_viewport_scale.Text)


                If CheckBox_USA_style.Checked = True Then
                    Hscale = Hscale * 1000
                    Vscale = Vscale * 1000
                End If

                Dim H1 As Double = CDbl(TextBox_H1.Text)
                Dim H2 As Double = CDbl(TextBox_H2.Text)
                Dim W1 As Double = CDbl(TextBox_W3.Text)
                Dim W2 As Double = CDbl(TextBox_W2.Text)


                Dim Decimals As Integer = 0
                If IsNumeric(TextBox_dec1.Text) Then
                    Decimals = CInt(TextBox_dec1.Text)
                End If

                If Hscale <= 0 Or Vscale <= 0 Or H1 <= 0 Or H2 <= 0 Or W1 <= 0 Or W2 <= 0 Or Viewport_scale <= 0 Then
                    MsgBox("Negative values not allowed")
                    Freeze_operations = False
                    Exit Sub
                End If

                If ListBox_DWG.Items.Count > 0 Then
                    Using lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument

                        Using Trans1 As Transaction = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.TransactionManager.StartTransaction
                            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                            Dim Database1 As Autodesk.AutoCAD.DatabaseServices.Database = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database

                            Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                            Dim Layoutdict As DBDictionary = Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead)
                            Dim BlockTable1 As BlockTable = Trans1.GetObject(Database1.BlockTableId, OpenMode.ForRead)

                            Creaza_layer_with_database(Database1, "VP", 4, "VIEWPORT", False)
                            Creaza_layer_with_database(Database1, "NO PLOT", 41, "no plot", False)
                            Creaza_layer_with_database(Database1, "GRID", 252, "grid", True)
                            Dim Layer_match_text As String = "0"
                            If Not TextBox_Match_layer.Text = "" Then
                                Layer_match_text = TextBox_Match_layer.Text
                            End If
                            If CheckBox_label_match.Checked = True And Not Layer_match_text = "0" Then
                                Creaza_layer_with_database(Database1, TextBox_Match_layer.Text, 7, Layer_match_text, True)
                            End If

                            Dim Text_style_table As Autodesk.AutoCAD.DatabaseServices.TextStyleTable = Trans1.GetObject(Database1.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                            For i = 0 To ListBox_DWG.Items.Count - 1
                                Dim Match_string As String = ListBox_sheet_numbers.Items(i)

                                If Match_string.Contains(" - ") = True Then
                                    Dim Poz1 As Integer = InStr(Match_string, " - ")
                                    Dim Part1 As String = Strings.Left(Match_string, Poz1 - 1)
                                    Dim Part2 As String = Strings.Mid(Match_string, Poz1 + 3)
                                    If IsNumeric(Part1) = True And IsNumeric(Part2) = True Then
                                        Sta1 = CDbl(Part1)
                                        Sta2 = CDbl(Part2)

                                        For Each entry As DBDictionaryEntry In Layoutdict
                                            Using Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForRead)
                                                If Layout1.LayoutName = ListBox_DWG.Items(i) Then
                                                    Using BTrecordPS As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(Layout1.BlockTableRecordId, OpenMode.ForWrite)

                                                        LayoutManager1.CurrentLayout = Layout1.LayoutName


                                                        Dim x As Double = CDbl(TextBox_x.Text)
                                                        Dim y As Double = CDbl(TextBox_y.Text)

                                                        Dim Point0 As Point3d
                                                        Dim Point1 As Point3d
                                                        Dim Point2 As Point3d
                                                        Dim Point3 As Point3d

                                                        If RadioButton_left_right.Checked = True Then
                                                            Point0 = New Point3d(Point_cunoscut.X - Chainage_cunoscuta * 1000 / Hscale, Point_cunoscut.Y - Elevatia_cunoscuta * 1000 / Vscale, 0)
                                                            Point1 = New Point3d(Point0.X + Sta1 * 1000 / Hscale, Point0.Y + Elevatia_cunoscuta * 1000 / Vscale, 0)
                                                            Point2 = New Point3d(Point0.X + Sta2 * 1000 / Hscale, Point0.Y + Elevatia_cunoscuta * 1000 / Vscale, 0)
                                                        Else
                                                            Point0 = New Point3d(Point_cunoscut.X + Chainage_cunoscuta * 1000 / Hscale, Point_cunoscut.Y - Elevatia_cunoscuta * 1000 / Vscale, 0)
                                                            Point1 = New Point3d(Point0.X - Sta1 * 1000 / Hscale, Point0.Y + Elevatia_cunoscuta * 1000 / Vscale, 0)
                                                            Point2 = New Point3d(Point0.X - Sta2 * 1000 / Hscale, Point0.Y + Elevatia_cunoscuta * 1000 / Vscale, 0)
                                                        End If
                                                        Point3 = New Point3d((Point1.X + Point2.X) / 2, (Point1.Y + Point2.Y) / 2, 0)

                                                        Dim Poly1 As New Polyline
                                                        For j = 0 To Data_table_poly.Rows.Count - 1
                                                            Poly1.AddVertexAt(j, New Point2d(Data_table_poly.Rows(j).Item("X"), Data_table_poly.Rows(j).Item("Y")), 0, 0, 0)
                                                        Next

                                                        Dim Linie1 As New Line(New Point3d(Point1.X, Point1.Y - 100000, 0), New Point3d(Point1.X, Point1.Y + 100000, 0))
                                                        Dim Linie2 As New Line(New Point3d(Point2.X, Point2.Y - 100000, 0), New Point3d(Point2.X, Point2.Y + 100000, 0))

                                                        Dim ColInt1 As New Point3dCollection
                                                        Dim ColInt2 As New Point3dCollection
                                                        Poly1.IntersectWith(Linie1, Intersect.OnBothOperands, ColInt1, IntPtr.Zero, IntPtr.Zero)
                                                        Poly1.IntersectWith(Linie2, Intersect.OnBothOperands, ColInt2, IntPtr.Zero, IntPtr.Zero)

                                                        If ColInt1.Count > 0 And ColInt2.Count > 0 Then
                                                            Point3 = New Point3d((Point1.X + Point2.X) / 2, (ColInt1(0).Y + ColInt2(0).Y) / 2, 0)
                                                        End If

                                                        Dim DeltaY2 As Double

                                                        If IsNumeric(TextBox_deltaY_Vw2.Text) = True Then
                                                            DeltaY2 = CDbl(TextBox_shiftY_viewport.Text)
                                                        End If
                                                        If Not DeltaY2 = 0 Then
                                                            Point3 = New Point3d(Point3.X, Point3.Y + DeltaY2, 0)
                                                        End If

                                                        Dim Point4 As New Point3d(Point3.X - Abs(W1 - W2) * (1000 / Hscale) / 2, Chainage_Y, 0)

                                                        If Elevation_X1 < Elevation_X2 Then
                                                            Dim Temp1 As Double = Elevation_X1
                                                            Elevation_X1 = Elevation_X2
                                                            Elevation_X2 = Temp1
                                                        End If

                                                        Dim Point5 As New Point3d(Elevation_X2, Point3.Y, 0)
                                                        Dim Point6 As New Point3d(Elevation_X1, Point3.Y, 0)

                                                        Dim DeltaX1, DeltaX2, DeltaY As Double
                                                        If IsNumeric(TextBox_deltaY_Vw2.Text) = True Then
                                                            DeltaY = CDbl(TextBox_deltaY_Vw2.Text)
                                                        End If
                                                        If Not DeltaY = 0 Then
                                                            Point4 = New Point3d(Point4.X, Point4.Y + DeltaY, 0)
                                                        End If



                                                        If IsNumeric(TextBox_delta_x1.Text) = True Then
                                                            DeltaX1 = CDbl(TextBox_delta_x1.Text)
                                                        End If
                                                        If IsNumeric(TextBox_delta_x2.Text) = True Then
                                                            DeltaX2 = CDbl(TextBox_delta_x2.Text)
                                                        End If


                                                        If Not DeltaX2 = 0 Then
                                                            Point5 = New Point3d(Point5.X + DeltaX2, Point5.Y, 0)
                                                        End If
                                                        If Not DeltaX1 = 0 Then
                                                            Point6 = New Point3d(Point6.X + DeltaX1, Point6.Y, 0)
                                                        End If


                                                        Dim ExtraL As Double = 10

                                                        If CheckBox_USA_style.Checked = True Then ExtraL = 0

                                                        Dim L1 As Double = (Abs(Point1.X - Point2.X) + 2 * ExtraL) * Viewport_scale

                                                        Dim Spacing1 As Double = 0
                                                        If IsNumeric(TextBox_Viewport_spacing1.Text) = True Then
                                                            Spacing1 = CDbl(TextBox_Viewport_spacing1.Text)
                                                        End If

                                                        Dim Spacing2 As Double = 0
                                                        If IsNumeric(TextBox_Viewport_spacing2.Text) = True Then
                                                            Spacing2 = CDbl(TextBox_Viewport_spacing2.Text)
                                                        End If

                                                        If CheckBox_pick_middle.Checked = True Then
                                                            x = x - (L1 / 2 + (W2 + W1) / 2 + (Spacing1 + Spacing2) / 2)
                                                        End If

                                                        Dim Viewport1 As New Viewport
                                                        Viewport1.SetDatabaseDefaults()
                                                        Viewport1.CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(x + L1 / 2 + (W2 + W1) / 2 + (Spacing1 + Spacing2) / 2, y + H1 + H2 / 2, 0) ' asta e pozitia viewport in paper space
                                                        Viewport1.Height = H2
                                                        Viewport1.Width = L1
                                                        Viewport1.Layer = "GRID"
                                                        'Viewport1.ColorIndex = 1
                                                        Viewport1.ViewDirection = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis
                                                        Viewport1.ViewTarget = Point3 ' asta e pozitia viewport in MODEL space
                                                        Viewport1.ViewCenter = Autodesk.AutoCAD.Geometry.Point2d.Origin
                                                        Viewport1.TwistAngle = 0 ' asta e PT TWIST

                                                        BTrecordPS.AppendEntity(Viewport1)
                                                        Trans1.AddNewlyCreatedDBObject(Viewport1, True)

                                                        Viewport1.On = True
                                                        Viewport1.CustomScale = Viewport_scale
                                                        Viewport1.Locked = True

                                                        Dim Viewport2 As New Viewport
                                                        Viewport2.SetDatabaseDefaults()
                                                        Viewport2.CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(x + L1 / 2 + (W2 + W1) / 2 + (Spacing1 + Spacing2) / 2, y + H1 / 2, 0) ' asta e pozitia viewport in paper space
                                                        Viewport2.Height = H1
                                                        Viewport2.Width = L1
                                                        Viewport2.Layer = "NO PLOT"
                                                        'Viewport2.ColorIndex = 2
                                                        Viewport2.ViewDirection = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis
                                                        Viewport2.ViewTarget = Point4 ' asta e pozitia viewport in MODEL space
                                                        Viewport2.ViewCenter = Autodesk.AutoCAD.Geometry.Point2d.Origin
                                                        Viewport2.TwistAngle = 0 ' asta e PT TWIST

                                                        BTrecordPS.AppendEntity(Viewport2)
                                                        Trans1.AddNewlyCreatedDBObject(Viewport2, True)

                                                        Viewport2.On = True
                                                        Viewport2.CustomScale = Viewport_scale
                                                        Viewport2.Locked = True

                                                        Dim Viewport3 As New Viewport
                                                        Viewport3.SetDatabaseDefaults()
                                                        Viewport3.CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(x + W2 / 2, y + H1 + H2 / 2, 0) ' asta e pozitia viewport in paper space
                                                        Viewport3.Height = H2
                                                        Viewport3.Width = W2
                                                        Viewport3.Layer = "NO PLOT"
                                                        'Viewport3.ColorIndex = 3
                                                        Viewport3.ViewDirection = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis
                                                        Viewport3.ViewTarget = Point5 ' asta e pozitia viewport in MODEL space
                                                        Viewport3.ViewCenter = Autodesk.AutoCAD.Geometry.Point2d.Origin
                                                        Viewport3.TwistAngle = 0 ' asta e PT TWIST

                                                        BTrecordPS.AppendEntity(Viewport3)
                                                        Trans1.AddNewlyCreatedDBObject(Viewport3, True)

                                                        Viewport3.On = True
                                                        Viewport3.CustomScale = Viewport_scale
                                                        Viewport3.Locked = True

                                                        Dim Viewport4 As New Viewport
                                                        Viewport4.SetDatabaseDefaults()
                                                        Viewport4.CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(x + W2 + L1 + W1 / 2 + Spacing1 + Spacing2, y + H1 + H2 / 2, 0) ' asta e pozitia viewport in paper space
                                                        Viewport4.Height = H2
                                                        Viewport4.Width = W1
                                                        Viewport4.Layer = "NO PLOT"
                                                        'Viewport4.ColorIndex = 4
                                                        Viewport4.ViewDirection = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis
                                                        Viewport4.ViewTarget = Point6 ' asta e pozitia viewport in MODEL space
                                                        Viewport4.ViewCenter = Autodesk.AutoCAD.Geometry.Point2d.Origin
                                                        Viewport4.TwistAngle = 0 ' asta e PT TWIST

                                                        BTrecordPS.AppendEntity(Viewport4)
                                                        Trans1.AddNewlyCreatedDBObject(Viewport4, True)

                                                        Viewport4.On = True
                                                        Viewport4.CustomScale = Viewport_scale
                                                        Viewport4.Locked = True


                                                        If CheckBox_label_match.Checked = True Then
                                                            Dim TextStyleID As ObjectId = Nothing
                                                            For Each Text_id As ObjectId In Text_style_table
                                                                Dim TextStyle1 As TextStyleTableRecord = Trans1.GetObject(Text_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                                                If TextStyle1.Name.ToUpper = TextBox_match_textstyle.Text.ToUpper Then
                                                                    TextStyleID = TextStyle1.ObjectId
                                                                    Exit For
                                                                End If
                                                            Next

                                                            Dim Mtext_sta1 As New MText
                                                            Mtext_sta1.Contents = "STA " & Get_chainage_feet_from_double(Sta1 + Get_equation_value(Sta1), Decimals)
                                                            Mtext_sta1.Attachment = AttachmentPoint.MiddleCenter
                                                            Mtext_sta1.Rotation = PI / 2

                                                            Dim th As Double = 16
                                                            If IsNumeric(TextBox_match_text_height.Text) = True Then
                                                                th = CDbl(TextBox_match_text_height.Text)
                                                            End If

                                                            Mtext_sta1.TextHeight = th
                                                            If IsNothing(TextStyleID) = False Then
                                                                Mtext_sta1.TextStyleId = TextStyleID
                                                            End If

                                                            Mtext_sta1.Layer = Layer_match_text

                                                            Dim dM As Double = Spacing1 / 2
                                                            If IsNumeric(TextBox_match_deltaX.Text) = True Then
                                                                dM = CDbl(TextBox_match_deltaX.Text)
                                                            End If

                                                            Dim Pt1 As New Point3d(x + W2 + dM, y + H1 + H2 / 2, 0)
                                                            Mtext_sta1.Location = Pt1

                                                            BTrecordPS.AppendEntity(Mtext_sta1)
                                                            Trans1.AddNewlyCreatedDBObject(Mtext_sta1, True)

                                                            If CheckBox_label_stationing.Checked = True Then
                                                                Dim Mtext_sta3 As New MText
                                                                Mtext_sta3.Contents = "ELEVATION"
                                                                Mtext_sta3.Attachment = AttachmentPoint.MiddleCenter
                                                                Mtext_sta3.Rotation = PI / 2

                                                                Mtext_sta3.TextHeight = th
                                                                If IsNothing(TextStyleID) = False Then
                                                                    Mtext_sta3.TextStyleId = TextStyleID
                                                                End If

                                                                Mtext_sta3.Layer = TextBox_Match_layer.Text

                                                                Dim Pt3 As New Point3d(x - th, y + H1 + H2 / 2, 0)
                                                                Mtext_sta3.Location = Pt3

                                                                BTrecordPS.AppendEntity(Mtext_sta3)
                                                                Trans1.AddNewlyCreatedDBObject(Mtext_sta3, True)
                                                            End If


                                                            Dim Mtext_sta2 As New MText
                                                            Mtext_sta2.Contents = "STA " & Get_chainage_feet_from_double(Sta2 + Get_equation_value(Sta2), Decimals)
                                                            Mtext_sta2.Attachment = AttachmentPoint.MiddleCenter
                                                            Mtext_sta2.Rotation = PI / 2

                                                            Mtext_sta2.TextHeight = th
                                                            If IsNothing(TextStyleID) = False Then
                                                                Mtext_sta2.TextStyleId = TextStyleID
                                                            End If

                                                            Mtext_sta2.Layer = Layer_match_text

                                                            Dim dM2 As Double = Spacing2 / 2
                                                            If IsNumeric(TextBox_match_deltaX.Text) = True Then
                                                                dM2 = CDbl(TextBox_match_deltaX.Text)
                                                            End If

                                                            Dim Pt2 As New Point3d(x + W2 + Spacing1 + L1 + dM2, y + H1 + H2 / 2, 0)
                                                            Mtext_sta2.Location = Pt2

                                                            BTrecordPS.AppendEntity(Mtext_sta2)
                                                            Trans1.AddNewlyCreatedDBObject(Mtext_sta2, True)


                                                            If CheckBox_label_stationing.Checked = True Then
                                                                Dim Mtext_sta3 As New MText
                                                                Mtext_sta3.Contents = "ELEVATION"
                                                                Mtext_sta3.Attachment = AttachmentPoint.MiddleCenter
                                                                Mtext_sta3.Rotation = PI / 2

                                                                Mtext_sta3.TextHeight = th
                                                                If IsNothing(TextStyleID) = False Then
                                                                    Mtext_sta3.TextStyleId = TextStyleID
                                                                End If

                                                                Mtext_sta3.Layer = TextBox_Match_layer.Text
                                                                Dim Pt3 As New Point3d(x + W2 + Spacing1 + L1 + Spacing2 + W1 + th, y + H1 + H2 / 2, 0)
                                                                Mtext_sta3.Location = Pt3


                                                                BTrecordPS.AppendEntity(Mtext_sta3)
                                                                Trans1.AddNewlyCreatedDBObject(Mtext_sta3, True)

                                                                Dim Mtext_sta4 As New MText
                                                                Mtext_sta4.Contents = "STATIONING"
                                                                Mtext_sta4.Attachment = AttachmentPoint.MiddleCenter
                                                                Mtext_sta4.Rotation = 0

                                                                Mtext_sta4.TextHeight = th
                                                                If IsNothing(TextStyleID) = False Then
                                                                    Mtext_sta4.TextStyleId = TextStyleID
                                                                End If

                                                                Mtext_sta4.Layer = TextBox_Match_layer.Text


                                                                Dim Pt4 As New Point3d(x + W2 + Spacing1 + L1 / 2, y + th / 4, 0)
                                                                Mtext_sta4.Location = Pt4

                                                                BTrecordPS.AppendEntity(Mtext_sta4)
                                                                Trans1.AddNewlyCreatedDBObject(Mtext_sta4, True)


                                                            End If


                                                        End If
                                                    End Using
                                                    Exit For
                                                End If
                                            End Using
                                        Next


                                    End If
                                End If
                            Next

                            Trans1.Commit()
                        End Using
                    End Using
                End If





            Catch ex As Exception
                MsgBox(ex.Message)

            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: DONE!")
            Freeze_operations = False

        End If
    End Sub


    Private Sub Button_load_equations_from_excel_Click(sender As Object, e As EventArgs) Handles Button_load_equations_from_excel.Click
        If Freeze_operations = False Then
            Freeze_operations = True

            Try

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Dim Start1 As Integer = 0
                Dim End1 As Integer = 0
                If IsNumeric(TextBox_Row_Start_eq.Text) = True Then
                    Start1 = CInt(TextBox_Row_Start_eq.Text)
                End If
                If IsNumeric(TextBox_Row_End_eq.Text) = True Then
                    End1 = CInt(TextBox_Row_End_eq.Text)
                End If

                If End1 = 0 Or Start1 = 0 Then
                    Freeze_operations = False
                    Exit Sub
                End If
                If End1 < Start1 Then
                    Freeze_operations = False
                    Exit Sub
                End If
                Dim Column_Sta_Back As String = ""
                Column_Sta_Back = TextBox_col_station_back.Text.ToUpper
                Dim Column_sta_ahead As String = ""
                Column_sta_ahead = TextBox_col_statation_ahead.Text.ToUpper

                Data_table_station_equation = New System.Data.DataTable
                Data_table_station_equation.Columns.Add("STATION_BACK", GetType(Double))
                Data_table_station_equation.Columns.Add("STATION_AHEAD", GetType(Double))


                Dim Index_data_table As Double



                For i = Start1 To End1
                    Dim Station_back As String = W1.Range(Column_Sta_Back & i).Value2
                    Dim Station_ahead As String = W1.Range(Column_sta_ahead & i).Value2
                    If IsNumeric(Station_ahead) = True And IsNumeric(Station_back) = True Then

                        Data_table_station_equation.Rows.Add()
                        Data_table_station_equation.Rows(Index_data_table).Item("STATION_BACK") = CDbl(Station_back)
                        Data_table_station_equation.Rows(Index_data_table).Item("STATION_AHEAD") = CDbl(Station_ahead)
                        Index_data_table = Index_data_table + 1

                    Else
                        MsgBox("non numerical values on row " & i)
                        W1.Rows(i).select()
                        Freeze_operations = False
                        Exit Sub

                    End If
                Next


                Data_table_station_equation = Sort_data_table(Data_table_station_equation, "STATION_BACK")

                'MsgBox(Data_table_Centerline.Rows.Count)



            Catch ex As Exception
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try
        End If
        Freeze_operations = False
    End Sub

    Private Sub Button_Clear_stat_eq_Click(sender As Object, e As EventArgs) Handles Button_Clear_stat_eq.Click
        If Freeze_operations = False Then
            Freeze_operations = True

            Try
                Data_table_station_equation = New System.Data.DataTable

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        Freeze_operations = False
    End Sub
    Private Sub Button_insert_blocks_paperspace_Click(sender As Object, e As EventArgs) Handles Button_insert_blocks_paperspace.Click



        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                If Not ListBox_DWG.Items.Count = ListBox_sheet_numbers.Items.Count Then
                    MsgBox("Please be sure numbers of DWG is equal to the number of bands")
                    Freeze_operations = False
                    Exit Sub
                End If


                If IsNumeric(TextBox_y_block_paperspace.Text) = False Then
                    MsgBox("Please specify the Y")
                    Freeze_operations = False
                    Exit Sub
                End If

                If Viewport_height = 0 Then
                    MsgBox("Not sample viewport loaded")
                    Freeze_operations = False
                    Exit Sub
                End If

                If IsNothing(Data_table_poly) = True Then
                    MsgBox("The centerline is not loaded")
                    Freeze_operations = False
                    Exit Sub
                End If


                Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
                ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
                Dim GraphicsManager As Autodesk.AutoCAD.GraphicsSystem.Manager = ThisDrawing.GraphicsManager
                If ListBox_DWG.Items.Count > 0 Then
                    Using lock1 As DocumentLock = ThisDrawing.LockDocument

                        Using Trans1 As Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                            Dim Database1 As Autodesk.AutoCAD.DatabaseServices.Database = ThisDrawing.Database
                            Dim Layoutdict As DBDictionary = Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead)
                            For Each entry As DBDictionaryEntry In Layoutdict
                                Using Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                    If Layout1.TabOrder > 0 Then
                                        If ListBox_DWG.Items.Contains(Layout1.LayoutName) = True Then
                                            LayoutManager1.CurrentLayout = Layout1.LayoutName
                                            Dim Tilemode1 As Integer = Application.GetSystemVariable("TILEMODE")
                                            Dim CVport1 As Integer = Application.GetSystemVariable("CVPORT")
                                            If Tilemode1 = 0 Then
                                                If Not CVport1 = 1 Then
                                                    Editor1.SwitchToPaperSpace()
                                                End If
                                            Else
                                                Application.SetSystemVariable("TILEMODE", 0)
                                            End If
                                            Using View1 As ViewTableRecord = ThisDrawing.Editor.GetCurrentView
                                                View1.SetUcs(Database1.Ucsorg, Database1.Ucsxdir, Database1.Ucsydir)
                                                View1.ViewDirection = New Point3d(0, 0, 1).GetAsVector
                                                View1.ViewTwist = 0
                                                Dim Ptmin As Point2d = New Point2d(ThisDrawing.Database.Pextmin.X, ThisDrawing.Database.Pextmin.Y)
                                                Dim Ptmax As Point2d = New Point2d(ThisDrawing.Database.Pextmax.X, ThisDrawing.Database.Pextmax.Y)
                                                Dim Size1 As Vector2d = Ptmax - Ptmin
                                                Dim Center1 As Point2d = Ptmin + Size1 / 2
                                                View1.CenterPoint = Center1
                                                View1.Width = Size1.X
                                                View1.Height = Size1.Y
                                                ThisDrawing.Editor.SetCurrentView(View1)
                                                ThisDrawing.Editor.Regen()
                                            End Using
                                        End If
                                    End If
                                End Using
                            Next
                            Trans1.Commit()
                        End Using


                        Using Trans1 As Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                            Dim Database1 As Autodesk.AutoCAD.DatabaseServices.Database = ThisDrawing.Database
                            Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                            Dim Layoutdict As DBDictionary = Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead)
                            Dim BlockTable1 As BlockTable = Trans1.GetObject(Database1.BlockTableId, OpenMode.ForRead)
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                            BTrecord = Trans1.GetObject(BlockTable1(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                            Dim Poly1 As New Polyline3d
                            BTrecord.AppendEntity(Poly1)
                            Trans1.AddNewlyCreatedDBObject(Poly1, True)

                            For i = 0 To Data_table_poly.Rows.Count - 1
                                Dim x1 As Double = Data_table_poly.Rows(i).Item("X")
                                Dim y1 As Double = Data_table_poly.Rows(i).Item("Y")
                                Dim z1 As Double = Data_table_poly.Rows(i).Item("Z")
                                Dim Vertex1 As PolylineVertex3d = New PolylineVertex3d(New Point3d(x1, y1, z1))
                                Poly1.AppendVertex(Vertex1)
                                Trans1.AddNewlyCreatedDBObject(Vertex1, True)
                            Next






                            For Each entry As DBDictionaryEntry In Layoutdict
                                Using Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                    If ListBox_DWG.Items.Contains(Layout1.LayoutName) = True Then
                                        Dim i As Integer = ListBox_DWG.Items.IndexOf(Layout1.LayoutName)

                                        Dim Match_string As String = ListBox_sheet_numbers.Items(i)
                                        If Match_string.Contains(" - ") = True Then
                                            Dim Poz1 As Integer = InStr(Match_string, " - ")
                                            Dim Part1 As String = Strings.Left(Match_string, Poz1 - 1)
                                            Dim Part2 As String = Strings.Mid(Match_string, Poz1 + 3)
                                            Dim Sta1 As Double = 0
                                            Dim Sta2 As Double = 0
                                            If IsNumeric(Part1) = True And IsNumeric(Part2) = True Then
                                                Sta1 = CDbl(Part1)
                                                Sta2 = CDbl(Part2)
                                                Using BTrecordPS As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(Layout1.BlockTableRecordId, OpenMode.ForWrite)
                                                    Dim y As Double = CDbl(TextBox_y_block_paperspace.Text)
                                                    Dim Scale1 As Double = 0
                                                    Dim Twist1 As Double = 0
                                                    Dim PointPS1 As New Point3d
                                                    Dim PointMS1 As New Point3d
                                                    Dim PointPS2 As New Point3d
                                                    Dim PointMS2 As New Point3d
                                                    Dim Viewport1 As Viewport
                                                    For Each objID1 As ObjectId In BTrecordPS
                                                        Dim Ent1 As Entity = Trans1.GetObject(objID1, OpenMode.ForRead)
                                                        If TypeOf Ent1 Is Viewport Then
                                                            Dim Vw1 As Viewport = Ent1
                                                            If Round(Vw1.Height, 0) = Round(Viewport_height, 0) Then
                                                                If Round(Vw1.Width, 0) = Round(Viewport_width, 0) Then
                                                                    If Round(Vw1.CustomScale, 2) = Round(Viewport_scale, 2) Then
                                                                        Twist1 = Vw1.TwistAngle
                                                                        Scale1 = Viewport_scale
                                                                        Dim Point_on_poly1 As New Point3d
                                                                        Dim Pt_cenPS As New Point3d
                                                                        Pt_cenPS = Vw1.CenterPoint
                                                                        If Poly1.Length >= Sta1 And Poly1.Length >= Sta2 Then
                                                                            LayoutManager1.CurrentLayout = Layout1.LayoutName
                                                                            PointMS1 = Poly1.GetPointAtDist(Sta1)
                                                                            PointMS2 = Poly1.GetPointAtDist(Sta2)
                                                                            Dim Point_target As New Point3d
                                                                            Editor1.SwitchToModelSpace()
                                                                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("CVPORT", Vw1.Number)
                                                                            Point_target = Application.GetSystemVariable("VIEWCTR")
                                                                            Editor1.SwitchToPaperSpace()
                                                                            PointPS1 = New Point3d(Pt_cenPS.X - (Point_target.X - PointMS1.X) * Scale1, Pt_cenPS.Y - (Point_target.Y - PointMS1.Y) * Scale1, 0)
                                                                            PointPS1 = PointPS1.TransformBy(Matrix3d.Rotation(Twist1, Vector3d.ZAxis, Pt_cenPS))
                                                                            PointPS2 = New Point3d(Pt_cenPS.X - (Point_target.X - PointMS2.X) * Scale1, Pt_cenPS.Y - (Point_target.Y - PointMS2.Y) * Scale1, 0)
                                                                            PointPS2 = PointPS2.TransformBy(Matrix3d.Rotation(Twist1, Vector3d.ZAxis, Pt_cenPS))
                                                                            InsertBlock_with_multiple_atributes("", ComboBox_blocks_left.Text, New Point3d(PointPS1.X, y, 0), 1, BTrecordPS, ComboBox_layer_blocks.Text, New Specialized.StringCollection, New Specialized.StringCollection)
                                                                            InsertBlock_with_multiple_atributes("", ComboBox_blocks_right.Text, New Point3d(PointPS2.X, y, 0), 1, BTrecordPS, ComboBox_layer_blocks.Text, New Specialized.StringCollection, New Specialized.StringCollection)
                                                                            Exit For
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    Next
                                                End Using
                                            End If
                                        End If
                                    End If
                                End Using
                            Next



                            Poly1.UpgradeOpen()
                            Poly1.Erase()
                            Trans1.Commit()
                        End Using
                    End Using
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: DONE!")
            Freeze_operations = False
        End If
    End Sub


    Private Sub Button_pick_Y_block_Click(sender As Object, e As EventArgs) Handles Button_pick_Y_block.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Try

                Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
                ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()


                Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                        Dim Pt_rezult As Autodesk.AutoCAD.EditorInput.PromptPointResult
                        Dim Prompt_pt As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify THE block insertion point")

                        Prompt_pt.AllowNone = True
                        Pt_rezult = Editor1.GetPoint(Prompt_pt)

                        If Pt_rezult.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            TextBox_y_block_paperspace.Text = Get_String_Rounded(Pt_rezult.Value.Y, 2)
                        End If



                    End Using
                End Using

            Catch ex As System.Exception
                MsgBox(ex.Message)

            End Try
        End If
        Freeze_operations = False
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
    End Sub

    Private Sub Button_load_Centerline_Click(sender As Object, e As EventArgs) Handles Button_load_Centerline.Click
        If Freeze_operations = False Then
            Freeze_operations = True


            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor


            Editor1.SetImpliedSelection(Empty_array)
            Try
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                Colectie1 = New Specialized.StringCollection


                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptEntityResult

                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select centerline:")

                Object_Prompt.SetRejectMessage(vbLf & "Please select a lightweight polyline or a 3d polyline")
                Object_Prompt.AddAllowedClass(GetType(Polyline), True)
                Object_Prompt.AddAllowedClass(GetType(Polyline3d), True)


                Rezultat1 = Editor1.GetEntity(Object_Prompt)


                If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.WriteMessage(vbLf & "Command:")
                    Freeze_operations = False
                    Editor1.SetImpliedSelection(Empty_array)
                    Exit Sub
                End If

                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                    If IsNothing(Rezultat1) = False Then

                        Using Lock_dwg As DocumentLock = ThisDrawing.LockDocument


                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction


                                Dim PolyCL_for_viewports As Polyline = TryCast(Trans1.GetObject(Rezultat1.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Polyline)
                                Dim PolyCL3D_for_viewports As Polyline3d = TryCast(Trans1.GetObject(Rezultat1.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Polyline3d)

                                If IsNothing(PolyCL_for_viewports) = False Or IsNothing(PolyCL3D_for_viewports) = False Then
                                    Data_table_poly = New System.Data.DataTable
                                    Data_table_poly.Columns.Add("X", GetType(Double))
                                    Data_table_poly.Columns.Add("Y", GetType(Double))
                                    Data_table_poly.Columns.Add("Z", GetType(Double))
                                End If




                                If IsNothing(PolyCL3D_for_viewports) = False Then
                                    Dim Index1 As Double = 0


                                    For Each vId As Autodesk.AutoCAD.DatabaseServices.ObjectId In PolyCL3D_for_viewports
                                        Dim v3d As Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d = DirectCast(Trans1.GetObject _
                                                (vId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d)

                                        Dim x1 As Double = v3d.Position.X
                                        Dim y1 As Double = v3d.Position.Y
                                        Dim z1 As Double = v3d.Position.Z
                                        Data_table_poly.Rows.Add()
                                        Data_table_poly.Rows(Index1).Item("X") = x1
                                        Data_table_poly.Rows(Index1).Item("Y") = y1
                                        Data_table_poly.Rows(Index1).Item("Z") = z1
                                        Index1 = Index1 + 1
                                    Next
                                End If

                                If IsNothing(PolyCL_for_viewports) = False Then


                                    For i = 0 To PolyCL_for_viewports.NumberOfVertices - 1
                                        Dim x1 As Double = PolyCL_for_viewports.GetPointAtParameter(i).X
                                        Dim y1 As Double = PolyCL_for_viewports.GetPointAtParameter(i).Y
                                        Data_table_poly.Rows.Add()
                                        Data_table_poly.Rows(i).Item("X") = x1
                                        Data_table_poly.Rows(i).Item("Y") = y1
                                        Data_table_poly.Rows(i).Item("Z") = 0

                                    Next
                                End If





                            End Using
                        End Using

                    End If
                End If



                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")

            Catch ex As Exception
                Editor1.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_read_sample_viewport_Click(sender As Object, e As EventArgs) Handles Button_read_sample_viewport.Click

        If Freeze_operations = False Then
            Freeze_operations = True


            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor


            Editor1.SetImpliedSelection(Empty_array)
            Try

                Colectie1 = New Specialized.StringCollection


                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptEntityResult

                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select the sample viewport:")

                Object_Prompt.SetRejectMessage(vbLf & "Please select a viewport")
                Object_Prompt.AddAllowedClass(GetType(Viewport), True)
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                Rezultat1 = Editor1.GetEntity(Object_Prompt)


                If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.WriteMessage(vbLf & "Command:")
                    Freeze_operations = False
                    Editor1.SetImpliedSelection(Empty_array)
                    Exit Sub
                End If

                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                    If IsNothing(Rezultat1) = False Then

                        Using Lock_dwg As DocumentLock = ThisDrawing.LockDocument


                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction


                                Dim Viewport1 As Viewport = TryCast(Trans1.GetObject(Rezultat1.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Viewport)


                                If IsNothing(Viewport1) = False Then
                                    Viewport_height = Viewport1.Height
                                    Viewport_width = Viewport1.Width
                                    Viewport_scale = Viewport1.CustomScale
                                End If



                            End Using
                        End Using

                    End If
                End If



                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")

            Catch ex As Exception
                Editor1.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try
            Freeze_operations = False
        End If
    End Sub
End Class