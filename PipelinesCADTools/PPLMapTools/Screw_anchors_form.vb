Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class Screw_anchors_form
    Dim Colectie1 As New Specialized.StringCollection

    Public Sub Button_pick_Click() Handles Button_pick.Click
        Dim Empty_array() As ObjectId

        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        Using lock As DocumentLock = ThisDrawing.LockDocument
            ascunde_butoanele_pentru_forms(Me, Colectie1)
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Try
                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select OLD polyline:"

                Object_Prompt.SingleOnly = True

                Rezultat1 = Editor1.GetSelection(Object_Prompt)


                If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.WriteMessage(vbLf & "Command:")
                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                    Exit Sub
                End If

                Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt2.MessageForAdding = vbLf & "Select NEW polyline:"

                Object_Prompt2.SingleOnly = True

                Rezultat2 = Editor1.GetSelection(Object_Prompt2)


                If Rezultat2.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.WriteMessage(vbLf & "Command:")
                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                    Exit Sub
                End If

                Dim Poly1 As Polyline
                Dim Poly3D As Polyline3d

                Dim Poly2 As Polyline
                Dim Poly3D2 As Polyline3d




                Dim Chainage_at_common_point_old As Double
                Dim Chainage_at_common_point_new As Double
                Dim Diferenta_chainage As Double


                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                    If IsNothing(Rezultat1) = False Then

                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(0)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                            Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj2 = Rezultat2.Value.Item(0)
                            Dim Ent2 As Entity
                            Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)


                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline And TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then


                                Poly1 = Ent1
                                Poly2 = Ent2

                                Dim Point0 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                                Dim Point_zero_old As New Point3d
                                Dim Point_zero_new As New Point3d

                                Dim PP0 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select reroute start point:")
                                PP0.AllowNone = True
                                Point0 = Editor1.GetPoint(PP0)
                                If Point0.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                    Point_zero_old = Poly1.GetClosestPointTo(Point0.Value.TransformBy(Editor1.CurrentUserCoordinateSystem), Editor1.GetCurrentView.ViewDirection, False)
                                Else
                                    Editor1.WriteMessage(vbLf & "Command:")
                                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                    Exit Sub
                                End If

                                Chainage_at_common_point_old = Poly1.GetDistAtPoint(Point_zero_old)
                                Point_zero_new = Poly2.GetClosestPointTo(Point0.Value.TransformBy(Editor1.CurrentUserCoordinateSystem), Editor1.GetCurrentView.ViewDirection, False)
                                Chainage_at_common_point_new = Poly2.GetDistAtPoint(Point_zero_new)
                                Diferenta_chainage = Chainage_at_common_point_new - Chainage_at_common_point_old

                                Trans1.Commit()

                            ElseIf TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline3d And TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then


                                Poly3D = Ent1
                                Poly3D2 = Ent2

                                Dim Point0 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                                Dim Point_zero_old As New Point3d
                                Dim Point_zero_new As New Point3d

                                Dim PP0 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select reroute start point:")
                                PP0.AllowNone = True
                                Point0 = Editor1.GetPoint(PP0)


                                If Point0.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                    Point_zero_old = Poly3D.GetClosestPointTo(Point0.Value.TransformBy(Editor1.CurrentUserCoordinateSystem), Editor1.GetCurrentView.ViewDirection, False)
                                Else
                                    Editor1.WriteMessage(vbLf & "Command:")
                                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                    Exit Sub
                                End If

                                Chainage_at_common_point_old = Poly3D.GetDistAtPoint(Point_zero_old)
                                Point_zero_new = Poly3D2.GetClosestPointTo(Point0.Value.TransformBy(Editor1.CurrentUserCoordinateSystem), Editor1.GetCurrentView.ViewDirection, False)
                                Chainage_at_common_point_new = Poly3D2.GetDistAtPoint(Point_zero_new)
                                Diferenta_chainage = Chainage_at_common_point_new - Chainage_at_common_point_old

                                Trans1.Commit()

                            Else
                                Editor1.WriteMessage("No Polylines")
                                Editor1.WriteMessage(vbLf & "Command:")
                                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                Exit Sub
                            End If
                        End Using
                    End If
                End If
1234:
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)

                    Dim Point_on_poly1 As New Point3d
                    Dim Point_on_poly2 As New Point3d


                    Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Please start point on the reroute:")
                    Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                    PP1.AllowNone = False
                    Point1 = Editor1.GetPoint(PP1)
                    If Not Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Editor1.SetImpliedSelection(Empty_array)
                        Trans1.Commit()
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        Exit Sub
                    End If

                    Dim Distanta_pana_la_xing1 As Double
                    If IsNothing(Poly2) = False Then
                        Point_on_poly1 = Poly2.GetClosestPointTo(Point1.Value.TransformBy(curent_ucs_matrix), Editor1.GetCurrentView.ViewDirection, False)
                        Distanta_pana_la_xing1 = Poly2.GetDistAtPoint(Point_on_poly1)
                    End If

                    If IsNothing(Poly3D2) = False Then
                        Point_on_poly1 = Poly3D2.GetClosestPointTo(Point1.Value.TransformBy(curent_ucs_matrix), Editor1.GetCurrentView.ViewDirection, False)
                        Distanta_pana_la_xing1 = Poly3D2.GetDistAtPoint(Point_on_poly1)
                    End If


                    Dim Chainage1 As Double = Distanta_pana_la_xing1 - Diferenta_chainage




                    Dim Chainage_string1 As String = Get_chainage_from_double(Chainage1, 1)
                    If Chainage_string1 = "-0+000.0" Then Chainage_string1 = "0+000.0"

                    Dim Mleader1 As New MLeader

                    If IsNothing(Point_on_poly1) = False Then
                        Mleader1 = Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly1, Chainage_string1, 0.5, 0.1, 0.5, 3, 3)
                    End If


                    Dim PP2 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Please end point on the reroute:")
                    Dim Point2 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                    PP2.AllowNone = False
                    PP2.UseBasePoint = True
                    PP2.BasePoint = Point_on_poly1
                    Point2 = Editor1.GetPoint(PP2)
                    If Not Point2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Editor1.SetImpliedSelection(Empty_array)
                        Trans1.Commit()
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        Exit Sub
                    End If

                    Dim Distanta_pana_la_xing2 As Double
                    If IsNothing(Poly2) = False Then
                        Point_on_poly2 = Poly2.GetClosestPointTo(Point2.Value.TransformBy(curent_ucs_matrix), Editor1.GetCurrentView.ViewDirection, False)
                        Distanta_pana_la_xing2 = Poly2.GetDistAtPoint(Point_on_poly2)
                    End If

                    If IsNothing(Poly3D2) = False Then
                        Point_on_poly2 = Poly3D2.GetClosestPointTo(Point2.Value.TransformBy(curent_ucs_matrix), Editor1.GetCurrentView.ViewDirection, False)
                        Distanta_pana_la_xing2 = Poly3D2.GetDistAtPoint(Point_on_poly2)
                    End If


                    Dim Chainage2 As Double = Distanta_pana_la_xing2 - Diferenta_chainage




                    Dim Chainage_string2 As String = Get_chainage_from_double(Chainage2, 1)
                    If Chainage_string2 = "-0+000.0" Then Chainage_string2 = "0+000.0"

                    Dim Mleader2 As New MLeader

                    If IsNothing(Point_on_poly2) = False Then
                        Mleader2 = Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly2, Chainage_string2, 0.5, 0.2, 0.5, 3, 3)
                    End If





                    Trans1.Commit()


                    If IsNumeric(TextBox_screw_anchors_spacing.Text) = True Then

                        Dim CH1_CH2 As Double = Abs(Chainage1 - Chainage2)
                        Dim Screw_space As Double = CDbl(TextBox_screw_anchors_spacing.Text)
                        Dim Multiplu As Double = Ceiling(Round(CH1_CH2, 2) / Screw_space)


                        Dim Number_screw As Integer = Floor(Multiplu) + 1
                        ListBox_no_of_screw_anchors.Items.Add(Number_screw)
                        ListBox_no_of_screw_anchors.Items.Add(Number_screw)
                        ListBox_no_of_screw_anchors.Items.Add("___")
                        Dim Chainage_middle As Double

                        If Chainage1 <= Chainage2 Then
                            Chainage_middle = Chainage1 + CH1_CH2 / 2
                            ListBox_Picked_chainages.Items.Add(Chainage_string1)
                            ListBox_Picked_chainages.Items.Add(Chainage_string2)
                            ListBox_Picked_chainages.Items.Add("___")
                        Else
                            Chainage_middle = Chainage2 + CH1_CH2 / 2
                            ListBox_Picked_chainages.Items.Add(Chainage_string2)
                            ListBox_Picked_chainages.Items.Add(Chainage_string1)
                            ListBox_Picked_chainages.Items.Add("___")
                        End If
                        Dim Chain1 As Double = Chainage_middle - (Multiplu * Screw_space) / 2
                        Dim Chain2 As Double = Chainage_middle + (Multiplu * Screw_space) / 2

                        ListBox_Recalculated_chainages.Items.Add(Get_chainage_from_double(Chain1, 1))
                        ListBox_Recalculated_chainages.Items.Add(Get_chainage_from_double(Chain2, 1))
                        ListBox_Recalculated_chainages.Items.Add("___")

                    End If




                    GoTo 1234

                End Using


                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
            Catch ex As Exception
                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                MsgBox(ex.Message)
            End Try
        End Using


    End Sub

    Private Sub Button_clear_text_Click(sender As Object, e As EventArgs) Handles Button_clear_text.Click
        ListBox_Picked_chainages.Items.Clear()
        ListBox_Recalculated_chainages.Items.Clear()
        ListBox_no_of_screw_anchors.Items.Clear()
    End Sub

    Private Sub ListBox_no_of_screw_anchors_Click(sender As Object, e As EventArgs) Handles ListBox_no_of_screw_anchors.Click
        Try
            Dim curent_index As Integer = ListBox_no_of_screw_anchors.SelectedIndex
            If curent_index >= 0 Then
                If ListBox_no_of_screw_anchors.Items.Count > 0 Then
                    If Not ListBox_no_of_screw_anchors.Items(curent_index).ToString = "___" Then
                        Dim Rezultat_msg As MsgBoxResult = MsgBox("Delete?", vbYesNo)
                        If Rezultat_msg = vbYes Then

                            If curent_index = 0 Or curent_index = 1 Then
                                ListBox_no_of_screw_anchors.Items.RemoveAt(0)
                                ListBox_no_of_screw_anchors.Items.RemoveAt(0)
                                ListBox_no_of_screw_anchors.Items.RemoveAt(0)

                                ListBox_Picked_chainages.Items.RemoveAt(0)
                                ListBox_Picked_chainages.Items.RemoveAt(0)
                                ListBox_Picked_chainages.Items.RemoveAt(0)

                                ListBox_Recalculated_chainages.Items.RemoveAt(0)
                                ListBox_Recalculated_chainages.Items.RemoveAt(0)
                                ListBox_Recalculated_chainages.Items.RemoveAt(0)
                            End If

                            If curent_index > 2 Then

                                If ListBox_no_of_screw_anchors.Items(curent_index - 1).ToString = "___" Then

                                    ListBox_no_of_screw_anchors.Items.RemoveAt(curent_index)
                                    ListBox_no_of_screw_anchors.Items.RemoveAt(curent_index)
                                    ListBox_no_of_screw_anchors.Items.RemoveAt(curent_index)

                                    ListBox_Picked_chainages.Items.RemoveAt(curent_index)
                                    ListBox_Picked_chainages.Items.RemoveAt(curent_index)
                                    ListBox_Picked_chainages.Items.RemoveAt(curent_index)

                                    ListBox_Recalculated_chainages.Items.RemoveAt(curent_index)
                                    ListBox_Recalculated_chainages.Items.RemoveAt(curent_index)
                                    ListBox_Recalculated_chainages.Items.RemoveAt(curent_index)

                                ElseIf ListBox_no_of_screw_anchors.Items(curent_index - 2).ToString = "___" Then

                                    ListBox_no_of_screw_anchors.Items.RemoveAt(curent_index - 1)
                                    ListBox_no_of_screw_anchors.Items.RemoveAt(curent_index - 1)
                                    ListBox_no_of_screw_anchors.Items.RemoveAt(curent_index - 1)

                                    ListBox_Picked_chainages.Items.RemoveAt(curent_index - 1)
                                    ListBox_Picked_chainages.Items.RemoveAt(curent_index - 1)
                                    ListBox_Picked_chainages.Items.RemoveAt(curent_index - 1)

                                    ListBox_Recalculated_chainages.Items.RemoveAt(curent_index - 1)
                                    ListBox_Recalculated_chainages.Items.RemoveAt(curent_index - 1)
                                    ListBox_Recalculated_chainages.Items.RemoveAt(curent_index - 1)
                                End If
                            End If

                        End If



                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_insert_block_recalculated_chainages_Click(sender As Object, e As EventArgs) Handles Button_insert_block_recalculated_chainages.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        Using lock As DocumentLock = ThisDrawing.LockDocument
            ascunde_butoanele_pentru_forms(Me, Colectie1)
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Try
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)


                    Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Please start point:")
                    Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                    PP1.AllowNone = False
                    Point1 = Editor1.GetPoint(PP1)
                    If Not Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        Exit Sub
                    End If

                    Insereaza_block_table_record_in_drawing("Screw_Anchor_Alignment.dwg", "SCREW_ANCHOR_ALIGNMENT")

                    Creazea_layer("TEXT", 2, "All text, notes baloons, section symbols, legends, door-room-wall numbers, text leader lines", True)
                    Dim Point_ins As New Point3d
                    Point_ins = Point1.Value.TransformBy(curent_ucs_matrix)

                    For i = 0 To ListBox_Recalculated_chainages.Items.Count - 1 Step 3
                        Dim Colectie_nume_atribute As New Specialized.StringCollection
                        Colectie_nume_atribute.Add("NO_TYPE")
                        Colectie_nume_atribute.Add("SPACING")
                        Colectie_nume_atribute.Add("BEGINSTA")
                        Colectie_nume_atribute.Add("ENDSTA")

                        Dim Colectie_valori_atribute As New Specialized.StringCollection
                        Colectie_valori_atribute.Add(ListBox_no_of_screw_anchors.Items(i) & " SA")
                        Colectie_valori_atribute.Add(TextBox_screw_anchors_spacing.Text & " C/C")
                        Colectie_valori_atribute.Add(ListBox_Recalculated_chainages.Items(i + 1))
                        Colectie_valori_atribute.Add(ListBox_Recalculated_chainages.Items(i))






                        InsertBlock_with_multiple_atributes("Screw_Anchor_Alignment.dwg", "SCREW_ANCHOR_ALIGNMENT", Point_ins, 1, BTrecord, "TEXT", Colectie_nume_atribute, Colectie_valori_atribute)
                        Point_ins = New Point3d(Point_ins.X + 30, Point_ins.Y, Point_ins.Z)

                    Next


                    Trans1.Commit()

                End Using



                Editor1.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
            Catch ex As Exception

                Editor1.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                MsgBox(ex.Message)
            End Try
        End Using
    End Sub
End Class