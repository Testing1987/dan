Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Public Class Read_from_acad_w_excel_form
    Dim Colectie1 As New Specialized.StringCollection

    Private Sub Global_change_Form_Load(sender As Object, e As System.EventArgs) Handles Me.Load

    End Sub






    Private Sub ListBox_FIND_Click(sender As Object, e As EventArgs) Handles ListBox_Text_from_Autocad.Click, ListBox_deflection.Click
        Try
            Dim curent_index As Integer = ListBox_Text_from_Autocad.SelectedIndex
            If curent_index >= 0 Then
                If ListBox_Text_from_Autocad.Items.Count > 0 Then
                    Dim Rezultat_msg As MsgBoxResult = MsgBox("Add?", vbYesNo)
                    If Rezultat_msg = vbYes Then
                        Dim Find1 As String = InputBox("Add the text to be added:")
                        If Not Find1 = "" Then
                            ListBox_Text_from_Autocad.Items.Add(Find1)
                        End If
                    ElseIf Rezultat_msg = vbNo Then
                        If MsgBox("Delete?", vbYesNo) = vbYes Then
                            ListBox_Text_from_Autocad.Items.RemoveAt(curent_index)
                        Else
                            Dim Find1 As String = InputBox("Specify new text:")
                            If Not Find1 = "" Then
                                ListBox_Text_from_Autocad.Items(curent_index) = Find1
                            End If
                        End If
                    End If
                End If
            Else
                Dim Rezultat_msg As MsgBoxResult = MsgBox("Add?", vbYesNo)
                If Rezultat_msg = vbYes Then
                    Dim Find1 As String = InputBox("Add the text to be added:")
                    If Not Find1 = "" Then
                        ListBox_Text_from_Autocad.Items.Add(Find1)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Button_load_text_to_combo_Click(sender As Object, e As EventArgs) Handles Button_load_text_to_combo.Click
        Try
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            ascunde_butoanele_pentru_forms(Me, Colectie1)
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
            Dim Database1 As Autodesk.AutoCAD.DatabaseServices.Database
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
            ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Database1 = ThisDrawing.Database
            Editor1 = ThisDrawing.Editor
            Dim Empty_array() As ObjectId

            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Rezultat1 = Editor1.SelectImplied

                    If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_Prompt.MessageForAdding = vbLf & "Select text objects:"
                        Object_Prompt.SingleOnly = False
                        Rezultat1 = Editor1.GetSelection(Object_Prompt)
                    End If

                    If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Editor1.SetImpliedSelection(Empty_array)
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        Exit Sub
                    End If

                    ListBox_Text_from_Autocad.Items.Clear()

                    For i = 0 To Rezultat1.Value.Count - 1
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat1.Value.Item(i)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                        If TypeOf Ent1 Is DBText Then

                            Dim Text1 As DBText = Ent1
                            Dim String1 As String = Text1.TextString
                            If ListBox_Text_from_Autocad.Items.Contains(String1) = False Then ListBox_Text_from_Autocad.Items.Add(String1)

                        End If
                        If TypeOf Ent1 Is MText Then

                            Dim MText1 As MText = Ent1
                            Dim String1 As String = MText1.Text
                            If ListBox_Text_from_Autocad.Items.Contains(String1) = False Then ListBox_Text_from_Autocad.Items.Add(String1)

                        End If

                        If TypeOf Ent1 Is MLeader Then

                            Dim Mleader1 As MLeader = Ent1
                            Dim MText1 As MText = Mleader1.MText
                            Dim String1 As String = MText1.Text
                            If ListBox_Text_from_Autocad.Items.Contains(String1) = False Then ListBox_Text_from_Autocad.Items.Add(String1)

                        End If

                        If TypeOf Ent1 Is BlockReference Then

                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                For Each Atid As ObjectId In Block1.AttributeCollection
                                    Dim Atr1 As AttributeReference = Atid.GetObject(OpenMode.ForRead)
                                    Dim String1 As String

                                    If Atr1.IsMTextAttribute = True Then
                                        String1 = Atr1.MTextAttribute.Text
                                    Else
                                        String1 = Atr1.TextString
                                    End If
                                    If ListBox_Text_from_Autocad.Items.Contains(String1) = False Then ListBox_Text_from_Autocad.Items.Add(String1)

                                Next

                            End If

                        End If

                    Next
                    Trans1.Commit()
                End Using ' asta e de la tranzactie
            End Using




            afiseaza_butoanele_pentru_forms(Me, Colectie1)
        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_load_chainages_from_text_Click(sender As Object, e As EventArgs) Handles Button_load_chainages_from_text.Click
        Try
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            ascunde_butoanele_pentru_forms(Me, Colectie1)
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
            Dim Database1 As Autodesk.AutoCAD.DatabaseServices.Database
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
            ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Database1 = ThisDrawing.Database
            Editor1 = ThisDrawing.Editor
            Dim Empty_array() As ObjectId

            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Rezultat1 = Editor1.SelectImplied

                    If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_Prompt.MessageForAdding = vbLf & "Select text containing chainages:"
                        Object_Prompt.SingleOnly = False
                        Rezultat1 = Editor1.GetSelection(Object_Prompt)
                    End If

                    If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Editor1.SetImpliedSelection(Empty_array)
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        Exit Sub
                    End If

                    ListBox_Text_from_Autocad.Items.Clear()

                    For i = 0 To Rezultat1.Value.Count - 1
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat1.Value.Item(i)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                        If TypeOf Ent1 Is DBText Then

                            Dim Text1 As DBText = Ent1
                            Dim String1 As String = Text1.TextString
                            If String1.Contains("+") = True Then
                                Dim valoare1 As String = extrage_chainage_din_text_de_la_inceputul_textului(String1)
                                If valoare1 = "" Then
                                    valoare1 = extrage_chainage_din_text_de_la_sfarsitul_textului(String1)
                                End If
                                If Not String1 = "" Then
                                    If ListBox_Text_from_Autocad.Items.Contains(valoare1) = False Then ListBox_Text_from_Autocad.Items.Add(valoare1)
                                End If
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then

                            Dim MText1 As MText = Ent1
                            Dim String1 As String = MText1.Contents
                            If String1.Contains("+") = True Then
                                Dim valoare1 As String = extrage_chainage_din_text_de_la_inceputul_textului(String1)
                                If valoare1 = "" Then
                                    valoare1 = extrage_chainage_din_text_de_la_sfarsitul_textului(String1)
                                End If
                                If Not String1 = "" Then
                                    If ListBox_Text_from_Autocad.Items.Contains(valoare1) = False Then ListBox_Text_from_Autocad.Items.Add(valoare1)
                                End If
                            End If

                        End If

                        If TypeOf Ent1 Is MLeader Then

                            Dim Mleader1 As MLeader = Ent1
                            Dim MText1 As MText = Mleader1.MText
                            Dim String1 As String = MText1.Contents
                            If String1.Contains("+") = True Then
                                Dim valoare1 As String = extrage_chainage_din_text_de_la_inceputul_textului(String1)
                                If valoare1 = "" Then
                                    valoare1 = extrage_chainage_din_text_de_la_sfarsitul_textului(String1)
                                End If
                                If Not String1 = "" Then
                                    If ListBox_Text_from_Autocad.Items.Contains(valoare1) = False Then ListBox_Text_from_Autocad.Items.Add(valoare1)
                                End If
                            End If

                        End If

                        If TypeOf Ent1 Is BlockReference Then

                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                For Each Atid As ObjectId In Block1.AttributeCollection
                                    Dim Atr1 As AttributeReference = Atid.GetObject(OpenMode.ForRead)
                                    Dim String1 As String

                                    If Atr1.IsMTextAttribute = True Then
                                        String1 = Atr1.MTextAttribute.Contents
                                    Else
                                        String1 = Atr1.TextString
                                    End If

                                    If String1.Contains("+") = True Then
                                        Dim valoare1 As String = extrage_chainage_din_text_de_la_inceputul_textului(String1)
                                        If valoare1 = "" Then
                                            valoare1 = extrage_chainage_din_text_de_la_sfarsitul_textului(String1)
                                        End If
                                        If Not String1 = "" Then
                                            If ListBox_Text_from_Autocad.Items.Contains(valoare1) = False Then ListBox_Text_from_Autocad.Items.Add(valoare1)
                                        End If
                                    End If
                                Next

                            End If

                        End If

                    Next
                    Trans1.Commit()
                End Using ' asta e de la tranzactie
            End Using




            afiseaza_butoanele_pentru_forms(Me, Colectie1)
        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_place_chainages_on_polyline_Click(sender As Object, e As EventArgs) Handles Button_place_chainages_on_polyline.Click
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try
            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt.MessageForAdding = vbLf & "Select polyline:"

            Object_Prompt.SingleOnly = True

            Rezultat1 = Editor1.GetSelection(Object_Prompt)


            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                Exit Sub
            End If
            Dim Poly1 As Polyline
            Dim Poly3D As Polyline3d

            Dim Point_on_poly As New Point3d

            Dim Dist_from_start_for_zero As Double

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then
                    Using lock1 As DocumentLock = ThisDrawing.LockDocument
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(0)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                Poly1 = Ent1
                                Dim Point_zero As New Point3d
                                Point_zero = Poly1.GetClosestPointTo(Poly1.StartPoint, Vector3d.ZAxis, False)
                                Dist_from_start_for_zero = Poly1.GetDistAtPoint(Point_zero)
                            ElseIf TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then
                                Poly3D = Ent1
                                Dim Point_zero As New Point3d
                                Point_zero = Poly3D.GetClosestPointTo(Poly3D.StartPoint, Vector3d.ZAxis, False)
                                Dist_from_start_for_zero = Poly3D.GetDistAtPoint(Point_zero)
                            Else
                                Editor1.WriteMessage("No Polyline")
                                Editor1.WriteMessage(vbLf & "Command:")
                                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                Exit Sub
                            End If

                            For i = 0 To ListBox_Text_from_Autocad.Items.Count - 1
                                Dim Ch_result As String = ListBox_Text_from_Autocad.Items(i)
                                Ch_result = Replace(Ch_result, "+", "")
                                Ch_result = Replace(Ch_result, " ", "")
                                If IsNumeric(Ch_result) = False Then
                                    Editor1.WriteMessage("Chainage is not specified correctly")
                                    Editor1.SetImpliedSelection(Empty_array)
                                    Editor1.WriteMessage(vbLf & "Command:")
                                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                    Exit Sub
                                End If
                                Dim Chainage As Double = CDbl(Ch_result)
                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                    Poly1 = Ent1
                                    Point_on_poly = Poly1.GetPointAtDist(Dist_from_start_for_zero + Chainage)
                                End If
                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then
                                    Poly3D = Ent1
                                    Point_on_poly = Poly3D.GetPointAtDist(Dist_from_start_for_zero + Chainage)
                                End If

                                Dim Chainage_string As String = ListBox_Text_from_Autocad.Items(i)
                                If Chainage_string = "-0+000.0" Then Chainage_string = "0+000.0"

                                If IsNothing(Point_on_poly) = False Then
                                    Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly, Chainage_string, 5, 2.5, 5, 10, 10)
                                End If

                            Next

                            Trans1.Commit()

                        End Using
                    End Using
                End If
            End If

            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub

    Private Sub Button_deflections_Click(sender As Object, e As EventArgs) Handles Button_deflections.Click
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try
            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt.MessageForAdding = vbLf & "Select the 2D polyline:"

            Object_Prompt.SingleOnly = True

            Rezultat1 = Editor1.GetSelection(Object_Prompt)


            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                Exit Sub
            End If
            Dim Poly1 As Polyline

            Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

            If CheckBox_3d_polyline.Checked = True Then

                Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt2.MessageForAdding = vbLf & "Select 3D polyline (NEW):"

                Object_Prompt2.SingleOnly = True

                Rezultat2 = Editor1.GetSelection(Object_Prompt2)


                If Rezultat2.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.WriteMessage(vbLf & "Command:")
                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                    Exit Sub
                End If
            End If

            Dim Poly3d As Polyline3d


            Dim Point_start As New Point3d
            Dim Point_end As New Point3d


            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then
                    Using lock1 As DocumentLock = ThisDrawing.LockDocument
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(0)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                Poly1 = Ent1
                            Else
                                Editor1.WriteMessage("No Polyline")
                                Editor1.WriteMessage(vbLf & "Command:")
                                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                Exit Sub
                            End If

                            If CheckBox_3d_polyline.Checked = True Then
                                Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj2 = Rezultat2.Value.Item(0)
                                Dim Ent2 As Entity
                                Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)

                                If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then
                                    Poly3d = Ent2
                                Else
                                    Editor1.WriteMessage("No 3d Polyline")
                                    Editor1.WriteMessage(vbLf & "Command:")
                                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                    Exit Sub
                                End If
                            End If



                            Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult

                            Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select First Point (Or press ENTER for start of the polyline):")
                            PP1.AllowNone = True
                            Point1 = Editor1.GetPoint(PP1)

                            If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                Editor1.WriteMessage(vbLf & "Command:")
                                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                Exit Sub
                            End If

                            If Not Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Point_start = Poly1.StartPoint
                            Else
                                'aici am tratat ucs-ul 
                                Point_start = Poly1.GetClosestPointTo(Point1.Value.TransformBy(curent_ucs_matrix), Vector3d.ZAxis, False)
                            End If

                            Dim Point2 As Autodesk.AutoCAD.EditorInput.PromptPointResult

                            Dim PP2 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select Second Point (Or press ENTER for End of the polyline):")
                            PP2.AllowNone = True
                            PP2.UseBasePoint = True
                            If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                PP2.BasePoint = Point1.Value
                            Else
                                PP2.BasePoint = Poly1.StartPoint
                            End If

                            Point2 = Editor1.GetPoint(PP2)

                            If Point2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                Editor1.WriteMessage(vbLf & "Command:")
                                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                Exit Sub
                            End If

                            If Not Point2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Point_end = Poly1.GetClosestPointTo(Poly1.EndPoint, Vector3d.ZAxis, False)
                            Else
                                'aici am tratat ucs-ul 
                                Point_end = Poly1.GetClosestPointTo(Point2.Value.TransformBy(Editor1.CurrentUserCoordinateSystem), Vector3d.ZAxis, False)
                            End If

                            Dim Parameter_start As Double = Poly1.GetParameterAtPoint(Point_start)
                            Dim Parameter_end As Double = Poly1.GetParameterAtPoint(Point_end)

                            If Parameter_start > Parameter_end Then
                                Dim temp_par As Double
                                Dim Temppt As Point3d
                                temp_par = Parameter_start
                                Parameter_start = Parameter_end
                                Parameter_end = temp_par
                                Temppt = Point_start
                                Point_start = Point_end
                                Point_end = Temppt
                            End If
                            Dim Vertex_start, Vertex_end As Double
                            Vertex_start = Ceiling(Parameter_start)
                            Vertex_end = Floor(Parameter_end)
                            If Vertex_start = 0 Then Vertex_start = 1
                            If Vertex_start > Vertex_end Then
                                Editor1.WriteMessage("No deflection can be calculated")
                                Editor1.WriteMessage(vbLf & "Command:")
                                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                Exit Sub
                            End If




                            Dim Planul_curent As Plane = New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)

                            If Vertex_start = Vertex_end Then Vertex_end = Vertex_end + 1
                            If CheckBox_populate_deflection.Checked = True Then
                                ListBox_deflection.Items.Clear()
                                ListBox_Text_from_Autocad.Items.Clear()
                            End If


                            Dim Curva_veche As Curve

                            Dim Point_zero_old As New Point3d

                            Dim Rezultat_vechi As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                            If CheckBox_REROUTE.Checked = True Then
                                Dim Object_Prompt_vechi As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                                Object_Prompt_vechi.MessageForAdding = vbLf & "Select OLD polyline:"

                                Object_Prompt_vechi.SingleOnly = True

                                Rezultat_vechi = Editor1.GetSelection(Object_Prompt_vechi)


                                If Rezultat_vechi.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                    Editor1.WriteMessage(vbLf & "Command:")
                                    Exit Sub
                                End If


                                Dim Obj_vechi As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj_vechi = Rezultat_vechi.Value.Item(0)
                                Dim Ent_vechi As Entity
                                Ent_vechi = Obj_vechi.ObjectId.GetObject(OpenMode.ForRead)

                                If TypeOf Ent_vechi Is Polyline Then
                                    Dim Pol2d As Polyline = Ent_vechi
                                    Dim Point0 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                                    Dim PP0 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select reroute start point:")
                                    PP0.AllowNone = True
                                    Point0 = Editor1.GetPoint(PP0)
                                    Point_zero_old = Pol2d.GetClosestPointTo(Point0.Value.TransformBy(Editor1.CurrentUserCoordinateSystem), Vector3d.ZAxis, False)
                                    Curva_veche = Ent_vechi
                                End If

                                If TypeOf Ent_vechi Is Polyline3d Then
                                    Dim Pol3d As Polyline3d = Ent_vechi
                                    Dim Point0 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                                    Dim PP0 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select reroute start point:")
                                    PP0.AllowNone = True
                                    Point0 = Editor1.GetPoint(PP0)
                                    Point_zero_old = Pol3d.GetClosestPointTo(Point0.Value.TransformBy(Editor1.CurrentUserCoordinateSystem), Vector3d.ZAxis, False)
                                    Curva_veche = Ent_vechi
                                End If




                            End If





                            For i = Vertex_start To Vertex_end - 1

                                Dim vector1 As Vector3d = Poly1.GetPoint3dAt(i - 1).GetVectorTo(Poly1.GetPoint3dAt(i))
                                If vector1.Length < 0.01 Then
                                    Dim K As Double = 2
                                    Do While vector1.Length < 0.01
                                        If i - K >= 0 Then
                                            vector1 = Poly1.GetPoint3dAt(i - K).GetVectorTo(Poly1.GetPoint3dAt(i))
                                        Else
                                            Exit Do
                                        End If
                                        K = K + 1
                                    Loop
                                End If

                                Dim vector2 As Vector3d = Poly1.GetPoint3dAt(i).GetVectorTo(Poly1.GetPoint3dAt(i + 1))
                                If vector2.Length < 0.01 Then
                                    Dim K As Double = 2
                                    Do While vector2.Length < 0.01
                                        If i + K <= Poly1.NumberOfVertices - 1 Then
                                            vector2 = Poly1.GetPoint3dAt(i).GetVectorTo(Poly1.GetPoint3dAt(i + K))
                                        Else
                                            Exit Do
                                        End If
                                        K = K + 1
                                    Loop
                                End If

                                Dim Bearing1 As Double = (vector1.AngleOnPlane(Planul_curent)) * 180 / PI
                                Dim Bearing2 As Double = (vector2.AngleOnPlane(Planul_curent)) * 180 / PI

                                Dim angle1 As Double = (vector2.GetAngleTo(vector1)) * 180 / PI
                                Dim Mleader1 As New MLeader
                                Dim LT_RT As String = ""


                                If Bearing1 < 180 Then
                                    If Bearing2 < Bearing1 + 180 And Bearing2 > Bearing1 Then
                                        LT_RT = " LT"
                                    Else
                                        LT_RT = " RT"
                                    End If
                                Else
                                    If Bearing2 < Bearing1 And Bearing2 > Bearing1 - 180 Then
                                        LT_RT = " RT"
                                    Else
                                        LT_RT = " LT"
                                    End If

                                End If

                                Dim AngleDMS As String = Floor(angle1) & "°"

                                Dim Minute As String = Round((angle1 - Floor(angle1)) * 60, 0) & "'"

                                If Round((angle1 - Floor(angle1)) * 60, 0) = 60 Then
                                    AngleDMS = Floor(angle1 + 1) & "°"
                                    Minute = "00'"
                                End If


                                If Len(Minute) = 2 Then Minute = "0" & Minute
                                AngleDMS = AngleDMS & Minute & "00" & Chr(34)



                                If Not AngleDMS = "0°00'00" & Chr(34) Then
                                    Mleader1 = Creaza_Mleader_nou_fara_UCS_transform(Poly1.GetPoint3dAt(i), AngleDMS & LT_RT, 0.5, 0.2, 0.2, 11, 7)
                                End If

                                If CheckBox_populate_deflection.Checked = True Then
                                    ListBox_deflection.Items.Add(AngleDMS & LT_RT)
                                    Dim Punct_vertex As New Point3d
                                    Punct_vertex = Poly1.GetPoint3dAt(i)
                                    Dim Chainage2d As Double = Poly1.GetDistanceAtParameter(i)
                                    If CheckBox_3d_polyline.Checked = True Then
                                        If IsNothing(Poly3d) = False Then
                                            Dim Point_on_3d As New Point3d
                                            Point_on_3d = Poly3d.GetClosestPointTo(Punct_vertex.TransformBy(Editor1.CurrentUserCoordinateSystem), Vector3d.ZAxis, False)
                                            Dim Chainage3d As Double = Poly3d.GetDistAtPoint(Point_on_3d)

                                            If CheckBox_REROUTE.Checked = True Then

                                                If TypeOf Curva_veche Is Polyline3d Then
                                                    Chainage3d = calculeaza_chainage_for_REROUTE(Poly3d, Curva_veche, Point_on_3d, Point_zero_old)
                                                Else
                                                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                                    Editor1.WriteMessage(vbLf & "Command:")
                                                    Exit Sub
                                                End If



                                            End If

                                            ListBox_Text_from_Autocad.Items.Add(Get_chainage_from_double(Chainage3d, 1))




                                        Else
                                            ListBox_Text_from_Autocad.Items.Add("NOPOLY3D")
                                        End If

                                    Else
                                        If CheckBox_REROUTE.Checked = True Then
                                            If TypeOf Curva_veche Is Polyline Then
                                                Chainage2d = calculeaza_chainage_for_REROUTE(Poly1, Curva_veche, Punct_vertex, Point_zero_old)
                                            Else
                                                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                                Editor1.WriteMessage(vbLf & "Command:")
                                                Exit Sub
                                            End If

                                            'Poly1 = calculeaza_chainage_for_REROUTE(Poly1, Punct_vertex)
                                        End If

                                        ListBox_Text_from_Autocad.Items.Add(Get_chainage_from_double(Chainage2d, 1))
                                    End If

                                End If

                            Next




                            Trans1.Commit()

                        End Using
                    End Using
                End If
            End If

            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub

   


    Private Sub Button_create_mtext_Click(sender As Object, e As EventArgs) Handles Button_create_mtext.Click
        If ListBox_Text_from_Autocad.Items.Count > 0 Then

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
            ascunde_butoanele_pentru_forms(Me, Colectie1)
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Try
                Dim Point_rezult As Autodesk.AutoCAD.EditorInput.PromptPointResult

                Dim PP0 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify insertion point")
                PP0.AllowNone = True
                Point_rezult = Editor1.GetPoint(PP0)
                If Not Point_rezult.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                    Exit Sub
                End If


                Using lock1 As DocumentLock = ThisDrawing.LockDocument
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                        Dim X As Double = Point_rezult.Value.TransformBy(curent_ucs_matrix).X
                        Dim y As Double = Point_rezult.Value.TransformBy(curent_ucs_matrix).Y
                        Dim z As Double = Point_rezult.Value.TransformBy(curent_ucs_matrix).Z

                        For i = 0 To ListBox_Text_from_Autocad.Items.Count - 1
                            Dim Text_string As String = ListBox_Text_from_Autocad.Items(i)
                            Dim Mtext1 As New MText
                            Mtext1.Contents = Text_string
                            Mtext1.TextHeight = 2.5
                            Mtext1.Location = New Point3d(X, y, z)
                            BTrecord.AppendEntity(Mtext1)
                            Trans1.AddNewlyCreatedDBObject(Mtext1, True)
                            y = y - 6
                        Next

                        Trans1.Commit()

                    End Using
                End Using



                Editor1.WriteMessage(vbLf & "Command:")
            Catch ex As Exception

                Editor1.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
            End Try
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

        End If
    End Sub


    Private Sub Button_W2XL_Click(sender As Object, e As EventArgs) Handles Button_W2XL.Click
        Try

            If ListBox_Text_from_Autocad.Items.Count > 0 Then

                If TextBox_column.Text = "" Then
                    MsgBox("Please specify the EXCEL COLUMN!")
                    Exit Sub
                End If
                If TextBox_ROW_START.Text = "" Then
                    MsgBox("Please specify the EXCEL START ROW!")
                    Exit Sub
                End If

                If IsNumeric(TextBox_ROW_START.Text) = False Then
                    With TextBox_ROW_START
                        .Text = ""
                        .Focus()
                    End With
                    MsgBox("Please specify start row")

                    Exit Sub
                End If
                If Val(TextBox_ROW_START.Text) < 1 Then
                    With TextBox_ROW_START
                        .Text = ""
                        .Focus()
                    End With
                    MsgBox("Start row can't be smaller than 1")

                    Exit Sub
                End If

            End If

            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
            Dim start1 As Integer = CInt(TextBox_ROW_START.Text)
            Dim Col1 As String = TextBox_column.Text.ToUpper
            For i = 0 To ListBox_Text_from_Autocad.Items.Count - 1
                W1.Range(Col1 & (start1 + i)).Value = ListBox_Text_from_Autocad.Items(i)
                If ListBox_deflection.Items.Count = ListBox_Text_from_Autocad.Items.Count Then
                    W1.Range(Chr(Asc(Col1) + 1) & (start1 + i)).Value = ListBox_deflection.Items(i)
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_LOAD_TEXT_FROM_EXCEL_Click(sender As Object, e As EventArgs) Handles Button_LOAD_TEXT_FROM_EXCEL.Click
        Try



            If TextBox_col_load.Text = "" Then
                MsgBox("Please specify the EXCEL COLUMN!")
                Exit Sub
            End If
            If TextBox_start_load.Text = "" Then
                MsgBox("Please specify the EXCEL START ROW!")
                Exit Sub
            End If

            If IsNumeric(TextBox_start_load.Text) = False Then
                With TextBox_start_load
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify start row")

                Exit Sub
            End If
            If TextBox_end_load.Text = "" Then
                MsgBox("Please specify the EXCEL END ROW!")
                Exit Sub
            End If

            If IsNumeric(TextBox_end_load.Text) = False Then
                With TextBox_end_load
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify END row")

                Exit Sub
            End If

            If Val(TextBox_start_load.Text) < 1 Then
                With TextBox_start_load
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Start row can't be smaller than 1")

                Exit Sub
            End If

            If Val(TextBox_end_load.Text) < 1 Then
                With TextBox_end_load
                    .Text = ""
                    .Focus()
                End With
                MsgBox("End row can't be smaller than 1")

                Exit Sub
            End If

            If Val(TextBox_end_load.Text) < Val(TextBox_start_load.Text) Then
                MsgBox("END row smaller than start row")

                Exit Sub
            End If


            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
            Dim start1 As Integer = CInt(TextBox_start_load.Text)
            Dim end1 As Integer = CInt(TextBox_end_load.Text)
            Dim Col1 As String = TextBox_column.Text.ToUpper
            ListBox_Text_from_Autocad.Items.Clear()
            For i = start1 To end1
                ListBox_Text_from_Autocad.Items.Add(W1.Range(TextBox_col_load.Text.ToUpper & i).Value)
            Next

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Sub Button_mtext_defl_chainage_Click(sender As Object, e As EventArgs) Handles Button_mtext_defl_chainage.Click
        If ListBox_Text_from_Autocad.Items.Count > 0 And ListBox_Text_from_Autocad.Items.Count = ListBox_deflection.Items.Count Then

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
            ascunde_butoanele_pentru_forms(Me, Colectie1)
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Try
                Dim Point_rezult As Autodesk.AutoCAD.EditorInput.PromptPointResult

                Dim PP0 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify insertion point")
                PP0.AllowNone = True
                Point_rezult = Editor1.GetPoint(PP0)
                If Not Point_rezult.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                    Exit Sub
                End If


                Using lock1 As DocumentLock = ThisDrawing.LockDocument
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                        Dim X As Double = Point_rezult.Value.TransformBy(curent_ucs_matrix).X
                        Dim y As Double = Point_rezult.Value.TransformBy(curent_ucs_matrix).Y
                        Dim z As Double = Point_rezult.Value.TransformBy(curent_ucs_matrix).Z

                        For i = 0 To ListBox_Text_from_Autocad.Items.Count - 1
                            Dim Text_string As String = ListBox_deflection.Items(i)
                            Dim Mtext1 As New MText
                            Mtext1.Contents = Text_string
                            Mtext1.TextHeight = 2.5
                            Mtext1.Rotation = PI / 2
                            Mtext1.Location = New Point3d(X, y, z)
                            BTrecord.AppendEntity(Mtext1)
                            Trans1.AddNewlyCreatedDBObject(Mtext1, True)

                            Dim Text_string2 As String = ListBox_Text_from_Autocad.Items(i)
                            Dim Mtext2 As New MText
                            Mtext2.Contents = Text_string2
                            Mtext2.TextHeight = 2.5
                            Mtext2.Rotation = PI / 2
                            Mtext2.Location = New Point3d(X, y - 34.38, z)
                            BTrecord.AppendEntity(Mtext2)
                            Trans1.AddNewlyCreatedDBObject(Mtext2, True)



                            Dim Poly_l As New Polyline
                            Poly_l.AddVertexAt(0, New Point2d(X + 2.435, y - 34.38 + 26.472 + 4.16), 0, 0, 0)
                            Poly_l.AddVertexAt(1, New Point2d(X + 2.435, y - 34.38 + 26.472), 0, 0, 0)
                            Poly_l.AddVertexAt(2, New Point2d(X + 2.435 - 2.8, y - 34.38 + 26.472 + 4.16), 0, 0, 0)
                            BTrecord.AppendEntity(Poly_l)
                            Trans1.AddNewlyCreatedDBObject(Poly_l, True)

                            Dim Poly_arc As New Polyline
                            Poly_arc.AddVertexAt(0, New Point2d(X + 2.435 - 3.377, y - 34.38 + 26.472 + 2.446), -Tan((117 * PI / 180) / 4), 0, 0)
                            Poly_arc.AddVertexAt(1, New Point2d(X + 2.435 + 0.997, y - 34.38 + 26.472 + 3.092), 0, 0, 0)

                            BTrecord.AppendEntity(Poly_arc)
                            Trans1.AddNewlyCreatedDBObject(Poly_arc, True)

                            X = X - 6
                        Next

                        Trans1.Commit()

                    End Using
                End Using



                Editor1.WriteMessage(vbLf & "Command:")
            Catch ex As Exception

                Editor1.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
            End Try
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

        End If
    End Sub


    Private Sub Button_clear_text_Click(sender As Object, e As EventArgs) Handles Button_clear_text.Click
        ListBox_Text_from_Autocad.Items.Clear()
    End Sub

    Private Sub Button_clear_defl_Click(sender As Object, e As EventArgs) Handles Button_clear_defl.Click
        ListBox_deflection.Items.Clear()
    End Sub
End Class