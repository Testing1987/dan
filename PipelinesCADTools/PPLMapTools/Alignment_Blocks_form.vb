Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Public Class Alignment_Blocks_form
    Dim Colectie1 As New Specialized.StringCollection



    Private Sub Alignment_Blocks_form_Load(sender As Object, e As EventArgs) Handles Me.Load
        Panel_CROSSING_PROFILE.Visible = False
        Panel_elbow.Visible = False
        Panel_HEAVY_WALL_Right.Visible = False
        Panel_heavy_wall_left.Visible = False
        Panel_pipe_heavy_wall.Visible = False
        Panel_screw.Visible = False
        Panel_SCREW_LEFT.Visible = False
        Panel_SCREW_RIGHT.Visible = False
        Panel_CROSSING_MATERIAL.Visible = True
        Panel_sand_bag.Visible = False
        Panel_concrete.Visible = False
        Panel_TEST_SECTION.Visible = False
        Me.Size = New System.Drawing.Size(950, 741)
        Incarca_existing_layers_to_combobox(ComboBox_layer)
    
    End Sub
    Private Sub Panel1_1_Click(sender As Object, e As EventArgs) Handles Panel1_1.Click
        Panel_CROSSING_PROFILE.Visible = False
        Panel_elbow.Visible = False
        Panel_HEAVY_WALL_Right.Visible = False
        Panel_heavy_wall_left.Visible = False
        Panel_pipe_heavy_wall.Visible = False
        Panel_screw.Visible = False
        Panel_SCREW_LEFT.Visible = False
        Panel_SCREW_RIGHT.Visible = False
        Panel_CROSSING_MATERIAL.Visible = True
        Panel_sand_bag.Visible = False
        Panel_concrete.Visible = False
        Panel_TEST_SECTION.Visible = False
        Panel_CROSSING_MATERIAL.Location = New System.Drawing.Point(9, 5)
    End Sub
    Private Sub Panel1_2_Click(sender As Object, e As EventArgs) Handles Panel1_2.Click
        Panel_CROSSING_PROFILE.Visible = False
        Panel_elbow.Visible = True
        Panel_HEAVY_WALL_Right.Visible = False
        Panel_heavy_wall_left.Visible = False
        Panel_pipe_heavy_wall.Visible = False
        Panel_screw.Visible = False
        Panel_SCREW_LEFT.Visible = False
        Panel_SCREW_RIGHT.Visible = False
        Panel_CROSSING_MATERIAL.Visible = False
        Panel_sand_bag.Visible = False
        Panel_concrete.Visible = False
        Panel_TEST_SECTION.Visible = False
        Panel_elbow.Location = New System.Drawing.Point(9, 5)
    End Sub
    Private Sub Panel1_3_Click(sender As Object, e As EventArgs) Handles Panel1_3.Click
        Panel_CROSSING_PROFILE.Visible = True
        Panel_elbow.Visible = False
        Panel_HEAVY_WALL_Right.Visible = False
        Panel_heavy_wall_left.Visible = False
        Panel_pipe_heavy_wall.Visible = False
        Panel_screw.Visible = False
        Panel_SCREW_LEFT.Visible = False
        Panel_SCREW_RIGHT.Visible = False
        Panel_CROSSING_MATERIAL.Visible = False
        Panel_sand_bag.Visible = False
        Panel_concrete.Visible = False
        Panel_TEST_SECTION.Visible = False
        Panel_CROSSING_PROFILE.Location = New System.Drawing.Point(9, 5)
    End Sub
    Private Sub Panel1_4_Click(sender As Object, e As EventArgs) Handles Panel1_4.Click
        Panel_CROSSING_PROFILE.Visible = False
        Panel_elbow.Visible = False
        Panel_HEAVY_WALL_Right.Visible = False
        Panel_heavy_wall_left.Visible = False
        Panel_pipe_heavy_wall.Visible = False
        Panel_screw.Visible = False
        Panel_SCREW_LEFT.Visible = False
        Panel_SCREW_RIGHT.Visible = False
        Panel_CROSSING_MATERIAL.Visible = False
        Panel_sand_bag.Visible = True
        Panel_concrete.Visible = False
        Panel_TEST_SECTION.Visible = False
        Panel_sand_bag.Location = New System.Drawing.Point(9, 5)
    End Sub
    Private Sub Panel2_1_Click(sender As Object, e As EventArgs) Handles Panel2_1.Click
        Panel_CROSSING_PROFILE.Visible = False
        Panel_elbow.Visible = False
        Panel_HEAVY_WALL_Right.Visible = False
        Panel_HEAVY_WALL_left.Visible = False
        Panel_pipe_heavy_wall.Visible = True
        Panel_screw.Visible = False
        Panel_SCREW_LEFT.Visible = False
        Panel_SCREW_RIGHT.Visible = False
        Panel_CROSSING_MATERIAL.Visible = False
        Panel_sand_bag.Visible = False
        Panel_concrete.Visible = False
        Panel_TEST_SECTION.Visible = False
        Panel_pipe_heavy_wall.Location = New System.Drawing.Point(9, 5)
    End Sub
    Private Sub Panel2_2_Click(sender As Object, e As EventArgs) Handles Panel2_2.Click
        Panel_CROSSING_PROFILE.Visible = False
        Panel_elbow.Visible = False
        Panel_HEAVY_WALL_Right.Visible = True
        Panel_heavy_wall_left.Visible = False
        Panel_pipe_heavy_wall.Visible = False
        Panel_screw.Visible = False
        Panel_SCREW_LEFT.Visible = False
        Panel_SCREW_RIGHT.Visible = False
        Panel_CROSSING_MATERIAL.Visible = False
        Panel_sand_bag.Visible = False
        Panel_concrete.Visible = False
        Panel_TEST_SECTION.Visible = False
        Panel_HEAVY_WALL_Right.Location = New System.Drawing.Point(9, 5)
    End Sub
    Private Sub Panel2_3_Click(sender As Object, e As EventArgs) Handles Panel2_3.Click
        Panel_CROSSING_PROFILE.Visible = False
        Panel_elbow.Visible = False
        Panel_HEAVY_WALL_Right.Visible = False
        Panel_heavy_wall_left.Visible = True
        Panel_pipe_heavy_wall.Visible = False
        Panel_screw.Visible = False
        Panel_SCREW_LEFT.Visible = False
        Panel_SCREW_RIGHT.Visible = False
        Panel_CROSSING_MATERIAL.Visible = False
        Panel_sand_bag.Visible = False
        Panel_concrete.Visible = False
        Panel_TEST_SECTION.Visible = False
        Panel_heavy_wall_left.Location = New System.Drawing.Point(9, 5)
    End Sub
    Private Sub Panel2_4_Click(sender As Object, e As EventArgs) Handles Panel2_4.Click
        Panel_CROSSING_PROFILE.Visible = False
        Panel_elbow.Visible = False
        Panel_HEAVY_WALL_Right.Visible = False
        Panel_heavy_wall_left.Visible = False
        Panel_pipe_heavy_wall.Visible = False
        Panel_screw.Visible = False
        Panel_SCREW_LEFT.Visible = False
        Panel_SCREW_RIGHT.Visible = False
        Panel_CROSSING_MATERIAL.Visible = False
        Panel_sand_bag.Visible = False
        Panel_concrete.Visible = True
        Panel_TEST_SECTION.Visible = False
        Panel_concrete.Location = New System.Drawing.Point(9, 5)
    End Sub
    Private Sub Panel3_1_Click(sender As Object, e As EventArgs) Handles Panel3_1.Click
        Panel_CROSSING_PROFILE.Visible = False
        Panel_elbow.Visible = False
        Panel_HEAVY_WALL_Right.Visible = False
        Panel_heavy_wall_left.Visible = False
        Panel_pipe_heavy_wall.Visible = False
        Panel_screw.Visible = True
        Panel_SCREW_LEFT.Visible = False
        Panel_SCREW_RIGHT.Visible = False
        Panel_CROSSING_MATERIAL.Visible = False
        Panel_sand_bag.Visible = False
        Panel_concrete.Visible = False
        Panel_TEST_SECTION.Visible = False
        Panel_screw.Location = New System.Drawing.Point(9, 5)
    End Sub
    Private Sub Panel3_2_Click(sender As Object, e As EventArgs) Handles Panel3_2.Click
        Panel_CROSSING_PROFILE.Visible = False
        Panel_elbow.Visible = False
        Panel_HEAVY_WALL_Right.Visible = False
        Panel_heavy_wall_left.Visible = False
        Panel_pipe_heavy_wall.Visible = False
        Panel_screw.Visible = False
        Panel_SCREW_LEFT.Visible = True
        Panel_SCREW_RIGHT.Visible = False
        Panel_CROSSING_MATERIAL.Visible = False
        Panel_sand_bag.Visible = False
        Panel_concrete.Visible = False
        Panel_TEST_SECTION.Visible = False
        Panel_SCREW_LEFT.Location = New System.Drawing.Point(9, 5)
    End Sub
    Private Sub Panel3_3_Click(sender As Object, e As EventArgs) Handles Panel3_3.Click
        Panel_CROSSING_PROFILE.Visible = False
        Panel_elbow.Visible = False
        Panel_HEAVY_WALL_Right.Visible = False
        Panel_heavy_wall_left.Visible = False
        Panel_pipe_heavy_wall.Visible = False
        Panel_screw.Visible = False
        Panel_SCREW_LEFT.Visible = False
        Panel_SCREW_RIGHT.Visible = True
        Panel_CROSSING_MATERIAL.Visible = False
        Panel_sand_bag.Visible = False
        Panel_concrete.Visible = False
        Panel_TEST_SECTION.Visible = False
        Panel_SCREW_RIGHT.Location = New System.Drawing.Point(9, 5)
    End Sub
    Private Sub Panel3_4_Click(sender As Object, e As EventArgs) Handles Panel3_4.Click
        Panel_CROSSING_PROFILE.Visible = False
        Panel_elbow.Visible = False
        Panel_HEAVY_WALL_Right.Visible = False
        Panel_heavy_wall_left.Visible = False
        Panel_pipe_heavy_wall.Visible = False
        Panel_screw.Visible = False
        Panel_SCREW_LEFT.Visible = False
        Panel_SCREW_RIGHT.Visible = False
        Panel_CROSSING_MATERIAL.Visible = False
        Panel_sand_bag.Visible = False
        Panel_concrete.Visible = False
        Panel_TEST_SECTION.Visible = True
        Panel_TEST_SECTION.Location = New System.Drawing.Point(9, 5)
    End Sub
    Private Sub Panel_LAYERS_Click(sender As Object, e As EventArgs) Handles Panel_LAYERS.Click
        Incarca_existing_layers_to_combobox(ComboBox_layer)
    End Sub



    Private Sub Button_crossing_material_clear_Click(sender As Object, e As EventArgs) Handles Button_CROSSING_MAT_CLEAR.Click
        ComboBox_crossing_material_COV.Items.Clear()
        ComboBox_crossing_material_DESC.Items.Clear()
        ComboBox_crossing_material_ID_NO.Items.Clear()
        ComboBox_crossing_material_STA.Items.Clear()
        Button_crossing_material_desc.BackColor = Drawing.Color.DimGray
        Button_crossing_material_id_no.BackColor = Drawing.Color.DimGray
        Button_crossing_material_cov.BackColor = Drawing.Color.DimGray
        Button_crossing_material_sta.BackColor = Drawing.Color.DimGray
    End Sub
    Private Sub Button_CROSSING_MATERIAL_PICK_Click(sender As Object, e As EventArgs) Handles Button_CROSSING_MATERIAL_PICK.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info {DESC}:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt2.MessageForAdding = vbLf & "Select info {ID_NO}:"

                    Object_Prompt2.SingleOnly = True
                    Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                    If Not Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    Dim Rezultat3 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt3 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt3.MessageForAdding = vbLf & "Select info {STA}:"

                    Object_Prompt3.SingleOnly = True
                    Rezultat3 = Editor1.GetSelection(Object_Prompt3)
                    If Not Rezultat3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    Dim Rezultat4 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt4 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt4.MessageForAdding = vbLf & "Select info {COV}:"

                    Object_Prompt4.SingleOnly = True
                    Rezultat4 = Editor1.GetSelection(Object_Prompt4)
                    If Not Rezultat4.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    ComboBox_crossing_material_DESC.Items.Clear()
                    ComboBox_crossing_material_DESC.Items.add(" ")
                    ComboBox_crossing_material_ID_NO.Items.Clear()
                    ComboBox_crossing_material_ID_NO.Items.add(" ")
                    ComboBox_crossing_material_STA.Items.Clear()
                    ComboBox_crossing_material_STA.Items.add(" ")
                    ComboBox_crossing_material_COV.Items.Clear()
                    ComboBox_crossing_material_COV.Items.add(" ")


                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent2 As Entity = Rezultat2.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent3 As Entity = Rezultat3.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent4 As Entity = Rezultat4.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)

                    Dim Continut1 As String
                    Dim Continut2 As String
                    Dim Continut3 As String
                    Dim Continut4 As String

                    If TypeOf ent1 Is DBText Then
                        Dim Text1 As DBText = ent1
                        Continut1 = Text1.TextString
                    End If
                    If TypeOf ent2 Is DBText Then
                        Dim Text2 As DBText = ent2
                        Continut2 = Text2.TextString
                    End If
                    If TypeOf ent3 Is DBText Then
                        Dim Text3 As DBText = ent3
                        Continut3 = Text3.TextString
                    End If
                    If TypeOf ent4 Is DBText Then
                        Dim Text4 As DBText = ent4
                        Continut4 = Text4.TextString
                    End If

                    If TypeOf ent1 Is MText Then
                        Dim MText1 As MText = ent1
                        Continut1 = MText1.Contents
                    End If
                    If TypeOf ent2 Is MText Then
                        Dim MText2 As MText = ent2
                        Continut2 = MText2.Contents
                    End If
                    If TypeOf ent3 Is MText Then
                        Dim MText3 As MText = ent3
                        Continut3 = MText3.Contents
                    End If
                    If TypeOf ent4 Is MText Then
                        Dim MText4 As MText = ent4
                        Continut4 = MText4.Contents
                    End If

                    If Not Continut1 = "" Then
                        ComboBox_crossing_material_DESC.Items.Add(Continut1)
                        ComboBox_crossing_material_DESC.SelectedIndex = ComboBox_crossing_material_DESC.Items.IndexOf(Continut1)
                        Button_crossing_material_desc.BackColor = Drawing.Color.Lime
                        ComboBox_crossing_material_ID_NO.Items.Add(Continut1)
                        ComboBox_crossing_material_STA.Items.Add(Continut1)
                        ComboBox_crossing_material_COV.Items.Add(Continut1)

                    End If
                    If Not Continut2 = "" Then
                        ComboBox_crossing_material_DESC.Items.Add(Continut2)
                        ComboBox_crossing_material_ID_NO.Items.Add(Continut2)
                        ComboBox_crossing_material_ID_NO.SelectedIndex = ComboBox_crossing_material_ID_NO.Items.IndexOf(Continut2)
                        Button_crossing_material_id_no.BackColor = Drawing.Color.Lime
                        ComboBox_crossing_material_STA.Items.Add(Continut2)
                        ComboBox_crossing_material_COV.Items.Add(Continut2)
                    End If
                    If Not Continut3 = "" Then
                        ComboBox_crossing_material_DESC.Items.Add(Continut3)
                        ComboBox_crossing_material_ID_NO.Items.Add(Continut3)
                        ComboBox_crossing_material_STA.Items.Add(Continut3)
                        ComboBox_crossing_material_STA.SelectedIndex = ComboBox_crossing_material_STA.Items.IndexOf(Continut3)
                        Button_crossing_material_sta.BackColor = Drawing.Color.Lime
                        ComboBox_crossing_material_COV.Items.Add(Continut3)
                    End If
                    If Not Continut4 = "" Then
                        If Continut4 = Continut3 Then Continut4 = " "
                        ComboBox_crossing_material_DESC.Items.Add(Continut4)
                        ComboBox_crossing_material_ID_NO.Items.Add(Continut4)
                        ComboBox_crossing_material_STA.Items.Add(Continut4)
                        ComboBox_crossing_material_COV.Items.Add(Continut4)
                        ComboBox_crossing_material_COV.SelectedIndex = ComboBox_crossing_material_COV.Items.IndexOf(Continut4)
                        Button_crossing_material_cov.BackColor = Drawing.Color.Lime
                    End If






                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")

    End Sub
    Private Sub Button_crossing_material_insert_Click(sender As Object, e As EventArgs) Handles Button_crossing_material_insert.Click

        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select objects to be deleted:"

                    Object_Prompt1.SingleOnly = False
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If
                    Dim Point_start As Autodesk.AutoCAD.EditorInput.PromptPointResult

                    Dim PP_start As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select insertion point:")
                    PP_start.AllowNone = True
                    Point_start = Editor1.GetPoint(PP_start)
                    If Not Point_start.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        Editor1.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If


                    Dim Colectie_atr_name As New Specialized.StringCollection
                    Dim Colectie_atr_value As New Specialized.StringCollection

                    If Not ComboBox_crossing_material_DESC.Text = "" Then
                        Colectie_atr_name.Add("DESC")
                        Colectie_atr_value.Add(ComboBox_crossing_material_DESC.Text)
                    End If

                    If Not ComboBox_crossing_material_ID_NO.Text = "" Then
                        Colectie_atr_name.Add("ID_NO")
                        Colectie_atr_value.Add(ComboBox_crossing_material_ID_NO.Text)
                    End If

                    If Not ComboBox_crossing_material_STA.Text = "" Then
                        Colectie_atr_name.Add("STA")
                        Colectie_atr_value.Add(ComboBox_crossing_material_STA.Text)
                    End If

                    If Not ComboBox_crossing_material_COV.Text = "" Then
                        Colectie_atr_name.Add("COV")
                        Colectie_atr_value.Add(ComboBox_crossing_material_COV.Text)
                    End If


                    InsertBlock_with_multiple_atributes("General_alignment_crossing.dwg", "General_alignment_crossing", Point_start.Value, 1, BTrecord, "TEXT", Colectie_atr_name, Colectie_atr_value)

                    For i = 0 To Rezultat1.Value.Count - 1






                        Dim ent1 As Entity = Rezultat1.Value.Item(i).ObjectId.GetObject(OpenMode.ForWrite)
                        ent1.Erase()
                    Next








                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)

                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_button_crossing_material_desc_Click(sender As Object, e As EventArgs) Handles Button_crossing_material_desc.Click

        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info {DESC}:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut1 As String


                    If TypeOf ent1 Is DBText Then
                        Dim Text1 As DBText = ent1
                        Continut1 = Text1.TextString
                    End If

                    If TypeOf ent1 Is MText Then
                        Dim MText1 As MText = ent1
                        Continut1 = MText1.Contents
                    End If


                    If Not Continut1 = "" Then
                        If ComboBox_crossing_material_DESC.Items.Count > 1 Then
                            ComboBox_crossing_material_DESC.Items(1) = Continut1
                            ComboBox_crossing_material_DESC.SelectedIndex = ComboBox_crossing_material_DESC.Items.IndexOf(Continut1)
                            Button_crossing_material_desc.BackColor = Drawing.Color.Lime
                            ComboBox_crossing_material_ID_NO.Items(1) = Continut1
                            ComboBox_crossing_material_STA.Items(1) = Continut1
                            ComboBox_crossing_material_COV.Items(1) = Continut1
                        End If
                        If ComboBox_crossing_material_DESC.Items.Count = 0 Then
                            ComboBox_crossing_material_DESC.Items.add(" ")
                            ComboBox_crossing_material_DESC.Items.Add(Continut1)
                            ComboBox_crossing_material_DESC.SelectedIndex = ComboBox_crossing_material_DESC.Items.IndexOf(Continut1)
                            Button_crossing_material_desc.BackColor = Drawing.Color.Lime
                            ComboBox_crossing_material_ID_NO.Items.add(" ")
                            ComboBox_crossing_material_ID_NO.Items.Add(Continut1)
                            ComboBox_crossing_material_STA.Items.add(" ")
                            ComboBox_crossing_material_STA.Items.Add(Continut1)
                            ComboBox_crossing_material_COV.Items.add(" ")
                            ComboBox_crossing_material_COV.Items.Add(Continut1)
                        End If

                    Else
                        If ComboBox_crossing_material_DESC.Items.Count = 0 Then
                            Button_crossing_material_desc.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")

    End Sub
    Private Sub button_crossing_material_id_no_Click(sender As Object, e As EventArgs) Handles Button_crossing_material_id_no.Click

        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt2.MessageForAdding = vbLf & "Select info {ID_NO}:"

                    Object_Prompt2.SingleOnly = True
                    Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                    If Not Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent2 As Entity = Rezultat2.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut2 As String


                    If TypeOf ent2 Is DBText Then
                        Dim Text2 As DBText = ent2
                        Continut2 = Text2.TextString
                    End If

                    If TypeOf ent2 Is MText Then
                        Dim MText2 As MText = ent2
                        Continut2 = MText2.Contents
                    End If


                    If Not Continut2 = "" Then
                        If ComboBox_crossing_material_ID_NO.Items.Count > 2 Then
                            ComboBox_crossing_material_DESC.Items(2) = Continut2
                            ComboBox_crossing_material_ID_NO.Items(2) = Continut2
                            ComboBox_crossing_material_ID_NO.SelectedIndex = ComboBox_crossing_material_ID_NO.Items.IndexOf(Continut2)
                            Button_crossing_material_id_no.BackColor = Drawing.Color.Lime
                            ComboBox_crossing_material_STA.Items(2) = Continut2
                            ComboBox_crossing_material_COV.Items(2) = Continut2
                        End If
                        If ComboBox_crossing_material_ID_NO.Items.Count = 0 Then
                            ComboBox_crossing_material_DESC.Items.add(" ")
                            ComboBox_crossing_material_DESC.Items.Add(Continut2)
                            ComboBox_crossing_material_ID_NO.Items.add(" ")
                            ComboBox_crossing_material_ID_NO.Items.Add(Continut2)
                            ComboBox_crossing_material_ID_NO.SelectedIndex = ComboBox_crossing_material_ID_NO.Items.IndexOf(Continut2)
                            Button_crossing_material_id_no.BackColor = Drawing.Color.Lime
                            ComboBox_crossing_material_STA.Items.add(" ")
                            ComboBox_crossing_material_STA.Items.Add(Continut2)
                            ComboBox_crossing_material_COV.Items.add(" ")
                            ComboBox_crossing_material_COV.Items.Add(Continut2)
                        End If
                        If ComboBox_crossing_material_ID_NO.Items.Count = 1 Or ComboBox_crossing_material_ID_NO.Items.Count = 2 Then
                            ComboBox_crossing_material_DESC.Items.Add(Continut2)
                            ComboBox_crossing_material_ID_NO.Items.Add(Continut2)
                            ComboBox_crossing_material_ID_NO.SelectedIndex = ComboBox_crossing_material_ID_NO.Items.IndexOf(Continut2)
                            Button_crossing_material_id_no.BackColor = Drawing.Color.Lime
                            ComboBox_crossing_material_STA.Items.Add(Continut2)
                            ComboBox_crossing_material_COV.Items.Add(Continut2)
                        End If
                    Else
                        If ComboBox_crossing_material_ID_NO.Items.Count = 0 Then
                            Button_crossing_material_id_no.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")

    End Sub
    Private Sub button_crossing_material_sta_Click(sender As Object, e As EventArgs) Handles Button_crossing_material_sta.Click

        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat3 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt3 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt3.MessageForAdding = vbLf & "Select info {STA}:"

                    Object_Prompt3.SingleOnly = True
                    Rezultat3 = Editor1.GetSelection(Object_Prompt3)
                    If Not Rezultat3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent3 As Entity = Rezultat3.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut3 As String


                    If TypeOf ent3 Is DBText Then
                        Dim Text3 As DBText = ent3
                        Continut3 = Text3.TextString
                    End If

                    If TypeOf ent3 Is MText Then
                        Dim MText3 As MText = ent3
                        Continut3 = MText3.Contents
                    End If


                    If Not Continut3 = "" Then
                        If ComboBox_crossing_material_STA.Items.Count > 3 Then
                            ComboBox_crossing_material_DESC.Items(3) = Continut3
                            ComboBox_crossing_material_ID_NO.Items(3) = Continut3
                            ComboBox_crossing_material_STA.Items(3) = Continut3
                            ComboBox_crossing_material_STA.SelectedIndex = ComboBox_crossing_material_STA.Items.IndexOf(Continut3)
                            Button_crossing_material_sta.BackColor = Drawing.Color.Lime
                            ComboBox_crossing_material_COV.Items(3) = Continut3
                        End If
                        If ComboBox_crossing_material_STA.Items.Count = 0 Then
                            ComboBox_crossing_material_DESC.Items.add(" ")
                            ComboBox_crossing_material_DESC.Items.Add(Continut3)
                            ComboBox_crossing_material_ID_NO.Items.add(" ")
                            ComboBox_crossing_material_ID_NO.Items.Add(Continut3)

                            ComboBox_crossing_material_STA.Items.add(" ")
                            ComboBox_crossing_material_STA.Items.Add(Continut3)
                            ComboBox_crossing_material_STA.SelectedIndex = ComboBox_crossing_material_STA.Items.IndexOf(Continut3)
                            Button_crossing_material_sta.BackColor = Drawing.Color.Lime
                            ComboBox_crossing_material_COV.Items.add(" ")
                            ComboBox_crossing_material_COV.Items.Add(Continut3)
                        End If
                        If ComboBox_crossing_material_STA.Items.Count = 1 Or ComboBox_crossing_material_STA.Items.Count = 2 Or ComboBox_crossing_material_STA.Items.Count = 3 Then
                            ComboBox_crossing_material_DESC.Items.Add(Continut3)
                            ComboBox_crossing_material_ID_NO.Items.Add(Continut3)

                            ComboBox_crossing_material_STA.Items.Add(Continut3)
                            ComboBox_crossing_material_STA.SelectedIndex = ComboBox_crossing_material_STA.Items.IndexOf(Continut3)
                            Button_crossing_material_sta.BackColor = Drawing.Color.Lime

                            ComboBox_crossing_material_COV.Items.Add(Continut3)
                        End If
                    Else
                        If ComboBox_crossing_material_STA.Items.Count = 0 Then
                            Button_crossing_material_sta.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")

    End Sub
    Private Sub button_crossing_material_cov_Click(sender As Object, e As EventArgs) Handles Button_crossing_material_cov.Click

        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat4 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt4 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt4.MessageForAdding = vbLf & "Select info {COV}:"

                    Object_Prompt4.SingleOnly = True
                    Rezultat4 = Editor1.GetSelection(Object_Prompt4)
                    If Not Rezultat4.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent4 As Entity = Rezultat4.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut4 As String


                    If TypeOf ent4 Is DBText Then
                        Dim Text4 As DBText = ent4
                        Continut4 = Text4.TextString
                    End If

                    If TypeOf ent4 Is MText Then
                        Dim MText3 As MText = ent4
                        Continut4 = MText3.Contents
                    End If


                    If Not Continut4 = "" Then
                        If ComboBox_crossing_material_COV.Items.Count > 4 Then
                            ComboBox_crossing_material_DESC.Items(4) = Continut4
                            ComboBox_crossing_material_ID_NO.Items(4) = Continut4
                            ComboBox_crossing_material_STA.Items(4) = Continut4
                            ComboBox_crossing_material_COV.Items(4) = Continut4
                            ComboBox_crossing_material_COV.SelectedIndex = ComboBox_crossing_material_COV.Items.IndexOf(Continut4)
                            Button_crossing_material_cov.BackColor = Drawing.Color.Lime
                        End If
                        If ComboBox_crossing_material_COV.Items.Count = 0 Then
                            ComboBox_crossing_material_DESC.Items.add(" ")
                            ComboBox_crossing_material_DESC.Items.Add(Continut4)
                            ComboBox_crossing_material_ID_NO.Items.add(" ")
                            ComboBox_crossing_material_ID_NO.Items.Add(Continut4)

                            ComboBox_crossing_material_STA.Items.add(" ")
                            ComboBox_crossing_material_STA.Items.Add(Continut4)

                            ComboBox_crossing_material_COV.Items.add(" ")
                            ComboBox_crossing_material_COV.Items.Add(Continut4)
                            ComboBox_crossing_material_COV.SelectedIndex = ComboBox_crossing_material_COV.Items.IndexOf(Continut4)
                            Button_crossing_material_cov.BackColor = Drawing.Color.Lime
                        End If
                        If ComboBox_crossing_material_COV.Items.Count = 1 Or ComboBox_crossing_material_COV.Items.Count = 2 Or ComboBox_crossing_material_COV.Items.Count = 3 Or ComboBox_crossing_material_COV.Items.Count = 4 Then
                            ComboBox_crossing_material_DESC.Items.Add(Continut4)
                            ComboBox_crossing_material_ID_NO.Items.Add(Continut4)

                            ComboBox_crossing_material_STA.Items.Add(Continut4)


                            ComboBox_crossing_material_COV.Items.Add(Continut4)
                            ComboBox_crossing_material_COV.SelectedIndex = ComboBox_crossing_material_COV.Items.IndexOf(Continut4)
                            Button_crossing_material_cov.BackColor = Drawing.Color.Lime
                        End If
                    Else
                        If ComboBox_crossing_material_COV.Items.Count = 0 Then
                            Button_crossing_material_cov.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")

    End Sub
    Private Sub Button_crossing_material_BlockPick_Click(sender As Object, e As EventArgs) Handles Button_crossing_material_BlockPick.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info from Block:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    ComboBox_crossing_material_DESC.Items.Clear()
                    ComboBox_crossing_material_DESC.Items.Add(" ")
                    ComboBox_crossing_material_ID_NO.Items.Clear()
                    ComboBox_crossing_material_ID_NO.Items.Add(" ")
                    ComboBox_crossing_material_STA.Items.Clear()
                    ComboBox_crossing_material_STA.Items.Add(" ")
                    ComboBox_crossing_material_COV.Items.Clear()
                    ComboBox_crossing_material_COV.Items.Add(" ")


                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut1 As String
                    Dim Continut2 As String
                    Dim Continut3 As String
                    Dim Continut4 As String

                    If TypeOf ent1 Is BlockReference Then
                        Dim Block1 As BlockReference = ent1
                        If Block1.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "DESC" Then
                                    Continut1 = attref.TextString
                                    If Continut1 = "" Then Continut1 = " "
                                End If
                                If attref.Tag = "ID_NO" Then
                                    Continut2 = attref.TextString
                                    If Continut2 = "" Then Continut2 = " "
                                End If
                                If attref.Tag = "STA" Then
                                    Continut3 = attref.TextString
                                    If Continut3 = "" Then Continut3 = " "
                                End If
                                If attref.Tag = "COV" Then
                                    Continut4 = attref.TextString
                                    If Continut4 = "" Then Continut4 = " "
                                End If
                            Next
                        End If
                    End If

                    If Not Continut1 = "" Then
                        ComboBox_crossing_material_DESC.Items.Add(Continut1)
                        ComboBox_crossing_material_DESC.SelectedIndex = ComboBox_crossing_material_DESC.Items.IndexOf(Continut1)
                        Button_crossing_material_desc.BackColor = Drawing.Color.Lime
                        ComboBox_crossing_material_ID_NO.Items.Add(Continut1)
                        ComboBox_crossing_material_STA.Items.Add(Continut1)
                        ComboBox_crossing_material_COV.Items.Add(Continut1)

                    End If
                    If Not Continut2 = "" Then
                        ComboBox_crossing_material_DESC.Items.Add(Continut2)
                        ComboBox_crossing_material_ID_NO.Items.Add(Continut2)
                        ComboBox_crossing_material_ID_NO.SelectedIndex = ComboBox_crossing_material_ID_NO.Items.IndexOf(Continut2)
                        Button_crossing_material_id_no.BackColor = Drawing.Color.Lime
                        ComboBox_crossing_material_STA.Items.Add(Continut2)
                        ComboBox_crossing_material_COV.Items.Add(Continut2)
                    End If
                    If Not Continut3 = "" Then
                        ComboBox_crossing_material_DESC.Items.Add(Continut3)
                        ComboBox_crossing_material_ID_NO.Items.Add(Continut3)
                        ComboBox_crossing_material_STA.Items.Add(Continut3)
                        ComboBox_crossing_material_STA.SelectedIndex = ComboBox_crossing_material_STA.Items.IndexOf(Continut3)
                        Button_crossing_material_sta.BackColor = Drawing.Color.Lime
                        ComboBox_crossing_material_COV.Items.Add(Continut3)
                    End If
                    If Not Continut4 = "" Then
                        ComboBox_crossing_material_DESC.Items.Add(Continut4)
                        ComboBox_crossing_material_ID_NO.Items.Add(Continut4)
                        ComboBox_crossing_material_STA.Items.Add(Continut4)
                        ComboBox_crossing_material_COV.Items.Add(Continut4)
                        ComboBox_crossing_material_COV.SelectedIndex = ComboBox_crossing_material_COV.Items.IndexOf(Continut4)
                        Button_crossing_material_cov.BackColor = Drawing.Color.Lime
                    End If






                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_crossing_material_pick_from_model_space_Click(sender As Object, e As EventArgs) Handles Button_crossing_material_pick_from_model_space.Click
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult


            Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt2.MessageForAdding = vbLf & "Select 3D polyline:"

            Object_Prompt2.SingleOnly = True

            Rezultat2 = Editor1.GetSelection(Object_Prompt2)


            If Rezultat2.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                Exit Sub
            End If


            Dim Poly3d As Polyline3d


            Dim Point_on_poly As New Point3d


            If Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat2) = False Then
                    Using lock1 As DocumentLock = ThisDrawing.LockDocument
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


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

                            Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult

                            Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select Point:")
                            PP1.AllowNone = True
                            Point1 = Editor1.GetPoint(PP1)

                            If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                Editor1.WriteMessage(vbLf & "Command:")
                                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                Exit Sub
                            End If

                            Point_on_poly = Poly3d.GetClosestPointTo(Point1.Value.TransformBy(curent_ucs_matrix), Editor1.GetCurrentView.ViewDirection, False)

                            Dim Parameter_picked As Double = Round(Poly3d.GetParameterAtPoint(Point_on_poly), 3)

                            Dim Parameter_start As Double = Floor(Parameter_picked)
                            Dim Parameter_end As Double = Ceiling(Parameter_picked)
                            If Parameter_picked = Round(Parameter_picked, 0) Then
                                Parameter_start = Parameter_picked
                                Parameter_end = Parameter_picked
                            End If


                            Dim Data_table1 As New System.Data.DataTable
                            Data_table1.Columns.Add("TEXT325", GetType(DBText))
                            Dim Index1 As Double = 0

                            Dim Data_table2 As New System.Data.DataTable
                            Data_table2.Columns.Add("TEXT0", GetType(DBText))
                            Dim Index2 As Double = 0

                            For Each ObjID In BTrecord
                                Dim DBobject As DBObject = Trans1.GetObject(ObjID, OpenMode.ForRead)
                                If TypeOf DBobject Is DBText Then
                                    Dim Text1 As DBText = DBobject
                                    If Text1.Layer = ComboBox_layer.Text Then
                                        If Text1.Rotation > 3 * PI / 2 Then
                                            Data_table1.Rows.Add()
                                            Data_table1.Rows(Index1).Item("TEXT325") = Text1
                                            Index1 = Index1 + 1
                                        End If
                                        If Text1.Rotation >= 0 And Text1.Rotation < PI / 4 Then
                                            Data_table2.Rows.Add()
                                            Data_table2.Rows(Index2).Item("TEXT0") = Text1
                                            Index2 = Index2 + 1
                                        End If

                                    End If
                                End If
                            Next
                            Dim Chainage_on_vertex As Double
                            Dim Distanta_pana_la_Vertex As Double
                            Dim CSF1, CSF2 As Double

                            If Data_table1.Rows.Count > 0 Then
                                Dim Point_CHAINAGE As New Point3d
                                Point_CHAINAGE = Poly3d.GetPointAtParameter(Parameter_start)
                                Distanta_pana_la_Vertex = Point_CHAINAGE.GetVectorTo(Point_on_poly).Length

                                For i = 0 To Data_table1.Rows.Count - 1
                                    Dim Text1 As DBText = Data_table1.Rows(i).Item("TEXT325")
                                    If Point_CHAINAGE.GetVectorTo(Text1.Position.TransformBy(Editor1.CurrentUserCoordinateSystem)).Length < 0.1 Then
                                        Dim String1 As String = Replace(Text1.TextString, "+", "")
                                        If IsNumeric(String1) = True Then
                                            Chainage_on_vertex = CDbl(String1)
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If

                            If Not Parameter_start = Parameter_end Then
                                If Data_table2.Rows.Count > 0 Then
                                    Dim Point_CHAINAGE1 As New Point3d
                                    Point_CHAINAGE1 = Poly3d.GetPointAtParameter(Parameter_start)
                                    Dim Point_CHAINAGE2 As New Point3d
                                    Point_CHAINAGE2 = Poly3d.GetPointAtParameter(Parameter_end)

                                    For i = 0 To Data_table2.Rows.Count - 1
                                        Dim Text1 As DBText = Data_table2.Rows(i).Item("TEXT0")
                                        Dim String1 As String = Text1.TextString
                                        String1 = extrage_numar_din_text_de_la_sfarsitul_textului(String1)
                                        If IsNumeric(String1) = True Then
                                            If Point_CHAINAGE1.GetVectorTo(Text1.Position.TransformBy(Editor1.CurrentUserCoordinateSystem)).Length < 0.1 Then
                                                If CDbl(String1) > 0.5 And CDbl(String1) < 1.5 Then
                                                    CSF1 = CDbl(String1)
                                                End If
                                            End If

                                            If Point_CHAINAGE2.GetVectorTo(Text1.Position.TransformBy(Editor1.CurrentUserCoordinateSystem)).Length < 0.1 Then
                                                If CDbl(String1) > 0.5 And CDbl(String1) < 1.5 Then
                                                    CSF2 = CDbl(String1)
                                                End If
                                            End If
                                        End If
                                    Next
                                End If
                            End If

                            Dim New_chainage As String
                            Dim New_ch As Double
                            If Not CSF1 + CSF2 = 0 And Not CSF1 = 0 And Not CSF2 = 0 Then
                                New_ch = Chainage_on_vertex + Distanta_pana_la_Vertex / ((CSF1 + CSF2) / 2)
                            Else
                                New_ch = Chainage_on_vertex + Distanta_pana_la_Vertex
                            End If


                            New_chainage = Get_chainage_from_double(New_ch, 1)

                            ComboBox_crossing_material_STA.Items.Add(New_chainage)
                            ComboBox_crossing_material_STA.SelectedIndex = ComboBox_crossing_material_STA.Items.IndexOf(New_chainage)
                            Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly, New_chainage, 0.5, 0.5, 0.5, 11, 3.5)

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

    Private Sub Button_EL_CLEAR_Click(sender As Object, e As EventArgs) Handles Button_EL_CLEAR.Click
        ComboBox_EL_desc.Items.Clear()
        ComboBox_el_id_no.Items.Clear()
        ComboBox_el_beg_sta.Items.Clear()
        ComboBox_EL_STA.Items.Clear()
        ComboBox_el_END_STA.Items.Clear()
        ComboBox_EL_LENGTH.Items.Clear()
        Button_EL_DESC.BackColor = Drawing.Color.DimGray
        Button_EL_ID_NO.BackColor = Drawing.Color.DimGray
        Button_EL_BEG_STA.BackColor = Drawing.Color.DimGray
        Button_EL_STA.BackColor = Drawing.Color.DimGray
        Button_EL_END_STA.BackColor = Drawing.Color.DimGray
        Button_EL_LENGTH.BackColor = Drawing.Color.DimGray
    End Sub
    Private Sub Button_EL_PICK_Click(sender As Object, e As EventArgs) Handles Button_EL_PICK.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info {DESC}:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt2.MessageForAdding = vbLf & "Select info {ID_NO}:"

                    Object_Prompt2.SingleOnly = True
                    Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                    If Not Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    Dim Rezultat3 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt3 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt3.MessageForAdding = vbLf & "Select info {BEGINSTA}:"

                    Object_Prompt3.SingleOnly = True
                    Rezultat3 = Editor1.GetSelection(Object_Prompt3)
                    If Not Rezultat3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    Dim Rezultat4 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt4 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt4.MessageForAdding = vbLf & "Select info {STA}:"

                    Object_Prompt4.SingleOnly = True
                    Rezultat4 = Editor1.GetSelection(Object_Prompt4)
                    If Not Rezultat4.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If


                    Dim Rezultat5 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt5 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt5.MessageForAdding = vbLf & "Select info {ENDSTA}:"

                    Object_Prompt5.SingleOnly = True
                    Rezultat5 = Editor1.GetSelection(Object_Prompt5)
                    If Not Rezultat5.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    Dim Rezultat6 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt6 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt6.MessageForAdding = vbLf & "Select info {LENGTH}:"

                    Object_Prompt6.SingleOnly = True
                    Rezultat6 = Editor1.GetSelection(Object_Prompt6)
                    If Not Rezultat6.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    ComboBox_EL_desc.Items.Clear()
                    ComboBox_EL_desc.Items.add(" ")
                    ComboBox_el_id_no.Items.Clear()
                    ComboBox_el_id_no.Items.add(" ")
                    ComboBox_el_beg_sta.Items.Clear()
                    ComboBox_el_beg_sta.Items.add(" ")
                    ComboBox_EL_STA.Items.Clear()
                    ComboBox_EL_STA.Items.add(" ")
                    ComboBox_el_END_STA.Items.Clear()
                    ComboBox_el_END_STA.Items.add(" ")
                    ComboBox_EL_LENGTH.Items.Clear()
                    ComboBox_EL_LENGTH.Items.add(" ")

                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent2 As Entity = Rezultat2.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent3 As Entity = Rezultat3.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent4 As Entity = Rezultat4.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent5 As Entity = Rezultat5.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent6 As Entity = Rezultat6.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)

                    Dim Continut1 As String
                    Dim Continut2 As String
                    Dim Continut3 As String
                    Dim Continut4 As String
                    Dim Continut5 As String
                    Dim Continut6 As String

                    If TypeOf ent1 Is DBText Then
                        Dim Text1 As DBText = ent1
                        Continut1 = Text1.TextString
                    End If
                    If TypeOf ent2 Is DBText Then
                        Dim Text2 As DBText = ent2
                        Continut2 = Text2.TextString
                    End If
                    If TypeOf ent3 Is DBText Then
                        Dim Text3 As DBText = ent3
                        Continut3 = Text3.TextString
                    End If
                    If TypeOf ent4 Is DBText Then
                        Dim Text4 As DBText = ent4
                        Continut4 = Text4.TextString
                    End If
                    If TypeOf ent5 Is DBText Then
                        Dim Text5 As DBText = ent5
                        Continut5 = Text5.TextString
                    End If
                    If TypeOf ent6 Is DBText Then
                        Dim Text6 As DBText = ent6
                        Continut6 = Text6.TextString
                    End If

                    If TypeOf ent1 Is MText Then
                        Dim MText1 As MText = ent1
                        Continut1 = MText1.Contents
                    End If
                    If TypeOf ent2 Is MText Then
                        Dim MText2 As MText = ent2
                        Continut2 = MText2.Contents
                    End If
                    If TypeOf ent3 Is MText Then
                        Dim MText3 As MText = ent3
                        Continut3 = MText3.Contents
                    End If
                    If TypeOf ent4 Is MText Then
                        Dim MText4 As MText = ent4
                        Continut4 = MText4.Contents
                    End If
                    If TypeOf ent5 Is MText Then
                        Dim MText5 As MText = ent5
                        Continut5 = MText5.Contents
                    End If

                    If TypeOf ent6 Is MText Then
                        Dim MText6 As MText = ent6
                        Continut6 = MText6.Contents
                    End If

                    If TypeOf ent1 Is BlockReference Then
                        Dim Block1 As BlockReference = ent1
                        If Block1.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "DESC" Then
                                    Continut1 = attref.TextString
                                End If
                            Next
                        End If
                    End If

                    If TypeOf ent2 Is BlockReference Then
                        Dim Block2 As BlockReference = ent2
                        If Block2.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block2.AttributeCollection
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "ID_NO" Then
                                    Continut2 = attref.TextString
                                End If
                            Next
                        End If
                    End If


                    If TypeOf ent3 Is BlockReference Then
                        Dim Block3 As BlockReference = ent3
                        If Block3.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block3.AttributeCollection
                            Dim ctemp1 As String
                            Dim ctemp2 As String
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "BEGINSTA" Then
                                    ctemp1 = attref.TextString
                                End If
                                If attref.Tag = "ENDSTA" Then
                                    ctemp2 = attref.TextString
                                End If
                            Next
                            ctemp1 = Replace(ctemp1, " ", "")
                            ctemp2 = Replace(ctemp2, " ", "")



                            If Continut5 = "" Then
                                If TypeOf ent5 Is BlockReference And Not ent3.ObjectId = ent5.ObjectId Then
                                    If Not ctemp1 = "" And Not ctemp2 = "" Then
                                        Continut3 = ctemp2
                                    End If
                                    If ctemp1 = "" And ctemp2 = "" Then
                                        Continut3 = " "
                                    End If
                                    If ctemp1 = "" And Not ctemp2 = "" Then
                                        Continut3 = ctemp2
                                    End If
                                    If Not ctemp1 = "" And ctemp2 = "" Then
                                        Continut3 = ctemp1
                                    End If
                                End If
                            Else
                                If Not ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut3 = ctemp2
                                End If
                                If ctemp1 = "" And ctemp2 = "" Then
                                    Continut3 = " "
                                End If
                                If ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut3 = ctemp2
                                End If
                                If Not ctemp1 = "" And ctemp2 = "" Then
                                    Continut3 = ctemp1
                                End If
                            End If
                            If ent1.ObjectId = ent2.ObjectId Then
                                Continut3 = ctemp1
                            End If

                        End If
                    End If


                    If TypeOf ent4 Is BlockReference Then
                        Dim Block4 As BlockReference = ent4
                        If Block4.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block4.AttributeCollection
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "STA" Then
                                    Continut4 = attref.TextString
                                End If
                            Next
                        End If
                    End If



                    If TypeOf ent5 Is BlockReference Then
                        Dim Block5 As BlockReference = ent5
                        If Block5.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block5.AttributeCollection
                            Dim ctemp1 As String
                            Dim ctemp2 As String
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "BEGINSTA" Then
                                    ctemp1 = attref.TextString
                                End If
                                If attref.Tag = "ENDSTA" Then
                                    ctemp2 = attref.TextString
                                End If
                            Next
                            ctemp1 = Replace(ctemp1, " ", "")
                            ctemp2 = Replace(ctemp2, " ", "")
                            If TypeOf ent3 Is BlockReference Then
                                If Not ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut5 = ctemp1
                                End If
                                If ctemp1 = "" And ctemp2 = "" Then
                                    Continut5 = " "
                                End If
                                If ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut5 = ctemp2
                                End If
                                If Not ctemp1 = "" And ctemp2 = "" Then
                                    Continut5 = ctemp1
                                End If
                            Else
                                Continut5 = ctemp1
                            End If
                            If ent3.ObjectId = ent5.ObjectId Then
                                Continut2 = ctemp2
                            End If


                        End If
                    End If

                    If TypeOf ent6 Is BlockReference Then
                        Dim Block6 As BlockReference = ent6
                        If Block6.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block6.AttributeCollection
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "LENGTH" Then
                                    Continut6 = attref.TextString
                                End If
                            Next
                        End If
                    End If


                    If Not Continut1 = "" Then
                        ComboBox_EL_desc.Items.Add(Continut1)
                        ComboBox_EL_desc.SelectedIndex = ComboBox_EL_desc.Items.IndexOf(Continut1)
                        Button_EL_DESC.BackColor = Drawing.Color.Lime
                        ComboBox_el_id_no.Items.Add(Continut1)
                        ComboBox_el_beg_sta.Items.Add(Continut1)
                        ComboBox_EL_STA.Items.Add(Continut1)
                        ComboBox_el_END_STA.Items.Add(Continut1)
                        ComboBox_EL_LENGTH.Items.Add(Continut1)
                    End If
                    If Not Continut2 = "" Then
                        ComboBox_EL_desc.Items.Add(Continut2)
                        ComboBox_el_id_no.Items.Add(Continut2)
                        ComboBox_el_id_no.SelectedIndex = ComboBox_el_id_no.Items.IndexOf(Continut2)
                        Button_EL_ID_NO.BackColor = Drawing.Color.Lime
                        ComboBox_el_beg_sta.Items.Add(Continut2)
                        ComboBox_EL_STA.Items.Add(Continut2)
                        ComboBox_el_END_STA.Items.Add(Continut2)
                        ComboBox_EL_LENGTH.Items.Add(Continut2)
                    End If
                    If Not Continut3 = "" Then
                        ComboBox_EL_desc.Items.Add(Continut3)
                        ComboBox_el_id_no.Items.Add(Continut3)
                        ComboBox_el_beg_sta.Items.Add(Continut3)
                        ComboBox_el_beg_sta.SelectedIndex = ComboBox_el_beg_sta.Items.IndexOf(Continut3)
                        Button_EL_BEG_STA.BackColor = Drawing.Color.Lime
                        ComboBox_EL_STA.Items.Add(Continut3)
                        ComboBox_el_END_STA.Items.Add(Continut3)
                        ComboBox_EL_LENGTH.Items.Add(Continut3)
                    End If
                    If Not Continut4 = "" Then
                        ComboBox_EL_desc.Items.Add(Continut4)
                        ComboBox_el_id_no.Items.Add(Continut4)
                        ComboBox_el_beg_sta.Items.Add(Continut4)
                        ComboBox_EL_STA.Items.Add(Continut4)
                        ComboBox_EL_STA.SelectedIndex = ComboBox_EL_STA.Items.IndexOf(Continut4)
                        Button_EL_STA.BackColor = Drawing.Color.Lime
                        ComboBox_el_END_STA.Items.Add(Continut4)
                        ComboBox_EL_LENGTH.Items.Add(Continut4)
                    End If

                    If Not Continut5 = "" Then
                        ComboBox_EL_desc.Items.Add(Continut5)
                        ComboBox_el_id_no.Items.Add(Continut5)
                        ComboBox_el_beg_sta.Items.Add(Continut5)
                        ComboBox_EL_STA.Items.Add(Continut5)
                        ComboBox_el_END_STA.Items.Add(Continut5)
                        ComboBox_el_END_STA.SelectedIndex = ComboBox_el_END_STA.Items.IndexOf(Continut5)
                        Button_EL_END_STA.BackColor = Drawing.Color.Lime
                        ComboBox_EL_LENGTH.Items.Add(Continut5)
                    End If

                    If Not Continut6 = "" Then
                        ComboBox_EL_desc.Items.Add(Continut6)
                        ComboBox_el_id_no.Items.Add(Continut6)
                        ComboBox_el_beg_sta.Items.Add(Continut6)
                        ComboBox_EL_STA.Items.Add(Continut6)
                        ComboBox_el_END_STA.Items.Add(Continut6)
                        ComboBox_EL_LENGTH.Items.Add(Continut6)
                        ComboBox_EL_LENGTH.SelectedIndex = ComboBox_EL_LENGTH.Items.IndexOf(Continut6)
                        Button_EL_LENGTH.BackColor = Drawing.Color.Lime
                    End If


                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_EL_INSERT_Click(sender As Object, e As EventArgs) Handles Button_EL_INSERT.Click

        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select objects to be deleted:"

                    Object_Prompt1.SingleOnly = False
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If
                    Dim Point_start As Autodesk.AutoCAD.EditorInput.PromptPointResult

                    Dim PP_start As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select insertion point:")
                    PP_start.AllowNone = True
                    Point_start = Editor1.GetPoint(PP_start)
                    If Not Point_start.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        Editor1.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If


                    Dim Colectie_atr_name As New Specialized.StringCollection
                    Dim Colectie_atr_value As New Specialized.StringCollection

                    If Not ComboBox_EL_desc.Text = "" Then
                        Colectie_atr_name.Add("DESC")
                        Colectie_atr_value.Add(ComboBox_EL_desc.Text)
                    End If
                    If Not ComboBox_el_id_no.Text = "" Then
                        Colectie_atr_name.Add("ID_NO")
                        Colectie_atr_value.Add(ComboBox_el_id_no.Text)
                    End If
                    If Not ComboBox_el_beg_sta.Text = "" Then
                        Colectie_atr_name.Add("BEGINSTA")
                        Colectie_atr_value.Add(ComboBox_el_beg_sta.Text)
                    End If

                    If Not ComboBox_EL_STA.Text = "" Then
                        Colectie_atr_name.Add("STA")
                        Colectie_atr_value.Add(ComboBox_EL_STA.Text)
                    End If

                    If Not ComboBox_el_END_STA.Text = "" Then
                        Colectie_atr_name.Add("ENDSTA")
                        Colectie_atr_value.Add(ComboBox_el_END_STA.Text)
                    End If

                    If Not ComboBox_el_END_STA.Text = "" Then
                        Colectie_atr_name.Add("LENGTH")
                        Colectie_atr_value.Add(ComboBox_EL_LENGTH.Text)
                    End If

                    InsertBlock_with_multiple_atributes("Elbow_alignment.dwg", "Elbow_alignment", Point_start.Value, 1, BTrecord, "TEXT", Colectie_atr_name, Colectie_atr_value)

                    For i = 0 To Rezultat1.Value.Count - 1






                        Dim ent1 As Entity = Rezultat1.Value.Item(i).ObjectId.GetObject(OpenMode.ForWrite)
                        ent1.Erase()
                    Next








                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub button_Panel_el_desc_Click(sender As Object, e As EventArgs) Handles Button_EL_DESC.Click

        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info {DESC}:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut1 As String


                    If TypeOf ent1 Is DBText Then
                        Dim Text1 As DBText = ent1
                        Continut1 = Text1.TextString
                    End If

                    If TypeOf ent1 Is MText Then
                        Dim MText1 As MText = ent1
                        Continut1 = MText1.Contents
                    End If


                    If Not Continut1 = "" Then
                        If ComboBox_EL_desc.Items.Count > 1 Then
                            ComboBox_EL_desc.Items(1) = Continut1
                            ComboBox_EL_desc.SelectedIndex = ComboBox_EL_desc.Items.IndexOf(Continut1)
                            Button_EL_DESC.BackColor = Drawing.Color.Lime
                            ComboBox_el_id_no.Items(1) = Continut1
                            ComboBox_el_beg_sta.Items(1) = Continut1
                            ComboBox_EL_STA.Items(1) = Continut1
                            ComboBox_el_END_STA.Items(1) = Continut1
                            ComboBox_EL_LENGTH.Items(1) = Continut1
                        End If
                        If ComboBox_EL_desc.Items.Count = 0 Then
                            ComboBox_EL_desc.Items.add(" ")
                            ComboBox_EL_desc.Items.Add(Continut1)
                            ComboBox_EL_desc.SelectedIndex = ComboBox_EL_desc.Items.IndexOf(Continut1)
                            Button_EL_DESC.BackColor = Drawing.Color.Lime
                            ComboBox_el_id_no.Items.add(" ")
                            ComboBox_el_id_no.Items.Add(Continut1)
                            ComboBox_el_beg_sta.Items.add(" ")
                            ComboBox_el_beg_sta.Items.Add(Continut1)
                            ComboBox_EL_STA.Items.add(" ")
                            ComboBox_EL_STA.Items.Add(Continut1)
                            ComboBox_el_END_STA.Items.add(" ")
                            ComboBox_el_END_STA.Items.Add(Continut1)
                            ComboBox_EL_LENGTH.Items.add(" ")
                            ComboBox_EL_LENGTH.Items.Add(Continut1)

                        End If

                    Else
                        If ComboBox_crossing_material_DESC.Items.Count = 0 Then
                            Button_EL_DESC.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")

    End Sub
    Private Sub button_Panel_EL_ID_NO_Click(sender As Object, e As EventArgs) Handles Button_EL_ID_NO.Click

        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt2.MessageForAdding = vbLf & "Select info {ID_NO}:"

                    Object_Prompt2.SingleOnly = True
                    Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                    If Not Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent2 As Entity = Rezultat2.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut2 As String


                    If TypeOf ent2 Is DBText Then
                        Dim Text2 As DBText = ent2
                        Continut2 = Text2.TextString
                    End If

                    If TypeOf ent2 Is MText Then
                        Dim MText2 As MText = ent2
                        Continut2 = MText2.Contents
                    End If


                    If Not Continut2 = "" Then
                        If ComboBox_el_id_no.Items.Count > 2 Then
                            ComboBox_EL_desc.Items(2) = Continut2
                            ComboBox_el_id_no.Items(2) = Continut2
                            ComboBox_el_id_no.SelectedIndex = ComboBox_el_id_no.Items.IndexOf(Continut2)
                            Button_EL_ID_NO.BackColor = Drawing.Color.Lime
                            ComboBox_el_beg_sta.Items(2) = Continut2
                            ComboBox_EL_STA.Items(2) = Continut2
                            ComboBox_el_END_STA.Items(2) = Continut2
                            ComboBox_EL_LENGTH.Items(2) = Continut2
                        End If
                        If ComboBox_el_id_no.Items.Count = 0 Then
                            ComboBox_EL_desc.Items.add(" ")
                            ComboBox_EL_desc.Items.Add(Continut2)
                            ComboBox_el_id_no.Items.add(" ")
                            ComboBox_el_id_no.Items.Add(Continut2)
                            ComboBox_el_id_no.SelectedIndex = ComboBox_el_id_no.Items.IndexOf(Continut2)
                            Button_EL_ID_NO.BackColor = Drawing.Color.Lime
                            ComboBox_el_beg_sta.Items.add(" ")
                            ComboBox_el_beg_sta.Items.Add(Continut2)
                            ComboBox_EL_STA.Items.add(" ")
                            ComboBox_EL_STA.Items.Add(Continut2)
                            ComboBox_el_END_STA.Items.add(" ")
                            ComboBox_el_END_STA.Items.Add(Continut2)
                            ComboBox_EL_LENGTH.Items.add(" ")
                            ComboBox_EL_LENGTH.Items.Add(Continut2)
                        End If
                        If ComboBox_el_id_no.Items.Count = 1 Or ComboBox_el_id_no.Items.Count = 2 Then
                            ComboBox_EL_desc.Items.Add(Continut2)
                            ComboBox_el_id_no.Items.Add(Continut2)
                            ComboBox_el_id_no.SelectedIndex = ComboBox_el_id_no.Items.IndexOf(Continut2)
                            Button_EL_ID_NO.BackColor = Drawing.Color.Lime
                            ComboBox_el_beg_sta.Items.Add(Continut2)
                            ComboBox_EL_STA.Items.Add(Continut2)
                            ComboBox_el_END_STA.Items.Add(Continut2)
                            ComboBox_EL_LENGTH.Items.Add(Continut2)
                        End If
                    Else
                        If ComboBox_el_id_no.Items.Count = 0 Then
                            Button_EL_ID_NO.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")

    End Sub
    Private Sub button_Panel_EL_BEG_STA_Click(sender As Object, e As EventArgs) Handles Button_EL_BEG_STA.Click

        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat3 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt3 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt3.MessageForAdding = vbLf & "Select info {STA}:"

                    Object_Prompt3.SingleOnly = True
                    Rezultat3 = Editor1.GetSelection(Object_Prompt3)
                    If Not Rezultat3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent3 As Entity = Rezultat3.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut3 As String


                    If TypeOf ent3 Is DBText Then
                        Dim Text3 As DBText = ent3
                        Continut3 = Text3.TextString
                    End If

                    If TypeOf ent3 Is MText Then
                        Dim MText3 As MText = ent3
                        Continut3 = MText3.Contents
                    End If


                    If Not Continut3 = "" Then
                        If ComboBox_el_beg_sta.Items.Count > 3 Then
                            ComboBox_EL_desc.Items(3) = Continut3
                            ComboBox_el_id_no.Items(3) = Continut3
                            ComboBox_el_beg_sta.Items(3) = Continut3
                            ComboBox_el_beg_sta.SelectedIndex = ComboBox_el_beg_sta.Items.IndexOf(Continut3)
                            Button_EL_BEG_STA.BackColor = Drawing.Color.Lime
                            ComboBox_EL_STA.Items(3) = Continut3
                            ComboBox_el_END_STA.Items(3) = Continut3
                            ComboBox_EL_LENGTH.Items(3) = Continut3
                        End If
                        If ComboBox_el_beg_sta.Items.Count = 0 Then
                            ComboBox_EL_desc.Items.add(" ")
                            ComboBox_EL_desc.Items.Add(Continut3)
                            ComboBox_el_id_no.Items.add(" ")
                            ComboBox_el_id_no.Items.Add(Continut3)

                            ComboBox_el_beg_sta.Items.add(" ")
                            ComboBox_el_beg_sta.Items.Add(Continut3)
                            ComboBox_el_beg_sta.SelectedIndex = ComboBox_el_beg_sta.Items.IndexOf(Continut3)
                            Button_EL_BEG_STA.BackColor = Drawing.Color.Lime
                            ComboBox_EL_STA.Items.add(" ")
                            ComboBox_EL_STA.Items.Add(Continut3)
                            ComboBox_el_END_STA.Items.add(" ")
                            ComboBox_el_END_STA.Items.Add(Continut3)
                            ComboBox_EL_LENGTH.Items.add(" ")
                            ComboBox_EL_LENGTH.Items.Add(Continut3)
                        End If
                        If ComboBox_el_beg_sta.Items.Count = 1 Or ComboBox_el_beg_sta.Items.Count = 2 Or ComboBox_el_beg_sta.Items.Count = 3 Then
                            ComboBox_EL_desc.Items.Add(Continut3)
                            ComboBox_el_id_no.Items.Add(Continut3)

                            ComboBox_el_beg_sta.Items.Add(Continut3)
                            ComboBox_el_beg_sta.SelectedIndex = ComboBox_el_beg_sta.Items.IndexOf(Continut3)
                            Button_EL_BEG_STA.BackColor = Drawing.Color.Lime

                            ComboBox_EL_STA.Items.Add(Continut3)
                            ComboBox_el_END_STA.Items.Add(Continut3)
                            ComboBox_EL_LENGTH.Items.Add(Continut3)
                        End If
                    Else
                        If ComboBox_el_beg_sta.Items.Count = 0 Then
                            Button_EL_BEG_STA.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")

    End Sub
    Private Sub Button_Panel_el_sta_Click(sender As Object, e As EventArgs) Handles Button_EL_STA.Click

        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat4 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt4 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt4.MessageForAdding = vbLf & "Select info {COV}:"

                    Object_Prompt4.SingleOnly = True
                    Rezultat4 = Editor1.GetSelection(Object_Prompt4)
                    If Not Rezultat4.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent4 As Entity = Rezultat4.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut4 As String


                    If TypeOf ent4 Is DBText Then
                        Dim Text4 As DBText = ent4
                        Continut4 = Text4.TextString
                    End If

                    If TypeOf ent4 Is MText Then
                        Dim MText4 As MText = ent4
                        Continut4 = MText4.Contents
                    End If


                    If Not Continut4 = "" Then
                        If ComboBox_EL_STA.Items.Count > 4 Then

                            ComboBox_EL_desc.Items(4) = Continut4
                            ComboBox_el_id_no.Items(4) = Continut4
                            ComboBox_el_beg_sta.Items(4) = Continut4
                            ComboBox_EL_STA.Items(4) = Continut4
                            ComboBox_EL_STA.SelectedIndex = ComboBox_EL_STA.Items.IndexOf(Continut4)
                            Button_EL_STA.BackColor = Drawing.Color.Lime
                            ComboBox_el_END_STA.Items(4) = Continut4
                            ComboBox_EL_LENGTH.Items(4) = Continut4
                        End If
                        If ComboBox_EL_STA.Items.Count = 0 Then
                            ComboBox_EL_desc.Items.add(" ")
                            ComboBox_EL_desc.Items.Add(Continut4)
                            ComboBox_el_id_no.Items.add(" ")
                            ComboBox_el_id_no.Items.Add(Continut4)

                            ComboBox_el_beg_sta.Items.add(" ")
                            ComboBox_el_beg_sta.Items.Add(Continut4)

                            ComboBox_EL_STA.Items.add(" ")
                            ComboBox_EL_STA.Items.Add(Continut4)
                            ComboBox_EL_STA.SelectedIndex = ComboBox_EL_STA.Items.IndexOf(Continut4)
                            Button_EL_STA.BackColor = Drawing.Color.Lime
                            ComboBox_el_END_STA.Items.add(" ")
                            ComboBox_el_END_STA.Items.Add(Continut4)
                            ComboBox_EL_LENGTH.Items.add(" ")
                            ComboBox_EL_LENGTH.Items.Add(Continut4)

                        End If
                        If ComboBox_EL_STA.Items.Count = 1 Or ComboBox_EL_STA.Items.Count = 2 Or ComboBox_EL_STA.Items.Count = 3 Or ComboBox_EL_STA.Items.Count = 4 Then
                            ComboBox_EL_desc.Items.Add(Continut4)
                            ComboBox_el_id_no.Items.Add(Continut4)

                            ComboBox_el_beg_sta.Items.Add(Continut4)


                            ComboBox_EL_STA.Items.Add(Continut4)
                            ComboBox_EL_STA.SelectedIndex = ComboBox_EL_STA.Items.IndexOf(Continut4)
                            Button_EL_STA.BackColor = Drawing.Color.Lime
                            ComboBox_el_END_STA.Items.Add(Continut4)
                            ComboBox_EL_LENGTH.Items.Add(Continut4)
                        End If
                    Else
                        If ComboBox_EL_STA.Items.Count = 0 Then
                            Button_EL_STA.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")

    End Sub
    Private Sub Button_Panel_el_end_sta_Click(sender As Object, e As EventArgs) Handles Button_EL_END_STA.Click

        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat5 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt5 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt5.MessageForAdding = vbLf & "Select info {ENDSTA}:"

                    Object_Prompt5.SingleOnly = True
                    Rezultat5 = Editor1.GetSelection(Object_Prompt5)
                    If Not Rezultat5.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent5 As Entity = Rezultat5.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut5 As String


                    If TypeOf ent5 Is DBText Then
                        Dim Text5 As DBText = ent5
                        Continut5 = Text5.TextString
                    End If

                    If TypeOf ent5 Is MText Then
                        Dim MText5 As MText = ent5
                        Continut5 = MText5.Contents
                    End If


                    If Not Continut5 = "" Then
                        If ComboBox_el_END_STA.Items.Count > 5 Then

                            ComboBox_EL_desc.Items(5) = Continut5
                            ComboBox_el_id_no.Items(5) = Continut5
                            ComboBox_el_beg_sta.Items(5) = Continut5
                            ComboBox_EL_STA.Items(5) = Continut5

                            ComboBox_el_END_STA.Items(5) = Continut5
                            ComboBox_el_END_STA.SelectedIndex = ComboBox_el_END_STA.Items.IndexOf(Continut5)
                            Button_EL_END_STA.BackColor = Drawing.Color.Lime
                            ComboBox_EL_LENGTH.Items(5) = Continut5
                        End If
                        If ComboBox_el_END_STA.Items.Count = 0 Then
                            ComboBox_EL_desc.Items.add(" ")
                            ComboBox_EL_desc.Items.Add(Continut5)
                            ComboBox_el_id_no.Items.add(" ")
                            ComboBox_el_id_no.Items.Add(Continut5)

                            ComboBox_el_beg_sta.Items.add(" ")
                            ComboBox_el_beg_sta.Items.Add(Continut5)

                            ComboBox_EL_STA.Items.add(" ")
                            ComboBox_EL_STA.Items.Add(Continut5)

                            ComboBox_el_END_STA.Items.add(" ")
                            ComboBox_el_END_STA.Items.Add(Continut5)
                            ComboBox_el_END_STA.SelectedIndex = ComboBox_el_END_STA.Items.IndexOf(Continut5)
                            Button_EL_END_STA.BackColor = Drawing.Color.Lime
                            ComboBox_EL_LENGTH.Items.add(" ")
                            ComboBox_EL_LENGTH.Items.Add(Continut5)

                        End If
                        If ComboBox_el_END_STA.Items.Count = 1 Or ComboBox_el_END_STA.Items.Count = 2 Or ComboBox_el_END_STA.Items.Count = 3 Or ComboBox_el_END_STA.Items.Count = 4 Or ComboBox_el_END_STA.Items.Count = 5 Then
                            ComboBox_EL_desc.Items.Add(Continut5)
                            ComboBox_el_id_no.Items.Add(Continut5)
                            ComboBox_el_beg_sta.Items.Add(Continut5)
                            ComboBox_EL_STA.Items.Add(Continut5)
                            ComboBox_el_END_STA.Items.Add(Continut5)
                            ComboBox_el_END_STA.SelectedIndex = ComboBox_el_END_STA.Items.IndexOf(Continut5)
                            Button_EL_END_STA.BackColor = Drawing.Color.Lime
                            ComboBox_EL_LENGTH.Items.Add(Continut5)
                        End If
                    Else
                        If ComboBox_el_END_STA.Items.Count = 0 Then
                            Button_EL_END_STA.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")

    End Sub
    Private Sub Button_Panel_el_length_Click(sender As Object, e As EventArgs) Handles Button_EL_LENGTH.Click

        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat6 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt6 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt6.MessageForAdding = vbLf & "Select info {LENGTH}:"

                    Object_Prompt6.SingleOnly = True
                    Rezultat6 = Editor1.GetSelection(Object_Prompt6)
                    If Not Rezultat6.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent6 As Entity = Rezultat6.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut6 As String


                    If TypeOf ent6 Is DBText Then
                        Dim Text6 As DBText = ent6
                        Continut6 = Text6.TextString
                    End If

                    If TypeOf ent6 Is MText Then
                        Dim MText3 As MText = ent6
                        Continut6 = MText3.Contents
                    End If


                    If Not Continut6 = "" Then
                        If ComboBox_EL_LENGTH.Items.Count > 6 Then

                            ComboBox_EL_desc.Items(6) = Continut6
                            ComboBox_el_id_no.Items(6) = Continut6
                            ComboBox_el_beg_sta.Items(6) = Continut6
                            ComboBox_EL_STA.Items(6) = Continut6
                            ComboBox_el_END_STA.Items(6) = Continut6
                            ComboBox_EL_LENGTH.Items(6) = Continut6
                            ComboBox_EL_LENGTH.SelectedIndex = ComboBox_EL_LENGTH.Items.IndexOf(Continut6)
                            Button_EL_LENGTH.BackColor = Drawing.Color.Lime
                        End If
                        If ComboBox_EL_LENGTH.Items.Count = 0 Then
                            ComboBox_EL_desc.Items.add(" ")
                            ComboBox_EL_desc.Items.Add(Continut6)
                            ComboBox_el_id_no.Items.add(" ")
                            ComboBox_el_id_no.Items.Add(Continut6)

                            ComboBox_el_beg_sta.Items.add(" ")
                            ComboBox_el_beg_sta.Items.Add(Continut6)

                            ComboBox_EL_STA.Items.add(" ")
                            ComboBox_EL_STA.Items.Add(Continut6)
                            ComboBox_el_END_STA.Items.add(" ")
                            ComboBox_el_END_STA.Items.Add(Continut6)
                            ComboBox_EL_LENGTH.Items.add(" ")
                            ComboBox_EL_LENGTH.Items.Add(Continut6)
                            ComboBox_EL_LENGTH.SelectedIndex = ComboBox_EL_LENGTH.Items.IndexOf(Continut6)
                            Button_EL_LENGTH.BackColor = Drawing.Color.Lime

                        End If
                        If ComboBox_EL_LENGTH.Items.Count = 1 Or ComboBox_EL_LENGTH.Items.Count = 2 Or ComboBox_EL_LENGTH.Items.Count = 3 Or ComboBox_EL_LENGTH.Items.Count = 4 Or ComboBox_EL_LENGTH.Items.Count = 5 Or ComboBox_EL_LENGTH.Items.Count = 6 Then
                            ComboBox_EL_desc.Items.Add(Continut6)
                            ComboBox_el_id_no.Items.Add(Continut6)
                            ComboBox_el_beg_sta.Items.Add(Continut6)
                            ComboBox_EL_STA.Items.Add(Continut6)
                            ComboBox_el_END_STA.Items.Add(Continut6)
                            ComboBox_EL_LENGTH.Items.Add(Continut6)
                            ComboBox_EL_LENGTH.SelectedIndex = ComboBox_EL_LENGTH.Items.IndexOf(Continut6)
                            Button_EL_LENGTH.BackColor = Drawing.Color.Lime
                        End If
                    Else
                        If ComboBox_EL_LENGTH.Items.Count = 0 Then
                            Button_EL_LENGTH.BackColor = Drawing.Color.DimGray
                        End If
                    End If
                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using
                afiseaza_butoanele_pentru_forms(Me, Colectie1)

                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")

    End Sub
    Private Sub Button_el_BlockPick_Click(sender As Object, e As EventArgs) Handles Button_el_BlockPick.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info from Block:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    ComboBox_EL_desc.Items.Clear()
                    ComboBox_EL_desc.Items.Add(" ")
                    ComboBox_el_id_no.Items.Clear()
                    ComboBox_el_id_no.Items.Add(" ")
                    ComboBox_el_beg_sta.Items.Clear()
                    ComboBox_el_beg_sta.Items.Add(" ")
                    ComboBox_EL_STA.Items.Clear()
                    ComboBox_EL_STA.Items.Add(" ")
                    ComboBox_el_END_STA.Items.Clear()
                    ComboBox_el_END_STA.Items.Add(" ")
                    ComboBox_EL_LENGTH.Items.Clear()
                    ComboBox_EL_LENGTH.Items.Add(" ")

                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut1 As String
                    Dim Continut2 As String
                    Dim Continut3 As String
                    Dim Continut4 As String
                    Dim Continut5 As String


                    If TypeOf ent1 Is BlockReference Then
                        Dim Block1 As BlockReference = ent1
                        If Block1.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "DESC" Then
                                    Continut1 = attref.TextString
                                    If Continut1 = "" Then Continut1 = " "
                                End If
                                If attref.Tag = "ID_NO" Then
                                    Continut2 = attref.TextString
                                    If Continut2 = "" Then Continut2 = " "
                                End If
                                If attref.Tag = "BEGINSTA" Then
                                    Continut3 = attref.TextString
                                    If Continut3 = "" Then Continut3 = " "
                                End If
                                If attref.Tag = "STA" Then
                                    Continut4 = attref.TextString
                                    If Continut4 = "" Then Continut4 = " "
                                End If
                                If attref.Tag = "ENDSTA" Then
                                    Continut5 = attref.TextString
                                    If Continut5 = "" Then Continut5 = " "
                                End If
                            Next
                        End If
                    End If

                    If Not Continut1 = "" Then
                        ComboBox_EL_desc.Items.Add(Continut1)
                        ComboBox_EL_desc.SelectedIndex = ComboBox_EL_desc.Items.IndexOf(Continut1)
                        Button_EL_DESC.BackColor = Drawing.Color.Lime
                        ComboBox_el_id_no.Items.Add(Continut1)
                        ComboBox_el_beg_sta.Items.Add(Continut1)
                        ComboBox_EL_STA.Items.Add(Continut1)
                        ComboBox_el_END_STA.Items.Add(Continut1)
                    End If
                    If Not Continut2 = "" Then
                        ComboBox_EL_desc.Items.Add(Continut2)
                        ComboBox_el_id_no.Items.Add(Continut2)
                        ComboBox_el_id_no.SelectedIndex = ComboBox_el_id_no.Items.IndexOf(Continut2)
                        Button_EL_ID_NO.BackColor = Drawing.Color.Lime
                        ComboBox_el_beg_sta.Items.Add(Continut2)
                        ComboBox_EL_STA.Items.Add(Continut2)
                        ComboBox_el_END_STA.Items.Add(Continut2)
                    End If
                    If Not Continut3 = "" Then
                        ComboBox_EL_desc.Items.Add(Continut3)
                        ComboBox_el_id_no.Items.Add(Continut3)
                        ComboBox_el_beg_sta.Items.Add(Continut3)
                        ComboBox_el_beg_sta.SelectedIndex = ComboBox_el_beg_sta.Items.IndexOf(Continut3)
                        Button_EL_BEG_STA.BackColor = Drawing.Color.Lime
                        ComboBox_EL_STA.Items.Add(Continut3)
                        ComboBox_el_END_STA.Items.Add(Continut3)
                    End If
                    If Not Continut4 = "" Then
                        ComboBox_EL_desc.Items.Add(Continut4)
                        ComboBox_el_id_no.Items.Add(Continut4)
                        ComboBox_el_beg_sta.Items.Add(Continut4)
                        ComboBox_EL_STA.Items.Add(Continut4)
                        ComboBox_EL_STA.SelectedIndex = ComboBox_EL_STA.Items.IndexOf(Continut4)
                        Button_EL_STA.BackColor = Drawing.Color.Lime
                        ComboBox_el_END_STA.Items.Add(Continut4)
                    End If

                    If Not Continut5 = "" Then
                        ComboBox_EL_desc.Items.Add(Continut5)
                        ComboBox_el_id_no.Items.Add(Continut5)
                        ComboBox_el_beg_sta.Items.Add(Continut5)
                        ComboBox_EL_STA.Items.Add(Continut5)
                        ComboBox_el_END_STA.Items.Add(Continut5)
                        ComboBox_el_END_STA.SelectedIndex = ComboBox_el_END_STA.Items.IndexOf(Continut5)
                        Button_EL_END_STA.BackColor = Drawing.Color.Lime
                    End If

                    If IsNumeric(Replace(Continut3, "+", "")) = True And IsNumeric(Replace(Continut5, "+", "")) = True Then
                        Dim Chain1 As Double = Replace(Continut3, "+", "")
                        Dim Chain2 As Double = Replace(Continut5, "+", "")
                        Dim Length As Double = Abs(Chain1 - Chain2)
                        Dim Length_string As String = Get_String_Rounded(Length, 1)
                        ComboBox_EL_desc.Items.Add(Length_string)
                        ComboBox_el_id_no.Items.Add(Length_string)
                        ComboBox_el_beg_sta.Items.Add(Length_string)
                        ComboBox_EL_STA.Items.Add(Length_string)
                        ComboBox_el_END_STA.Items.Add(Length_string)
                        ComboBox_EL_LENGTH.Items.Add(Length_string)
                        ComboBox_EL_LENGTH.SelectedIndex = ComboBox_EL_LENGTH.Items.IndexOf(Length_string)
                        Button_EL_LENGTH.BackColor = Drawing.Color.Lime
                    End If



                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_el_pick_from_model_space_Click(sender As Object, e As EventArgs) Handles Button_el_pick_from_model_space.Click
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult


            Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt2.MessageForAdding = vbLf & "Select 3D polyline:"

            Object_Prompt2.SingleOnly = True

            Rezultat2 = Editor1.GetSelection(Object_Prompt2)


            If Rezultat2.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                Exit Sub
            End If


            Dim Poly3d As Polyline3d


            Dim Point_on_poly As New Point3d


            If Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat2) = False Then
                    Using lock1 As DocumentLock = ThisDrawing.LockDocument
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


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

                            Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult

                            Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select Point:")
                            PP1.AllowNone = True
                            Point1 = Editor1.GetPoint(PP1)

                            If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                Editor1.WriteMessage(vbLf & "Command:")
                                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                Exit Sub
                            End If

                            Point_on_poly = Poly3d.GetClosestPointTo(Point1.Value.TransformBy(curent_ucs_matrix), Editor1.GetCurrentView.ViewDirection, False)

                            Dim Parameter_picked As Double = Round(Poly3d.GetParameterAtPoint(Point_on_poly), 3)

                            Dim Parameter_start As Double = Floor(Parameter_picked)
                            Dim Parameter_end As Double = Ceiling(Parameter_picked)
                            If Parameter_picked = Round(Parameter_picked, 0) Then
                                Parameter_start = Parameter_picked
                                Parameter_end = Parameter_picked
                            End If

                            Dim Data_table1 As New System.Data.DataTable
                            Data_table1.Columns.Add("TEXT325", GetType(DBText))
                            Dim Index1 As Double = 0

                            Dim Data_table2 As New System.Data.DataTable
                            Data_table2.Columns.Add("TEXT0", GetType(DBText))
                            Dim Index2 As Double = 0

                            For Each ObjID In BTrecord
                                Dim DBobject As DBObject = Trans1.GetObject(ObjID, OpenMode.ForRead)
                                If TypeOf DBobject Is DBText Then
                                    Dim Text1 As DBText = DBobject
                                    If Text1.Layer = ComboBox_layer.Text Then
                                        If Text1.Rotation > 3 * PI / 2 Then
                                            Data_table1.Rows.Add()
                                            Data_table1.Rows(Index1).Item("TEXT325") = Text1
                                            Index1 = Index1 + 1
                                        End If
                                        If Text1.Rotation >= 0 And Text1.Rotation < PI / 4 Then
                                            Data_table2.Rows.Add()
                                            Data_table2.Rows(Index2).Item("TEXT0") = Text1
                                            Index2 = Index2 + 1
                                        End If

                                    End If
                                End If
                            Next
                            Dim Chainage_on_vertex As Double
                            Dim Distanta_pana_la_Vertex As Double
                            Dim CSF1, CSF2 As Double

                            If Data_table1.Rows.Count > 0 Then
                                Dim Point_CHAINAGE As New Point3d
                                Point_CHAINAGE = Poly3d.GetPointAtParameter(Parameter_start)
                                Distanta_pana_la_Vertex = Point_CHAINAGE.GetVectorTo(Point_on_poly).Length

                                For i = 0 To Data_table1.Rows.Count - 1
                                    Dim Text1 As DBText = Data_table1.Rows(i).Item("TEXT325")
                                    If Point_CHAINAGE.GetVectorTo(Text1.Position.TransformBy(Editor1.CurrentUserCoordinateSystem)).Length < 0.1 Then
                                        Dim String1 As String = Replace(Text1.TextString, "+", "")
                                        If IsNumeric(String1) = True Then
                                            Chainage_on_vertex = CDbl(String1)
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                            If Not Parameter_start = Parameter_end Then
                                If Data_table2.Rows.Count > 0 Then
                                    Dim Point_CHAINAGE1 As New Point3d
                                    Point_CHAINAGE1 = Poly3d.GetPointAtParameter(Parameter_start)
                                    Dim Point_CHAINAGE2 As New Point3d
                                    Point_CHAINAGE2 = Poly3d.GetPointAtParameter(Parameter_end)

                                    For i = 0 To Data_table2.Rows.Count - 1
                                        Dim Text1 As DBText = Data_table2.Rows(i).Item("TEXT0")
                                        Dim String1 As String = Text1.TextString
                                        String1 = extrage_numar_din_text_de_la_sfarsitul_textului(String1)
                                        If IsNumeric(String1) = True Then
                                            If Point_CHAINAGE1.GetVectorTo(Text1.Position.TransformBy(Editor1.CurrentUserCoordinateSystem)).Length < 0.1 Then
                                                If CDbl(String1) > 0.5 And CDbl(String1) < 1.5 Then
                                                    CSF1 = CDbl(String1)
                                                End If
                                            End If

                                            If Point_CHAINAGE2.GetVectorTo(Text1.Position.TransformBy(Editor1.CurrentUserCoordinateSystem)).Length < 0.1 Then
                                                If CDbl(String1) > 0.5 And CDbl(String1) < 1.5 Then
                                                    CSF2 = CDbl(String1)
                                                End If
                                            End If
                                        End If
                                    Next
                                End If
                            End If


                            Dim New_ch As Double
                            Dim Old_ch As Double
                            Dim Diferenta As Double
                            Dim New_chainage As String

                            If Not CSF1 + CSF2 = 0 And Not CSF1 = 0 And Not CSF2 = 0 Then
                                New_ch = Chainage_on_vertex + Distanta_pana_la_Vertex / ((CSF1 + CSF2) / 2)
                            Else
                                New_ch = Chainage_on_vertex + Distanta_pana_la_Vertex
                            End If


                            New_chainage = Get_chainage_from_double(New_ch, 1)
                            Dim Old_chainage As String = ComboBox_EL_STA.Text

                            If IsNumeric(Replace(Old_chainage, "+", "")) = True And IsNumeric(Replace(New_chainage, "+", "")) = True Then
                                New_ch = CDbl(Replace(New_chainage, "+", ""))
                                Old_ch = CDbl(Replace(Old_chainage, "+", ""))
                                Diferenta = New_ch - Old_ch
                                Dim Old_chainage1 As String = ComboBox_el_beg_sta.Text
                                Dim Old_chainage2 As String = ComboBox_el_END_STA.Text
                                If IsNumeric(Replace(Old_chainage1, "+", "")) = True And IsNumeric(Replace(Old_chainage2, "+", "")) = True Then
                                    Dim New_ch1 As Double = CDbl(Replace(Old_chainage1, "+", "")) + Diferenta
                                    Dim New_ch2 As Double = CDbl(Replace(Old_chainage2, "+", "")) + Diferenta
                                    Dim New_chainage1 As String = Get_chainage_from_double(New_ch1, 1)
                                    Dim New_chainage2 As String = Get_chainage_from_double(New_ch2, 1)
                                    ComboBox_el_beg_sta.Items.Add(New_chainage1)
                                    ComboBox_el_beg_sta.SelectedIndex = ComboBox_el_beg_sta.Items.IndexOf(New_chainage1)
                                    ComboBox_el_END_STA.Items.Add(New_chainage2)
                                    ComboBox_el_END_STA.SelectedIndex = ComboBox_el_END_STA.Items.IndexOf(New_chainage2)
                                Else
                                    If ComboBox_el_beg_sta.Items.Count > 0 Then ComboBox_el_beg_sta.SelectedIndex = 0
                                    If ComboBox_el_END_STA.Items.Count > 0 Then ComboBox_el_END_STA.SelectedIndex = 0
                                End If

                            End If




                            ComboBox_EL_STA.Items.Add(New_chainage)
                            ComboBox_EL_STA.SelectedIndex = ComboBox_EL_STA.Items.IndexOf(New_chainage)



                            Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly, New_chainage, 0.5, 0.5, 0.5, 11, 3.5)

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

    Private Sub Button_PIPE_CLEAR_Click(sender As Object, e As EventArgs) Handles Button_PIPE_CLEAR.Click
        ComboBox_PIPE_MAT.Items.Clear()
        ComboBox_PIPE_BEGINSTA.Items.Clear()
        ComboBox_PIPE_ENDSTA.Items.Clear()
        ComboBox_PIPE_LENGTH.Items.Clear()
        Button_PIPE_BEGINSTA.BackColor = Drawing.Color.DimGray
        Button_PIPE_ENDSTA.BackColor = Drawing.Color.DimGray
        Button_PIPE_MAT.BackColor = Drawing.Color.DimGray
        Button_PIPE_LENGTH.BackColor = Drawing.Color.DimGray
    End Sub
    Private Sub Button_PIPE_PICK_Click(sender As Object, e As EventArgs) Handles Button_PIPE_PICK.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info {BEGINSTA}:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt2.MessageForAdding = vbLf & "Select info {ENDSTA}:"

                    Object_Prompt2.SingleOnly = True
                    Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                    If Not Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    Dim Rezultat3 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt3 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt3.MessageForAdding = vbLf & "Select info {LENGTH}:"

                    Object_Prompt3.SingleOnly = True
                    Rezultat3 = Editor1.GetSelection(Object_Prompt3)
                    If Not Rezultat3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    Dim Rezultat4 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt4 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt4.MessageForAdding = vbLf & "Select info {MAT}:"

                    Object_Prompt4.SingleOnly = True
                    Rezultat4 = Editor1.GetSelection(Object_Prompt4)
                    If Not Rezultat4.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    ComboBox_PIPE_BEGINSTA.Items.Clear()
                    ComboBox_PIPE_BEGINSTA.Items.Add(" ")
                    ComboBox_PIPE_ENDSTA.Items.Clear()
                    ComboBox_PIPE_ENDSTA.Items.Add(" ")
                    ComboBox_PIPE_LENGTH.Items.Clear()
                    ComboBox_PIPE_LENGTH.Items.Add(" ")
                    ComboBox_PIPE_MAT.Items.Clear()
                    ComboBox_PIPE_MAT.Items.Add(" ")


                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent2 As Entity = Rezultat2.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent3 As Entity = Rezultat3.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent4 As Entity = Rezultat4.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)

                    Dim Continut1 As String
                    Dim Continut2 As String
                    Dim Continut3 As String
                    Dim Continut4 As String

                    If TypeOf ent1 Is DBText Then
                        Dim Text1 As DBText = ent1
                        Continut1 = Text1.TextString
                    End If
                    If TypeOf ent2 Is DBText Then
                        Dim Text2 As DBText = ent2
                        Continut2 = Text2.TextString
                    End If
                    If TypeOf ent3 Is DBText Then
                        Dim Text3 As DBText = ent3
                        Continut3 = Text3.TextString
                    End If
                    If TypeOf ent4 Is DBText Then
                        Dim Text4 As DBText = ent4
                        Continut4 = Text4.TextString
                    End If

                    If TypeOf ent1 Is MText Then
                        Dim MText1 As MText = ent1
                        Continut1 = MText1.Contents
                    End If
                    If TypeOf ent2 Is MText Then
                        Dim MText2 As MText = ent2
                        Continut2 = MText2.Contents
                    End If
                    If TypeOf ent3 Is MText Then
                        Dim MText3 As MText = ent3
                        Continut3 = MText3.Contents
                    End If
                    If TypeOf ent4 Is MText Then
                        Dim MText4 As MText = ent4
                        Continut4 = MText4.Contents
                    End If


                    If TypeOf ent1 Is BlockReference Then
                        Dim Block1 As BlockReference = ent1
                        If Block1.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                            Dim ctemp1 As String
                            Dim ctemp2 As String
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "BEGINSTA" Then
                                    ctemp1 = attref.TextString
                                End If
                                If attref.Tag = "ENDSTA" Then
                                    ctemp2 = attref.TextString
                                End If
                            Next
                            ctemp1 = Replace(ctemp1, " ", "")
                            ctemp2 = Replace(ctemp2, " ", "")



                            If Continut2 = "" Then
                                If TypeOf ent2 Is BlockReference And Not ent1.ObjectId = ent2.ObjectId Then
                                    If Not ctemp1 = "" And Not ctemp2 = "" Then
                                        Continut1 = ctemp2
                                    End If
                                    If ctemp1 = "" And ctemp2 = "" Then
                                        Continut1 = " "
                                    End If
                                    If ctemp1 = "" And Not ctemp2 = "" Then
                                        Continut1 = ctemp2
                                    End If
                                    If Not ctemp1 = "" And ctemp2 = "" Then
                                        Continut1 = ctemp1
                                    End If
                                End If
                            Else
                                If Not ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut1 = ctemp2
                                End If
                                If ctemp1 = "" And ctemp2 = "" Then
                                    Continut1 = " "
                                End If
                                If ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut1 = ctemp2
                                End If
                                If Not ctemp1 = "" And ctemp2 = "" Then
                                    Continut1 = ctemp1
                                End If
                            End If
                            If ent1.ObjectId = ent2.ObjectId Then
                                Continut1 = ctemp1
                            End If

                        End If
                    End If


                    If TypeOf ent2 Is BlockReference Then
                        Dim Block2 As BlockReference = ent2
                        If Block2.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block2.AttributeCollection
                            Dim ctemp1 As String
                            Dim ctemp2 As String
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "BEGINSTA" Then
                                    ctemp1 = attref.TextString
                                End If
                                If attref.Tag = "ENDSTA" Then
                                    ctemp2 = attref.TextString
                                End If
                            Next
                            ctemp1 = Replace(ctemp1, " ", "")
                            ctemp2 = Replace(ctemp2, " ", "")
                            If TypeOf ent1 Is BlockReference Then
                                If Not ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut2 = ctemp1
                                End If
                                If ctemp1 = "" And ctemp2 = "" Then
                                    Continut2 = " "
                                End If
                                If ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut2 = ctemp2
                                End If
                                If Not ctemp1 = "" And ctemp2 = "" Then
                                    Continut2 = ctemp1
                                End If
                            Else
                                Continut2 = ctemp1
                            End If
                            If ent1.ObjectId = ent2.ObjectId Then
                                Continut2 = ctemp2
                            End If


                        End If
                    End If


                    If TypeOf ent3 Is BlockReference Then
                        Dim Block3 As BlockReference = ent3
                        If Block3.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block3.AttributeCollection
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "LENGTH" Then
                                    Continut3 = attref.TextString
                                End If
                            Next
                        End If
                    End If


                    If TypeOf ent4 Is BlockReference Then
                        Dim Block4 As BlockReference = ent4
                        If Block4.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block4.AttributeCollection
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "MAT" Then
                                    Continut4 = attref.TextString
                                End If
                            Next
                        End If
                    End If



                    If Not Continut1 = "" Then
                        ComboBox_PIPE_BEGINSTA.Items.Add(Continut1)
                        ComboBox_PIPE_BEGINSTA.SelectedIndex = ComboBox_PIPE_BEGINSTA.Items.IndexOf(Continut1)
                        Button_PIPE_BEGINSTA.BackColor = Drawing.Color.Lime
                        ComboBox_PIPE_ENDSTA.Items.Add(Continut1)
                        ComboBox_PIPE_LENGTH.Items.Add(Continut1)
                        ComboBox_PIPE_MAT.Items.Add(Continut1)

                    End If
                    If Not Continut2 = "" Then
                        ComboBox_PIPE_BEGINSTA.Items.Add(Continut2)
                        ComboBox_PIPE_ENDSTA.Items.Add(Continut2)
                        ComboBox_PIPE_ENDSTA.SelectedIndex = ComboBox_PIPE_ENDSTA.Items.IndexOf(Continut2)
                        Button_PIPE_ENDSTA.BackColor = Drawing.Color.Lime
                        ComboBox_PIPE_LENGTH.Items.Add(Continut2)
                        ComboBox_PIPE_MAT.Items.Add(Continut2)
                    End If
                    If Not Continut3 = "" Then
                        ComboBox_PIPE_BEGINSTA.Items.Add(Continut3)
                        ComboBox_PIPE_ENDSTA.Items.Add(Continut3)
                        ComboBox_PIPE_LENGTH.Items.Add(Continut3)
                        ComboBox_PIPE_LENGTH.SelectedIndex = ComboBox_PIPE_LENGTH.Items.IndexOf(Continut3)
                        Button_PIPE_LENGTH.BackColor = Drawing.Color.Lime
                        ComboBox_PIPE_MAT.Items.Add(Continut3)
                    End If
                    If Not Continut4 = "" Then
                        ComboBox_PIPE_BEGINSTA.Items.Add(Continut4)
                        ComboBox_PIPE_ENDSTA.Items.Add(Continut4)
                        ComboBox_PIPE_LENGTH.Items.Add(Continut4)
                        ComboBox_PIPE_MAT.Items.Add(Continut4)
                        ComboBox_PIPE_MAT.SelectedIndex = ComboBox_PIPE_MAT.Items.IndexOf(Continut4)
                        Button_PIPE_MAT.BackColor = Drawing.Color.Lime
                    End If






                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_PIPE_INSERT_Click(sender As Object, e As EventArgs) Handles Button_PIPE_INSERT.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select objects to be deleted:"

                    Object_Prompt1.SingleOnly = False
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If
                    Dim Point_start As Autodesk.AutoCAD.EditorInput.PromptPointResult

                    Dim PP_start As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select insertion point:")
                    PP_start.AllowNone = True
                    Point_start = Editor1.GetPoint(PP_start)
                    If Not Point_start.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        Editor1.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If

                    Dim Colectie_atr_name As New Specialized.StringCollection
                    Dim Colectie_atr_value As New Specialized.StringCollection

                    If Not ComboBox_PIPE_BEGINSTA.Text = "" Then
                        Colectie_atr_name.Add("BEGINSTA")
                        Colectie_atr_value.Add(ComboBox_PIPE_BEGINSTA.Text)
                    End If

                    If Not ComboBox_PIPE_ENDSTA.Text = "" Then
                        Colectie_atr_name.Add("ENDSTA")
                        Colectie_atr_value.Add(ComboBox_PIPE_ENDSTA.Text)
                    End If

                    If Not ComboBox_PIPE_LENGTH.Text = "" Then
                        Colectie_atr_name.Add("LENGTH")
                        Colectie_atr_value.Add(ComboBox_PIPE_LENGTH.Text)
                    End If

                    If Not ComboBox_PIPE_MAT.Text = "" Then
                        Colectie_atr_name.Add("MAT")
                        Colectie_atr_value.Add(ComboBox_PIPE_MAT.Text)
                    End If


                    InsertBlock_with_multiple_atributes("heavy_wall_1.dwg", "heavy_wall_1", Point_start.Value, 1, BTrecord, "TEXT", Colectie_atr_name, Colectie_atr_value)

                    For i = 0 To Rezultat1.Value.Count - 1






                        Dim ent1 As Entity = Rezultat1.Value.Item(i).ObjectId.GetObject(OpenMode.ForWrite)
                        ent1.Erase()
                    Next








                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)

                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_PIPE_BEGINSTA_Click(sender As Object, e As EventArgs) Handles Button_PIPE_BEGINSTA.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info {BEGINSTA}:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut1 As String


                    If TypeOf ent1 Is DBText Then
                        Dim Text1 As DBText = ent1
                        Continut1 = Text1.TextString
                    End If

                    If TypeOf ent1 Is MText Then
                        Dim MText1 As MText = ent1
                        Continut1 = MText1.Contents
                    End If


                    If Not Continut1 = "" Then
                        If ComboBox_PIPE_BEGINSTA.Items.Count > 1 Then
                            ComboBox_PIPE_BEGINSTA.Items(1) = Continut1
                            ComboBox_PIPE_BEGINSTA.SelectedIndex = ComboBox_PIPE_BEGINSTA.Items.IndexOf(Continut1)
                            Button_PIPE_BEGINSTA.BackColor = Drawing.Color.Lime
                            ComboBox_PIPE_ENDSTA.Items(1) = Continut1
                            ComboBox_PIPE_LENGTH.Items(1) = Continut1
                            ComboBox_PIPE_MAT.Items(1) = Continut1
                        End If
                        If ComboBox_PIPE_BEGINSTA.Items.Count = 0 Then
                            ComboBox_PIPE_BEGINSTA.Items.Add(" ")
                            ComboBox_PIPE_BEGINSTA.Items.Add(Continut1)
                            ComboBox_PIPE_BEGINSTA.SelectedIndex = ComboBox_PIPE_BEGINSTA.Items.IndexOf(Continut1)
                            Button_PIPE_BEGINSTA.BackColor = Drawing.Color.Lime
                            ComboBox_PIPE_ENDSTA.Items.Add(" ")
                            ComboBox_PIPE_ENDSTA.Items.Add(Continut1)
                            ComboBox_PIPE_LENGTH.Items.Add(" ")
                            ComboBox_PIPE_LENGTH.Items.Add(Continut1)
                            ComboBox_PIPE_MAT.Items.Add(" ")
                            ComboBox_PIPE_MAT.Items.Add(Continut1)
                        End If

                    Else
                        If ComboBox_PIPE_BEGINSTA.Items.Count = 0 Then
                            Button_PIPE_BEGINSTA.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_PIPE_ENDSTA_Click(sender As Object, e As EventArgs) Handles Button_PIPE_ENDSTA.Click

        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt2.MessageForAdding = vbLf & "Select info {ENDSTA}:"

                    Object_Prompt2.SingleOnly = True
                    Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                    If Not Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent2 As Entity = Rezultat2.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut2 As String


                    If TypeOf ent2 Is DBText Then
                        Dim Text2 As DBText = ent2
                        Continut2 = Text2.TextString
                    End If

                    If TypeOf ent2 Is MText Then
                        Dim MText2 As MText = ent2
                        Continut2 = MText2.Contents
                    End If


                    If Not Continut2 = "" Then
                        If ComboBox_PIPE_ENDSTA.Items.Count > 2 Then
                            ComboBox_PIPE_BEGINSTA.Items(2) = Continut2
                            ComboBox_PIPE_ENDSTA.Items(2) = Continut2
                            ComboBox_PIPE_ENDSTA.SelectedIndex = ComboBox_PIPE_ENDSTA.Items.IndexOf(Continut2)
                            Button_PIPE_ENDSTA.BackColor = Drawing.Color.Lime
                            ComboBox_PIPE_LENGTH.Items(2) = Continut2
                            ComboBox_PIPE_MAT.Items(2) = Continut2
                        End If
                        If ComboBox_PIPE_ENDSTA.Items.Count = 0 Then
                            ComboBox_PIPE_BEGINSTA.Items.Add(" ")
                            ComboBox_PIPE_BEGINSTA.Items.Add(Continut2)
                            ComboBox_PIPE_ENDSTA.Items.Add(" ")
                            ComboBox_PIPE_ENDSTA.Items.Add(Continut2)
                            ComboBox_PIPE_ENDSTA.SelectedIndex = ComboBox_PIPE_ENDSTA.Items.IndexOf(Continut2)
                            Button_PIPE_ENDSTA.BackColor = Drawing.Color.Lime
                            ComboBox_PIPE_LENGTH.Items.Add(" ")
                            ComboBox_PIPE_LENGTH.Items.Add(Continut2)
                            ComboBox_PIPE_MAT.Items.Add(" ")
                            ComboBox_PIPE_MAT.Items.Add(Continut2)
                        End If
                        If ComboBox_PIPE_ENDSTA.Items.Count = 1 Or ComboBox_PIPE_ENDSTA.Items.Count = 2 Then
                            ComboBox_PIPE_BEGINSTA.Items.Add(Continut2)
                            ComboBox_PIPE_ENDSTA.Items.Add(Continut2)
                            ComboBox_PIPE_ENDSTA.SelectedIndex = ComboBox_PIPE_ENDSTA.Items.IndexOf(Continut2)
                            Button_PIPE_ENDSTA.BackColor = Drawing.Color.Lime
                            ComboBox_PIPE_LENGTH.Items.Add(Continut2)
                            ComboBox_PIPE_MAT.Items.Add(Continut2)
                        End If
                    Else
                        If ComboBox_PIPE_ENDSTA.Items.Count = 0 Then
                            Button_PIPE_ENDSTA.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_PIPE_LENGTH_Click(sender As Object, e As EventArgs) Handles Button_PIPE_LENGTH.Click

        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat5 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt5 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt5.MessageForAdding = vbLf & "Select info {LENGTH}:"

                    Object_Prompt5.SingleOnly = True
                    Rezultat5 = Editor1.GetSelection(Object_Prompt5)
                    If Not Rezultat5.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent5 As Entity = Rezultat5.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut5 As String


                    If TypeOf ent5 Is DBText Then
                        Dim Text5 As DBText = ent5
                        Continut5 = Text5.TextString
                    End If

                    If TypeOf ent5 Is MText Then
                        Dim MText5 As MText = ent5
                        Continut5 = MText5.Contents
                    End If


                    If Not Continut5 = "" Then
                        If ComboBox_PIPE_LENGTH.Items.Count > 5 Then
                            ComboBox_PIPE_BEGINSTA.Items(5) = Continut5
                            ComboBox_PIPE_ENDSTA.Items(5) = Continut5
                            ComboBox_PIPE_LENGTH.Items(5) = Continut5
                            ComboBox_PIPE_LENGTH.SelectedIndex = ComboBox_PIPE_LENGTH.Items.IndexOf(Continut5)
                            Button_PIPE_LENGTH.BackColor = Drawing.Color.Lime
                            ComboBox_PIPE_MAT.Items(5) = Continut5
                        End If
                        If ComboBox_PIPE_LENGTH.Items.Count = 0 Then
                            ComboBox_PIPE_BEGINSTA.Items.Add(" ")
                            ComboBox_PIPE_BEGINSTA.Items.Add(Continut5)
                            ComboBox_PIPE_ENDSTA.Items.Add(" ")
                            ComboBox_PIPE_ENDSTA.Items.Add(Continut5)

                            ComboBox_PIPE_LENGTH.Items.Add(" ")
                            ComboBox_PIPE_LENGTH.Items.Add(Continut5)
                            ComboBox_PIPE_LENGTH.SelectedIndex = ComboBox_PIPE_LENGTH.Items.IndexOf(Continut5)
                            Button_PIPE_LENGTH.BackColor = Drawing.Color.Lime
                            ComboBox_PIPE_MAT.Items.Add(" ")
                            ComboBox_PIPE_MAT.Items.Add(Continut5)
                        End If
                        If ComboBox_PIPE_LENGTH.Items.Count = 1 Or ComboBox_PIPE_LENGTH.Items.Count = 2 Or ComboBox_PIPE_LENGTH.Items.Count = 3 Then
                            ComboBox_PIPE_BEGINSTA.Items.Add(Continut5)
                            ComboBox_PIPE_ENDSTA.Items.Add(Continut5)

                            ComboBox_PIPE_LENGTH.Items.Add(Continut5)
                            ComboBox_PIPE_LENGTH.SelectedIndex = ComboBox_PIPE_LENGTH.Items.IndexOf(Continut5)
                            Button_PIPE_LENGTH.BackColor = Drawing.Color.Lime

                            ComboBox_PIPE_MAT.Items.Add(Continut5)
                        End If
                    Else
                        If ComboBox_PIPE_LENGTH.Items.Count = 0 Then
                            Button_PIPE_LENGTH.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")

    End Sub
    Private Sub Button_PIPE_MAT_Click(sender As Object, e As EventArgs) Handles Button_PIPE_MAT.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat6 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt6 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt6.MessageForAdding = vbLf & "Select info {PIPE_MAT}:"

                    Object_Prompt6.SingleOnly = True
                    Rezultat6 = Editor1.GetSelection(Object_Prompt6)
                    If Not Rezultat6.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If
                    Dim ent6 As Entity = Rezultat6.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim Continut6 As String
                    If TypeOf ent6 Is DBText Then
                        Dim Text6 As DBText = ent6
                        Continut6 = Text6.TextString
                    End If

                    If TypeOf ent6 Is MText Then
                        Dim MText3 As MText = ent6
                        Continut6 = MText3.Contents
                    End If


                    If Not Continut6 = "" Then
                        If ComboBox_PIPE_MAT.Items.Count > 6 Then
                            ComboBox_PIPE_BEGINSTA.Items(6) = Continut6
                            ComboBox_PIPE_ENDSTA.Items(6) = Continut6
                            ComboBox_PIPE_LENGTH.Items(6) = Continut6
                            ComboBox_PIPE_MAT.Items(6) = Continut6
                            ComboBox_PIPE_MAT.SelectedIndex = ComboBox_PIPE_MAT.Items.IndexOf(Continut6)
                            Button_PIPE_MAT.BackColor = Drawing.Color.Lime
                        End If
                        If ComboBox_PIPE_MAT.Items.Count = 0 Then
                            ComboBox_PIPE_BEGINSTA.Items.Add(" ")
                            ComboBox_PIPE_BEGINSTA.Items.Add(Continut6)
                            ComboBox_PIPE_ENDSTA.Items.Add(" ")
                            ComboBox_PIPE_ENDSTA.Items.Add(Continut6)

                            ComboBox_PIPE_LENGTH.Items.Add(" ")
                            ComboBox_PIPE_LENGTH.Items.Add(Continut6)

                            ComboBox_PIPE_MAT.Items.Add(" ")
                            ComboBox_PIPE_MAT.Items.Add(Continut6)
                            ComboBox_PIPE_MAT.SelectedIndex = ComboBox_PIPE_MAT.Items.IndexOf(Continut6)
                            Button_PIPE_MAT.BackColor = Drawing.Color.Lime
                        End If
                        If ComboBox_PIPE_MAT.Items.Count = 1 Or ComboBox_PIPE_MAT.Items.Count = 2 Or ComboBox_PIPE_MAT.Items.Count = 3 Or ComboBox_PIPE_MAT.Items.Count = 4 Then
                            ComboBox_PIPE_BEGINSTA.Items.Add(Continut6)
                            ComboBox_PIPE_ENDSTA.Items.Add(Continut6)

                            ComboBox_PIPE_LENGTH.Items.Add(Continut6)


                            ComboBox_PIPE_MAT.Items.Add(Continut6)
                            ComboBox_PIPE_MAT.SelectedIndex = ComboBox_PIPE_MAT.Items.IndexOf(Continut6)
                            Button_PIPE_MAT.BackColor = Drawing.Color.Lime
                        End If
                    Else
                        If ComboBox_PIPE_MAT.Items.Count = 0 Then
                            Button_PIPE_MAT.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")

    End Sub
    Private Sub Button_pipe_blockpick_Click(sender As Object, e As EventArgs) Handles Button_pipe_blockpick.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info from Block:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    ComboBox_PIPE_BEGINSTA.Items.Clear()
                    ComboBox_PIPE_BEGINSTA.Items.Add(" ")
                    ComboBox_PIPE_ENDSTA.Items.Clear()
                    ComboBox_PIPE_ENDSTA.Items.Add(" ")
                    ComboBox_PIPE_LENGTH.Items.Clear()
                    ComboBox_PIPE_LENGTH.Items.Add(" ")
                    ComboBox_PIPE_MAT.Items.Clear()
                    ComboBox_PIPE_MAT.Items.Add(" ")

                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut1 As String
                    Dim Continut2 As String
                    Dim Continut4 As String



                    If TypeOf ent1 Is BlockReference Then
                        Dim Block1 As BlockReference = ent1
                        If Block1.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "BEGINSTA" Then
                                    Continut1 = attref.TextString
                                    If Continut1 = "" Then Continut1 = " "
                                End If
                                If attref.Tag = "ENDSTA" Then
                                    Continut2 = attref.TextString
                                    If Continut2 = "" Then Continut2 = " "
                                End If

                                If attref.Tag = "MAT" Then
                                    Continut4 = attref.TextString
                                    If Continut4 = "" Then Continut4 = " "
                                End If

                            Next
                        End If
                    End If

                    If Not Continut1 = "" Then
                        ComboBox_PIPE_BEGINSTA.Items.Add(Continut1)
                        ComboBox_PIPE_BEGINSTA.SelectedIndex = ComboBox_PIPE_BEGINSTA.Items.IndexOf(Continut1)
                        Button_PIPE_BEGINSTA.BackColor = Drawing.Color.Lime
                        ComboBox_PIPE_ENDSTA.Items.Add(Continut1)
                        ComboBox_PIPE_MAT.Items.Add(Continut1)
                    End If
                    If Not Continut2 = "" Then
                        ComboBox_PIPE_BEGINSTA.Items.Add(Continut2)
                        ComboBox_PIPE_ENDSTA.Items.Add(Continut2)
                        ComboBox_PIPE_ENDSTA.SelectedIndex = ComboBox_PIPE_ENDSTA.Items.IndexOf(Continut2)
                        Button_PIPE_ENDSTA.BackColor = Drawing.Color.Lime
                        ComboBox_PIPE_MAT.Items.Add(Continut2)
                    End If
                    If Not Continut4 = "" Then
                        ComboBox_PIPE_BEGINSTA.Items.Add(Continut4)
                        ComboBox_PIPE_ENDSTA.Items.Add(Continut4)
                        ComboBox_PIPE_MAT.Items.Add(Continut4)
                        ComboBox_PIPE_MAT.SelectedIndex = ComboBox_PIPE_MAT.Items.IndexOf(Continut4)
                        Button_PIPE_MAT.BackColor = Drawing.Color.Lime

                    End If

                    If IsNumeric(Replace(Continut1, "+", "")) = True And IsNumeric(Replace(Continut2, "+", "")) = True Then
                        Dim Chain1 As Double = Replace(Continut1, "+", "")
                        Dim Chain2 As Double = Replace(Continut2, "+", "")
                        Dim Length As Double = Abs(Chain1 - Chain2)
                        Dim Length_string As String = Get_String_Rounded(Length, 1)
                        ComboBox_PIPE_BEGINSTA.Items.Add(Length_string)
                        ComboBox_PIPE_ENDSTA.Items.Add(Length_string)
                        ComboBox_PIPE_LENGTH.Items.Add(Length_string)
                        ComboBox_PIPE_LENGTH.SelectedIndex = ComboBox_PIPE_LENGTH.Items.IndexOf(Length_string)
                        Button_PIPE_LENGTH.BackColor = Drawing.Color.Lime
                        ComboBox_PIPE_MAT.Items.Add(Length_string)
                    End If



                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub

    Private Sub Button_cross_clear_Click(sender As Object, e As EventArgs) Handles Button_cross_clear.Click

        ComboBox_CROSS_DESC.Items.Clear()
        ComboBox_CROSS_STA.Items.Clear()

        Button_CROSS_DESC.BackColor = Drawing.Color.DimGray
        Button_CROSS_STA.BackColor = Drawing.Color.DimGray

    End Sub
    Private Sub Button_cross_pick_Click(sender As Object, e As EventArgs) Handles Button_cross_pick.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info {DESC}:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt2.MessageForAdding = vbLf & "Select info {STA}:"

                    Object_Prompt2.SingleOnly = True
                    Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                    If Not Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    ComboBox_CROSS_DESC.Items.Clear()
                    ComboBox_CROSS_DESC.Items.add(" ")
                    ComboBox_CROSS_STA.Items.Clear()
                    ComboBox_CROSS_STA.Items.add(" ")


                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent2 As Entity = Rezultat2.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut1 As String
                    Dim Continut2 As String


                    If TypeOf ent1 Is DBText Then
                        Dim Text1 As DBText = ent1
                        Continut1 = Text1.TextString
                    End If
                    If TypeOf ent2 Is DBText Then
                        Dim Text2 As DBText = ent2
                        Continut2 = Text2.TextString
                    End If

                    If TypeOf ent1 Is MText Then
                        Dim MText1 As MText = ent1
                        Continut1 = MText1.Contents
                    End If
                    If TypeOf ent2 Is MText Then
                        Dim MText2 As MText = ent2
                        Continut2 = MText2.Contents
                    End If


                    If Not Continut1 = "" Then
                        ComboBox_CROSS_DESC.Items.Add(Continut1)
                        ComboBox_CROSS_DESC.SelectedIndex = ComboBox_CROSS_DESC.Items.IndexOf(Continut1)
                        Button_CROSS_DESC.BackColor = Drawing.Color.Lime
                        ComboBox_CROSS_STA.Items.Add(Continut1)

                    End If
                    If Not Continut2 = "" Then
                        ComboBox_CROSS_DESC.Items.Add(Continut2)
                        ComboBox_CROSS_STA.Items.Add(Continut2)
                        ComboBox_CROSS_STA.SelectedIndex = ComboBox_CROSS_STA.Items.IndexOf(Continut2)
                        Button_CROSS_STA.BackColor = Drawing.Color.Lime

                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_cross_insert_Click(sender As Object, e As EventArgs) Handles Button_cross_insert.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select objects to be deleted:"

                    Object_Prompt1.SingleOnly = False
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If
                    Dim Point_start As Autodesk.AutoCAD.EditorInput.PromptPointResult

                    Dim PP_start As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select insertion point:")
                    PP_start.AllowNone = True
                    Point_start = Editor1.GetPoint(PP_start)
                    If Not Point_start.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        Editor1.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If


                    Dim Colectie_atr_name As New Specialized.StringCollection
                    Dim Colectie_atr_value As New Specialized.StringCollection

                    If Not ComboBox_CROSS_DESC.Text = "" Then
                        Colectie_atr_name.Add("DESC")
                        Colectie_atr_value.Add(ComboBox_CROSS_DESC.Text)
                    End If

                    If Not ComboBox_CROSS_STA.Text = "" Then
                        Colectie_atr_name.Add("STA")
                        Colectie_atr_value.Add(ComboBox_CROSS_STA.Text)
                    End If



                    InsertBlock_with_multiple_atributes("Crossing_Profile_al.dwg", "AL_CROSSING_PROFILE", Point_start.Value, 1, BTrecord, "TEXT", Colectie_atr_name, Colectie_atr_value)

                    For i = 0 To Rezultat1.Value.Count - 1






                        Dim ent1 As Entity = Rezultat1.Value.Item(i).ObjectId.GetObject(OpenMode.ForWrite)
                        ent1.Erase()
                    Next








                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)

                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_CROSS_DESC_Click(sender As Object, e As EventArgs) Handles Button_CROSS_DESC.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info {DESC}:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut1 As String


                    If TypeOf ent1 Is DBText Then
                        Dim Text1 As DBText = ent1
                        Continut1 = Text1.TextString
                    End If

                    If TypeOf ent1 Is MText Then
                        Dim MText1 As MText = ent1
                        Continut1 = MText1.Contents
                    End If


                    If Not Continut1 = "" Then
                        If ComboBox_CROSS_DESC.Items.Count > 1 Then
                            ComboBox_CROSS_DESC.Items(1) = Continut1
                            ComboBox_CROSS_DESC.SelectedIndex = ComboBox_CROSS_DESC.Items.IndexOf(Continut1)
                            Button_CROSS_DESC.BackColor = Drawing.Color.Lime
                            ComboBox_CROSS_STA.Items(1) = Continut1

                        End If
                        If ComboBox_CROSS_DESC.Items.Count = 0 Then
                            ComboBox_CROSS_DESC.Items.add(" ")
                            ComboBox_CROSS_DESC.Items.Add(Continut1)
                            ComboBox_CROSS_DESC.SelectedIndex = ComboBox_CROSS_DESC.Items.IndexOf(Continut1)
                            Button_CROSS_DESC.BackColor = Drawing.Color.Lime
                            ComboBox_CROSS_STA.Items.add(" ")
                            ComboBox_CROSS_STA.Items.Add(Continut1)
                        End If

                    Else
                        If ComboBox_CROSS_DESC.Items.Count = 0 Then
                            Button_CROSS_DESC.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_CROSS_STA_Click(sender As Object, e As EventArgs) Handles Button_CROSS_STA.Click

        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt2.MessageForAdding = vbLf & "Select info {STA}:"

                    Object_Prompt2.SingleOnly = True
                    Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                    If Not Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent2 As Entity = Rezultat2.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut2 As String


                    If TypeOf ent2 Is DBText Then
                        Dim Text2 As DBText = ent2
                        Continut2 = Text2.TextString
                    End If

                    If TypeOf ent2 Is MText Then
                        Dim MText2 As MText = ent2
                        Continut2 = MText2.Contents
                    End If


                    If Not Continut2 = "" Then
                        If ComboBox_CROSS_STA.Items.Count > 2 Then
                            ComboBox_CROSS_DESC.Items(2) = Continut2
                            ComboBox_CROSS_STA.Items(2) = Continut2
                            ComboBox_CROSS_STA.SelectedIndex = ComboBox_CROSS_STA.Items.IndexOf(Continut2)
                            Button_CROSS_STA.BackColor = Drawing.Color.Lime
                        End If
                        If ComboBox_CROSS_STA.Items.Count = 0 Then
                            ComboBox_CROSS_DESC.Items.add(" ")
                            ComboBox_CROSS_DESC.Items.Add(Continut2)
                            ComboBox_CROSS_STA.Items.add(" ")
                            ComboBox_CROSS_STA.Items.Add(Continut2)
                            ComboBox_CROSS_STA.SelectedIndex = ComboBox_CROSS_STA.Items.IndexOf(Continut2)
                            Button_CROSS_STA.BackColor = Drawing.Color.Lime
                        End If
                        If ComboBox_CROSS_STA.Items.Count = 1 Or ComboBox_CROSS_STA.Items.Count = 2 Then
                            ComboBox_CROSS_DESC.Items.Add(Continut2)
                            ComboBox_CROSS_STA.Items.Add(Continut2)
                            ComboBox_CROSS_STA.SelectedIndex = ComboBox_CROSS_STA.Items.IndexOf(Continut2)
                            Button_CROSS_STA.BackColor = Drawing.Color.Lime
                        End If
                    Else
                        If ComboBox_CROSS_STA.Items.Count = 0 Then
                            Button_CROSS_STA.BackColor = Drawing.Color.DimGray
                        End If
                    End If

                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using

                afiseaza_butoanele_pentru_forms(Me, Colectie1)

                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_cross_blockpick_Click(sender As Object, e As EventArgs) Handles Button_cross_blockpick.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info from Block:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    ComboBox_CROSS_DESC.Items.Clear()
                    ComboBox_CROSS_DESC.Items.Add(" ")
                    ComboBox_CROSS_STA.Items.Clear()
                    ComboBox_CROSS_STA.Items.Add(" ")


                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut1 As String
                    Dim Continut2 As String



                    If TypeOf ent1 Is BlockReference Then
                        Dim Block1 As BlockReference = ent1
                        If Block1.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "DESC" Then
                                    Continut1 = attref.TextString
                                    If Continut1 = "" Then Continut1 = " "
                                End If
                                If attref.Tag = "STA" Then
                                    Continut2 = attref.TextString
                                    If Continut2 = "" Then Continut2 = " "
                                End If
                            Next
                        End If
                    End If

                    If Not Continut1 = "" Then
                        ComboBox_CROSS_DESC.Items.Add(Continut1)
                        ComboBox_CROSS_DESC.SelectedIndex = ComboBox_CROSS_DESC.Items.IndexOf(Continut1)
                        Button_CROSS_DESC.BackColor = Drawing.Color.Lime
                        ComboBox_CROSS_STA.Items.Add(Continut1)
                    End If
                    If Not Continut2 = "" Then
                        ComboBox_CROSS_DESC.Items.Add(Continut2)
                        ComboBox_CROSS_STA.Items.Add(Continut2)
                        ComboBox_CROSS_STA.SelectedIndex = ComboBox_CROSS_STA.Items.IndexOf(Continut2)
                        Button_CROSS_STA.BackColor = Drawing.Color.Lime
                    End If

                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_cross_pick_from_model_space_Click(sender As Object, e As EventArgs) Handles Button_cross_pick_from_model_space.Click
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult


            Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt2.MessageForAdding = vbLf & "Select 3D polyline:"

            Object_Prompt2.SingleOnly = True

            Rezultat2 = Editor1.GetSelection(Object_Prompt2)


            If Rezultat2.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                Exit Sub
            End If


            Dim Poly3d As Polyline3d


            Dim Point_on_poly As New Point3d


            If Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat2) = False Then
                    Using lock1 As DocumentLock = ThisDrawing.LockDocument
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


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

                            Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult

                            Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select Point:")
                            PP1.AllowNone = True
                            Point1 = Editor1.GetPoint(PP1)

                            If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                Editor1.WriteMessage(vbLf & "Command:")
                                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                Exit Sub
                            End If

                            Point_on_poly = Poly3d.GetClosestPointTo(Point1.Value.TransformBy(curent_ucs_matrix), Editor1.GetCurrentView.ViewDirection, False)

                            Dim Parameter_picked As Double = Round(Poly3d.GetParameterAtPoint(Point_on_poly), 3)

                            Dim Parameter_start As Double = Floor(Parameter_picked)
                            Dim Parameter_end As Double = Ceiling(Parameter_picked)
                            If Parameter_picked = Round(Parameter_picked, 0) Then
                                Parameter_start = Parameter_picked
                                Parameter_end = Parameter_picked
                            End If

                            Dim Data_table1 As New System.Data.DataTable
                            Data_table1.Columns.Add("TEXT325", GetType(DBText))
                            Dim Index1 As Double = 0

                            Dim Data_table2 As New System.Data.DataTable
                            Data_table2.Columns.Add("TEXT0", GetType(DBText))
                            Dim Index2 As Double = 0

                            For Each ObjID In BTrecord
                                Dim DBobject As DBObject = Trans1.GetObject(ObjID, OpenMode.ForRead)
                                If TypeOf DBobject Is DBText Then
                                    Dim Text1 As DBText = DBobject
                                    If Text1.Layer = ComboBox_layer.Text Then
                                        If Text1.Rotation > 3 * PI / 2 Then
                                            Data_table1.Rows.Add()
                                            Data_table1.Rows(Index1).Item("TEXT325") = Text1
                                            Index1 = Index1 + 1
                                        End If
                                        If Text1.Rotation >= 0 And Text1.Rotation < PI / 4 Then
                                            Data_table2.Rows.Add()
                                            Data_table2.Rows(Index2).Item("TEXT0") = Text1
                                            Index2 = Index2 + 1
                                        End If

                                    End If
                                End If
                            Next
                            Dim Chainage_on_vertex As Double
                            Dim Distanta_pana_la_Vertex As Double
                            Dim CSF1, CSF2 As Double

                            If Data_table1.Rows.Count > 0 Then
                                Dim Point_CHAINAGE As New Point3d
                                Point_CHAINAGE = Poly3d.GetPointAtParameter(Parameter_start)
                                Distanta_pana_la_Vertex = Point_CHAINAGE.GetVectorTo(Point_on_poly).Length

                                For i = 0 To Data_table1.Rows.Count - 1
                                    Dim Text1 As DBText = Data_table1.Rows(i).Item("TEXT325")
                                    If Point_CHAINAGE.GetVectorTo(Text1.Position.TransformBy(Editor1.CurrentUserCoordinateSystem)).Length < 0.1 Then
                                        Dim String1 As String = Replace(Text1.TextString, "+", "")
                                        If IsNumeric(String1) = True Then
                                            Chainage_on_vertex = CDbl(String1)
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                            If Not Parameter_start = Parameter_end Then
                                If Data_table2.Rows.Count > 0 Then
                                    Dim Point_CHAINAGE1 As New Point3d
                                    Point_CHAINAGE1 = Poly3d.GetPointAtParameter(Parameter_start)
                                    Dim Point_CHAINAGE2 As New Point3d
                                    Point_CHAINAGE2 = Poly3d.GetPointAtParameter(Parameter_end)

                                    For i = 0 To Data_table2.Rows.Count - 1
                                        Dim Text1 As DBText = Data_table2.Rows(i).Item("TEXT0")
                                        Dim String1 As String = Text1.TextString
                                        String1 = extrage_numar_din_text_de_la_sfarsitul_textului(String1)
                                        If IsNumeric(String1) = True Then
                                            If Point_CHAINAGE1.GetVectorTo(Text1.Position.TransformBy(Editor1.CurrentUserCoordinateSystem)).Length < 0.1 Then
                                                If CDbl(String1) > 0.5 And CDbl(String1) < 1.5 Then
                                                    CSF1 = CDbl(String1)
                                                End If
                                            End If

                                            If Point_CHAINAGE2.GetVectorTo(Text1.Position.TransformBy(Editor1.CurrentUserCoordinateSystem)).Length < 0.1 Then
                                                If CDbl(String1) > 0.5 And CDbl(String1) < 1.5 Then
                                                    CSF2 = CDbl(String1)
                                                End If
                                            End If
                                        End If
                                    Next
                                End If
                            End If


                            Dim New_ch As Double
                            Dim Old_ch As Double
                            Dim Diferenta As Double
                            Dim New_chainage As String

                            If Not CSF1 + CSF2 = 0 And Not CSF1 = 0 And Not CSF2 = 0 Then
                                New_ch = Chainage_on_vertex + Distanta_pana_la_Vertex / ((CSF1 + CSF2) / 2)
                            Else
                                New_ch = Chainage_on_vertex + Distanta_pana_la_Vertex
                            End If


                            New_chainage = Get_chainage_from_double(New_ch, 1)

                            ComboBox_CROSS_STA.Items.Add(New_chainage)
                            ComboBox_CROSS_STA.SelectedIndex = ComboBox_CROSS_STA.Items.IndexOf(New_chainage)



                            Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly, New_chainage, 0.5, 0.5, 0.5, 11, 3.5)

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

    Private Sub Button_PIPEMR_CLEAR_Click(sender As Object, e As EventArgs) Handles Button_PIPEMR_CLEAR.Click
        ComboBox_PIPEMR_MAT.Items.Clear()
        ComboBox_PIPEMR_BEGINSTA.Items.Clear()
        ComboBox_PIPEMR_LENGTH.Items.Clear()
        Button_PIPEMR_BEGINSTA.BackColor = Drawing.Color.DimGray
        Button_PIPEMR_MAT.BackColor = Drawing.Color.DimGray
        Button_PIPEMR_LENGTH.BackColor = Drawing.Color.DimGray
    End Sub
    Private Sub Button_PIPEMR_PICK_Click(sender As Object, e As EventArgs) Handles Button_PIPEMR_PICK.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info {BEGINSTA}:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt2.MessageForAdding = vbLf & "Select info {LENGTH}:"

                    Object_Prompt2.SingleOnly = True
                    Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                    If Not Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    Dim Rezultat3 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt3 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt3.MessageForAdding = vbLf & "Select info {MAT}:"

                    Object_Prompt3.SingleOnly = True
                    Rezultat3 = Editor1.GetSelection(Object_Prompt3)
                    If Not Rezultat3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If


                    ComboBox_PIPEMR_BEGINSTA.Items.Clear()
                    ComboBox_PIPEMR_BEGINSTA.Items.Add(" ")
                    ComboBox_PIPEMR_LENGTH.Items.Clear()
                    ComboBox_PIPEMR_LENGTH.Items.Add(" ")
                    ComboBox_PIPEMR_MAT.Items.Clear()
                    ComboBox_PIPEMR_MAT.Items.Add(" ")


                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent2 As Entity = Rezultat2.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent3 As Entity = Rezultat3.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)

                    Dim Continut1 As String
                    Dim Continut2 As String
                    Dim Continut3 As String


                    If TypeOf ent1 Is DBText Then
                        Dim Text1 As DBText = ent1
                        Continut1 = Text1.TextString
                    End If
                    If TypeOf ent2 Is DBText Then
                        Dim Text2 As DBText = ent2
                        Continut2 = Text2.TextString
                    End If
                    If TypeOf ent3 Is DBText Then
                        Dim Text3 As DBText = ent3
                        Continut3 = Text3.TextString
                    End If

                    If TypeOf ent1 Is MText Then
                        Dim MText1 As MText = ent1
                        Continut1 = MText1.Contents
                    End If
                    If TypeOf ent2 Is MText Then
                        Dim MText2 As MText = ent2
                        Continut2 = MText2.Contents
                    End If
                    If TypeOf ent3 Is MText Then
                        Dim MText3 As MText = ent3
                        Continut3 = MText3.Contents
                    End If


                    If TypeOf ent1 Is BlockReference Then
                        Dim Block1 As BlockReference = ent1
                        If Block1.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                            Dim ctemp1 As String
                            Dim ctemp2 As String
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "BEGINSTA" Then
                                    ctemp1 = attref.TextString
                                End If
                                If attref.Tag = "ENDSTA" Then
                                    ctemp2 = attref.TextString
                                End If
                            Next
                            ctemp1 = Replace(ctemp1, " ", "")
                            ctemp2 = Replace(ctemp2, " ", "")



                            If Continut2 = "" Then
                                If TypeOf ent2 Is BlockReference And Not ent1.ObjectId = ent2.ObjectId Then
                                    If Not ctemp1 = "" And Not ctemp2 = "" Then
                                        Continut1 = ctemp2
                                    End If
                                    If ctemp1 = "" And ctemp2 = "" Then
                                        Continut1 = " "
                                    End If
                                    If ctemp1 = "" And Not ctemp2 = "" Then
                                        Continut1 = ctemp2
                                    End If
                                    If Not ctemp1 = "" And ctemp2 = "" Then
                                        Continut1 = ctemp1
                                    End If
                                End If
                            Else
                                If Not ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut1 = ctemp2
                                End If
                                If ctemp1 = "" And ctemp2 = "" Then
                                    Continut1 = " "
                                End If
                                If ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut1 = ctemp2
                                End If
                                If Not ctemp1 = "" And ctemp2 = "" Then
                                    Continut1 = ctemp1
                                End If
                            End If
                            If ent1.ObjectId = ent2.ObjectId Then
                                Continut1 = ctemp1
                            End If

                        End If
                    End If


                    If TypeOf ent2 Is BlockReference Then
                        Dim Block2 As BlockReference = ent2
                        If Block2.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block2.AttributeCollection
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "LENGTH" Then
                                    Continut2 = attref.TextString
                                End If
                            Next
                        End If
                    End If


                    If TypeOf ent3 Is BlockReference Then
                        Dim Block3 As BlockReference = ent3
                        If Block3.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block3.AttributeCollection
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "MAT" Then
                                    Continut3 = attref.TextString
                                End If
                            Next
                        End If
                    End If

                    If Not Continut1 = "" Then
                        ComboBox_PIPEMR_BEGINSTA.Items.Add(Continut1)
                        ComboBox_PIPEMR_BEGINSTA.SelectedIndex = ComboBox_PIPEMR_BEGINSTA.Items.IndexOf(Continut1)
                        Button_PIPEMR_BEGINSTA.BackColor = Drawing.Color.Lime
                        ComboBox_PIPEMR_LENGTH.Items.Add(Continut1)
                        ComboBox_PIPEMR_MAT.Items.Add(Continut1)

                    End If
                    If Not Continut2 = "" Then
                        ComboBox_PIPEMR_BEGINSTA.Items.Add(Continut2)
                        ComboBox_PIPEMR_LENGTH.Items.Add(Continut2)
                        ComboBox_PIPEMR_LENGTH.SelectedIndex = ComboBox_PIPEMR_LENGTH.Items.IndexOf(Continut2)
                        Button_PIPEMR_LENGTH.BackColor = Drawing.Color.Lime
                        ComboBox_PIPEMR_MAT.Items.Add(Continut2)
                    End If
                    If Not Continut3 = "" Then
                        ComboBox_PIPEMR_BEGINSTA.Items.Add(Continut3)
                        ComboBox_PIPEMR_LENGTH.Items.Add(Continut3)
                        ComboBox_PIPEMR_MAT.Items.Add(Continut3)
                        ComboBox_PIPEMR_MAT.SelectedIndex = ComboBox_PIPEMR_MAT.Items.IndexOf(Continut3)
                        Button_PIPEMR_MAT.BackColor = Drawing.Color.Lime
                    End If






                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_PIPEMR_INSERT_Click(sender As Object, e As EventArgs) Handles Button_PIPEMR_INSERT.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select objects to be deleted:"

                    Object_Prompt1.SingleOnly = False
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If
                    Dim Point_start As Autodesk.AutoCAD.EditorInput.PromptPointResult

                    Dim PP_start As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select insertion point:")
                    PP_start.AllowNone = True
                    Point_start = Editor1.GetPoint(PP_start)
                    If Not Point_start.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        Editor1.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If

                    Dim Colectie_atr_name As New Specialized.StringCollection
                    Dim Colectie_atr_value As New Specialized.StringCollection

                    If Not ComboBox_PIPEMR_BEGINSTA.Text = "" Then
                        Colectie_atr_name.Add("BEGINSTA")
                        Colectie_atr_value.Add(ComboBox_PIPEMR_BEGINSTA.Text)
                    End If

                    If Not ComboBox_PIPEMR_LENGTH.Text = "" Then
                        Colectie_atr_name.Add("LENGTH")
                        Colectie_atr_value.Add(ComboBox_PIPEMR_LENGTH.Text)
                    End If

                    If Not ComboBox_PIPEMR_MAT.Text = "" Then
                        Colectie_atr_name.Add("MAT")
                        Colectie_atr_value.Add(ComboBox_PIPEMR_MAT.Text)
                    End If


                    InsertBlock_with_multiple_atributes("heavy_wall_match_right_1.dwg", "heavy_wall_match_right_1", Point_start.Value, 1, BTrecord, "TEXT", Colectie_atr_name, Colectie_atr_value)

                    For i = 0 To Rezultat1.Value.Count - 1






                        Dim ent1 As Entity = Rezultat1.Value.Item(i).ObjectId.GetObject(OpenMode.ForWrite)
                        ent1.Erase()
                    Next








                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)

                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_PIPEMR_BEGINSTA_Click(sender As Object, e As EventArgs) Handles Button_PIPEMR_BEGINSTA.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info {BEGINSTA}:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut1 As String


                    If TypeOf ent1 Is DBText Then
                        Dim Text1 As DBText = ent1
                        Continut1 = Text1.TextString
                    End If

                    If TypeOf ent1 Is MText Then
                        Dim MText1 As MText = ent1
                        Continut1 = MText1.Contents
                    End If


                    If Not Continut1 = "" Then
                        If ComboBox_PIPEMR_BEGINSTA.Items.Count > 1 Then
                            ComboBox_PIPEMR_BEGINSTA.Items(1) = Continut1
                            ComboBox_PIPEMR_BEGINSTA.SelectedIndex = ComboBox_PIPEMR_BEGINSTA.Items.IndexOf(Continut1)
                            Button_PIPEMR_BEGINSTA.BackColor = Drawing.Color.Lime
                            ComboBox_PIPEMR_LENGTH.Items(1) = Continut1
                            ComboBox_PIPEMR_MAT.Items(1) = Continut1
                        End If
                        If ComboBox_PIPEMR_BEGINSTA.Items.Count = 0 Then
                            ComboBox_PIPEMR_BEGINSTA.Items.Add(" ")
                            ComboBox_PIPEMR_BEGINSTA.Items.Add(Continut1)
                            ComboBox_PIPEMR_BEGINSTA.SelectedIndex = ComboBox_PIPEMR_BEGINSTA.Items.IndexOf(Continut1)
                            Button_PIPEMR_BEGINSTA.BackColor = Drawing.Color.Lime
                            ComboBox_PIPEMR_LENGTH.Items.Add(" ")
                            ComboBox_PIPEMR_LENGTH.Items.Add(Continut1)
                            ComboBox_PIPEMR_MAT.Items.Add(" ")
                            ComboBox_PIPEMR_MAT.Items.Add(Continut1)
                        End If

                    Else
                        If ComboBox_PIPEMR_BEGINSTA.Items.Count = 0 Then
                            Button_PIPEMR_BEGINSTA.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_PIPEMR_LENGTH_Click(sender As Object, e As EventArgs) Handles Button_PIPEMR_LENGTH.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt2.MessageForAdding = vbLf & "Select info {LENGTH}:"

                    Object_Prompt2.SingleOnly = True
                    Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                    If Not Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent2 As Entity = Rezultat2.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut2 As String


                    If TypeOf ent2 Is DBText Then
                        Dim Text2 As DBText = ent2
                        Continut2 = Text2.TextString
                    End If

                    If TypeOf ent2 Is MText Then
                        Dim MText2 As MText = ent2
                        Continut2 = MText2.Contents
                    End If


                    If Not Continut2 = "" Then
                        If ComboBox_PIPEMR_LENGTH.Items.Count > 2 Then
                            ComboBox_PIPEMR_BEGINSTA.Items(2) = Continut2
                            ComboBox_PIPEMR_LENGTH.Items(2) = Continut2
                            ComboBox_PIPEMR_LENGTH.SelectedIndex = ComboBox_PIPEMR_LENGTH.Items.IndexOf(Continut2)
                            Button_PIPEMR_LENGTH.BackColor = Drawing.Color.Lime
                            ComboBox_PIPEMR_MAT.Items(2) = Continut2
                        End If
                        If ComboBox_PIPEMR_LENGTH.Items.Count = 0 Then
                            ComboBox_PIPEMR_BEGINSTA.Items.Add(" ")
                            ComboBox_PIPEMR_BEGINSTA.Items.Add(Continut2)
                            ComboBox_PIPEMR_LENGTH.Items.Add(" ")
                            ComboBox_PIPEMR_LENGTH.Items.Add(Continut2)
                            ComboBox_PIPEMR_LENGTH.SelectedIndex = ComboBox_PIPEMR_LENGTH.Items.IndexOf(Continut2)
                            Button_PIPEMR_LENGTH.BackColor = Drawing.Color.Lime
                            ComboBox_PIPEMR_MAT.Items.Add(" ")
                            ComboBox_PIPEMR_MAT.Items.Add(Continut2)
                        End If
                        If ComboBox_PIPEMR_LENGTH.Items.Count = 1 Or ComboBox_PIPEMR_LENGTH.Items.Count = 2 Then
                            ComboBox_PIPEMR_BEGINSTA.Items.Add(Continut2)
                            ComboBox_PIPEMR_LENGTH.Items.Add(Continut2)
                            ComboBox_PIPEMR_LENGTH.SelectedIndex = ComboBox_PIPEMR_LENGTH.Items.IndexOf(Continut2)
                            Button_PIPEMR_LENGTH.BackColor = Drawing.Color.Lime

                            ComboBox_PIPEMR_MAT.Items.Add(Continut2)
                        End If
                    Else
                        If ComboBox_PIPEMR_LENGTH.Items.Count = 0 Then
                            Button_PIPEMR_LENGTH.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_PIPEMR_MAT_Click(sender As Object, e As EventArgs) Handles Button_PIPEMR_MAT.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat3 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt3 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt3.MessageForAdding = vbLf & "Select info {MAT}:"

                    Object_Prompt3.SingleOnly = True
                    Rezultat3 = Editor1.GetSelection(Object_Prompt3)
                    If Not Rezultat3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If
                    Dim ent3 As Entity = Rezultat3.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim Continut3 As String
                    If TypeOf ent3 Is DBText Then
                        Dim Text3 As DBText = ent3
                        Continut3 = Text3.TextString
                    End If

                    If TypeOf ent3 Is MText Then
                        Dim MText3 As MText = ent3
                        Continut3 = MText3.Contents
                    End If


                    If Not Continut3 = "" Then
                        If ComboBox_PIPEMR_MAT.Items.Count > 3 Then
                            ComboBox_PIPEMR_BEGINSTA.Items(3) = Continut3
                            ComboBox_PIPEMR_LENGTH.Items(3) = Continut3
                            ComboBox_PIPEMR_MAT.Items(3) = Continut3
                            ComboBox_PIPEMR_MAT.SelectedIndex = ComboBox_PIPEMR_MAT.Items.IndexOf(Continut3)
                            Button_PIPEMR_MAT.BackColor = Drawing.Color.Lime
                        End If
                        If ComboBox_PIPEMR_MAT.Items.Count = 0 Then
                            ComboBox_PIPEMR_BEGINSTA.Items.Add(" ")
                            ComboBox_PIPEMR_BEGINSTA.Items.Add(Continut3)
                            ComboBox_PIPEMR_LENGTH.Items.Add(" ")
                            ComboBox_PIPEMR_LENGTH.Items.Add(Continut3)

                            ComboBox_PIPEMR_MAT.Items.Add(" ")
                            ComboBox_PIPEMR_MAT.Items.Add(Continut3)
                            ComboBox_PIPEMR_MAT.SelectedIndex = ComboBox_PIPEMR_MAT.Items.IndexOf(Continut3)
                            Button_PIPEMR_MAT.BackColor = Drawing.Color.Lime
                        End If
                        If ComboBox_PIPEMR_MAT.Items.Count = 1 Or ComboBox_PIPEMR_MAT.Items.Count = 2 Or ComboBox_PIPEMR_MAT.Items.Count = 3 Then
                            ComboBox_PIPEMR_BEGINSTA.Items.Add(Continut3)


                            ComboBox_PIPEMR_LENGTH.Items.Add(Continut3)


                            ComboBox_PIPEMR_MAT.Items.Add(Continut3)
                            ComboBox_PIPEMR_MAT.SelectedIndex = ComboBox_PIPEMR_MAT.Items.IndexOf(Continut3)
                            Button_PIPEMR_MAT.BackColor = Drawing.Color.Lime
                        End If
                    Else
                        If ComboBox_PIPEMR_MAT.Items.Count = 0 Then
                            Button_PIPEMR_MAT.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_PIPEMR_blockpick_Click(sender As Object, e As EventArgs) Handles Button_pipeMR_blockpick.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info from Block:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    ComboBox_PIPEMR_BEGINSTA.Items.Clear()
                    ComboBox_PIPEMR_BEGINSTA.Items.Add(" ")
                    ComboBox_PIPEMR_LENGTH.Items.Clear()
                    ComboBox_PIPEMR_LENGTH.Items.Add(" ")
                    ComboBox_PIPEMR_MAT.Items.Clear()
                    ComboBox_PIPEMR_MAT.Items.Add(" ")

                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut1 As String
                    Dim Continut2 As String
                    Dim Continut3 As String



                    If TypeOf ent1 Is BlockReference Then
                        Dim Block1 As BlockReference = ent1
                        If Block1.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "BEGINSTA" Then
                                    Continut1 = attref.TextString
                                    If Continut1 = "" Then Continut1 = " "
                                End If
                                If attref.Tag = "LENGTH" Then
                                    Continut2 = attref.TextString
                                    If Continut2 = "" Then Continut2 = " "
                                End If
                                If attref.Tag = "MAT" Then
                                    Continut3 = attref.TextString
                                    If Continut3 = "" Then Continut3 = " "
                                End If
                            Next
                        End If
                    End If

                    If Not Continut1 = "" Then
                        ComboBox_PIPEMR_BEGINSTA.Items.Add(Continut1)
                        ComboBox_PIPEMR_BEGINSTA.SelectedIndex = ComboBox_PIPEMR_BEGINSTA.Items.IndexOf(Continut1)
                        Button_PIPEMR_BEGINSTA.BackColor = Drawing.Color.Lime
                        ComboBox_PIPEMR_LENGTH.Items.Add(Continut1)
                        ComboBox_PIPEMR_MAT.Items.Add(Continut1)
                    End If
                    If Not Continut2 = "" Then
                        ComboBox_PIPEMR_BEGINSTA.Items.Add(Continut2)
                        ComboBox_PIPEMR_LENGTH.Items.Add(Continut2)
                        ComboBox_PIPEMR_LENGTH.SelectedIndex = ComboBox_PIPEMR_LENGTH.Items.IndexOf(Continut2)
                        Button_PIPEMR_LENGTH.BackColor = Drawing.Color.Lime
                        ComboBox_PIPEMR_MAT.Items.Add(Continut2)
                    End If
                    If Not Continut3 = "" Then
                        ComboBox_PIPEMR_BEGINSTA.Items.Add(Continut3)
                        ComboBox_PIPEMR_LENGTH.Items.Add(Continut3)
                        ComboBox_PIPEMR_MAT.Items.Add(Continut3)
                        ComboBox_PIPEMR_MAT.SelectedIndex = ComboBox_PIPEMR_MAT.Items.IndexOf(Continut3)
                        Button_PIPEMR_MAT.BackColor = Drawing.Color.Lime

                    End If

                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub

    Private Sub Button_screw_clear_Click(sender As Object, e As EventArgs) Handles Button_screw_clear.Click
        ComboBox_screw_spacing.Items.Clear()
        ComboBox_screw_beginsta.Items.Clear()
        ComboBox_screw_endsta.Items.Clear()
        ComboBox_screw_no_type.Items.Clear()
        Button_screw_Beginsta.BackColor = Drawing.Color.DimGray
        Button_screw_Endsta.BackColor = Drawing.Color.DimGray
        Button_screw_spacing.BackColor = Drawing.Color.DimGray
        Button_screw_no_type.BackColor = Drawing.Color.DimGray
    End Sub
    Private Sub Button_screw_pick_Click(sender As Object, e As EventArgs) Handles Button_screw_pick.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info {BEGINSTA}:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt2.MessageForAdding = vbLf & "Select info {ENDSTA}:"

                    Object_Prompt2.SingleOnly = True
                    Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                    If Not Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    Dim Rezultat3 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt3 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt3.MessageForAdding = vbLf & "Select info {NO_TYPE}:"

                    Object_Prompt3.SingleOnly = True
                    Rezultat3 = Editor1.GetSelection(Object_Prompt3)
                    If Not Rezultat3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    Dim Rezultat4 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt4 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt4.MessageForAdding = vbLf & "Select info {SPACING}:"

                    Object_Prompt4.SingleOnly = True
                    Rezultat4 = Editor1.GetSelection(Object_Prompt4)
                    If Not Rezultat4.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    ComboBox_screw_beginsta.Items.Clear()
                    ComboBox_screw_beginsta.Items.Add(" ")
                    ComboBox_screw_endsta.Items.Clear()
                    ComboBox_screw_endsta.Items.Add(" ")
                    ComboBox_screw_no_type.Items.Clear()
                    ComboBox_screw_no_type.Items.Add(" ")
                    ComboBox_screw_spacing.Items.Clear()
                    ComboBox_screw_spacing.Items.Add(" ")


                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent2 As Entity = Rezultat2.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent3 As Entity = Rezultat3.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent4 As Entity = Rezultat4.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)

                    Dim Continut1 As String
                    Dim Continut2 As String
                    Dim Continut3 As String
                    Dim Continut4 As String

                    If TypeOf ent1 Is DBText Then
                        Dim Text1 As DBText = ent1
                        Continut1 = Text1.TextString
                    End If
                    If TypeOf ent2 Is DBText Then
                        Dim Text2 As DBText = ent2
                        Continut2 = Text2.TextString
                    End If
                    If TypeOf ent3 Is DBText Then
                        Dim Text3 As DBText = ent3
                        Continut3 = Text3.TextString
                    End If
                    If TypeOf ent4 Is DBText Then
                        Dim Text4 As DBText = ent4
                        Continut4 = Text4.TextString
                    End If

                    If TypeOf ent1 Is MText Then
                        Dim MText1 As MText = ent1
                        Continut1 = MText1.Contents
                    End If
                    If TypeOf ent2 Is MText Then
                        Dim MText2 As MText = ent2
                        Continut2 = MText2.Contents
                    End If
                    If TypeOf ent3 Is MText Then
                        Dim MText3 As MText = ent3
                        Continut3 = MText3.Contents
                    End If
                    If TypeOf ent4 Is MText Then
                        Dim MText4 As MText = ent4
                        Continut4 = MText4.Contents
                    End If


                    If TypeOf ent1 Is BlockReference Then
                        Dim Block1 As BlockReference = ent1
                        If Block1.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                            Dim ctemp1 As String
                            Dim ctemp2 As String
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "BEGINSTA" Then
                                    ctemp1 = attref.TextString
                                End If
                                If attref.Tag = "ENDSTA" Then
                                    ctemp2 = attref.TextString
                                End If
                            Next
                            ctemp1 = Replace(ctemp1, " ", "")
                            ctemp2 = Replace(ctemp2, " ", "")
                            If TypeOf ent2 Is BlockReference Then
                                If Not ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut1 = ctemp2
                                End If
                                If ctemp1 = "" And ctemp2 = "" Then
                                    Continut1 = " "
                                End If
                                If ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut1 = ctemp2
                                End If
                                If Not ctemp1 = "" And ctemp2 = "" Then
                                    Continut1 = ctemp1
                                End If
                            Else
                                Continut1 = ctemp1
                            End If
                        End If
                    End If


                    If TypeOf ent2 Is BlockReference Then
                        Dim Block2 As BlockReference = ent2
                        If Block2.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block2.AttributeCollection
                            Dim ctemp1 As String
                            Dim ctemp2 As String
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "BEGINSTA" Then
                                    ctemp1 = attref.TextString
                                End If
                                If attref.Tag = "ENDSTA" Then
                                    ctemp2 = attref.TextString
                                End If
                            Next
                            ctemp1 = Replace(ctemp1, " ", "")
                            ctemp2 = Replace(ctemp2, " ", "")
                            If TypeOf ent1 Is BlockReference Then
                                If Not ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut2 = ctemp1
                                End If
                                If ctemp1 = "" And ctemp2 = "" Then
                                    Continut2 = " "
                                End If
                                If ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut2 = ctemp2
                                End If
                                If Not ctemp1 = "" And ctemp2 = "" Then
                                    Continut2 = ctemp1
                                End If
                            Else
                                Continut2 = ctemp2
                            End If

                        End If
                    End If


                    If TypeOf ent3 Is BlockReference Then
                        Dim Block3 As BlockReference = ent3
                        If Block3.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block3.AttributeCollection
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "NO_TYPE" Then
                                    Continut3 = attref.TextString
                                End If
                            Next
                        End If
                    End If


                    If TypeOf ent4 Is BlockReference Then
                        Dim Block4 As BlockReference = ent4
                        If Block4.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block4.AttributeCollection
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "SPACING" Then
                                    Continut4 = attref.TextString
                                End If
                            Next
                        End If
                    End If



                    If Not Continut1 = "" Then
                        ComboBox_screw_beginsta.Items.Add(Continut1)
                        ComboBox_screw_beginsta.SelectedIndex = ComboBox_screw_beginsta.Items.IndexOf(Continut1)
                        Button_screw_Beginsta.BackColor = Drawing.Color.Lime
                        ComboBox_screw_endsta.Items.Add(Continut1)
                        ComboBox_screw_no_type.Items.Add(Continut1)
                        ComboBox_screw_spacing.Items.Add(Continut1)

                    End If
                    If Not Continut2 = "" Then
                        ComboBox_screw_beginsta.Items.Add(Continut2)
                        ComboBox_screw_endsta.Items.Add(Continut2)
                        ComboBox_screw_endsta.SelectedIndex = ComboBox_screw_endsta.Items.IndexOf(Continut2)
                        Button_screw_Endsta.BackColor = Drawing.Color.Lime
                        ComboBox_screw_no_type.Items.Add(Continut2)
                        ComboBox_screw_spacing.Items.Add(Continut2)
                    End If
                    If Not Continut3 = "" Then
                        ComboBox_screw_beginsta.Items.Add(Continut3)
                        ComboBox_screw_endsta.Items.Add(Continut3)
                        ComboBox_screw_no_type.Items.Add(Continut3)
                        ComboBox_screw_no_type.SelectedIndex = ComboBox_screw_no_type.Items.IndexOf(Continut3)
                        Button_screw_no_type.BackColor = Drawing.Color.Lime
                        ComboBox_screw_spacing.Items.Add(Continut3)
                    End If
                    If Not Continut4 = "" Then
                        ComboBox_screw_beginsta.Items.Add(Continut4)
                        ComboBox_screw_endsta.Items.Add(Continut4)
                        ComboBox_screw_no_type.Items.Add(Continut4)
                        ComboBox_screw_spacing.Items.Add(Continut4)
                        ComboBox_screw_spacing.SelectedIndex = ComboBox_screw_spacing.Items.IndexOf(Continut4)
                        Button_screw_spacing.BackColor = Drawing.Color.Lime
                    End If






                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_screw_insert_Click(sender As Object, e As EventArgs) Handles Button_screw_insert.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select objects to be deleted:"

                    Object_Prompt1.SingleOnly = False
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If
                    Dim Point_start As Autodesk.AutoCAD.EditorInput.PromptPointResult

                    Dim PP_start As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select insertion point:")
                    PP_start.AllowNone = True
                    Point_start = Editor1.GetPoint(PP_start)
                    If Not Point_start.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        Editor1.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If

                    Dim Colectie_atr_name As New Specialized.StringCollection
                    Dim Colectie_atr_value As New Specialized.StringCollection

                    If Not ComboBox_screw_beginsta.Text = "" Then
                        Colectie_atr_name.Add("BEGINSTA")
                        Colectie_atr_value.Add(ComboBox_screw_beginsta.Text)
                    End If

                    If Not ComboBox_screw_endsta.Text = "" Then
                        Colectie_atr_name.Add("ENDSTA")
                        Colectie_atr_value.Add(ComboBox_screw_endsta.Text)
                    End If

                    If Not ComboBox_screw_no_type.Text = "" Then
                        Colectie_atr_name.Add("NO_TYPE")
                        Colectie_atr_value.Add(ComboBox_screw_no_type.Text)
                    End If

                    If Not ComboBox_screw_spacing.Text = "" Then
                        Colectie_atr_name.Add("SPACING")
                        Colectie_atr_value.Add(ComboBox_screw_spacing.Text)
                    End If


                    InsertBlock_with_multiple_atributes("Screw_anchor_alignment1.dwg", "Screw_anchor_alignment1", Point_start.Value, 1, BTrecord, "TEXT", Colectie_atr_name, Colectie_atr_value)

                    For i = 0 To Rezultat1.Value.Count - 1






                        Dim ent1 As Entity = Rezultat1.Value.Item(i).ObjectId.GetObject(OpenMode.ForWrite)
                        ent1.Erase()
                    Next








                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)

                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_screw_Beginsta_Click(sender As Object, e As EventArgs) Handles Button_screw_Beginsta.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info {BEGINSTA}:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut1 As String


                    If TypeOf ent1 Is DBText Then
                        Dim Text1 As DBText = ent1
                        Continut1 = Text1.TextString
                    End If

                    If TypeOf ent1 Is MText Then
                        Dim MText1 As MText = ent1
                        Continut1 = MText1.Contents
                    End If


                    If Not Continut1 = "" Then
                        If ComboBox_screw_beginsta.Items.Count > 1 Then
                            ComboBox_screw_beginsta.Items(1) = Continut1
                            ComboBox_screw_beginsta.SelectedIndex = ComboBox_screw_beginsta.Items.IndexOf(Continut1)
                            Button_screw_Beginsta.BackColor = Drawing.Color.Lime
                            ComboBox_screw_endsta.Items(1) = Continut1
                            ComboBox_screw_no_type.Items(1) = Continut1
                            ComboBox_screw_spacing.Items(1) = Continut1
                        End If
                        If ComboBox_screw_beginsta.Items.Count = 0 Then
                            ComboBox_screw_beginsta.Items.Add(" ")
                            ComboBox_screw_beginsta.Items.Add(Continut1)
                            ComboBox_screw_beginsta.SelectedIndex = ComboBox_screw_beginsta.Items.IndexOf(Continut1)
                            Button_screw_Beginsta.BackColor = Drawing.Color.Lime
                            ComboBox_screw_endsta.Items.Add(" ")
                            ComboBox_screw_endsta.Items.Add(Continut1)
                            ComboBox_screw_no_type.Items.Add(" ")
                            ComboBox_screw_no_type.Items.Add(Continut1)
                            ComboBox_screw_spacing.Items.Add(" ")
                            ComboBox_screw_spacing.Items.Add(Continut1)
                        End If

                    Else
                        If ComboBox_screw_beginsta.Items.Count = 0 Then
                            Button_screw_Beginsta.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_screw_Endsta_Click(sender As Object, e As EventArgs) Handles Button_screw_Endsta.Click

        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt2.MessageForAdding = vbLf & "Select info {ENDSTA}:"

                    Object_Prompt2.SingleOnly = True
                    Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                    If Not Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent2 As Entity = Rezultat2.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut2 As String


                    If TypeOf ent2 Is DBText Then
                        Dim Text2 As DBText = ent2
                        Continut2 = Text2.TextString
                    End If

                    If TypeOf ent2 Is MText Then
                        Dim MText2 As MText = ent2
                        Continut2 = MText2.Contents
                    End If


                    If Not Continut2 = "" Then
                        If ComboBox_screw_endsta.Items.Count > 2 Then
                            ComboBox_screw_beginsta.Items(2) = Continut2
                            ComboBox_screw_endsta.Items(2) = Continut2
                            ComboBox_screw_endsta.SelectedIndex = ComboBox_screw_endsta.Items.IndexOf(Continut2)
                            Button_screw_Endsta.BackColor = Drawing.Color.Lime
                            ComboBox_screw_no_type.Items(2) = Continut2
                            ComboBox_screw_spacing.Items(2) = Continut2
                        End If
                        If ComboBox_screw_endsta.Items.Count = 0 Then
                            ComboBox_screw_beginsta.Items.Add(" ")
                            ComboBox_screw_beginsta.Items.Add(Continut2)
                            ComboBox_screw_endsta.Items.Add(" ")
                            ComboBox_screw_endsta.Items.Add(Continut2)
                            ComboBox_screw_endsta.SelectedIndex = ComboBox_screw_endsta.Items.IndexOf(Continut2)
                            Button_screw_Endsta.BackColor = Drawing.Color.Lime
                            ComboBox_screw_no_type.Items.Add(" ")
                            ComboBox_screw_no_type.Items.Add(Continut2)
                            ComboBox_screw_spacing.Items.Add(" ")
                            ComboBox_screw_spacing.Items.Add(Continut2)
                        End If
                        If ComboBox_screw_endsta.Items.Count = 1 Or ComboBox_screw_endsta.Items.Count = 2 Then
                            ComboBox_screw_beginsta.Items.Add(Continut2)
                            ComboBox_screw_endsta.Items.Add(Continut2)
                            ComboBox_screw_endsta.SelectedIndex = ComboBox_screw_endsta.Items.IndexOf(Continut2)
                            Button_screw_Endsta.BackColor = Drawing.Color.Lime
                            ComboBox_screw_no_type.Items.Add(Continut2)
                            ComboBox_screw_spacing.Items.Add(Continut2)
                        End If
                    Else
                        If ComboBox_screw_endsta.Items.Count = 0 Then
                            Button_screw_Endsta.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_screw_no_type_Click(sender As Object, e As EventArgs) Handles Button_screw_no_type.Click

        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat5 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt5 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt5.MessageForAdding = vbLf & "Select info {NO_TYPE}:"

                    Object_Prompt5.SingleOnly = True
                    Rezultat5 = Editor1.GetSelection(Object_Prompt5)
                    If Not Rezultat5.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent5 As Entity = Rezultat5.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut5 As String


                    If TypeOf ent5 Is DBText Then
                        Dim Text5 As DBText = ent5
                        Continut5 = Text5.TextString
                    End If

                    If TypeOf ent5 Is MText Then
                        Dim MText5 As MText = ent5
                        Continut5 = MText5.Contents
                    End If


                    If Not Continut5 = "" Then
                        If ComboBox_screw_no_type.Items.Count > 5 Then
                            ComboBox_screw_beginsta.Items(5) = Continut5
                            ComboBox_screw_endsta.Items(5) = Continut5
                            ComboBox_screw_no_type.Items(5) = Continut5
                            ComboBox_screw_no_type.SelectedIndex = ComboBox_screw_no_type.Items.IndexOf(Continut5)
                            Button_screw_no_type.BackColor = Drawing.Color.Lime
                            ComboBox_screw_spacing.Items(5) = Continut5
                        End If
                        If ComboBox_screw_no_type.Items.Count = 0 Then
                            ComboBox_screw_beginsta.Items.Add(" ")
                            ComboBox_screw_beginsta.Items.Add(Continut5)
                            ComboBox_screw_endsta.Items.Add(" ")
                            ComboBox_screw_endsta.Items.Add(Continut5)

                            ComboBox_screw_no_type.Items.Add(" ")
                            ComboBox_screw_no_type.Items.Add(Continut5)
                            ComboBox_screw_no_type.SelectedIndex = ComboBox_screw_no_type.Items.IndexOf(Continut5)
                            Button_screw_no_type.BackColor = Drawing.Color.Lime
                            ComboBox_screw_spacing.Items.Add(" ")
                            ComboBox_screw_spacing.Items.Add(Continut5)
                        End If
                        If ComboBox_screw_no_type.Items.Count = 1 Or ComboBox_screw_no_type.Items.Count = 2 Or ComboBox_screw_no_type.Items.Count = 3 Then
                            ComboBox_screw_beginsta.Items.Add(Continut5)
                            ComboBox_screw_endsta.Items.Add(Continut5)

                            ComboBox_screw_no_type.Items.Add(Continut5)
                            ComboBox_screw_no_type.SelectedIndex = ComboBox_screw_no_type.Items.IndexOf(Continut5)
                            Button_screw_no_type.BackColor = Drawing.Color.Lime

                            ComboBox_screw_spacing.Items.Add(Continut5)
                        End If
                    Else
                        If ComboBox_screw_no_type.Items.Count = 0 Then
                            Button_screw_no_type.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_screw_spacing_Click(sender As Object, e As EventArgs) Handles Button_screw_spacing.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat6 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt6 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt6.MessageForAdding = vbLf & "Select info {SPACING}:"

                    Object_Prompt6.SingleOnly = True
                    Rezultat6 = Editor1.GetSelection(Object_Prompt6)
                    If Not Rezultat6.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If
                    Dim ent6 As Entity = Rezultat6.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim Continut6 As String
                    If TypeOf ent6 Is DBText Then
                        Dim Text6 As DBText = ent6
                        Continut6 = Text6.TextString
                    End If

                    If TypeOf ent6 Is MText Then
                        Dim MText3 As MText = ent6
                        Continut6 = MText3.Contents
                    End If


                    If Not Continut6 = "" Then
                        If ComboBox_screw_spacing.Items.Count > 6 Then
                            ComboBox_screw_beginsta.Items(6) = Continut6
                            ComboBox_screw_endsta.Items(6) = Continut6
                            ComboBox_screw_no_type.Items(6) = Continut6
                            ComboBox_screw_spacing.Items(6) = Continut6
                            ComboBox_screw_spacing.SelectedIndex = ComboBox_screw_spacing.Items.IndexOf(Continut6)
                            Button_screw_spacing.BackColor = Drawing.Color.Lime
                        End If
                        If ComboBox_screw_spacing.Items.Count = 0 Then
                            ComboBox_screw_beginsta.Items.Add(" ")
                            ComboBox_screw_beginsta.Items.Add(Continut6)
                            ComboBox_screw_endsta.Items.Add(" ")
                            ComboBox_screw_endsta.Items.Add(Continut6)

                            ComboBox_screw_no_type.Items.Add(" ")
                            ComboBox_screw_no_type.Items.Add(Continut6)

                            ComboBox_screw_spacing.Items.Add(" ")
                            ComboBox_screw_spacing.Items.Add(Continut6)
                            ComboBox_screw_spacing.SelectedIndex = ComboBox_screw_spacing.Items.IndexOf(Continut6)
                            Button_screw_spacing.BackColor = Drawing.Color.Lime
                        End If
                        If ComboBox_screw_spacing.Items.Count = 1 Or ComboBox_screw_spacing.Items.Count = 2 Or ComboBox_screw_spacing.Items.Count = 3 Or ComboBox_screw_spacing.Items.Count = 4 Then
                            ComboBox_screw_beginsta.Items.Add(Continut6)
                            ComboBox_screw_endsta.Items.Add(Continut6)

                            ComboBox_screw_no_type.Items.Add(Continut6)


                            ComboBox_screw_spacing.Items.Add(Continut6)
                            ComboBox_screw_spacing.SelectedIndex = ComboBox_screw_spacing.Items.IndexOf(Continut6)
                            Button_screw_spacing.BackColor = Drawing.Color.Lime
                        End If
                    Else
                        If ComboBox_screw_spacing.Items.Count = 0 Then
                            Button_screw_spacing.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub

    Private Sub Button_ScrewL_clear_Click(sender As Object, e As EventArgs) Handles Button_ScrewL_clear.Click
        ComboBox_screwL_endsta.Items.Clear()
        ComboBox_screwL_no_type.Items.Clear()
        ComboBox_screwL_spacing.Items.Clear()
        ComboBox_screwL_endsta.BackColor = Drawing.Color.DimGray
        ComboBox_screwL_no_type.BackColor = Drawing.Color.DimGray
        ComboBox_screwL_spacing.BackColor = Drawing.Color.DimGray
    End Sub
    Private Sub Button__ScrewL_Pick_Click(sender As Object, e As EventArgs) Handles Button__ScrewL_Pick.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info {ENDSTA}:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt2.MessageForAdding = vbLf & "Select info {NO_TYPE}:"

                    Object_Prompt2.SingleOnly = True
                    Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                    If Not Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    Dim Rezultat3 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt3 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt3.MessageForAdding = vbLf & "Select info {SPACING}:"

                    Object_Prompt3.SingleOnly = True
                    Rezultat3 = Editor1.GetSelection(Object_Prompt3)
                    If Not Rezultat3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If


                    ComboBox_screwL_endsta.Items.Clear()
                    ComboBox_screwL_endsta.Items.Add(" ")
                    ComboBox_screwL_no_type.Items.Clear()
                    ComboBox_screwL_no_type.Items.Add(" ")
                    ComboBox_screwL_spacing.Items.Clear()
                    ComboBox_screwL_spacing.Items.Add(" ")


                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent2 As Entity = Rezultat2.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent3 As Entity = Rezultat3.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)

                    Dim Continut1 As String
                    Dim Continut2 As String
                    Dim Continut3 As String


                    If TypeOf ent1 Is DBText Then
                        Dim Text1 As DBText = ent1
                        Continut1 = Text1.TextString
                    End If
                    If TypeOf ent2 Is DBText Then
                        Dim Text2 As DBText = ent2
                        Continut2 = Text2.TextString
                    End If
                    If TypeOf ent3 Is DBText Then
                        Dim Text3 As DBText = ent3
                        Continut3 = Text3.TextString
                    End If

                    If TypeOf ent1 Is MText Then
                        Dim MText1 As MText = ent1
                        Continut1 = MText1.Contents
                    End If
                    If TypeOf ent2 Is MText Then
                        Dim MText2 As MText = ent2
                        Continut2 = MText2.Contents
                    End If
                    If TypeOf ent3 Is MText Then
                        Dim MText3 As MText = ent3
                        Continut3 = MText3.Contents
                    End If


                    If TypeOf ent1 Is BlockReference Then
                        Dim Block1 As BlockReference = ent1
                        If Block1.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                            Dim ctemp1 As String
                            Dim ctemp2 As String
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "BEGINSTA" Then
                                    ctemp1 = attref.TextString
                                End If
                                If attref.Tag = "ENDSTA" Then
                                    ctemp2 = attref.TextString
                                End If
                            Next
                            ctemp1 = Replace(ctemp1, " ", "")
                            ctemp2 = Replace(ctemp2, " ", "")
                            If TypeOf ent2 Is BlockReference Then
                                If Not ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut1 = ctemp2
                                End If
                                If ctemp1 = "" And ctemp2 = "" Then
                                    Continut1 = " "
                                End If
                                If ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut1 = ctemp2
                                End If
                                If Not ctemp1 = "" And ctemp2 = "" Then
                                    Continut1 = ctemp1
                                End If
                            Else
                                Continut1 = ctemp1
                            End If
                        End If
                    End If


                    If TypeOf ent2 Is BlockReference Then
                        Dim Block2 As BlockReference = ent2
                        If Block2.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block2.AttributeCollection
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "NO_TYPE" Then
                                    Continut2 = attref.TextString
                                End If
                            Next
                        End If
                    End If


                    If TypeOf ent3 Is BlockReference Then
                        Dim Block3 As BlockReference = ent3
                        If Block3.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block3.AttributeCollection
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "SPACING" Then
                                    Continut3 = attref.TextString
                                End If
                            Next
                        End If
                    End If

                    If Not Continut1 = "" Then
                        ComboBox_screwL_endsta.Items.Add(Continut1)
                        ComboBox_screwL_endsta.SelectedIndex = ComboBox_screwL_endsta.Items.IndexOf(Continut1)
                        Button_ScrewL_endsta.BackColor = Drawing.Color.Lime
                        ComboBox_screwL_no_type.Items.Add(Continut1)
                        ComboBox_screwL_spacing.Items.Add(Continut1)

                    End If
                    If Not Continut2 = "" Then
                        ComboBox_screwL_endsta.Items.Add(Continut2)
                        ComboBox_screwL_no_type.Items.Add(Continut2)
                        ComboBox_screwL_no_type.SelectedIndex = ComboBox_screwL_no_type.Items.IndexOf(Continut2)
                        Button_ScrewL_no_type.BackColor = Drawing.Color.Lime
                        ComboBox_screwL_spacing.Items.Add(Continut2)
                    End If
                    If Not Continut3 = "" Then
                        ComboBox_screwL_endsta.Items.Add(Continut3)
                        ComboBox_screwL_no_type.Items.Add(Continut3)
                        ComboBox_screwL_spacing.Items.Add(Continut3)
                        ComboBox_screwL_spacing.SelectedIndex = ComboBox_screwL_spacing.Items.IndexOf(Continut3)
                        Button_ScrewL_spacing.BackColor = Drawing.Color.Lime
                    End If






                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_ScrewL_insert_Click(sender As Object, e As EventArgs) Handles Button_ScrewL_insert.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select objects to be deleted:"

                    Object_Prompt1.SingleOnly = False
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If
                    Dim Point_start As Autodesk.AutoCAD.EditorInput.PromptPointResult

                    Dim PP_start As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select insertion point:")
                    PP_start.AllowNone = True
                    Point_start = Editor1.GetPoint(PP_start)
                    If Not Point_start.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        Editor1.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If

                    Dim Colectie_atr_name As New Specialized.StringCollection
                    Dim Colectie_atr_value As New Specialized.StringCollection

                    If Not ComboBox_screwL_endsta.Text = "" Then
                        Colectie_atr_name.Add("ENDSTA")
                        Colectie_atr_value.Add(ComboBox_screwL_endsta.Text)
                    End If

                    If Not ComboBox_screwL_no_type.Text = "" Then
                        Colectie_atr_name.Add("NO_TYPE")
                        Colectie_atr_value.Add(ComboBox_screwL_no_type.Text)
                    End If

                    If Not ComboBox_screwL_spacing.Text = "" Then
                        Colectie_atr_name.Add("SPACING")
                        Colectie_atr_value.Add(ComboBox_screwL_spacing.Text)
                    End If


                    InsertBlock_with_multiple_atributes("Screw_anchor_alignment_match_left1.dwg", "Screw_anchor_alignment_match_left1", Point_start.Value, 1, BTrecord, "TEXT", Colectie_atr_name, Colectie_atr_value)

                    For i = 0 To Rezultat1.Value.Count - 1






                        Dim ent1 As Entity = Rezultat1.Value.Item(i).ObjectId.GetObject(OpenMode.ForWrite)
                        ent1.Erase()
                    Next








                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)

                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_ScrewL_endsta_Click(sender As Object, e As EventArgs) Handles Button_ScrewL_endsta.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info {ENDSTA}:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut1 As String


                    If TypeOf ent1 Is DBText Then
                        Dim Text1 As DBText = ent1
                        Continut1 = Text1.TextString
                    End If

                    If TypeOf ent1 Is MText Then
                        Dim MText1 As MText = ent1
                        Continut1 = MText1.Contents
                    End If


                    If Not Continut1 = "" Then
                        If ComboBox_screwL_endsta.Items.Count > 1 Then
                            ComboBox_screwL_endsta.Items(1) = Continut1
                            ComboBox_screwL_endsta.SelectedIndex = ComboBox_screwL_endsta.Items.IndexOf(Continut1)
                            Button_ScrewL_endsta.BackColor = Drawing.Color.Lime
                            ComboBox_screwL_no_type.Items(1) = Continut1
                            ComboBox_screwL_spacing.Items(1) = Continut1
                        End If
                        If ComboBox_screwL_endsta.Items.Count = 0 Then
                            ComboBox_screwL_endsta.Items.Add(" ")
                            ComboBox_screwL_endsta.Items.Add(Continut1)
                            ComboBox_screwL_endsta.SelectedIndex = ComboBox_screwL_endsta.Items.IndexOf(Continut1)
                            Button_ScrewL_endsta.BackColor = Drawing.Color.Lime
                            ComboBox_screwL_no_type.Items.Add(" ")
                            ComboBox_screwL_no_type.Items.Add(Continut1)
                            ComboBox_screwL_spacing.Items.Add(" ")
                            ComboBox_screwL_spacing.Items.Add(Continut1)
                        End If

                    Else
                        If ComboBox_screwL_endsta.Items.Count = 0 Then
                            Button_ScrewL_endsta.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_ScrewL_no_type_Click(sender As Object, e As EventArgs) Handles Button_ScrewL_no_type.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt2.MessageForAdding = vbLf & "Select info {NO_TYPE}:"

                    Object_Prompt2.SingleOnly = True
                    Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                    If Not Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent2 As Entity = Rezultat2.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut2 As String


                    If TypeOf ent2 Is DBText Then
                        Dim Text2 As DBText = ent2
                        Continut2 = Text2.TextString
                    End If

                    If TypeOf ent2 Is MText Then
                        Dim MText2 As MText = ent2
                        Continut2 = MText2.Contents
                    End If


                    If Not Continut2 = "" Then
                        If ComboBox_screwL_no_type.Items.Count > 2 Then
                            ComboBox_screwL_endsta.Items(2) = Continut2
                            ComboBox_screwL_no_type.Items(2) = Continut2
                            ComboBox_screwL_no_type.SelectedIndex = ComboBox_screwL_no_type.Items.IndexOf(Continut2)
                            Button_ScrewL_no_type.BackColor = Drawing.Color.Lime
                            ComboBox_screwL_spacing.Items(2) = Continut2
                        End If
                        If ComboBox_screwL_no_type.Items.Count = 0 Then
                            ComboBox_screwL_endsta.Items.Add(" ")
                            ComboBox_screwL_endsta.Items.Add(Continut2)
                            ComboBox_screwL_no_type.Items.Add(" ")
                            ComboBox_screwL_no_type.Items.Add(Continut2)
                            ComboBox_screwL_no_type.SelectedIndex = ComboBox_screwL_no_type.Items.IndexOf(Continut2)
                            Button_ScrewL_no_type.BackColor = Drawing.Color.Lime
                            ComboBox_screwL_spacing.Items.Add(" ")
                            ComboBox_screwL_spacing.Items.Add(Continut2)
                        End If
                        If ComboBox_screwL_no_type.Items.Count = 1 Or ComboBox_screwL_no_type.Items.Count = 2 Then
                            ComboBox_screwL_endsta.Items.Add(Continut2)
                            ComboBox_screwL_no_type.Items.Add(Continut2)
                            ComboBox_screwL_no_type.SelectedIndex = ComboBox_screwL_no_type.Items.IndexOf(Continut2)
                            Button_ScrewL_no_type.BackColor = Drawing.Color.Lime

                            ComboBox_screwL_spacing.Items.Add(Continut2)
                        End If
                    Else
                        If ComboBox_screwL_no_type.Items.Count = 0 Then
                            Button_ScrewL_no_type.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_ScrewL_spacing_Click(sender As Object, e As EventArgs) Handles Button_ScrewL_spacing.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat3 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt3 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt3.MessageForAdding = vbLf & "Select info {SPACING}:"

                    Object_Prompt3.SingleOnly = True
                    Rezultat3 = Editor1.GetSelection(Object_Prompt3)
                    If Not Rezultat3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If
                    Dim ent3 As Entity = Rezultat3.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim Continut3 As String
                    If TypeOf ent3 Is DBText Then
                        Dim Text3 As DBText = ent3
                        Continut3 = Text3.TextString
                    End If

                    If TypeOf ent3 Is MText Then
                        Dim MText3 As MText = ent3
                        Continut3 = MText3.Contents
                    End If


                    If Not Continut3 = "" Then
                        If ComboBox_screwL_spacing.Items.Count > 3 Then
                            ComboBox_screwL_endsta.Items(3) = Continut3
                            ComboBox_screwL_no_type.Items(3) = Continut3
                            ComboBox_screwL_spacing.Items(3) = Continut3
                            ComboBox_screwL_spacing.SelectedIndex = ComboBox_screwL_spacing.Items.IndexOf(Continut3)
                            Button_ScrewL_spacing.BackColor = Drawing.Color.Lime
                        End If
                        If ComboBox_screwL_spacing.Items.Count = 0 Then
                            ComboBox_screwL_endsta.Items.Add(" ")
                            ComboBox_screwL_endsta.Items.Add(Continut3)
                            ComboBox_screwL_no_type.Items.Add(" ")
                            ComboBox_screwL_no_type.Items.Add(Continut3)

                            ComboBox_screwL_spacing.Items.Add(" ")
                            ComboBox_screwL_spacing.Items.Add(Continut3)
                            ComboBox_screwL_spacing.SelectedIndex = ComboBox_screwL_spacing.Items.IndexOf(Continut3)
                            Button_ScrewL_spacing.BackColor = Drawing.Color.Lime
                        End If
                        If ComboBox_screwL_spacing.Items.Count = 1 Or ComboBox_screwL_spacing.Items.Count = 2 Or ComboBox_screwL_spacing.Items.Count = 3 Then
                            ComboBox_screwL_endsta.Items.Add(Continut3)


                            ComboBox_screwL_no_type.Items.Add(Continut3)


                            ComboBox_screwL_spacing.Items.Add(Continut3)
                            ComboBox_screwL_spacing.SelectedIndex = ComboBox_screwL_spacing.Items.IndexOf(Continut3)
                            Button_ScrewL_spacing.BackColor = Drawing.Color.Lime
                        End If
                    Else
                        If ComboBox_screwL_spacing.Items.Count = 0 Then
                            Button_ScrewL_spacing.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub

    Private Sub Button_SCREWR_clear_Click(sender As Object, e As EventArgs) Handles Button_ScrewR_clear.Click
        ComboBox_screwR_BEGINsta.Items.Clear()
        ComboBox_screwR_no_type.Items.Clear()
        ComboBox_screwR_spacing.Items.Clear()
        ComboBox_screwR_BEGINsta.BackColor = Drawing.Color.DimGray
        ComboBox_screwR_no_type.BackColor = Drawing.Color.DimGray
        ComboBox_screwR_spacing.BackColor = Drawing.Color.DimGray
    End Sub
    Private Sub Button__SCREWR_Pick_Click(sender As Object, e As EventArgs) Handles Button__ScrewR_Pick.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info {ENDSTA}:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt2.MessageForAdding = vbLf & "Select info {NO_TYPE}:"

                    Object_Prompt2.SingleOnly = True
                    Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                    If Not Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    Dim Rezultat3 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt3 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt3.MessageForAdding = vbLf & "Select info {SPACING}:"

                    Object_Prompt3.SingleOnly = True
                    Rezultat3 = Editor1.GetSelection(Object_Prompt3)
                    If Not Rezultat3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If


                    ComboBox_screwR_BEGINsta.Items.Clear()
                    ComboBox_screwR_BEGINsta.Items.Add(" ")
                    ComboBox_screwR_no_type.Items.Clear()
                    ComboBox_screwR_no_type.Items.Add(" ")
                    ComboBox_screwR_spacing.Items.Clear()
                    ComboBox_screwR_spacing.Items.Add(" ")


                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent2 As Entity = Rezultat2.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent3 As Entity = Rezultat3.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)

                    Dim Continut1 As String
                    Dim Continut2 As String
                    Dim Continut3 As String


                    If TypeOf ent1 Is DBText Then
                        Dim Text1 As DBText = ent1
                        Continut1 = Text1.TextString
                    End If
                    If TypeOf ent2 Is DBText Then
                        Dim Text2 As DBText = ent2
                        Continut2 = Text2.TextString
                    End If
                    If TypeOf ent3 Is DBText Then
                        Dim Text3 As DBText = ent3
                        Continut3 = Text3.TextString
                    End If

                    If TypeOf ent1 Is MText Then
                        Dim MText1 As MText = ent1
                        Continut1 = MText1.Contents
                    End If
                    If TypeOf ent2 Is MText Then
                        Dim MText2 As MText = ent2
                        Continut2 = MText2.Contents
                    End If
                    If TypeOf ent3 Is MText Then
                        Dim MText3 As MText = ent3
                        Continut3 = MText3.Contents
                    End If


                    If TypeOf ent1 Is BlockReference Then
                        Dim Block1 As BlockReference = ent1
                        If Block1.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                            Dim ctemp1 As String
                            Dim ctemp2 As String
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "BEGINSTA" Then
                                    ctemp1 = attref.TextString
                                End If
                                If attref.Tag = "ENDSTA" Then
                                    ctemp2 = attref.TextString
                                End If
                            Next
                            ctemp1 = Replace(ctemp1, " ", "")
                            ctemp2 = Replace(ctemp2, " ", "")
                            If TypeOf ent2 Is BlockReference Then
                                If Not ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut1 = ctemp2
                                End If
                                If ctemp1 = "" And ctemp2 = "" Then
                                    Continut1 = " "
                                End If
                                If ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut1 = ctemp2
                                End If
                                If Not ctemp1 = "" And ctemp2 = "" Then
                                    Continut1 = ctemp1
                                End If
                            Else
                                Continut1 = ctemp1
                            End If
                        End If
                    End If


                    If TypeOf ent2 Is BlockReference Then
                        Dim Block2 As BlockReference = ent2
                        If Block2.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block2.AttributeCollection
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "NO_TYPE" Then
                                    Continut2 = attref.TextString
                                End If
                            Next
                        End If
                    End If


                    If TypeOf ent3 Is BlockReference Then
                        Dim Block3 As BlockReference = ent3
                        If Block3.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block3.AttributeCollection
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "SPACING" Then
                                    Continut3 = attref.TextString
                                End If
                            Next
                        End If
                    End If

                    If Not Continut1 = "" Then
                        ComboBox_screwR_BEGINsta.Items.Add(Continut1)
                        ComboBox_screwR_BEGINsta.SelectedIndex = ComboBox_screwR_BEGINsta.Items.IndexOf(Continut1)
                        Button_ScrewR_beginsta.BackColor = Drawing.Color.Lime
                        ComboBox_screwR_no_type.Items.Add(Continut1)
                        ComboBox_screwR_spacing.Items.Add(Continut1)

                    End If
                    If Not Continut2 = "" Then
                        ComboBox_screwR_BEGINsta.Items.Add(Continut2)
                        ComboBox_screwR_no_type.Items.Add(Continut2)
                        ComboBox_screwR_no_type.SelectedIndex = ComboBox_screwR_no_type.Items.IndexOf(Continut2)
                        Button_ScrewR_no_type.BackColor = Drawing.Color.Lime
                        ComboBox_screwR_spacing.Items.Add(Continut2)
                    End If
                    If Not Continut3 = "" Then
                        ComboBox_screwR_BEGINsta.Items.Add(Continut3)
                        ComboBox_screwR_no_type.Items.Add(Continut3)
                        ComboBox_screwR_spacing.Items.Add(Continut3)
                        ComboBox_screwR_spacing.SelectedIndex = ComboBox_screwR_spacing.Items.IndexOf(Continut3)
                        Button_ScrewR_spacing.BackColor = Drawing.Color.Lime
                    End If






                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_SCREWR_insert_Click(sender As Object, e As EventArgs) Handles Button_ScrewR_insert.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select objects to be deleted:"

                    Object_Prompt1.SingleOnly = False
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If
                    Dim Point_start As Autodesk.AutoCAD.EditorInput.PromptPointResult

                    Dim PP_start As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select insertion point:")
                    PP_start.AllowNone = True
                    Point_start = Editor1.GetPoint(PP_start)
                    If Not Point_start.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        Editor1.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If

                    Dim Colectie_atr_name As New Specialized.StringCollection
                    Dim Colectie_atr_value As New Specialized.StringCollection

                    If Not ComboBox_screwR_BEGINsta.Text = "" Then
                        Colectie_atr_name.Add("BEGINSTA")
                        Colectie_atr_value.Add(ComboBox_screwR_BEGINsta.Text)
                    End If

                    If Not ComboBox_screwR_no_type.Text = "" Then
                        Colectie_atr_name.Add("NO_TYPE")
                        Colectie_atr_value.Add(ComboBox_screwR_no_type.Text)
                    End If

                    If Not ComboBox_screwR_spacing.Text = "" Then
                        Colectie_atr_name.Add("SPACING")
                        Colectie_atr_value.Add(ComboBox_screwR_spacing.Text)
                    End If


                    InsertBlock_with_multiple_atributes("Screw_anchor_alignment_match_right1.dwg", "Screw_anchor_alignment_match_right1", Point_start.Value, 1, BTrecord, "TEXT", Colectie_atr_name, Colectie_atr_value)

                    For i = 0 To Rezultat1.Value.Count - 1






                        Dim ent1 As Entity = Rezultat1.Value.Item(i).ObjectId.GetObject(OpenMode.ForWrite)
                        ent1.Erase()
                    Next








                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)

                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_SCREWR_endsta_Click(sender As Object, e As EventArgs) Handles Button_ScrewR_beginsta.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info {ENDSTA}:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut1 As String


                    If TypeOf ent1 Is DBText Then
                        Dim Text1 As DBText = ent1
                        Continut1 = Text1.TextString
                    End If

                    If TypeOf ent1 Is MText Then
                        Dim MText1 As MText = ent1
                        Continut1 = MText1.Contents
                    End If


                    If Not Continut1 = "" Then
                        If ComboBox_screwR_BEGINsta.Items.Count > 1 Then
                            ComboBox_screwR_BEGINsta.Items(1) = Continut1
                            ComboBox_screwR_BEGINsta.SelectedIndex = ComboBox_screwR_BEGINsta.Items.IndexOf(Continut1)
                            Button_ScrewR_beginsta.BackColor = Drawing.Color.Lime
                            ComboBox_screwR_no_type.Items(1) = Continut1
                            ComboBox_screwR_spacing.Items(1) = Continut1
                        End If
                        If ComboBox_screwR_BEGINsta.Items.Count = 0 Then
                            ComboBox_screwR_BEGINsta.Items.Add(" ")
                            ComboBox_screwR_BEGINsta.Items.Add(Continut1)
                            ComboBox_screwR_BEGINsta.SelectedIndex = ComboBox_screwR_BEGINsta.Items.IndexOf(Continut1)
                            Button_ScrewR_beginsta.BackColor = Drawing.Color.Lime
                            ComboBox_screwR_no_type.Items.Add(" ")
                            ComboBox_screwR_no_type.Items.Add(Continut1)
                            ComboBox_screwR_spacing.Items.Add(" ")
                            ComboBox_screwR_spacing.Items.Add(Continut1)
                        End If

                    Else
                        If ComboBox_screwR_BEGINsta.Items.Count = 0 Then
                            Button_ScrewR_beginsta.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_SCREWR_no_type_Click(sender As Object, e As EventArgs) Handles Button_ScrewR_no_type.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt2.MessageForAdding = vbLf & "Select info {NO_TYPE}:"

                    Object_Prompt2.SingleOnly = True
                    Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                    If Not Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent2 As Entity = Rezultat2.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut2 As String


                    If TypeOf ent2 Is DBText Then
                        Dim Text2 As DBText = ent2
                        Continut2 = Text2.TextString
                    End If

                    If TypeOf ent2 Is MText Then
                        Dim MText2 As MText = ent2
                        Continut2 = MText2.Contents
                    End If


                    If Not Continut2 = "" Then
                        If ComboBox_screwR_no_type.Items.Count > 2 Then
                            ComboBox_screwR_BEGINsta.Items(2) = Continut2
                            ComboBox_screwR_no_type.Items(2) = Continut2
                            ComboBox_screwR_no_type.SelectedIndex = ComboBox_screwR_no_type.Items.IndexOf(Continut2)
                            Button_ScrewR_no_type.BackColor = Drawing.Color.Lime
                            ComboBox_screwR_spacing.Items(2) = Continut2
                        End If
                        If ComboBox_screwR_no_type.Items.Count = 0 Then
                            ComboBox_screwR_BEGINsta.Items.Add(" ")
                            ComboBox_screwR_BEGINsta.Items.Add(Continut2)
                            ComboBox_screwR_no_type.Items.Add(" ")
                            ComboBox_screwR_no_type.Items.Add(Continut2)
                            ComboBox_screwR_no_type.SelectedIndex = ComboBox_screwR_no_type.Items.IndexOf(Continut2)
                            Button_ScrewR_no_type.BackColor = Drawing.Color.Lime
                            ComboBox_screwR_spacing.Items.Add(" ")
                            ComboBox_screwR_spacing.Items.Add(Continut2)
                        End If
                        If ComboBox_screwR_no_type.Items.Count = 1 Or ComboBox_screwR_no_type.Items.Count = 2 Then
                            ComboBox_screwR_BEGINsta.Items.Add(Continut2)
                            ComboBox_screwR_no_type.Items.Add(Continut2)
                            ComboBox_screwR_no_type.SelectedIndex = ComboBox_screwR_no_type.Items.IndexOf(Continut2)
                            Button_ScrewR_no_type.BackColor = Drawing.Color.Lime

                            ComboBox_screwR_spacing.Items.Add(Continut2)
                        End If
                    Else
                        If ComboBox_screwR_no_type.Items.Count = 0 Then
                            Button_ScrewR_no_type.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_SCREWR_spacing_Click(sender As Object, e As EventArgs) Handles Button_ScrewR_spacing.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat3 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt3 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt3.MessageForAdding = vbLf & "Select info {SPACING}:"

                    Object_Prompt3.SingleOnly = True
                    Rezultat3 = Editor1.GetSelection(Object_Prompt3)
                    If Not Rezultat3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If
                    Dim ent3 As Entity = Rezultat3.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim Continut3 As String
                    If TypeOf ent3 Is DBText Then
                        Dim Text3 As DBText = ent3
                        Continut3 = Text3.TextString
                    End If

                    If TypeOf ent3 Is MText Then
                        Dim MText3 As MText = ent3
                        Continut3 = MText3.Contents
                    End If


                    If Not Continut3 = "" Then
                        If ComboBox_screwR_spacing.Items.Count > 3 Then
                            ComboBox_screwR_BEGINsta.Items(3) = Continut3
                            ComboBox_screwR_no_type.Items(3) = Continut3
                            ComboBox_screwR_spacing.Items(3) = Continut3
                            ComboBox_screwR_spacing.SelectedIndex = ComboBox_screwR_spacing.Items.IndexOf(Continut3)
                            Button_ScrewR_spacing.BackColor = Drawing.Color.Lime
                        End If
                        If ComboBox_screwR_spacing.Items.Count = 0 Then
                            ComboBox_screwR_BEGINsta.Items.Add(" ")
                            ComboBox_screwR_BEGINsta.Items.Add(Continut3)
                            ComboBox_screwR_no_type.Items.Add(" ")
                            ComboBox_screwR_no_type.Items.Add(Continut3)

                            ComboBox_screwR_spacing.Items.Add(" ")
                            ComboBox_screwR_spacing.Items.Add(Continut3)
                            ComboBox_screwR_spacing.SelectedIndex = ComboBox_screwR_spacing.Items.IndexOf(Continut3)
                            Button_ScrewR_spacing.BackColor = Drawing.Color.Lime
                        End If
                        If ComboBox_screwR_spacing.Items.Count = 1 Or ComboBox_screwR_spacing.Items.Count = 2 Or ComboBox_screwR_spacing.Items.Count = 3 Then
                            ComboBox_screwR_BEGINsta.Items.Add(Continut3)


                            ComboBox_screwR_no_type.Items.Add(Continut3)


                            ComboBox_screwR_spacing.Items.Add(Continut3)
                            ComboBox_screwR_spacing.SelectedIndex = ComboBox_screwR_spacing.Items.IndexOf(Continut3)
                            Button_ScrewR_spacing.BackColor = Drawing.Color.Lime
                        End If
                    Else
                        If ComboBox_screwR_spacing.Items.Count = 0 Then
                            Button_ScrewR_spacing.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub

    Private Sub Button_pipemLCLEAR_Click(sender As Object, e As EventArgs) Handles Button_PIPEML_CLEAR.Click
        ComboBox_PIPEML_MAT.Items.Clear()
        ComboBox_PIPEML_ENDSTA.Items.Clear()
        ComboBox_PIPEML_LENGTH.Items.Clear()
        Button_PIPEML_ENDSTA.BackColor = Drawing.Color.DimGray
        Button_PIPEML_MAT.BackColor = Drawing.Color.DimGray
        Button_PIPEML_LENGTH.BackColor = Drawing.Color.DimGray
    End Sub
    Private Sub Button_pipemLPICK_Click(sender As Object, e As EventArgs) Handles Button_PIPEML_PICK.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info {ENDSTA}:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt2.MessageForAdding = vbLf & "Select info {LENGTH}:"

                    Object_Prompt2.SingleOnly = True
                    Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                    If Not Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    Dim Rezultat3 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt3 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt3.MessageForAdding = vbLf & "Select info {MAT}:"

                    Object_Prompt3.SingleOnly = True
                    Rezultat3 = Editor1.GetSelection(Object_Prompt3)
                    If Not Rezultat3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If


                    ComboBox_PIPEML_ENDSTA.Items.Clear()
                    ComboBox_PIPEML_ENDSTA.Items.Add(" ")
                    ComboBox_PIPEML_LENGTH.Items.Clear()
                    ComboBox_PIPEML_LENGTH.Items.Add(" ")
                    ComboBox_PIPEML_MAT.Items.Clear()
                    ComboBox_PIPEML_MAT.Items.Add(" ")


                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent2 As Entity = Rezultat2.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent3 As Entity = Rezultat3.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)

                    Dim Continut1 As String
                    Dim Continut2 As String
                    Dim Continut3 As String


                    If TypeOf ent1 Is DBText Then
                        Dim Text1 As DBText = ent1
                        Continut1 = Text1.TextString
                    End If
                    If TypeOf ent2 Is DBText Then
                        Dim Text2 As DBText = ent2
                        Continut2 = Text2.TextString
                    End If
                    If TypeOf ent3 Is DBText Then
                        Dim Text3 As DBText = ent3
                        Continut3 = Text3.TextString
                    End If

                    If TypeOf ent1 Is MText Then
                        Dim MText1 As MText = ent1
                        Continut1 = MText1.Contents
                    End If
                    If TypeOf ent2 Is MText Then
                        Dim MText2 As MText = ent2
                        Continut2 = MText2.Contents
                    End If
                    If TypeOf ent3 Is MText Then
                        Dim MText3 As MText = ent3
                        Continut3 = MText3.Contents
                    End If


                    If TypeOf ent1 Is BlockReference Then
                        Dim Block1 As BlockReference = ent1
                        If Block1.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                            Dim ctemp1 As String
                            Dim ctemp2 As String
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "BEGINSTA" Then
                                    ctemp1 = attref.TextString
                                End If
                                If attref.Tag = "ENDSTA" Then
                                    ctemp2 = attref.TextString
                                End If
                            Next
                            ctemp1 = Replace(ctemp1, " ", "")
                            ctemp2 = Replace(ctemp2, " ", "")



                            If Continut2 = "" Then
                                If TypeOf ent2 Is BlockReference And Not ent1.ObjectId = ent2.ObjectId Then
                                    If Not ctemp1 = "" And Not ctemp2 = "" Then
                                        Continut1 = ctemp2
                                    End If
                                    If ctemp1 = "" And ctemp2 = "" Then
                                        Continut1 = " "
                                    End If
                                    If ctemp1 = "" And Not ctemp2 = "" Then
                                        Continut1 = ctemp2
                                    End If
                                    If Not ctemp1 = "" And ctemp2 = "" Then
                                        Continut1 = ctemp1
                                    End If
                                End If
                            Else
                                If Not ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut1 = ctemp2
                                End If
                                If ctemp1 = "" And ctemp2 = "" Then
                                    Continut1 = " "
                                End If
                                If ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut1 = ctemp2
                                End If
                                If Not ctemp1 = "" And ctemp2 = "" Then
                                    Continut1 = ctemp1
                                End If
                            End If
                            If ent1.ObjectId = ent2.ObjectId Then
                                Continut1 = ctemp1
                            End If

                        End If
                    End If


                    If TypeOf ent2 Is BlockReference Then
                        Dim Block2 As BlockReference = ent2
                        If Block2.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block2.AttributeCollection
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "LENGTH" Then
                                    Continut2 = attref.TextString
                                End If
                            Next
                        End If
                    End If


                    If TypeOf ent3 Is BlockReference Then
                        Dim Block3 As BlockReference = ent3
                        If Block3.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block3.AttributeCollection
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "MAT" Then
                                    Continut3 = attref.TextString
                                End If
                            Next
                        End If
                    End If

                    If Not Continut1 = "" Then
                        ComboBox_PIPEML_ENDSTA.Items.Add(Continut1)
                        ComboBox_PIPEML_ENDSTA.SelectedIndex = ComboBox_PIPEML_ENDSTA.Items.IndexOf(Continut1)
                        Button_PIPEML_ENDSTA.BackColor = Drawing.Color.Lime
                        ComboBox_PIPEML_LENGTH.Items.Add(Continut1)
                        ComboBox_PIPEML_MAT.Items.Add(Continut1)

                    End If
                    If Not Continut2 = "" Then
                        ComboBox_PIPEML_ENDSTA.Items.Add(Continut2)
                        ComboBox_PIPEML_LENGTH.Items.Add(Continut2)
                        ComboBox_PIPEML_LENGTH.SelectedIndex = ComboBox_PIPEML_LENGTH.Items.IndexOf(Continut2)
                        Button_PIPEML_LENGTH.BackColor = Drawing.Color.Lime
                        ComboBox_PIPEML_MAT.Items.Add(Continut2)
                    End If
                    If Not Continut3 = "" Then
                        ComboBox_PIPEML_ENDSTA.Items.Add(Continut3)
                        ComboBox_PIPEML_LENGTH.Items.Add(Continut3)
                        ComboBox_PIPEML_MAT.Items.Add(Continut3)
                        ComboBox_PIPEML_MAT.SelectedIndex = ComboBox_PIPEML_MAT.Items.IndexOf(Continut3)
                        Button_PIPEML_MAT.BackColor = Drawing.Color.Lime
                    End If






                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_pipemLINSERT_Click(sender As Object, e As EventArgs) Handles Button_PIPEML_INSERT.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select objects to be deleted:"

                    Object_Prompt1.SingleOnly = False
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If
                    Dim Point_start As Autodesk.AutoCAD.EditorInput.PromptPointResult

                    Dim PP_start As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select insertion point:")
                    PP_start.AllowNone = True
                    Point_start = Editor1.GetPoint(PP_start)
                    If Not Point_start.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        Editor1.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If

                    Dim Colectie_atr_name As New Specialized.StringCollection
                    Dim Colectie_atr_value As New Specialized.StringCollection

                    If Not ComboBox_PIPEML_ENDSTA.Text = "" Then
                        Colectie_atr_name.Add("ENDSTA")
                        Colectie_atr_value.Add(ComboBox_PIPEML_ENDSTA.Text)
                    End If

                    If Not ComboBox_PIPEML_LENGTH.Text = "" Then
                        Colectie_atr_name.Add("LENGTH")
                        Colectie_atr_value.Add(ComboBox_PIPEML_LENGTH.Text)
                    End If

                    If Not ComboBox_PIPEML_MAT.Text = "" Then
                        Colectie_atr_name.Add("MAT")
                        Colectie_atr_value.Add(ComboBox_PIPEML_MAT.Text)
                    End If


                    InsertBlock_with_multiple_atributes("heavy_wall_match_left_1.dwg", "heavy_wall_match_left_1", Point_start.Value, 1, BTrecord, "TEXT", Colectie_atr_name, Colectie_atr_value)

                    For i = 0 To Rezultat1.Value.Count - 1






                        Dim ent1 As Entity = Rezultat1.Value.Item(i).ObjectId.GetObject(OpenMode.ForWrite)
                        ent1.Erase()
                    Next








                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)

                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_pipemLENDSTA_Click(sender As Object, e As EventArgs) Handles Button_PIPEML_ENDSTA.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info {ENDSTA}:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut1 As String


                    If TypeOf ent1 Is DBText Then
                        Dim Text1 As DBText = ent1
                        Continut1 = Text1.TextString
                    End If

                    If TypeOf ent1 Is MText Then
                        Dim MText1 As MText = ent1
                        Continut1 = MText1.Contents
                    End If


                    If Not Continut1 = "" Then
                        If ComboBox_PIPEML_ENDSTA.Items.Count > 1 Then
                            ComboBox_PIPEML_ENDSTA.Items(1) = Continut1
                            ComboBox_PIPEML_ENDSTA.SelectedIndex = ComboBox_PIPEML_ENDSTA.Items.IndexOf(Continut1)
                            Button_PIPEML_ENDSTA.BackColor = Drawing.Color.Lime
                            ComboBox_PIPEML_LENGTH.Items(1) = Continut1
                            ComboBox_PIPEML_MAT.Items(1) = Continut1
                        End If
                        If ComboBox_PIPEML_ENDSTA.Items.Count = 0 Then
                            ComboBox_PIPEML_ENDSTA.Items.Add(" ")
                            ComboBox_PIPEML_ENDSTA.Items.Add(Continut1)
                            ComboBox_PIPEML_ENDSTA.SelectedIndex = ComboBox_PIPEML_ENDSTA.Items.IndexOf(Continut1)
                            Button_PIPEML_ENDSTA.BackColor = Drawing.Color.Lime
                            ComboBox_PIPEML_LENGTH.Items.Add(" ")
                            ComboBox_PIPEML_LENGTH.Items.Add(Continut1)
                            ComboBox_PIPEML_MAT.Items.Add(" ")
                            ComboBox_PIPEML_MAT.Items.Add(Continut1)
                        End If

                    Else
                        If ComboBox_PIPEML_ENDSTA.Items.Count = 0 Then
                            Button_PIPEML_ENDSTA.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_pipemLLENGTH_Click(sender As Object, e As EventArgs) Handles Button_PIPEML_LENGTH.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt2.MessageForAdding = vbLf & "Select info {LENGTH}:"

                    Object_Prompt2.SingleOnly = True
                    Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                    If Not Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent2 As Entity = Rezultat2.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut2 As String


                    If TypeOf ent2 Is DBText Then
                        Dim Text2 As DBText = ent2
                        Continut2 = Text2.TextString
                    End If

                    If TypeOf ent2 Is MText Then
                        Dim MText2 As MText = ent2
                        Continut2 = MText2.Contents
                    End If


                    If Not Continut2 = "" Then
                        If ComboBox_PIPEML_LENGTH.Items.Count > 2 Then
                            ComboBox_PIPEML_ENDSTA.Items(2) = Continut2
                            ComboBox_PIPEML_LENGTH.Items(2) = Continut2
                            ComboBox_PIPEML_LENGTH.SelectedIndex = ComboBox_PIPEML_LENGTH.Items.IndexOf(Continut2)
                            Button_PIPEML_LENGTH.BackColor = Drawing.Color.Lime
                            ComboBox_PIPEML_MAT.Items(2) = Continut2
                        End If
                        If ComboBox_PIPEML_LENGTH.Items.Count = 0 Then
                            ComboBox_PIPEML_ENDSTA.Items.Add(" ")
                            ComboBox_PIPEML_ENDSTA.Items.Add(Continut2)
                            ComboBox_PIPEML_LENGTH.Items.Add(" ")
                            ComboBox_PIPEML_LENGTH.Items.Add(Continut2)
                            ComboBox_PIPEML_LENGTH.SelectedIndex = ComboBox_PIPEML_LENGTH.Items.IndexOf(Continut2)
                            Button_PIPEML_LENGTH.BackColor = Drawing.Color.Lime
                            ComboBox_PIPEML_MAT.Items.Add(" ")
                            ComboBox_PIPEML_MAT.Items.Add(Continut2)
                        End If
                        If ComboBox_PIPEML_LENGTH.Items.Count = 1 Or ComboBox_PIPEML_LENGTH.Items.Count = 2 Then
                            ComboBox_PIPEML_ENDSTA.Items.Add(Continut2)
                            ComboBox_PIPEML_LENGTH.Items.Add(Continut2)
                            ComboBox_PIPEML_LENGTH.SelectedIndex = ComboBox_PIPEML_LENGTH.Items.IndexOf(Continut2)
                            Button_PIPEML_LENGTH.BackColor = Drawing.Color.Lime

                            ComboBox_PIPEML_MAT.Items.Add(Continut2)
                        End If
                    Else
                        If ComboBox_PIPEML_LENGTH.Items.Count = 0 Then
                            Button_PIPEML_LENGTH.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_pipemLMAT_Click(sender As Object, e As EventArgs) Handles Button_PIPEML_MAT.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat3 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt3 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt3.MessageForAdding = vbLf & "Select info {MAT}:"

                    Object_Prompt3.SingleOnly = True
                    Rezultat3 = Editor1.GetSelection(Object_Prompt3)
                    If Not Rezultat3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If
                    Dim ent3 As Entity = Rezultat3.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim Continut3 As String
                    If TypeOf ent3 Is DBText Then
                        Dim Text3 As DBText = ent3
                        Continut3 = Text3.TextString
                    End If

                    If TypeOf ent3 Is MText Then
                        Dim MText3 As MText = ent3
                        Continut3 = MText3.Contents
                    End If


                    If Not Continut3 = "" Then
                        If ComboBox_PIPEML_MAT.Items.Count > 3 Then
                            ComboBox_PIPEML_ENDSTA.Items(3) = Continut3
                            ComboBox_PIPEML_LENGTH.Items(3) = Continut3
                            ComboBox_PIPEML_MAT.Items(3) = Continut3
                            ComboBox_PIPEML_MAT.SelectedIndex = ComboBox_PIPEML_MAT.Items.IndexOf(Continut3)
                            Button_PIPEML_MAT.BackColor = Drawing.Color.Lime
                        End If
                        If ComboBox_PIPEML_MAT.Items.Count = 0 Then
                            ComboBox_PIPEML_ENDSTA.Items.Add(" ")
                            ComboBox_PIPEML_ENDSTA.Items.Add(Continut3)
                            ComboBox_PIPEML_LENGTH.Items.Add(" ")
                            ComboBox_PIPEML_LENGTH.Items.Add(Continut3)

                            ComboBox_PIPEML_MAT.Items.Add(" ")
                            ComboBox_PIPEML_MAT.Items.Add(Continut3)
                            ComboBox_PIPEML_MAT.SelectedIndex = ComboBox_PIPEML_MAT.Items.IndexOf(Continut3)
                            Button_PIPEML_MAT.BackColor = Drawing.Color.Lime
                        End If
                        If ComboBox_PIPEML_MAT.Items.Count = 1 Or ComboBox_PIPEML_MAT.Items.Count = 2 Or ComboBox_PIPEML_MAT.Items.Count = 3 Then
                            ComboBox_PIPEML_ENDSTA.Items.Add(Continut3)


                            ComboBox_PIPEML_LENGTH.Items.Add(Continut3)


                            ComboBox_PIPEML_MAT.Items.Add(Continut3)
                            ComboBox_PIPEML_MAT.SelectedIndex = ComboBox_PIPEML_MAT.Items.IndexOf(Continut3)
                            Button_PIPEML_MAT.BackColor = Drawing.Color.Lime
                        End If
                    Else
                        If ComboBox_PIPEML_MAT.Items.Count = 0 Then
                            Button_PIPEML_MAT.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_pipeML_blockpick_Click(sender As Object, e As EventArgs) Handles Button_pipeML_blockpick.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info from Block:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    ComboBox_PIPEML_ENDSTA.Items.Clear()
                    ComboBox_PIPEML_ENDSTA.Items.Add(" ")
                    ComboBox_PIPEML_LENGTH.Items.Clear()
                    ComboBox_PIPEML_LENGTH.Items.Add(" ")
                    ComboBox_PIPEML_MAT.Items.Clear()
                    ComboBox_PIPEML_MAT.Items.Add(" ")

                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut1 As String
                    Dim Continut2 As String
                    Dim Continut3 As String



                    If TypeOf ent1 Is BlockReference Then
                        Dim Block1 As BlockReference = ent1
                        If Block1.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "ENDSTA" Then
                                    Continut1 = attref.TextString
                                    If Continut1 = "" Then Continut1 = " "
                                End If
                                If attref.Tag = "LENGTH" Then
                                    Continut2 = attref.TextString
                                    If Continut2 = "" Then Continut2 = " "
                                End If
                                If attref.Tag = "MAT" Then
                                    Continut3 = attref.TextString
                                    If Continut3 = "" Then Continut3 = " "
                                End If
                            Next
                        End If
                    End If

                    If Not Continut1 = "" Then
                        ComboBox_PIPEML_ENDSTA.Items.Add(Continut1)
                        ComboBox_PIPEML_ENDSTA.SelectedIndex = ComboBox_PIPEML_ENDSTA.Items.IndexOf(Continut1)
                        Button_PIPEML_ENDSTA.BackColor = Drawing.Color.Lime
                        ComboBox_PIPEML_LENGTH.Items.Add(Continut1)
                        ComboBox_PIPEML_MAT.Items.Add(Continut1)
                    End If
                    If Not Continut2 = "" Then
                        ComboBox_PIPEML_ENDSTA.Items.Add(Continut2)
                        ComboBox_PIPEML_LENGTH.Items.Add(Continut2)
                        ComboBox_PIPEML_LENGTH.SelectedIndex = ComboBox_PIPEML_LENGTH.Items.IndexOf(Continut2)
                        Button_PIPEML_LENGTH.BackColor = Drawing.Color.Lime
                        ComboBox_PIPEML_MAT.Items.Add(Continut2)
                    End If
                    If Not Continut3 = "" Then
                        ComboBox_PIPEML_ENDSTA.Items.Add(Continut3)
                        ComboBox_PIPEML_LENGTH.Items.Add(Continut3)
                        ComboBox_PIPEML_MAT.Items.Add(Continut3)
                        ComboBox_PIPEML_MAT.SelectedIndex = ComboBox_PIPEML_MAT.Items.IndexOf(Continut3)
                        Button_PIPEML_MAT.BackColor = Drawing.Color.Lime

                    End If

                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub

    Private Sub Button_SAND_clear_Click(sender As Object, e As EventArgs) Handles Button_sand_clear.Click
        ComboBox_SAND_spacing.Items.Clear()
        ComboBox_SAND_beginsta.Items.Clear()
        ComboBox_SAND_endsta.Items.Clear()
        ComboBox_SAND_no_type.Items.Clear()
        Button_SAND_Beginsta.BackColor = Drawing.Color.DimGray
        Button_SAND_Endsta.BackColor = Drawing.Color.DimGray
        Button_SAND_spacing.BackColor = Drawing.Color.DimGray
        Button_SAND_no_type.BackColor = Drawing.Color.DimGray
    End Sub
    Private Sub Button_SAND_pick_Click(sender As Object, e As EventArgs) Handles Button_sand_pick.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info {BEGINSTA}:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt2.MessageForAdding = vbLf & "Select info {ENDSTA}:"

                    Object_Prompt2.SingleOnly = True
                    Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                    If Not Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    Dim Rezultat3 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt3 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt3.MessageForAdding = vbLf & "Select info {NO_TYPE}:"

                    Object_Prompt3.SingleOnly = True
                    Rezultat3 = Editor1.GetSelection(Object_Prompt3)
                    If Not Rezultat3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    Dim Rezultat4 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt4 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt4.MessageForAdding = vbLf & "Select info {SPACING}:"

                    Object_Prompt4.SingleOnly = True
                    Rezultat4 = Editor1.GetSelection(Object_Prompt4)
                    If Not Rezultat4.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    ComboBox_SAND_beginsta.Items.Clear()
                    ComboBox_SAND_beginsta.Items.Add(" ")
                    ComboBox_SAND_endsta.Items.Clear()
                    ComboBox_SAND_endsta.Items.Add(" ")
                    ComboBox_SAND_no_type.Items.Clear()
                    ComboBox_SAND_no_type.Items.Add(" ")
                    ComboBox_SAND_spacing.Items.Clear()
                    ComboBox_SAND_spacing.Items.Add(" ")


                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent2 As Entity = Rezultat2.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent3 As Entity = Rezultat3.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent4 As Entity = Rezultat4.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)

                    Dim Continut1 As String
                    Dim Continut2 As String
                    Dim Continut3 As String
                    Dim Continut4 As String

                    If TypeOf ent1 Is DBText Then
                        Dim Text1 As DBText = ent1
                        Continut1 = Text1.TextString
                    End If
                    If TypeOf ent2 Is DBText Then
                        Dim Text2 As DBText = ent2
                        Continut2 = Text2.TextString
                    End If
                    If TypeOf ent3 Is DBText Then
                        Dim Text3 As DBText = ent3
                        Continut3 = Text3.TextString
                    End If
                    If TypeOf ent4 Is DBText Then
                        Dim Text4 As DBText = ent4
                        Continut4 = Text4.TextString
                    End If

                    If TypeOf ent1 Is MText Then
                        Dim MText1 As MText = ent1
                        Continut1 = MText1.Contents
                    End If
                    If TypeOf ent2 Is MText Then
                        Dim MText2 As MText = ent2
                        Continut2 = MText2.Contents
                    End If
                    If TypeOf ent3 Is MText Then
                        Dim MText3 As MText = ent3
                        Continut3 = MText3.Contents
                    End If
                    If TypeOf ent4 Is MText Then
                        Dim MText4 As MText = ent4
                        Continut4 = MText4.Contents
                    End If


                    If TypeOf ent1 Is BlockReference Then
                        Dim Block1 As BlockReference = ent1
                        If Block1.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                            Dim ctemp1 As String
                            Dim ctemp2 As String
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "BEGINSTA" Then
                                    ctemp1 = attref.TextString
                                End If
                                If attref.Tag = "ENDSTA" Then
                                    ctemp2 = attref.TextString
                                End If
                            Next
                            ctemp1 = Replace(ctemp1, " ", "")
                            ctemp2 = Replace(ctemp2, " ", "")
                            If TypeOf ent2 Is BlockReference Then
                                If Not ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut1 = ctemp2
                                End If
                                If ctemp1 = "" And ctemp2 = "" Then
                                    Continut1 = " "
                                End If
                                If ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut1 = ctemp2
                                End If
                                If Not ctemp1 = "" And ctemp2 = "" Then
                                    Continut1 = ctemp1
                                End If
                            Else
                                Continut1 = ctemp1
                            End If
                        End If
                    End If


                    If TypeOf ent2 Is BlockReference Then
                        Dim Block2 As BlockReference = ent2
                        If Block2.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block2.AttributeCollection
                            Dim ctemp1 As String
                            Dim ctemp2 As String
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "BEGINSTA" Then
                                    ctemp1 = attref.TextString
                                End If
                                If attref.Tag = "ENDSTA" Then
                                    ctemp2 = attref.TextString
                                End If
                            Next
                            ctemp1 = Replace(ctemp1, " ", "")
                            ctemp2 = Replace(ctemp2, " ", "")
                            If TypeOf ent1 Is BlockReference Then
                                If Not ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut2 = ctemp1
                                End If
                                If ctemp1 = "" And ctemp2 = "" Then
                                    Continut2 = " "
                                End If
                                If ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut2 = ctemp2
                                End If
                                If Not ctemp1 = "" And ctemp2 = "" Then
                                    Continut2 = ctemp1
                                End If
                            Else
                                Continut2 = ctemp2
                            End If

                        End If
                    End If


                    If TypeOf ent3 Is BlockReference Then
                        Dim Block3 As BlockReference = ent3
                        If Block3.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block3.AttributeCollection
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "NO_TYPE" Then
                                    Continut3 = attref.TextString
                                End If
                            Next
                        End If
                    End If


                    If TypeOf ent4 Is BlockReference Then
                        Dim Block4 As BlockReference = ent4
                        If Block4.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block4.AttributeCollection
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "SPACING" Then
                                    Continut4 = attref.TextString
                                End If
                            Next
                        End If
                    End If



                    If Not Continut1 = "" Then
                        ComboBox_SAND_beginsta.Items.Add(Continut1)
                        ComboBox_SAND_beginsta.SelectedIndex = ComboBox_SAND_beginsta.Items.IndexOf(Continut1)
                        Button_SAND_Beginsta.BackColor = Drawing.Color.Lime
                        ComboBox_SAND_endsta.Items.Add(Continut1)
                        ComboBox_SAND_no_type.Items.Add(Continut1)
                        ComboBox_SAND_spacing.Items.Add(Continut1)

                    End If
                    If Not Continut2 = "" Then
                        ComboBox_SAND_beginsta.Items.Add(Continut2)
                        ComboBox_SAND_endsta.Items.Add(Continut2)
                        ComboBox_SAND_endsta.SelectedIndex = ComboBox_SAND_endsta.Items.IndexOf(Continut2)
                        Button_SAND_Endsta.BackColor = Drawing.Color.Lime
                        ComboBox_SAND_no_type.Items.Add(Continut2)
                        ComboBox_SAND_spacing.Items.Add(Continut2)
                    End If
                    If Not Continut3 = "" Then
                        ComboBox_SAND_beginsta.Items.Add(Continut3)
                        ComboBox_SAND_endsta.Items.Add(Continut3)
                        ComboBox_SAND_no_type.Items.Add(Continut3)
                        ComboBox_SAND_no_type.SelectedIndex = ComboBox_SAND_no_type.Items.IndexOf(Continut3)
                        Button_SAND_no_type.BackColor = Drawing.Color.Lime
                        ComboBox_SAND_spacing.Items.Add(Continut3)
                    End If
                    If Not Continut4 = "" Then
                        ComboBox_SAND_beginsta.Items.Add(Continut4)
                        ComboBox_SAND_endsta.Items.Add(Continut4)
                        ComboBox_SAND_no_type.Items.Add(Continut4)
                        ComboBox_SAND_spacing.Items.Add(Continut4)
                        ComboBox_SAND_spacing.SelectedIndex = ComboBox_SAND_spacing.Items.IndexOf(Continut4)
                        Button_SAND_spacing.BackColor = Drawing.Color.Lime
                    End If






                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_SAND_insert_Click(sender As Object, e As EventArgs) Handles Button_sand_insert.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select objects to be deleted:"

                    Object_Prompt1.SingleOnly = False
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If
                    Dim Point_start As Autodesk.AutoCAD.EditorInput.PromptPointResult

                    Dim PP_start As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select insertion point:")
                    PP_start.AllowNone = True
                    Point_start = Editor1.GetPoint(PP_start)
                    If Not Point_start.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        Editor1.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If

                    Dim Colectie_atr_name As New Specialized.StringCollection
                    Dim Colectie_atr_value As New Specialized.StringCollection

                    If Not ComboBox_SAND_beginsta.Text = "" Then
                        Colectie_atr_name.Add("BEGINSTA")
                        Colectie_atr_value.Add(ComboBox_SAND_beginsta.Text)
                    End If

                    If Not ComboBox_SAND_endsta.Text = "" Then
                        Colectie_atr_name.Add("ENDSTA")
                        Colectie_atr_value.Add(ComboBox_SAND_endsta.Text)
                    End If

                    If Not ComboBox_SAND_no_type.Text = "" Then
                        Colectie_atr_name.Add("NO_TYPE")
                        Colectie_atr_value.Add(ComboBox_SAND_no_type.Text)
                    End If

                    If Not ComboBox_SAND_spacing.Text = "" Then
                        Colectie_atr_name.Add("SPACING")
                        Colectie_atr_value.Add(ComboBox_SAND_spacing.Text)
                    End If


                    InsertBlock_with_multiple_atributes("Sand_Bag_alignment1.dwg", "Sand_Bag_alignment1", Point_start.Value, 1, BTrecord, "TEXT", Colectie_atr_name, Colectie_atr_value)

                    For i = 0 To Rezultat1.Value.Count - 1






                        Dim ent1 As Entity = Rezultat1.Value.Item(i).ObjectId.GetObject(OpenMode.ForWrite)
                        ent1.Erase()
                    Next








                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)

                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_SAND_Beginsta_Click(sender As Object, e As EventArgs) Handles Button_sand_Beginsta.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info {BEGINSTA}:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut1 As String


                    If TypeOf ent1 Is DBText Then
                        Dim Text1 As DBText = ent1
                        Continut1 = Text1.TextString
                    End If

                    If TypeOf ent1 Is MText Then
                        Dim MText1 As MText = ent1
                        Continut1 = MText1.Contents
                    End If


                    If Not Continut1 = "" Then
                        If ComboBox_SAND_beginsta.Items.Count > 1 Then
                            ComboBox_SAND_beginsta.Items(1) = Continut1
                            ComboBox_SAND_beginsta.SelectedIndex = ComboBox_SAND_beginsta.Items.IndexOf(Continut1)
                            Button_SAND_Beginsta.BackColor = Drawing.Color.Lime
                            ComboBox_SAND_endsta.Items(1) = Continut1
                            ComboBox_SAND_no_type.Items(1) = Continut1
                            ComboBox_SAND_spacing.Items(1) = Continut1
                        End If
                        If ComboBox_SAND_beginsta.Items.Count = 0 Then
                            ComboBox_SAND_beginsta.Items.Add(" ")
                            ComboBox_SAND_beginsta.Items.Add(Continut1)
                            ComboBox_SAND_beginsta.SelectedIndex = ComboBox_SAND_beginsta.Items.IndexOf(Continut1)
                            Button_SAND_Beginsta.BackColor = Drawing.Color.Lime
                            ComboBox_SAND_endsta.Items.Add(" ")
                            ComboBox_SAND_endsta.Items.Add(Continut1)
                            ComboBox_SAND_no_type.Items.Add(" ")
                            ComboBox_SAND_no_type.Items.Add(Continut1)
                            ComboBox_SAND_spacing.Items.Add(" ")
                            ComboBox_SAND_spacing.Items.Add(Continut1)
                        End If

                    Else
                        If ComboBox_SAND_beginsta.Items.Count = 0 Then
                            Button_SAND_Beginsta.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_SAND_Endsta_Click(sender As Object, e As EventArgs) Handles Button_sand_Endsta.Click

        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt2.MessageForAdding = vbLf & "Select info {ENDSTA}:"

                    Object_Prompt2.SingleOnly = True
                    Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                    If Not Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent2 As Entity = Rezultat2.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut2 As String


                    If TypeOf ent2 Is DBText Then
                        Dim Text2 As DBText = ent2
                        Continut2 = Text2.TextString
                    End If

                    If TypeOf ent2 Is MText Then
                        Dim MText2 As MText = ent2
                        Continut2 = MText2.Contents
                    End If


                    If Not Continut2 = "" Then
                        If ComboBox_SAND_endsta.Items.Count > 2 Then
                            ComboBox_SAND_beginsta.Items(2) = Continut2
                            ComboBox_SAND_endsta.Items(2) = Continut2
                            ComboBox_SAND_endsta.SelectedIndex = ComboBox_SAND_endsta.Items.IndexOf(Continut2)
                            Button_SAND_Endsta.BackColor = Drawing.Color.Lime
                            ComboBox_SAND_no_type.Items(2) = Continut2
                            ComboBox_SAND_spacing.Items(2) = Continut2
                        End If
                        If ComboBox_SAND_endsta.Items.Count = 0 Then
                            ComboBox_SAND_beginsta.Items.Add(" ")
                            ComboBox_SAND_beginsta.Items.Add(Continut2)
                            ComboBox_SAND_endsta.Items.Add(" ")
                            ComboBox_SAND_endsta.Items.Add(Continut2)
                            ComboBox_SAND_endsta.SelectedIndex = ComboBox_SAND_endsta.Items.IndexOf(Continut2)
                            Button_SAND_Endsta.BackColor = Drawing.Color.Lime
                            ComboBox_SAND_no_type.Items.Add(" ")
                            ComboBox_SAND_no_type.Items.Add(Continut2)
                            ComboBox_SAND_spacing.Items.Add(" ")
                            ComboBox_SAND_spacing.Items.Add(Continut2)
                        End If
                        If ComboBox_SAND_endsta.Items.Count = 1 Or ComboBox_SAND_endsta.Items.Count = 2 Then
                            ComboBox_SAND_beginsta.Items.Add(Continut2)
                            ComboBox_SAND_endsta.Items.Add(Continut2)
                            ComboBox_SAND_endsta.SelectedIndex = ComboBox_SAND_endsta.Items.IndexOf(Continut2)
                            Button_SAND_Endsta.BackColor = Drawing.Color.Lime
                            ComboBox_SAND_no_type.Items.Add(Continut2)
                            ComboBox_SAND_spacing.Items.Add(Continut2)
                        End If
                    Else
                        If ComboBox_SAND_endsta.Items.Count = 0 Then
                            Button_SAND_Endsta.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_SAND_no_type_Click(sender As Object, e As EventArgs) Handles Button_sand_no_type.Click

        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat5 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt5 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt5.MessageForAdding = vbLf & "Select info {NO_TYPE}:"

                    Object_Prompt5.SingleOnly = True
                    Rezultat5 = Editor1.GetSelection(Object_Prompt5)
                    If Not Rezultat5.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent5 As Entity = Rezultat5.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut5 As String


                    If TypeOf ent5 Is DBText Then
                        Dim Text5 As DBText = ent5
                        Continut5 = Text5.TextString
                    End If

                    If TypeOf ent5 Is MText Then
                        Dim MText5 As MText = ent5
                        Continut5 = MText5.Contents
                    End If


                    If Not Continut5 = "" Then
                        If ComboBox_SAND_no_type.Items.Count > 5 Then
                            ComboBox_SAND_beginsta.Items(5) = Continut5
                            ComboBox_SAND_endsta.Items(5) = Continut5
                            ComboBox_SAND_no_type.Items(5) = Continut5
                            ComboBox_SAND_no_type.SelectedIndex = ComboBox_SAND_no_type.Items.IndexOf(Continut5)
                            Button_SAND_no_type.BackColor = Drawing.Color.Lime
                            ComboBox_SAND_spacing.Items(5) = Continut5
                        End If
                        If ComboBox_SAND_no_type.Items.Count = 0 Then
                            ComboBox_SAND_beginsta.Items.Add(" ")
                            ComboBox_SAND_beginsta.Items.Add(Continut5)
                            ComboBox_SAND_endsta.Items.Add(" ")
                            ComboBox_SAND_endsta.Items.Add(Continut5)

                            ComboBox_SAND_no_type.Items.Add(" ")
                            ComboBox_SAND_no_type.Items.Add(Continut5)
                            ComboBox_SAND_no_type.SelectedIndex = ComboBox_SAND_no_type.Items.IndexOf(Continut5)
                            Button_SAND_no_type.BackColor = Drawing.Color.Lime
                            ComboBox_SAND_spacing.Items.Add(" ")
                            ComboBox_SAND_spacing.Items.Add(Continut5)
                        End If
                        If ComboBox_SAND_no_type.Items.Count = 1 Or ComboBox_SAND_no_type.Items.Count = 2 Or ComboBox_SAND_no_type.Items.Count = 3 Then
                            ComboBox_SAND_beginsta.Items.Add(Continut5)
                            ComboBox_SAND_endsta.Items.Add(Continut5)

                            ComboBox_SAND_no_type.Items.Add(Continut5)
                            ComboBox_SAND_no_type.SelectedIndex = ComboBox_SAND_no_type.Items.IndexOf(Continut5)
                            Button_SAND_no_type.BackColor = Drawing.Color.Lime

                            ComboBox_SAND_spacing.Items.Add(Continut5)
                        End If
                    Else
                        If ComboBox_SAND_no_type.Items.Count = 0 Then
                            Button_SAND_no_type.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_SAND_spacing_Click(sender As Object, e As EventArgs) Handles Button_sand_spacing.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat6 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt6 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt6.MessageForAdding = vbLf & "Select info {SPACING}:"

                    Object_Prompt6.SingleOnly = True
                    Rezultat6 = Editor1.GetSelection(Object_Prompt6)
                    If Not Rezultat6.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If
                    Dim ent6 As Entity = Rezultat6.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim Continut6 As String
                    If TypeOf ent6 Is DBText Then
                        Dim Text6 As DBText = ent6
                        Continut6 = Text6.TextString
                    End If

                    If TypeOf ent6 Is MText Then
                        Dim MText3 As MText = ent6
                        Continut6 = MText3.Contents
                    End If


                    If Not Continut6 = "" Then
                        If ComboBox_SAND_spacing.Items.Count > 6 Then
                            ComboBox_SAND_beginsta.Items(6) = Continut6
                            ComboBox_SAND_endsta.Items(6) = Continut6
                            ComboBox_SAND_no_type.Items(6) = Continut6
                            ComboBox_SAND_spacing.Items(6) = Continut6
                            ComboBox_SAND_spacing.SelectedIndex = ComboBox_SAND_spacing.Items.IndexOf(Continut6)
                            Button_SAND_spacing.BackColor = Drawing.Color.Lime
                        End If
                        If ComboBox_SAND_spacing.Items.Count = 0 Then
                            ComboBox_SAND_beginsta.Items.Add(" ")
                            ComboBox_SAND_beginsta.Items.Add(Continut6)
                            ComboBox_SAND_endsta.Items.Add(" ")
                            ComboBox_SAND_endsta.Items.Add(Continut6)

                            ComboBox_SAND_no_type.Items.Add(" ")
                            ComboBox_SAND_no_type.Items.Add(Continut6)

                            ComboBox_SAND_spacing.Items.Add(" ")
                            ComboBox_SAND_spacing.Items.Add(Continut6)
                            ComboBox_SAND_spacing.SelectedIndex = ComboBox_SAND_spacing.Items.IndexOf(Continut6)
                            Button_SAND_spacing.BackColor = Drawing.Color.Lime
                        End If
                        If ComboBox_SAND_spacing.Items.Count = 1 Or ComboBox_SAND_spacing.Items.Count = 2 Or ComboBox_SAND_spacing.Items.Count = 3 Or ComboBox_SAND_spacing.Items.Count = 4 Then
                            ComboBox_SAND_beginsta.Items.Add(Continut6)
                            ComboBox_SAND_endsta.Items.Add(Continut6)

                            ComboBox_SAND_no_type.Items.Add(Continut6)


                            ComboBox_SAND_spacing.Items.Add(Continut6)
                            ComboBox_SAND_spacing.SelectedIndex = ComboBox_SAND_spacing.Items.IndexOf(Continut6)
                            Button_SAND_spacing.BackColor = Drawing.Color.Lime
                        End If
                    Else
                        If ComboBox_SAND_spacing.Items.Count = 0 Then
                            Button_SAND_spacing.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub

    Private Sub Button_conc_CLEAR_Click(sender As Object, e As EventArgs) Handles Button_CONC_CLEAR.Click
        ComboBox_CONC_BEGINSTA.Items.Clear()
        ComboBox_CONC_ENDSTA.Items.Clear()
        ComboBox_CONC_LENGTH.Items.Clear()
        Button_CONC_BEGINSTA.BackColor = Drawing.Color.DimGray
        Button_CONC_ENDSTA.BackColor = Drawing.Color.DimGray
        Button_CONC_LENGTH.BackColor = Drawing.Color.DimGray
    End Sub
    Private Sub Button_conc_PICK_Click(sender As Object, e As EventArgs) Handles Button_CONC_PICK.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info {BEGINSTA}:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt2.MessageForAdding = vbLf & "Select info {ENDSTA}:"

                    Object_Prompt2.SingleOnly = True
                    Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                    If Not Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    Dim Rezultat3 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt3 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt3.MessageForAdding = vbLf & "Select info {LENGTH}:"

                    Object_Prompt3.SingleOnly = True
                    Rezultat3 = Editor1.GetSelection(Object_Prompt3)
                    If Not Rezultat3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    ComboBox_CONC_BEGINSTA.Items.Clear()
                    ComboBox_CONC_BEGINSTA.Items.Add(" ")
                    ComboBox_CONC_ENDSTA.Items.Clear()
                    ComboBox_CONC_ENDSTA.Items.Add(" ")
                    ComboBox_CONC_LENGTH.Items.Clear()
                    ComboBox_CONC_LENGTH.Items.Add(" ")



                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent2 As Entity = Rezultat2.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent3 As Entity = Rezultat3.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut1 As String
                    Dim Continut2 As String
                    Dim Continut3 As String


                    If TypeOf ent1 Is DBText Then
                        Dim Text1 As DBText = ent1
                        Continut1 = Text1.TextString
                    End If
                    If TypeOf ent2 Is DBText Then
                        Dim Text2 As DBText = ent2
                        Continut2 = Text2.TextString
                    End If
                    If TypeOf ent3 Is DBText Then
                        Dim Text3 As DBText = ent3
                        Continut3 = Text3.TextString
                    End If


                    If TypeOf ent1 Is MText Then
                        Dim MText1 As MText = ent1
                        Continut1 = MText1.Contents
                    End If
                    If TypeOf ent2 Is MText Then
                        Dim MText2 As MText = ent2
                        Continut2 = MText2.Contents
                    End If
                    If TypeOf ent3 Is MText Then
                        Dim MText3 As MText = ent3
                        Continut3 = MText3.Contents
                    End If



                    If TypeOf ent1 Is BlockReference Then
                        Dim Block1 As BlockReference = ent1
                        If Block1.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                            Dim ctemp1 As String
                            Dim ctemp2 As String
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "BEGINSTA" Then
                                    ctemp1 = attref.TextString
                                End If
                                If attref.Tag = "ENDSTA" Then
                                    ctemp2 = attref.TextString
                                End If
                            Next
                            ctemp1 = Replace(ctemp1, " ", "")
                            ctemp2 = Replace(ctemp2, " ", "")



                            If Continut2 = "" Then
                                If TypeOf ent2 Is BlockReference And Not ent1.ObjectId = ent2.ObjectId Then
                                    If Not ctemp1 = "" And Not ctemp2 = "" Then
                                        Continut1 = ctemp2
                                    End If
                                    If ctemp1 = "" And ctemp2 = "" Then
                                        Continut1 = " "
                                    End If
                                    If ctemp1 = "" And Not ctemp2 = "" Then
                                        Continut1 = ctemp2
                                    End If
                                    If Not ctemp1 = "" And ctemp2 = "" Then
                                        Continut1 = ctemp1
                                    End If
                                End If
                            Else
                                If Not ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut1 = ctemp2
                                End If
                                If ctemp1 = "" And ctemp2 = "" Then
                                    Continut1 = " "
                                End If
                                If ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut1 = ctemp2
                                End If
                                If Not ctemp1 = "" And ctemp2 = "" Then
                                    Continut1 = ctemp1
                                End If
                            End If
                            If ent1.ObjectId = ent2.ObjectId Then
                                Continut1 = ctemp1
                            End If

                        End If
                    End If


                    If TypeOf ent2 Is BlockReference Then
                        Dim Block2 As BlockReference = ent2
                        If Block2.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block2.AttributeCollection
                            Dim ctemp1 As String
                            Dim ctemp2 As String
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "BEGINSTA" Then
                                    ctemp1 = attref.TextString
                                End If
                                If attref.Tag = "ENDSTA" Then
                                    ctemp2 = attref.TextString
                                End If
                            Next
                            ctemp1 = Replace(ctemp1, " ", "")
                            ctemp2 = Replace(ctemp2, " ", "")
                            If TypeOf ent1 Is BlockReference Then
                                If Not ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut2 = ctemp1
                                End If
                                If ctemp1 = "" And ctemp2 = "" Then
                                    Continut2 = " "
                                End If
                                If ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut2 = ctemp2
                                End If
                                If Not ctemp1 = "" And ctemp2 = "" Then
                                    Continut2 = ctemp1
                                End If
                            Else
                                Continut2 = ctemp1
                            End If
                            If ent1.ObjectId = ent2.ObjectId Then
                                Continut2 = ctemp2
                            End If


                        End If
                    End If


                    If TypeOf ent3 Is BlockReference Then
                        Dim Block3 As BlockReference = ent3
                        If Block3.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block3.AttributeCollection
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "LENGTH" Then
                                    Continut3 = attref.TextString
                                End If
                            Next
                        End If
                    End If



                    If Not Continut1 = "" Then
                        ComboBox_CONC_BEGINSTA.Items.Add(Continut1)
                        ComboBox_CONC_BEGINSTA.SelectedIndex = ComboBox_CONC_BEGINSTA.Items.IndexOf(Continut1)
                        Button_CONC_BEGINSTA.BackColor = Drawing.Color.Lime
                        ComboBox_CONC_ENDSTA.Items.Add(Continut1)
                        ComboBox_CONC_LENGTH.Items.Add(Continut1)
                    End If

                    If Not Continut2 = "" Then
                        ComboBox_CONC_BEGINSTA.Items.Add(Continut2)
                        ComboBox_CONC_ENDSTA.Items.Add(Continut2)
                        ComboBox_CONC_ENDSTA.SelectedIndex = ComboBox_CONC_ENDSTA.Items.IndexOf(Continut2)
                        Button_CONC_ENDSTA.BackColor = Drawing.Color.Lime
                        ComboBox_CONC_LENGTH.Items.Add(Continut2)
                    End If

                    If Not Continut3 = "" Then
                        ComboBox_CONC_BEGINSTA.Items.Add(Continut3)
                        ComboBox_CONC_ENDSTA.Items.Add(Continut3)
                        ComboBox_CONC_LENGTH.Items.Add(Continut3)
                        ComboBox_CONC_LENGTH.SelectedIndex = ComboBox_CONC_LENGTH.Items.IndexOf(Continut3)
                        Button_CONC_LENGTH.BackColor = Drawing.Color.Lime
                    End If

                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_conc_INSERT_Click(sender As Object, e As EventArgs) Handles Button_CONC_INSERT.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select objects to be deleted:"

                    Object_Prompt1.SingleOnly = False
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If
                    Dim Point_start As Autodesk.AutoCAD.EditorInput.PromptPointResult

                    Dim PP_start As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select insertion point:")
                    PP_start.AllowNone = True
                    Point_start = Editor1.GetPoint(PP_start)
                    If Not Point_start.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        Editor1.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If

                    Dim Colectie_atr_name As New Specialized.StringCollection
                    Dim Colectie_atr_value As New Specialized.StringCollection

                    If Not ComboBox_CONC_BEGINSTA.Text = "" Then
                        Colectie_atr_name.Add("BEGINSTA")
                        Colectie_atr_value.Add(ComboBox_CONC_BEGINSTA.Text)
                    End If

                    If Not ComboBox_CONC_ENDSTA.Text = "" Then
                        Colectie_atr_name.Add("ENDSTA")
                        Colectie_atr_value.Add(ComboBox_CONC_ENDSTA.Text)
                    End If

                    If Not ComboBox_CONC_LENGTH.Text = "" Then
                        Colectie_atr_name.Add("LENGTH")
                        Colectie_atr_value.Add(ComboBox_CONC_LENGTH.Text)
                    End If



                    InsertBlock_with_multiple_atributes("Concrete_alignment1.dwg", "Concrete_alignment1", Point_start.Value, 1, BTrecord, "TEXT", Colectie_atr_name, Colectie_atr_value)

                    For i = 0 To Rezultat1.Value.Count - 1






                        Dim ent1 As Entity = Rezultat1.Value.Item(i).ObjectId.GetObject(OpenMode.ForWrite)
                        ent1.Erase()
                    Next








                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)

                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_conc_BEGINSTA_Click(sender As Object, e As EventArgs) Handles Button_CONC_BEGINSTA.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info {BEGINSTA}:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut1 As String


                    If TypeOf ent1 Is DBText Then
                        Dim Text1 As DBText = ent1
                        Continut1 = Text1.TextString
                    End If

                    If TypeOf ent1 Is MText Then
                        Dim MText1 As MText = ent1
                        Continut1 = MText1.Contents
                    End If


                    If Not Continut1 = "" Then
                        If ComboBox_CONC_BEGINSTA.Items.Count > 1 Then
                            ComboBox_CONC_BEGINSTA.Items(1) = Continut1
                            ComboBox_CONC_BEGINSTA.SelectedIndex = ComboBox_CONC_BEGINSTA.Items.IndexOf(Continut1)
                            Button_CONC_BEGINSTA.BackColor = Drawing.Color.Lime
                            ComboBox_CONC_ENDSTA.Items(1) = Continut1
                            ComboBox_CONC_LENGTH.Items(1) = Continut1
                        End If
                        If ComboBox_CONC_BEGINSTA.Items.Count = 0 Then
                            ComboBox_CONC_BEGINSTA.Items.Add(" ")
                            ComboBox_CONC_BEGINSTA.Items.Add(Continut1)
                            ComboBox_CONC_BEGINSTA.SelectedIndex = ComboBox_CONC_BEGINSTA.Items.IndexOf(Continut1)
                            Button_CONC_BEGINSTA.BackColor = Drawing.Color.Lime
                            ComboBox_CONC_ENDSTA.Items.Add(" ")
                            ComboBox_CONC_ENDSTA.Items.Add(Continut1)
                            ComboBox_CONC_LENGTH.Items.Add(" ")
                            ComboBox_CONC_LENGTH.Items.Add(Continut1)
                        End If

                    Else
                        If ComboBox_CONC_BEGINSTA.Items.Count = 0 Then
                            Button_CONC_BEGINSTA.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_conc_ENDSTA_Click(sender As Object, e As EventArgs) Handles Button_CONC_ENDSTA.Click

        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt2.MessageForAdding = vbLf & "Select info {ENDSTA}:"

                    Object_Prompt2.SingleOnly = True
                    Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                    If Not Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent2 As Entity = Rezultat2.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut2 As String


                    If TypeOf ent2 Is DBText Then
                        Dim Text2 As DBText = ent2
                        Continut2 = Text2.TextString
                    End If

                    If TypeOf ent2 Is MText Then
                        Dim MText2 As MText = ent2
                        Continut2 = MText2.Contents
                    End If


                    If Not Continut2 = "" Then
                        If ComboBox_CONC_ENDSTA.Items.Count > 2 Then
                            ComboBox_CONC_BEGINSTA.Items(2) = Continut2
                            ComboBox_CONC_ENDSTA.Items(2) = Continut2
                            ComboBox_CONC_ENDSTA.SelectedIndex = ComboBox_CONC_ENDSTA.Items.IndexOf(Continut2)
                            Button_CONC_ENDSTA.BackColor = Drawing.Color.Lime
                            ComboBox_CONC_LENGTH.Items(2) = Continut2
                        End If
                        If ComboBox_CONC_ENDSTA.Items.Count = 0 Then
                            ComboBox_CONC_BEGINSTA.Items.Add(" ")
                            ComboBox_CONC_BEGINSTA.Items.Add(Continut2)
                            ComboBox_CONC_ENDSTA.Items.Add(" ")
                            ComboBox_CONC_ENDSTA.Items.Add(Continut2)
                            ComboBox_CONC_ENDSTA.SelectedIndex = ComboBox_CONC_ENDSTA.Items.IndexOf(Continut2)
                            Button_CONC_ENDSTA.BackColor = Drawing.Color.Lime
                            ComboBox_CONC_LENGTH.Items.Add(" ")
                            ComboBox_CONC_LENGTH.Items.Add(Continut2)
                        End If
                        If ComboBox_CONC_ENDSTA.Items.Count = 1 Or ComboBox_CONC_ENDSTA.Items.Count = 2 Then
                            ComboBox_CONC_BEGINSTA.Items.Add(Continut2)
                            ComboBox_CONC_ENDSTA.Items.Add(Continut2)
                            ComboBox_CONC_ENDSTA.SelectedIndex = ComboBox_CONC_ENDSTA.Items.IndexOf(Continut2)
                            Button_CONC_ENDSTA.BackColor = Drawing.Color.Lime
                            ComboBox_CONC_LENGTH.Items.Add(Continut2)
                        End If
                    Else
                        If ComboBox_CONC_ENDSTA.Items.Count = 0 Then
                            Button_CONC_ENDSTA.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_conc_LENGTH_Click(sender As Object, e As EventArgs) Handles Button_CONC_LENGTH.Click

        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat5 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt5 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt5.MessageForAdding = vbLf & "Select info {LENGTH}:"

                    Object_Prompt5.SingleOnly = True
                    Rezultat5 = Editor1.GetSelection(Object_Prompt5)
                    If Not Rezultat5.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent5 As Entity = Rezultat5.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut5 As String


                    If TypeOf ent5 Is DBText Then
                        Dim Text5 As DBText = ent5
                        Continut5 = Text5.TextString
                    End If

                    If TypeOf ent5 Is MText Then
                        Dim MText5 As MText = ent5
                        Continut5 = MText5.Contents
                    End If


                    If Not Continut5 = "" Then
                        If ComboBox_CONC_LENGTH.Items.Count > 5 Then
                            ComboBox_CONC_BEGINSTA.Items(5) = Continut5
                            ComboBox_CONC_ENDSTA.Items(5) = Continut5
                            ComboBox_CONC_LENGTH.Items(5) = Continut5
                            ComboBox_CONC_LENGTH.SelectedIndex = ComboBox_CONC_LENGTH.Items.IndexOf(Continut5)
                            Button_CONC_LENGTH.BackColor = Drawing.Color.Lime
                        End If
                        If ComboBox_CONC_LENGTH.Items.Count = 0 Then
                            ComboBox_CONC_BEGINSTA.Items.Add(" ")
                            ComboBox_CONC_BEGINSTA.Items.Add(Continut5)
                            ComboBox_CONC_ENDSTA.Items.Add(" ")
                            ComboBox_CONC_ENDSTA.Items.Add(Continut5)

                            ComboBox_CONC_LENGTH.Items.Add(" ")
                            ComboBox_CONC_LENGTH.Items.Add(Continut5)
                            ComboBox_CONC_LENGTH.SelectedIndex = ComboBox_CONC_LENGTH.Items.IndexOf(Continut5)
                            Button_CONC_LENGTH.BackColor = Drawing.Color.Lime
                        End If
                        If ComboBox_CONC_LENGTH.Items.Count = 1 Or ComboBox_CONC_LENGTH.Items.Count = 2 Or ComboBox_CONC_LENGTH.Items.Count = 3 Then
                            ComboBox_CONC_BEGINSTA.Items.Add(Continut5)
                            ComboBox_CONC_ENDSTA.Items.Add(Continut5)

                            ComboBox_CONC_LENGTH.Items.Add(Continut5)
                            ComboBox_CONC_LENGTH.SelectedIndex = ComboBox_CONC_LENGTH.Items.IndexOf(Continut5)
                            Button_CONC_LENGTH.BackColor = Drawing.Color.Lime
                        End If
                    Else
                        If ComboBox_CONC_LENGTH.Items.Count = 0 Then
                            Button_CONC_LENGTH.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")

    End Sub
    Private Sub Button_CONC_blockpick_Click(sender As Object, e As EventArgs) Handles Button_CONC_blockpick.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info from Block:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    ComboBox_CONC_BEGINSTA.Items.Clear()
                    ComboBox_CONC_BEGINSTA.Items.Add(" ")
                    ComboBox_CONC_ENDSTA.Items.Clear()
                    ComboBox_CONC_ENDSTA.Items.Add(" ")
                    ComboBox_CONC_LENGTH.Items.Clear()
                    ComboBox_CONC_LENGTH.Items.Add(" ")


                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut1 As String
                    Dim Continut2 As String




                    If TypeOf ent1 Is BlockReference Then
                        Dim Block1 As BlockReference = ent1
                        If Block1.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "BEGINSTA" Then
                                    Continut1 = attref.TextString
                                    If Continut1 = "" Then Continut1 = " "
                                End If
                                If attref.Tag = "ENDSTA" Then
                                    Continut2 = attref.TextString
                                    If Continut2 = "" Then Continut2 = " "
                                End If
                            Next
                        End If
                    End If

                    If Not Continut1 = "" Then
                        ComboBox_CONC_BEGINSTA.Items.Add(Continut1)
                        ComboBox_CONC_BEGINSTA.SelectedIndex = ComboBox_CONC_BEGINSTA.Items.IndexOf(Continut1)
                        Button_CONC_BEGINSTA.BackColor = Drawing.Color.Lime
                        ComboBox_CONC_ENDSTA.Items.Add(Continut1)
                    End If
                    If Not Continut2 = "" Then
                        ComboBox_CONC_BEGINSTA.Items.Add(Continut2)
                        ComboBox_CONC_ENDSTA.Items.Add(Continut2)
                        ComboBox_CONC_ENDSTA.SelectedIndex = ComboBox_CONC_ENDSTA.Items.IndexOf(Continut2)
                        Button_CONC_ENDSTA.BackColor = Drawing.Color.Lime
                    End If


                    If IsNumeric(Replace(Continut1, "+", "")) = True And IsNumeric(Replace(Continut2, "+", "")) = True Then
                        Dim Chain1 As Double = Replace(Continut1, "+", "")
                        Dim Chain2 As Double = Replace(Continut2, "+", "")
                        Dim Length As Double = Abs(Chain1 - Chain2)
                        Dim Length_string As String = Get_String_Rounded(Length, 1)
                        ComboBox_CONC_BEGINSTA.Items.Add(Length_string)
                        ComboBox_CONC_ENDSTA.Items.Add(Length_string)
                        ComboBox_CONC_LENGTH.Items.Add(Length_string)
                        ComboBox_CONC_LENGTH.SelectedIndex = ComboBox_CONC_LENGTH.Items.IndexOf(Length_string)
                        Button_CONC_LENGTH.BackColor = Drawing.Color.Lime

                    End If



                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub

    Private Sub Button_test_CLEAR_Click(sender As Object, e As EventArgs) Handles Button_test_clear.Click
        ComboBox_test_desc.Items.Clear()
        ComboBox_test_id_no.Items.Clear()
        ComboBox_test_sta.Items.Clear()
        Button_test_desc.BackColor = Drawing.Color.DimGray
        Button_test_id_no.BackColor = Drawing.Color.DimGray
        Button_test_sta.BackColor = Drawing.Color.DimGray
    End Sub
    Private Sub Button_test_PICK_Click(sender As Object, e As EventArgs) Handles Button_test_pick.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info {desc}:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt2.MessageForAdding = vbLf & "Select info {id_no}:"

                    Object_Prompt2.SingleOnly = True
                    Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                    If Not Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    Dim Rezultat3 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt3 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt3.MessageForAdding = vbLf & "Select info {sta}:"

                    Object_Prompt3.SingleOnly = True
                    Rezultat3 = Editor1.GetSelection(Object_Prompt3)
                    If Not Rezultat3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    ComboBox_test_desc.Items.Clear()
                    ComboBox_test_desc.Items.Add(" ")
                    ComboBox_test_id_no.Items.Clear()
                    ComboBox_test_id_no.Items.Add(" ")
                    ComboBox_test_sta.Items.Clear()
                    ComboBox_test_sta.Items.Add(" ")



                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent2 As Entity = Rezultat2.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
                    Dim ent3 As Entity = Rezultat3.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut1 As String
                    Dim Continut2 As String
                    Dim Continut3 As String


                    If TypeOf ent1 Is DBText Then
                        Dim Text1 As DBText = ent1
                        Continut1 = Text1.TextString
                    End If
                    If TypeOf ent2 Is DBText Then
                        Dim Text2 As DBText = ent2
                        Continut2 = Text2.TextString
                    End If
                    If TypeOf ent3 Is DBText Then
                        Dim Text3 As DBText = ent3
                        Continut3 = Text3.TextString
                    End If


                    If TypeOf ent1 Is MText Then
                        Dim MText1 As MText = ent1
                        Continut1 = MText1.Contents
                    End If
                    If TypeOf ent2 Is MText Then
                        Dim MText2 As MText = ent2
                        Continut2 = MText2.Contents
                    End If
                    If TypeOf ent3 Is MText Then
                        Dim MText3 As MText = ent3
                        Continut3 = MText3.Contents
                    End If



                    If TypeOf ent1 Is BlockReference Then
                        Dim Block1 As BlockReference = ent1
                        If Block1.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                            Dim ctemp1 As String
                            Dim ctemp2 As String
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "DESC" Then
                                    ctemp1 = attref.TextString
                                End If
                                If attref.Tag = "ID_NO" Then
                                    ctemp2 = attref.TextString
                                End If
                            Next
                            ctemp1 = Replace(ctemp1, " ", "")
                            ctemp2 = Replace(ctemp2, " ", "")



                            If Continut2 = "" Then
                                If TypeOf ent2 Is BlockReference And Not ent1.ObjectId = ent2.ObjectId Then
                                    If Not ctemp1 = "" And Not ctemp2 = "" Then
                                        Continut1 = ctemp2
                                    End If
                                    If ctemp1 = "" And ctemp2 = "" Then
                                        Continut1 = " "
                                    End If
                                    If ctemp1 = "" And Not ctemp2 = "" Then
                                        Continut1 = ctemp2
                                    End If
                                    If Not ctemp1 = "" And ctemp2 = "" Then
                                        Continut1 = ctemp1
                                    End If
                                End If
                            Else
                                If Not ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut1 = ctemp2
                                End If
                                If ctemp1 = "" And ctemp2 = "" Then
                                    Continut1 = " "
                                End If
                                If ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut1 = ctemp2
                                End If
                                If Not ctemp1 = "" And ctemp2 = "" Then
                                    Continut1 = ctemp1
                                End If
                            End If
                            If ent1.ObjectId = ent2.ObjectId Then
                                Continut1 = ctemp1
                            End If

                        End If
                    End If


                    If TypeOf ent2 Is BlockReference Then
                        Dim Block2 As BlockReference = ent2
                        If Block2.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block2.AttributeCollection
                            Dim ctemp1 As String
                            Dim ctemp2 As String
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "DESC" Then
                                    ctemp1 = attref.TextString
                                End If
                                If attref.Tag = "ID_NO" Then
                                    ctemp2 = attref.TextString
                                End If
                            Next
                            ctemp1 = Replace(ctemp1, " ", "")
                            ctemp2 = Replace(ctemp2, " ", "")
                            If TypeOf ent1 Is BlockReference Then
                                If Not ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut2 = ctemp1
                                End If
                                If ctemp1 = "" And ctemp2 = "" Then
                                    Continut2 = " "
                                End If
                                If ctemp1 = "" And Not ctemp2 = "" Then
                                    Continut2 = ctemp2
                                End If
                                If Not ctemp1 = "" And ctemp2 = "" Then
                                    Continut2 = ctemp1
                                End If
                            Else
                                Continut2 = ctemp1
                            End If
                            If ent1.ObjectId = ent2.ObjectId Then
                                Continut2 = ctemp2
                            End If


                        End If
                    End If


                    If TypeOf ent3 Is BlockReference Then
                        Dim Block3 As BlockReference = ent3
                        If Block3.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block3.AttributeCollection
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "STA" Then
                                    Continut3 = attref.TextString
                                End If
                            Next
                        End If
                    End If



                    If Not Continut1 = "" Then
                        ComboBox_test_desc.Items.Add(Continut1)
                        ComboBox_test_desc.SelectedIndex = ComboBox_test_desc.Items.IndexOf(Continut1)
                        Button_test_desc.BackColor = Drawing.Color.Lime
                        ComboBox_test_id_no.Items.Add(Continut1)
                        ComboBox_test_sta.Items.Add(Continut1)
                    End If

                    If Not Continut2 = "" Then
                        ComboBox_test_desc.Items.Add(Continut2)
                        ComboBox_test_id_no.Items.Add(Continut2)
                        ComboBox_test_id_no.SelectedIndex = ComboBox_test_id_no.Items.IndexOf(Continut2)
                        Button_test_id_no.BackColor = Drawing.Color.Lime
                        ComboBox_test_sta.Items.Add(Continut2)
                    End If

                    If Not Continut3 = "" Then
                        ComboBox_test_desc.Items.Add(Continut3)
                        ComboBox_test_id_no.Items.Add(Continut3)
                        ComboBox_test_sta.Items.Add(Continut3)
                        ComboBox_test_sta.SelectedIndex = ComboBox_test_sta.Items.IndexOf(Continut3)
                        Button_test_sta.BackColor = Drawing.Color.Lime
                    End If

                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_test_INSERT_Click(sender As Object, e As EventArgs) Handles Button_test_insert.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select objects to be deleted:"

                    Object_Prompt1.SingleOnly = False
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If
                    Dim Point_start As Autodesk.AutoCAD.EditorInput.PromptPointResult

                    Dim PP_start As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select insertion point:")
                    PP_start.AllowNone = True
                    Point_start = Editor1.GetPoint(PP_start)
                    If Not Point_start.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        Editor1.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If

                    Dim Colectie_atr_name As New Specialized.StringCollection
                    Dim Colectie_atr_value As New Specialized.StringCollection

                    If Not ComboBox_test_desc.Text = "" Then
                        Colectie_atr_name.Add("DESC")
                        Colectie_atr_value.Add(ComboBox_test_desc.Text)
                    End If

                    If Not ComboBox_test_id_no.Text = "" Then
                        Colectie_atr_name.Add("ID_NO")
                        Colectie_atr_value.Add(ComboBox_test_id_no.Text)
                    End If

                    If Not ComboBox_test_sta.Text = "" Then
                        Colectie_atr_name.Add("STA")
                        Colectie_atr_value.Add(ComboBox_test_sta.Text)
                    End If



                    InsertBlock_with_multiple_atributes("Test_section_alignment_crossing.dwg", "Test_section_alignment_crossing", Point_start.Value, 1, BTrecord, "TEXT", Colectie_atr_name, Colectie_atr_value)

                    For i = 0 To Rezultat1.Value.Count - 1
                        Dim ent1 As Entity = Rezultat1.Value.Item(i).ObjectId.GetObject(OpenMode.ForWrite)
                        ent1.Erase()
                    Next








                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)

                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_test_desc_Click(sender As Object, e As EventArgs) Handles Button_test_desc.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info {desc}:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut1 As String


                    If TypeOf ent1 Is DBText Then
                        Dim Text1 As DBText = ent1
                        Continut1 = Text1.TextString
                    End If

                    If TypeOf ent1 Is MText Then
                        Dim MText1 As MText = ent1
                        Continut1 = MText1.Contents
                    End If


                    If Not Continut1 = "" Then
                        If ComboBox_test_desc.Items.Count > 1 Then
                            ComboBox_test_desc.Items(1) = Continut1
                            ComboBox_test_desc.SelectedIndex = ComboBox_test_desc.Items.IndexOf(Continut1)
                            Button_test_desc.BackColor = Drawing.Color.Lime
                            ComboBox_test_id_no.Items(1) = Continut1
                            ComboBox_test_sta.Items(1) = Continut1
                        End If
                        If ComboBox_test_desc.Items.Count = 0 Then
                            ComboBox_test_desc.Items.Add(" ")
                            ComboBox_test_desc.Items.Add(Continut1)
                            ComboBox_test_desc.SelectedIndex = ComboBox_test_desc.Items.IndexOf(Continut1)
                            Button_test_desc.BackColor = Drawing.Color.Lime
                            ComboBox_test_id_no.Items.Add(" ")
                            ComboBox_test_id_no.Items.Add(Continut1)
                            ComboBox_test_sta.Items.Add(" ")
                            ComboBox_test_sta.Items.Add(Continut1)
                        End If

                    Else
                        If ComboBox_test_desc.Items.Count = 0 Then
                            Button_test_desc.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_test_id_no_Click(sender As Object, e As EventArgs) Handles Button_test_id_no.Click

        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt2.MessageForAdding = vbLf & "Select info {id_no}:"

                    Object_Prompt2.SingleOnly = True
                    Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                    If Not Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent2 As Entity = Rezultat2.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut2 As String


                    If TypeOf ent2 Is DBText Then
                        Dim Text2 As DBText = ent2
                        Continut2 = Text2.TextString
                    End If

                    If TypeOf ent2 Is MText Then
                        Dim MText2 As MText = ent2
                        Continut2 = MText2.Contents
                    End If


                    If Not Continut2 = "" Then
                        If ComboBox_test_id_no.Items.Count > 2 Then
                            ComboBox_test_desc.Items(2) = Continut2
                            ComboBox_test_id_no.Items(2) = Continut2
                            ComboBox_test_id_no.SelectedIndex = ComboBox_test_id_no.Items.IndexOf(Continut2)
                            Button_test_id_no.BackColor = Drawing.Color.Lime
                            ComboBox_test_sta.Items(2) = Continut2
                        End If
                        If ComboBox_test_id_no.Items.Count = 0 Then
                            ComboBox_test_desc.Items.Add(" ")
                            ComboBox_test_desc.Items.Add(Continut2)
                            ComboBox_test_id_no.Items.Add(" ")
                            ComboBox_test_id_no.Items.Add(Continut2)
                            ComboBox_test_id_no.SelectedIndex = ComboBox_test_id_no.Items.IndexOf(Continut2)
                            Button_test_id_no.BackColor = Drawing.Color.Lime
                            ComboBox_test_sta.Items.Add(" ")
                            ComboBox_test_sta.Items.Add(Continut2)
                        End If
                        If ComboBox_test_id_no.Items.Count = 1 Or ComboBox_test_id_no.Items.Count = 2 Then
                            ComboBox_test_desc.Items.Add(Continut2)
                            ComboBox_test_id_no.Items.Add(Continut2)
                            ComboBox_test_id_no.SelectedIndex = ComboBox_test_id_no.Items.IndexOf(Continut2)
                            Button_test_id_no.BackColor = Drawing.Color.Lime
                            ComboBox_test_sta.Items.Add(Continut2)
                        End If
                    Else
                        If ComboBox_test_id_no.Items.Count = 0 Then
                            Button_test_id_no.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub
    Private Sub Button_test_sta_Click(sender As Object, e As EventArgs) Handles Button_test_sta.Click

        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat5 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt5 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt5.MessageForAdding = vbLf & "Select info {sta}:"

                    Object_Prompt5.SingleOnly = True
                    Rezultat5 = Editor1.GetSelection(Object_Prompt5)
                    If Not Rezultat5.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If



                    Dim ent5 As Entity = Rezultat5.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut5 As String


                    If TypeOf ent5 Is DBText Then
                        Dim Text5 As DBText = ent5
                        Continut5 = Text5.TextString
                    End If

                    If TypeOf ent5 Is MText Then
                        Dim MText5 As MText = ent5
                        Continut5 = MText5.Contents
                    End If


                    If Not Continut5 = "" Then
                        If ComboBox_test_sta.Items.Count > 5 Then
                            ComboBox_test_desc.Items(5) = Continut5
                            ComboBox_test_id_no.Items(5) = Continut5
                            ComboBox_test_sta.Items(5) = Continut5
                            ComboBox_test_sta.SelectedIndex = ComboBox_test_sta.Items.IndexOf(Continut5)
                            Button_test_sta.BackColor = Drawing.Color.Lime
                        End If
                        If ComboBox_test_sta.Items.Count = 0 Then
                            ComboBox_test_desc.Items.Add(" ")
                            ComboBox_test_desc.Items.Add(Continut5)
                            ComboBox_test_id_no.Items.Add(" ")
                            ComboBox_test_id_no.Items.Add(Continut5)

                            ComboBox_test_sta.Items.Add(" ")
                            ComboBox_test_sta.Items.Add(Continut5)
                            ComboBox_test_sta.SelectedIndex = ComboBox_test_sta.Items.IndexOf(Continut5)
                            Button_test_sta.BackColor = Drawing.Color.Lime
                        End If
                        If ComboBox_test_sta.Items.Count = 1 Or ComboBox_test_sta.Items.Count = 2 Or ComboBox_test_sta.Items.Count = 3 Then
                            ComboBox_test_desc.Items.Add(Continut5)
                            ComboBox_test_id_no.Items.Add(Continut5)

                            ComboBox_test_sta.Items.Add(Continut5)
                            ComboBox_test_sta.SelectedIndex = ComboBox_test_sta.Items.IndexOf(Continut5)
                            Button_test_sta.BackColor = Drawing.Color.Lime
                        End If
                    Else
                        If ComboBox_test_sta.Items.Count = 0 Then
                            Button_test_sta.BackColor = Drawing.Color.DimGray
                        End If
                    End If







                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")

    End Sub
    Private Sub Button_test_blockpick_Click(sender As Object, e As EventArgs) Handles Button_test_blockpick.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt1.MessageForAdding = vbLf & "Select info from Block:"

                    Object_Prompt1.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt1)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)

                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub

                    End If

                    ComboBox_test_desc.Items.Clear()
                    ComboBox_test_desc.Items.Add(" ")
                    ComboBox_test_id_no.Items.Clear()
                    ComboBox_test_id_no.Items.Add(" ")
                    ComboBox_test_sta.Items.Clear()
                    ComboBox_test_sta.Items.Add(" ")


                    Dim ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    Dim Continut1 As String
                    Dim Continut2 As String




                    If TypeOf ent1 Is BlockReference Then
                        Dim Block1 As BlockReference = ent1
                        If Block1.AttributeCollection.Count > 0 Then
                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                            For Each id In attColl
                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                If attref.Tag = "DESC" Then
                                    Continut1 = attref.TextString
                                    If Continut1 = "" Then Continut1 = " "
                                End If
                                If attref.Tag = "ID_NO" Then
                                    Continut2 = attref.TextString
                                    If Continut2 = "" Then Continut2 = " "
                                End If
                            Next
                        End If
                    End If

                    If Not Continut1 = "" Then
                        ComboBox_test_desc.Items.Add(Continut1)
                        ComboBox_test_desc.SelectedIndex = ComboBox_test_desc.Items.IndexOf(Continut1)
                        Button_test_desc.BackColor = Drawing.Color.Lime
                        ComboBox_test_id_no.Items.Add(Continut1)
                    End If
                    If Not Continut2 = "" Then
                        ComboBox_test_desc.Items.Add(Continut2)
                        ComboBox_test_id_no.Items.Add(Continut2)
                        ComboBox_test_id_no.SelectedIndex = ComboBox_test_id_no.Items.IndexOf(Continut2)
                        Button_test_id_no.BackColor = Drawing.Color.Lime
                    End If


                    If IsNumeric(Replace(Continut1, "+", "")) = True And IsNumeric(Replace(Continut2, "+", "")) = True Then
                        Dim Chain1 As Double = Replace(Continut1, "+", "")
                        Dim Chain2 As Double = Replace(Continut2, "+", "")
                        Dim sta As Double = Abs(Chain1 - Chain2)
                        Dim sta_string As String = Get_String_Rounded(sta, 1)
                        ComboBox_test_desc.Items.Add(sta_string)
                        ComboBox_test_id_no.Items.Add(sta_string)
                        ComboBox_test_sta.Items.Add(sta_string)
                        ComboBox_test_sta.SelectedIndex = ComboBox_test_sta.Items.IndexOf(sta_string)
                        Button_test_sta.BackColor = Drawing.Color.Lime

                    End If



                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                afiseaza_butoanele_pentru_forms(Me, Colectie1)




                ' asta e de la lock
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
    End Sub

    Private Sub Button_pick_multiple_chainages_Click(sender As Object, e As EventArgs) Handles Button_pick_multiple_chainages.Click
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult


            Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt2.MessageForAdding = vbLf & "Select 3D polyline:"

            Object_Prompt2.SingleOnly = True

            Rezultat2 = Editor1.GetSelection(Object_Prompt2)


            If Rezultat2.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                Exit Sub
            End If


            Dim Poly3d As Polyline3d


            Dim Point_on_poly As New Point3d


            If Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat2) = False Then
                    Using lock1 As DocumentLock = ThisDrawing.LockDocument


                        Dim Data_table1 As New System.Data.DataTable
                        Data_table1.Columns.Add("TEXT325", GetType(DBText))
                        Dim Index1 As Double = 0

                        Dim Data_table2 As New System.Data.DataTable
                        Data_table2.Columns.Add("TEXT0", GetType(DBText))
                        Dim Index2 As Double = 0

                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                            For Each ObjID In BTrecord
                                Dim DBobject As DBObject = Trans1.GetObject(ObjID, OpenMode.ForRead)
                                If TypeOf DBobject Is DBText Then
                                    Dim Text1 As DBText = DBobject
                                    If Text1.Layer = ComboBox_layer.Text Then
                                        If Text1.Rotation > 3 * PI / 2 Then
                                            Data_table1.Rows.Add()
                                            Data_table1.Rows(Index1).Item("TEXT325") = Text1
                                            Index1 = Index1 + 1
                                        End If
                                        If Text1.Rotation >= 0 And Text1.Rotation < PI / 4 Then
                                            Data_table2.Rows.Add()
                                            Data_table2.Rows(Index2).Item("TEXT0") = Text1
                                            Index2 = Index2 + 1
                                        End If

                                    End If
                                End If
                            Next
                        End Using

123:

                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


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

                            Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult

                            Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select Point:")
                            PP1.AllowNone = True
                            Point1 = Editor1.GetPoint(PP1)

                            If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                Editor1.WriteMessage(vbLf & "Command:")
                                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                Trans1.Commit()
                                Exit Sub
                            End If

                            Point_on_poly = Poly3d.GetClosestPointTo(Point1.Value.TransformBy(curent_ucs_matrix), Editor1.GetCurrentView.ViewDirection, False)

                            Dim Parameter_picked As Double = Round(Poly3d.GetParameterAtPoint(Point_on_poly), 3)

                            Dim Parameter_start As Double = Floor(Parameter_picked)
                            Dim Parameter_end As Double = Ceiling(Parameter_picked)
                            If Parameter_picked = Round(Parameter_picked, 0) Then
                                Parameter_start = Parameter_picked
                                Parameter_end = Parameter_picked
                            End If

                            Dim Chainage_on_vertex As Double
                            Dim Distanta_pana_la_Vertex As Double
                            Dim CSF1, CSF2 As Double

                            If Data_table1.Rows.Count > 0 Then
                                Dim Point_CHAINAGE As New Point3d
                                Point_CHAINAGE = Poly3d.GetPointAtParameter(Parameter_start)
                                Distanta_pana_la_Vertex = Point_CHAINAGE.GetVectorTo(Point_on_poly).Length

                                For i = 0 To Data_table1.Rows.Count - 1
                                    Dim Text1 As DBText = Data_table1.Rows(i).Item("TEXT325")
                                    If Point_CHAINAGE.GetVectorTo(Text1.Position.TransformBy(Editor1.CurrentUserCoordinateSystem)).Length < 0.1 Then
                                        Dim String1 As String = Replace(Text1.TextString, "+", "")
                                        If IsNumeric(String1) = True Then
                                            Chainage_on_vertex = CDbl(String1)
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If

                            If Not Parameter_start = Parameter_end Then
                                If Data_table2.Rows.Count > 0 Then
                                    Dim Point_CHAINAGE1 As New Point3d
                                    Point_CHAINAGE1 = Poly3d.GetPointAtParameter(Parameter_start)
                                    Dim Point_CHAINAGE2 As New Point3d
                                    Point_CHAINAGE2 = Poly3d.GetPointAtParameter(Parameter_end)

                                    For i = 0 To Data_table2.Rows.Count - 1
                                        Dim Text1 As DBText = Data_table2.Rows(i).Item("TEXT0")
                                        Dim String1 As String = Text1.TextString
                                        String1 = extrage_numar_din_text_de_la_sfarsitul_textului(String1)
                                        If IsNumeric(String1) = True Then
                                            If Point_CHAINAGE1.GetVectorTo(Text1.Position.TransformBy(Editor1.CurrentUserCoordinateSystem)).Length < 0.1 Then
                                                If CDbl(String1) > 0.5 And CDbl(String1) < 1.5 Then
                                                    CSF1 = CDbl(String1)
                                                End If
                                            End If

                                            If Point_CHAINAGE2.GetVectorTo(Text1.Position.TransformBy(Editor1.CurrentUserCoordinateSystem)).Length < 0.1 Then
                                                If CDbl(String1) > 0.5 And CDbl(String1) < 1.5 Then
                                                    CSF2 = CDbl(String1)
                                                End If
                                            End If
                                        End If
                                    Next
                                End If
                            End If

                            Dim New_chainage As String
                            Dim New_ch As Double
                            If Not CSF1 + CSF2 = 0 And Not CSF1 = 0 And Not CSF2 = 0 Then
                                New_ch = Chainage_on_vertex + Distanta_pana_la_Vertex / ((CSF1 + CSF2) / 2)
                            Else
                                New_ch = Chainage_on_vertex + Distanta_pana_la_Vertex
                            End If


                            New_chainage = Get_chainage_from_double(New_ch, 1)

                            ComboBox_PICK_CHAINAGE.Items.Add(New_chainage)
                            ComboBox_PICK_CHAINAGE.SelectedIndex = ComboBox_PICK_CHAINAGE.Items.IndexOf(New_chainage)
                            Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly, New_chainage, 0.5, 0.5, 0.5, 11, 3.5)

                            Trans1.Commit()

                        End Using
                        GoTo 123
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
    Private Sub Button_multiple_chainages_clear_Click(sender As Object, e As EventArgs) Handles Button_multiple_chainages_clear.Click
        ComboBox_PICK_CHAINAGE.Items.Clear()
    End Sub
    Private Sub Button_UPDATE_2_Blocks_Click(sender As Object, e As EventArgs) Handles Button_UPDATE_2_Blocks.Click
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult


            Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt2.MessageForAdding = vbLf & "Select the blocks:"

            Object_Prompt2.SingleOnly = False

            Rezultat2 = Editor1.GetSelection(Object_Prompt2)


            If Rezultat2.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                Exit Sub
            End If




            If Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat2) = False Then
                    Using lock1 As DocumentLock = ThisDrawing.LockDocument
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)

                            If Rezultat2.Value.Count > 0 Then
                                Dim Old_Chainage_station As String
                                Dim Diferenta As Double
                                For i = 0 To Rezultat2.Value.Count - 1
                                    Dim Block1 As BlockReference
                                    Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                    Obj1 = Rezultat2.Value.Item(i)
                                    Dim Ent1 As Entity
                                    Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                                    If TypeOf Ent1 Is BlockReference Then
                                        Block1 = Ent1
                                        If Block1.AttributeCollection.Count > 0 Then
                                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                            For Each id In attColl
                                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                                If attref.Tag = "STA" Then
                                                    Old_Chainage_station = attref.TextString
                                                End If
                                            Next
                                            Dim New_Chainage_station As String
                                            If IsNumeric(Replace(ComboBox_PICK_CHAINAGE.Text, "+", "")) = True Then
                                                New_Chainage_station = ComboBox_PICK_CHAINAGE.Text
                                            End If
                                            If Not Old_Chainage_station = "" And Not New_Chainage_station = "" Then
                                                Dim New_ch As Double
                                                Dim Old_ch As Double
                                                New_ch = CDbl(Replace(New_Chainage_station, "+", ""))
                                                Old_ch = CDbl(Replace(Old_Chainage_station, "+", ""))
                                                Diferenta = New_ch - Old_ch
                                                Exit For
                                            End If
                                        End If
                                    End If
                                Next

                                For i = 0 To Rezultat2.Value.Count - 1
                                    Dim Block1 As BlockReference
                                    Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                    Obj1 = Rezultat2.Value.Item(i)
                                    Dim Ent1 As Entity
                                    Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)


                                    If TypeOf Ent1 Is BlockReference Then
                                        Block1 = Ent1
                                        If Block1.AttributeCollection.Count > 0 Then
                                            Block1.UpgradeOpen()
                                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection

                                            For Each id In attColl
                                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForWrite)
                                                If attref.Tag = "STA" Or attref.Tag = "BEGINSTA" Or attref.Tag = "ENDSTA" Then
                                                    If IsNumeric(Replace(attref.TextString, "+", "")) = True Then
                                                        Dim Old_ch As Double
                                                        Old_ch = CDbl(Replace(attref.TextString, "+", ""))
                                                        Dim New_ch As Double
                                                        New_ch = Old_ch + Diferenta
                                                        If Not Diferenta = 0 Then
                                                            attref.TextString = Get_chainage_from_double(New_ch, 1)
                                                        End If
                                                    End If
                                                End If
                                            Next
                                        End If
                                    End If
                                Next

                            End If


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

    Private Sub Button_update_match_Click(sender As Object, e As EventArgs) Handles Button_update_match.Click
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt1.MessageForAdding = vbLf & "Select the block to be updated:"
            Object_Prompt1.SingleOnly = True
            Rezultat1 = Editor1.GetSelection(Object_Prompt1)
            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                Exit Sub
            End If


            Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt2.MessageForAdding = vbLf & "Select the block for {BEGINSTA}:"
            Object_Prompt2.SingleOnly = True
            Rezultat2 = Editor1.GetSelection(Object_Prompt2)


            Dim Rezultat3 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Dim Object_Prompt3 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt3.MessageForAdding = vbLf & "Select the block for {ENDSTA}:"
            Object_Prompt3.SingleOnly = True
            Rezultat3 = Editor1.GetSelection(Object_Prompt3)



            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                If IsNothing(Rezultat1) = False Then

                    Using lock1 As DocumentLock = ThisDrawing.LockDocument
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                            Dim Begin_station As String
                            Dim End_station As String


                            If Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                If IsNothing(Rezultat2) = False Then
                                    If Rezultat2.Value.Count > 0 Then
                                        For i = 0 To Rezultat2.Value.Count - 1
                                            Dim Block2 As BlockReference
                                            Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                            Obj2 = Rezultat2.Value.Item(i)
                                            Dim Ent2 As Entity
                                            Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)

                                            If TypeOf Ent2 Is BlockReference Then
                                                Block2 = Ent2
                                                If Block2.AttributeCollection.Count > 0 Then
                                                    Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block2.AttributeCollection
                                                    For Each id In attColl
                                                        Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                                        If attref.Tag = "ENDSTA" Then
                                                            Begin_station = attref.TextString
                                                        End If
                                                    Next
                                                End If
                                            End If
                                        Next
                                    End If
                                End If
                            End If

                            If Rezultat3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                If IsNothing(Rezultat3) = False Then
                                    If Rezultat3.Value.Count > 0 Then
                                        For i = 0 To Rezultat3.Value.Count - 1
                                            Dim Block3 As BlockReference
                                            Dim Obj3 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                            Obj3 = Rezultat3.Value.Item(i)
                                            Dim Ent3 As Entity
                                            Ent3 = Obj3.ObjectId.GetObject(OpenMode.ForRead)

                                            If TypeOf Ent3 Is BlockReference Then
                                                Block3 = Ent3
                                                If Block3.AttributeCollection.Count > 0 Then
                                                    Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block3.AttributeCollection
                                                    For Each id In attColl
                                                        Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                                        If attref.Tag = "BEGINSTA" Then
                                                            End_station = attref.TextString
                                                        End If
                                                    Next
                                                End If
                                            End If
                                        Next
                                    End If
                                End If
                            End If

                            For i = 0 To Rezultat1.Value.Count - 1
                                Dim Block1 As BlockReference
                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Rezultat1.Value.Item(i)
                                Dim Ent1 As Entity
                                Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)


                                If TypeOf Ent1 Is BlockReference Then
                                    Block1 = Ent1
                                    If Block1.AttributeCollection.Count > 0 Then
                                        Block1.UpgradeOpen()
                                        Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection

                                        For Each id In attColl
                                            Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForWrite)
                                            If attref.Tag = "BEGINSTA" Then
                                                If Not Begin_station = "" Then attref.TextString = Begin_station
                                            End If
                                            If attref.Tag = "ENDSTA" Then
                                                If Not End_station = "" Then attref.TextString = End_station
                                            End If
                                        Next
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


End Class