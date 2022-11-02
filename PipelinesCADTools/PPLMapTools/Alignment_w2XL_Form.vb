Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class Alignment_w2XL_Form
    Dim Colectie1 As New Specialized.StringCollection
    Dim Index_row_excel As Integer
    Dim ENDSTA_screw_anchors As Double = -1
    Dim ENDSTA_RW As Double = -1
    Dim ENDSTA_SBW As Double = -1
    Dim Spacing_screw_anchors As String = ""
    Dim Spacing_RW As String = ""
    Dim Spacing_SBW As String = ""
    Dim Curent_mat As Integer = -1
    Private Sub Panel_Click(sender As Object, e As EventArgs) Handles Panel1.Click, Panel4.Click, Panel3.Click, Panel2.Click
        Button_SA_end.Visible = False
        Button_SA_start.Visible = True
        Button_RW_end.Visible = False
        Button_RW_start.Visible = True
        Button_SBW_End.Visible = False
        Button_SBW_start.Visible = True
    End Sub

    Private Sub Button_TRANSITION_end_Click(sender As Object, e As EventArgs) Handles Button_transition.Click
        If IsNumeric(TextBox_ROW_START_XL.Text) = False Then
            Exit Sub
        Else
            Index_row_excel = CInt(TextBox_ROW_START_XL.Text)
        End If
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()

        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                    Editor1.SetImpliedSelection(Empty_array)



                    Dim Rezultat_End_sta As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_End_sta As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_End_sta.MessageForAdding = vbLf & "Select the Chainage:"
                    Object_Prompt_End_sta.SingleOnly = True
                    Rezultat_End_sta = Editor1.GetSelection(Object_Prompt_End_sta)
                    If Not Rezultat_End_sta.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If


                    Dim Rezultat_Material As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_Material As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_Material.MessageForAdding = vbLf & "Select the Material (before):"
                    Object_Prompt_Material.SingleOnly = True
                    Rezultat_Material = Editor1.GetSelection(Object_Prompt_Material)
                    If Not Rezultat_Material.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If

                    Dim Rezultat_Material1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_Material1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_Material1.MessageForAdding = vbLf & "Select the Material (after):"
                    Object_Prompt_Material1.SingleOnly = True
                    Rezultat_Material1 = Editor1.GetSelection(Object_Prompt_Material1)
                    If Not Rezultat_Material1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If


                    Dim ENDSTA As Double = -1
                    Dim MAT As Double = -1



                    If IsNothing(Rezultat_End_sta) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_End_sta.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            If IsNumeric(Replace(Text1.TextString, "+", "")) = True Then
                                ENDSTA = CDbl(Replace(Text1.TextString, "+", ""))
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            If IsNumeric(Replace(Text1.Text, "+", "")) = True Then
                                ENDSTA = CDbl(Replace(Text1.Text, "+", ""))
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "ENDSTA" Then
                                        Dim Continut As String = attref.TextString
                                        If IsNumeric(Replace(Continut, "+", "")) = True Then
                                            ENDSTA = CDbl(Replace(Continut, "+", ""))
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            If IsNumeric(Replace(Continut, "+", "")) = True Then
                                ENDSTA = CDbl(Replace(Continut, "+", ""))

                            End If
                        End If

                    End If

                    If IsNothing(Rezultat_Material) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_Material.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            If IsNumeric(Text1.TextString) = True Then
                                MAT = Round(CDbl(Text1.TextString), 0)
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            If IsNumeric(Text1.Text) = True Then
                                MAT = Round(CDbl(Text1.Text), 0)
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "MAT" Then
                                        Dim Continut As String = attref.TextString
                                        If IsNumeric(Continut) = True Then
                                            MAT = Round(CDbl(Continut), 0)
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            If IsNumeric(Continut) = True Then
                                MAT = Round(CDbl(Continut), 0)
                            End If
                        End If

                    End If

                    Dim Mat1 As Integer = -1
                    If IsNothing(Rezultat_Material1) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_Material1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            If IsNumeric(Text1.TextString) = True Then
                                Mat1 = CInt(Text1.TextString)
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            If IsNumeric(Text1.Text) = True Then
                                Mat1 = CInt(Text1.Text)
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "MAT" Then
                                        Dim Continut As String = attref.TextString
                                        If IsNumeric(Continut) = True Then
                                            Mat1 = CInt(Continut)
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            If IsNumeric(Continut) = True Then
                                Mat1 = CInt(Continut)
                            End If
                        End If

                    End If

                    If Not MAT = -1 And Not ENDSTA = -1 And Not Mat1 = -1 Then
                        If Not Curent_mat = -1 Then
                            W1.Range("G" & Index_row_excel).Value = MAT
                            W1.Range("A" & Index_row_excel & ":J" & Index_row_excel).Interior.Color = 12040422
                            Index_row_excel = Index_row_excel + 1
                        End If



                        Curent_mat = Mat1
                        W1.Range("B" & Index_row_excel).Value = "TRANSITION"
                        W1.Range("E" & Index_row_excel).Value = ENDSTA
                        W1.Range("G" & Index_row_excel).Value = "T"
                        W1.Range("A" & Index_row_excel & ":J" & Index_row_excel).Interior.ColorIndex = 37
                        TextBox_ROW_START_XL.Text = (Index_row_excel + 1).ToString

                    End If





                    Trans1.Commit()
                End Using
            End Using

            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub

    Private Sub Button_PICK_water_Click(sender As Object, e As EventArgs) Handles Button_PICK_Water.Click
        If IsNumeric(TextBox_ROW_START_XL.Text) = False Then
            Exit Sub
        Else
            Index_row_excel = CInt(TextBox_ROW_START_XL.Text)
        End If
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()

        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                    Editor1.SetImpliedSelection(Empty_array)



                    Dim Rezultat_desc1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_desc1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_desc1.MessageForAdding = vbLf & "Select Description Line 1 :"
                    Object_Prompt_desc1.SingleOnly = True

                    Rezultat_desc1 = Editor1.GetSelection(Object_Prompt_desc1)
                    If Not Rezultat_desc1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If


                    Dim Rezultat_desc2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_desc2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_desc2.MessageForAdding = vbLf & "Select Description Line 2 :"
                    Object_Prompt_desc2.SingleOnly = True

                    Rezultat_desc2 = Editor1.GetSelection(Object_Prompt_desc2)

                    Dim Rezultat_desc_ref As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_desc_ref As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_desc_ref.MessageForAdding = vbLf & "Select Reference Drawing :"
                    Object_Prompt_desc_ref.SingleOnly = True

                    Rezultat_desc_ref = Editor1.GetSelection(Object_Prompt_desc_ref)

                    Dim Rezultat_sta As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_sta As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_sta.MessageForAdding = vbLf & "Select Station Chainage:"
                    Object_Prompt_sta.SingleOnly = True

                    Rezultat_sta = Editor1.GetSelection(Object_Prompt_sta)
                    If Not Rezultat_sta.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If


                    Dim STA As Double = -1

                    If IsNothing(Rezultat_sta) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_sta.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            If IsNumeric(Replace(Text1.TextString, "+", "")) = True Then
                                STA = CDbl(Replace(Text1.TextString, "+", ""))
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            If IsNumeric(Replace(Text1.Text, "+", "")) = True Then
                                STA = CDbl(Replace(Text1.Text, "+", ""))
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "STA" Then
                                        Dim Continut As String = attref.TextString
                                        If IsNumeric(Replace(Continut, "+", "")) = True Then
                                            STA = CDbl(Replace(Continut, "+", ""))
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            If IsNumeric(Replace(Continut, "+", "")) = True Then
                                STA = CDbl(Replace(Continut, "+", ""))

                            End If
                        End If

                    End If

                    Dim Description1 As String = ""
                    If IsNothing(Rezultat_desc1) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_desc1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Description1 = Text1.TextString
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Description1 = Text1.Contents
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "DESC" Then
                                        If attref.IsMTextAttribute = False Then
                                            Dim Continut As String = attref.TextString
                                            Description1 = Continut
                                        Else
                                            Dim Continut As String = attref.MTextAttribute.Contents
                                            Description1 = Continut
                                        End If

                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            Description1 = Continut
                        End If

                    End If

                    Dim Description2 As String = ""
                    If Rezultat_desc2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_desc2.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Description2 = Text1.TextString
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Description2 = Text1.Contents
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "DESC" Then
                                        If attref.IsMTextAttribute = False Then
                                            Dim Continut As String = attref.TextString
                                            Description2 = Continut
                                        Else
                                            Dim Continut As String = attref.MTextAttribute.Contents
                                            Description2 = Continut
                                        End If

                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            Description2 = Continut
                        End If

                    End If

                    Dim Ref_ID As String = ""
                    If Rezultat_desc_ref.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_desc_ref.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Ref_ID = Text1.TextString
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Ref_ID = Text1.Contents
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "ID_NO" Then
                                        If attref.IsMTextAttribute = False Then
                                            Dim Continut As String = attref.TextString
                                            Ref_ID = Continut
                                        Else
                                            Dim Continut As String = attref.MTextAttribute.Contents
                                            Ref_ID = Continut
                                        End If

                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            Ref_ID = Continut
                        End If

                    End If



                    If Not STA = -1 And Not Description1 = "" Then


                        W1.Range("B" & Index_row_excel).Value = Description1
                        If Not Description2 = "" Then
                            W1.Range("C" & Index_row_excel).Value = Description2
                        End If
                        If Not Ref_ID = "" Then
                            W1.Range("D" & Index_row_excel).Value = Ref_ID
                        End If

                        W1.Range("E" & Index_row_excel).Value = STA
                        W1.Range("G" & Index_row_excel).Value = Curent_mat
                        W1.Range("J" & Index_row_excel).Value = "WC"
                        W1.Range("A" & Index_row_excel & ":J" & Index_row_excel).Interior.Color = 12040422
                        TextBox_ROW_START_XL.Text = (Index_row_excel + 1).ToString
                    End If

                    Trans1.Commit()
                End Using
            End Using

            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub

    Private Sub Button_PICK_road_Click(sender As Object, e As EventArgs) Handles Button_PICK_road.Click
        If IsNumeric(TextBox_ROW_START_XL.Text) = False Then
            Exit Sub
        Else
            Index_row_excel = CInt(TextBox_ROW_START_XL.Text)
        End If
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()

        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                    Editor1.SetImpliedSelection(Empty_array)



                    Dim Rezultat_desc1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_desc1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_desc1.MessageForAdding = vbLf & "Select Description Line 1 :"
                    Object_Prompt_desc1.SingleOnly = True

                    Rezultat_desc1 = Editor1.GetSelection(Object_Prompt_desc1)
                    If Not Rezultat_desc1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If


                    Dim Rezultat_desc2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_desc2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_desc2.MessageForAdding = vbLf & "Select Description Line 2 :"
                    Object_Prompt_desc2.SingleOnly = True

                    Rezultat_desc2 = Editor1.GetSelection(Object_Prompt_desc2)

                    Dim Rezultat_desc_ref As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_desc_ref As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_desc_ref.MessageForAdding = vbLf & "Select Reference Drawing :"
                    Object_Prompt_desc_ref.SingleOnly = True

                    Rezultat_desc_ref = Editor1.GetSelection(Object_Prompt_desc_ref)

                    Dim Rezultat_sta As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_sta As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_sta.MessageForAdding = vbLf & "Select Station Chainage:"
                    Object_Prompt_sta.SingleOnly = True

                    Rezultat_sta = Editor1.GetSelection(Object_Prompt_sta)
                    If Not Rezultat_sta.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If


                    Dim STA As Double = -1

                    If IsNothing(Rezultat_sta) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_sta.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            If IsNumeric(Replace(Text1.TextString, "+", "")) = True Then
                                STA = CDbl(Replace(Text1.TextString, "+", ""))
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            If IsNumeric(Replace(Text1.Text, "+", "")) = True Then
                                STA = CDbl(Replace(Text1.Text, "+", ""))
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "STA" Then
                                        Dim Continut As String = attref.TextString
                                        If IsNumeric(Replace(Continut, "+", "")) = True Then
                                            STA = CDbl(Replace(Continut, "+", ""))
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            If IsNumeric(Replace(Continut, "+", "")) = True Then
                                STA = CDbl(Replace(Continut, "+", ""))

                            End If
                        End If

                    End If

                    Dim Description1 As String = ""
                    If IsNothing(Rezultat_desc1) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_desc1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Description1 = Text1.TextString
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Description1 = Text1.Contents
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "DESC" Then
                                        If attref.IsMTextAttribute = False Then
                                            Dim Continut As String = attref.TextString
                                            Description1 = Continut
                                        Else
                                            Dim Continut As String = attref.MTextAttribute.Contents
                                            Description1 = Continut
                                        End If

                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            Description1 = Continut
                        End If

                    End If

                    Dim Description2 As String = ""
                    If Rezultat_desc2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_desc2.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Description2 = Text1.TextString
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Description2 = Text1.Contents
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "DESC" Then
                                        If attref.IsMTextAttribute = False Then
                                            Dim Continut As String = attref.TextString
                                            Description2 = Continut
                                        Else
                                            Dim Continut As String = attref.MTextAttribute.Contents
                                            Description2 = Continut
                                        End If

                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            Description2 = Continut
                        End If

                    End If

                    Dim Ref_ID As String = ""
                    If Rezultat_desc_ref.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_desc_ref.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Ref_ID = Text1.TextString
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Ref_ID = Text1.Contents
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "ID_NO" Then
                                        If attref.IsMTextAttribute = False Then
                                            Dim Continut As String = attref.TextString
                                            Ref_ID = Continut
                                        Else
                                            Dim Continut As String = attref.MTextAttribute.Contents
                                            Ref_ID = Continut
                                        End If

                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            Ref_ID = Continut
                        End If

                    End If



                    If Not STA = -1 And Not Description1 = "" Then

                        W1.Range("B" & Index_row_excel).Value = Description1
                        If Not Description2 = "" Then
                            W1.Range("C" & Index_row_excel).Value = Description2
                        End If
                        If Not Ref_ID = "" Then
                            W1.Range("D" & Index_row_excel).Value = Ref_ID
                        End If

                        W1.Range("E" & Index_row_excel).Value = STA
                        W1.Range("G" & Index_row_excel).Value = Curent_mat
                        W1.Range("J" & Index_row_excel).Value = "RD"
                        W1.Range("A" & Index_row_excel & ":J" & Index_row_excel).Interior.Color = 12040422
                        TextBox_ROW_START_XL.Text = (Index_row_excel + 1).ToString
                    End If

                    Trans1.Commit()
                End Using
            End Using

            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub

    Private Sub Button_PICK_rail_Click(sender As Object, e As EventArgs) Handles Button_pick_rail.Click
        If IsNumeric(TextBox_ROW_START_XL.Text) = False Then
            Exit Sub
        Else
            Index_row_excel = CInt(TextBox_ROW_START_XL.Text)
        End If
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()

        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                    Editor1.SetImpliedSelection(Empty_array)



                    Dim Rezultat_desc1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_desc1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_desc1.MessageForAdding = vbLf & "Select Description Line 1 :"
                    Object_Prompt_desc1.SingleOnly = True

                    Rezultat_desc1 = Editor1.GetSelection(Object_Prompt_desc1)
                    If Not Rezultat_desc1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If


                    Dim Rezultat_desc2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_desc2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_desc2.MessageForAdding = vbLf & "Select Description Line 2 :"
                    Object_Prompt_desc2.SingleOnly = True

                    Rezultat_desc2 = Editor1.GetSelection(Object_Prompt_desc2)

                    Dim Rezultat_desc_ref As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_desc_ref As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_desc_ref.MessageForAdding = vbLf & "Select Reference Drawing :"
                    Object_Prompt_desc_ref.SingleOnly = True

                    Rezultat_desc_ref = Editor1.GetSelection(Object_Prompt_desc_ref)

                    Dim Rezultat_sta As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_sta As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_sta.MessageForAdding = vbLf & "Select Station Chainage:"
                    Object_Prompt_sta.SingleOnly = True

                    Rezultat_sta = Editor1.GetSelection(Object_Prompt_sta)
                    If Not Rezultat_sta.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If


                    Dim STA As Double = -1

                    If IsNothing(Rezultat_sta) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_sta.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            If IsNumeric(Replace(Text1.TextString, "+", "")) = True Then
                                STA = CDbl(Replace(Text1.TextString, "+", ""))
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            If IsNumeric(Replace(Text1.Text, "+", "")) = True Then
                                STA = CDbl(Replace(Text1.Text, "+", ""))
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "STA" Then
                                        Dim Continut As String = attref.TextString
                                        If IsNumeric(Replace(Continut, "+", "")) = True Then
                                            STA = CDbl(Replace(Continut, "+", ""))
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            If IsNumeric(Replace(Continut, "+", "")) = True Then
                                STA = CDbl(Replace(Continut, "+", ""))

                            End If
                        End If

                    End If

                    Dim Description1 As String = ""
                    If IsNothing(Rezultat_desc1) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_desc1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Description1 = Text1.TextString
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Description1 = Text1.Contents
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "DESC" Then
                                        If attref.IsMTextAttribute = False Then
                                            Dim Continut As String = attref.TextString
                                            Description1 = Continut
                                        Else
                                            Dim Continut As String = attref.MTextAttribute.Contents
                                            Description1 = Continut
                                        End If

                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            Description1 = Continut
                        End If

                    End If

                    Dim Description2 As String = ""
                    If Rezultat_desc2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_desc2.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Description2 = Text1.TextString
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Description2 = Text1.Contents
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "DESC" Then
                                        If attref.IsMTextAttribute = False Then
                                            Dim Continut As String = attref.TextString
                                            Description2 = Continut
                                        Else
                                            Dim Continut As String = attref.MTextAttribute.Contents
                                            Description2 = Continut
                                        End If

                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            Description2 = Continut
                        End If

                    End If

                    Dim Ref_ID As String = ""
                    If Rezultat_desc_ref.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_desc_ref.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Ref_ID = Text1.TextString
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Ref_ID = Text1.Contents
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "ID_NO" Then
                                        If attref.IsMTextAttribute = False Then
                                            Dim Continut As String = attref.TextString
                                            Ref_ID = Continut
                                        Else
                                            Dim Continut As String = attref.MTextAttribute.Contents
                                            Ref_ID = Continut
                                        End If

                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            Ref_ID = Continut
                        End If

                    End If



                    If Not STA = -1 And Not Description1 = "" Then


                        W1.Range("B" & Index_row_excel).Value = Description1
                        If Not Description2 = "" Then
                            W1.Range("C" & Index_row_excel).Value = Description2
                        End If
                        If Not Ref_ID = "" Then
                            W1.Range("D" & Index_row_excel).Value = Ref_ID
                        End If

                        W1.Range("E" & Index_row_excel).Value = STA
                        W1.Range("G" & Index_row_excel).Value = Curent_mat
                        W1.Range("J" & Index_row_excel).Value = "RR"
                        W1.Range("A" & Index_row_excel & ":J" & Index_row_excel).Interior.Color = 12040422
                        TextBox_ROW_START_XL.Text = (Index_row_excel + 1).ToString
                    End If

                    Trans1.Commit()
                End Using
            End Using

            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub

    Private Sub Button_PICK_power_Click(sender As Object, e As EventArgs) Handles Button_pick_power.Click
        If IsNumeric(TextBox_ROW_START_XL.Text) = False Then
            Exit Sub
        Else
            Index_row_excel = CInt(TextBox_ROW_START_XL.Text)
        End If
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()

        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                    Editor1.SetImpliedSelection(Empty_array)



                    Dim Rezultat_desc1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_desc1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_desc1.MessageForAdding = vbLf & "Select Description Line 1 :"
                    Object_Prompt_desc1.SingleOnly = True

                    Rezultat_desc1 = Editor1.GetSelection(Object_Prompt_desc1)
                    If Not Rezultat_desc1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If


                    Dim Rezultat_desc2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_desc2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_desc2.MessageForAdding = vbLf & "Select Description Line 2 :"
                    Object_Prompt_desc2.SingleOnly = True

                    Rezultat_desc2 = Editor1.GetSelection(Object_Prompt_desc2)

                    Dim Rezultat_desc_ref As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_desc_ref As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_desc_ref.MessageForAdding = vbLf & "Select Reference Drawing :"
                    Object_Prompt_desc_ref.SingleOnly = True

                    Rezultat_desc_ref = Editor1.GetSelection(Object_Prompt_desc_ref)

                    Dim Rezultat_sta As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_sta As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_sta.MessageForAdding = vbLf & "Select Station Chainage:"
                    Object_Prompt_sta.SingleOnly = True

                    Rezultat_sta = Editor1.GetSelection(Object_Prompt_sta)
                    If Not Rezultat_sta.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If

                    Dim Rezultat_cover As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_cover As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_cover.MessageForAdding = vbLf & "Select clearance :"
                    Object_Prompt_cover.SingleOnly = True

                    Rezultat_cover = Editor1.GetSelection(Object_Prompt_cover)

                    Dim STA As Double = -1

                    If IsNothing(Rezultat_sta) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_sta.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            If IsNumeric(Replace(Text1.TextString, "+", "")) = True Then
                                STA = CDbl(Replace(Text1.TextString, "+", ""))
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            If IsNumeric(Replace(Text1.Text, "+", "")) = True Then
                                STA = CDbl(Replace(Text1.Text, "+", ""))
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "STA" Then
                                        Dim Continut As String = attref.TextString
                                        If IsNumeric(Replace(Continut, "+", "")) = True Then
                                            STA = CDbl(Replace(Continut, "+", ""))
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            If IsNumeric(Replace(Continut, "+", "")) = True Then
                                STA = CDbl(Replace(Continut, "+", ""))

                            End If
                        End If

                    End If

                    Dim Description1 As String = ""
                    If IsNothing(Rezultat_desc1) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_desc1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Description1 = Text1.TextString
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Description1 = Text1.Contents
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "DESC" Then
                                        If attref.IsMTextAttribute = False Then
                                            Dim Continut As String = attref.TextString
                                            Description1 = Continut
                                        Else
                                            Dim Continut As String = attref.MTextAttribute.Contents
                                            Description1 = Continut
                                        End If

                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            Description1 = Continut
                        End If

                    End If

                    Dim Description2 As String = ""
                    If Rezultat_desc2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_desc2.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Description2 = Text1.TextString
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Description2 = Text1.Contents
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "DESC" Then
                                        If attref.IsMTextAttribute = False Then
                                            Dim Continut As String = attref.TextString
                                            Description2 = Continut
                                        Else
                                            Dim Continut As String = attref.MTextAttribute.Contents
                                            Description2 = Continut
                                        End If

                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            Description2 = Continut
                        End If

                    End If

                    Dim Ref_ID As String = ""
                    If Rezultat_desc_ref.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_desc_ref.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Ref_ID = Text1.TextString
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Ref_ID = Text1.Contents
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "ID_NO" Then
                                        If attref.IsMTextAttribute = False Then
                                            Dim Continut As String = attref.TextString
                                            Ref_ID = Continut
                                        Else
                                            Dim Continut As String = attref.MTextAttribute.Contents
                                            Ref_ID = Continut
                                        End If

                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            Ref_ID = Continut
                        End If

                    End If

                    Dim Cover As Double = -1
                    If Rezultat_cover.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_cover.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Dim Continut As String = extrage_numar_din_text(Text1.TextString)
                            If IsNumeric(Continut) = True Then
                                Cover = CDbl(Continut)
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Dim Continut As String = extrage_numar_din_text(Text1.Text)
                            If IsNumeric(Continut) = True Then
                                Cover = CDbl(Continut)
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "ID_NO" Then
                                        If attref.IsMTextAttribute = False Then
                                            Dim Continut As String = extrage_numar_din_text(attref.TextString)
                                            If IsNumeric(Continut) = True Then
                                                Cover = CDbl(Continut)
                                            End If
                                        Else
                                            Dim Continut As String = extrage_numar_din_text(attref.MTextAttribute.Contents)
                                            If IsNumeric(Continut) = True Then
                                                Cover = CDbl(Continut)
                                            End If
                                        End If

                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = extrage_numar_din_text(attref.Tag.ToString)
                            If IsNumeric(Continut) = True Then
                                Cover = CDbl(Continut)
                            End If
                        End If

                    End If

                    If Not STA = -1 And Not Description1 = "" Then


                        W1.Range("B" & Index_row_excel).Value = Description1
                        If Not Description2 = "" Then
                            W1.Range("C" & Index_row_excel).Value = Description2
                        End If
                        If Not Ref_ID = "" Then
                            W1.Range("D" & Index_row_excel).Value = Ref_ID
                        End If

                        If Not Cover = -1 Then
                            W1.Range("I" & Index_row_excel).Value = Cover
                        End If

                        W1.Range("E" & Index_row_excel).Value = STA
                        W1.Range("G" & Index_row_excel).Value = Curent_mat
                        W1.Range("J" & Index_row_excel).Value = "PW"
                        W1.Range("A" & Index_row_excel & ":J" & Index_row_excel).Interior.Color = 12040422
                        TextBox_ROW_START_XL.Text = (Index_row_excel + 1).ToString
                    End If

                    Trans1.Commit()
                End Using
            End Using

            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub

    Private Sub Button_PICK_corosion_Click(sender As Object, e As EventArgs) Handles Button_COROSION.Click
        If IsNumeric(TextBox_ROW_START_XL.Text) = False Then
            Exit Sub
        Else
            Index_row_excel = CInt(TextBox_ROW_START_XL.Text)
        End If
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()

        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                    Editor1.SetImpliedSelection(Empty_array)




                    Dim Rezultat_sta As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_sta As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_sta.MessageForAdding = vbLf & "Select Station Chainage:"
                    Object_Prompt_sta.SingleOnly = True

                    Rezultat_sta = Editor1.GetSelection(Object_Prompt_sta)
                    If Not Rezultat_sta.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If


                    Dim STA As Double = -1

                    If IsNothing(Rezultat_sta) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_sta.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            If IsNumeric(Replace(Text1.TextString, "+", "")) = True Then
                                STA = CDbl(Replace(Text1.TextString, "+", ""))
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            If IsNumeric(Replace(Text1.Text, "+", "")) = True Then
                                STA = CDbl(Replace(Text1.Text, "+", ""))
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "STA" Then
                                        Dim Continut As String = attref.TextString
                                        If IsNumeric(Replace(Continut, "+", "")) = True Then
                                            STA = CDbl(Replace(Continut, "+", ""))
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            If IsNumeric(Replace(Continut, "+", "")) = True Then
                                STA = CDbl(Replace(Continut, "+", ""))

                            End If
                        End If

                    End If



                    If Not STA = -1 Then


                       

                        W1.Range("E" & Index_row_excel).Value = STA

                        W1.Range("G" & Index_row_excel).Value = "CP"
                        W1.Range("A" & Index_row_excel & ":J" & Index_row_excel).Interior.Color = 65535
                        TextBox_ROW_START_XL.Text = (Index_row_excel + 1).ToString
                    End If

                    Trans1.Commit()
                End Using
            End Using

            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub

    Private Sub Button_PICK_pipe_Click(sender As Object, e As EventArgs) Handles Button_pick_pipe.Click
        If IsNumeric(TextBox_ROW_START_XL.Text) = False Then
            Exit Sub
        Else
            Index_row_excel = CInt(TextBox_ROW_START_XL.Text)
        End If
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()

        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                    Editor1.SetImpliedSelection(Empty_array)



                    Dim Rezultat_desc1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_desc1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_desc1.MessageForAdding = vbLf & "Select Description Line 1 :"
                    Object_Prompt_desc1.SingleOnly = True

                    Rezultat_desc1 = Editor1.GetSelection(Object_Prompt_desc1)
                    If Not Rezultat_desc1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If


                    Dim Rezultat_desc2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_desc2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_desc2.MessageForAdding = vbLf & "Select Description Line 2 :"
                    Object_Prompt_desc2.SingleOnly = True

                    Rezultat_desc2 = Editor1.GetSelection(Object_Prompt_desc2)

                    Dim Rezultat_desc_ref As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_desc_ref As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_desc_ref.MessageForAdding = vbLf & "Select Reference Drawing :"
                    Object_Prompt_desc_ref.SingleOnly = True

                    Rezultat_desc_ref = Editor1.GetSelection(Object_Prompt_desc_ref)

                    Dim Rezultat_sta As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_sta As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_sta.MessageForAdding = vbLf & "Select Station Chainage:"
                    Object_Prompt_sta.SingleOnly = True

                    Rezultat_sta = Editor1.GetSelection(Object_Prompt_sta)
                    If Not Rezultat_sta.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If
                    Dim Rezultat_cover As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_cover As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_cover.MessageForAdding = vbLf & "Select Cover :"
                    Object_Prompt_cover.SingleOnly = True

                    Rezultat_cover = Editor1.GetSelection(Object_Prompt_cover)

                    Dim STA As Double = -1

                    If IsNothing(Rezultat_sta) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_sta.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            If IsNumeric(Replace(Text1.TextString, "+", "")) = True Then
                                STA = CDbl(Replace(Text1.TextString, "+", ""))
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            If IsNumeric(Replace(Text1.Text, "+", "")) = True Then
                                STA = CDbl(Replace(Text1.Text, "+", ""))
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "STA" Then
                                        Dim Continut As String = attref.TextString
                                        If IsNumeric(Replace(Continut, "+", "")) = True Then
                                            STA = CDbl(Replace(Continut, "+", ""))
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            If IsNumeric(Replace(Continut, "+", "")) = True Then
                                STA = CDbl(Replace(Continut, "+", ""))

                            End If
                        End If

                    End If

                    Dim Description1 As String = ""
                    If IsNothing(Rezultat_desc1) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_desc1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Description1 = Text1.TextString
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Description1 = Text1.Contents
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "DESC" Then
                                        If attref.IsMTextAttribute = False Then
                                            Dim Continut As String = attref.TextString
                                            Description1 = Continut
                                        Else
                                            Dim Continut As String = attref.MTextAttribute.Contents
                                            Description1 = Continut
                                        End If

                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            Description1 = Continut
                        End If

                    End If

                    Dim Description2 As String = ""
                    If Rezultat_desc2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_desc2.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Description2 = Text1.TextString
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Description2 = Text1.Contents
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "DESC" Then
                                        If attref.IsMTextAttribute = False Then
                                            Dim Continut As String = attref.TextString
                                            Description2 = Continut
                                        Else
                                            Dim Continut As String = attref.MTextAttribute.Contents
                                            Description2 = Continut
                                        End If

                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            Description2 = Continut
                        End If

                    End If

                    Dim Ref_ID As String = ""
                    If Rezultat_desc_ref.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_desc_ref.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Ref_ID = Text1.TextString
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Ref_ID = Text1.Contents
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "ID_NO" Then
                                        If attref.IsMTextAttribute = False Then
                                            Dim Continut As String = attref.TextString
                                            Ref_ID = Continut
                                        Else
                                            Dim Continut As String = attref.MTextAttribute.Contents
                                            Ref_ID = Continut
                                        End If

                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            Ref_ID = Continut
                        End If

                    End If

                    Dim Cover As Double = -1
                    If Rezultat_cover.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_cover.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Dim Continut As String = extrage_numar_din_text(Text1.TextString)
                            If IsNumeric(Continut) = True Then
                                Cover = CDbl(Continut)
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Dim Continut As String = extrage_numar_din_text(Text1.Text)
                            If IsNumeric(Continut) = True Then
                                Cover = CDbl(Continut)
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "ID_NO" Then
                                        If attref.IsMTextAttribute = False Then
                                            Dim Continut As String = extrage_numar_din_text(attref.TextString)
                                            If IsNumeric(Continut) = True Then
                                                Cover = CDbl(Continut)
                                            End If
                                        Else
                                            Dim Continut As String = extrage_numar_din_text(attref.MTextAttribute.Contents)
                                            If IsNumeric(Continut) = True Then
                                                Cover = CDbl(Continut)
                                            End If
                                        End If

                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = extrage_numar_din_text(attref.Tag.ToString)
                            If IsNumeric(Continut) = True Then
                                Cover = CDbl(Continut)
                            End If
                        End If

                    End If


                    If Not STA = -1 And Not Description1 = "" Then


                        W1.Range("B" & Index_row_excel).Value = Description1
                        If Not Description2 = "" Then
                            W1.Range("C" & Index_row_excel).Value = Description2
                        End If
                        If Not Ref_ID = "" Then
                            W1.Range("D" & Index_row_excel).Value = Ref_ID
                        End If
                        If Not Cover = -1 Then
                            W1.Range("I" & Index_row_excel).Value = Round(Cover, 1)
                        End If
                        W1.Range("E" & Index_row_excel).Value = STA
                        W1.Range("G" & Index_row_excel).Value = Curent_mat
                        W1.Range("J" & Index_row_excel).Value = "PL"
                        W1.Range("A" & Index_row_excel & ":J" & Index_row_excel).Interior.Color = 12040422
                        TextBox_ROW_START_XL.Text = (Index_row_excel + 1).ToString
                    End If

                    Trans1.Commit()
                End Using
            End Using

            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub

    Private Sub Button_CABLE_Click(sender As Object, e As EventArgs) Handles Button_CABLE.Click
        If IsNumeric(TextBox_ROW_START_XL.Text) = False Then
            Exit Sub
        Else
            Index_row_excel = CInt(TextBox_ROW_START_XL.Text)
        End If
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()

        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                    Editor1.SetImpliedSelection(Empty_array)



                    Dim Rezultat_desc1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_desc1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_desc1.MessageForAdding = vbLf & "Select Description Line 1 :"
                    Object_Prompt_desc1.SingleOnly = True

                    Rezultat_desc1 = Editor1.GetSelection(Object_Prompt_desc1)
                    If Not Rezultat_desc1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If


                    Dim Rezultat_desc2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_desc2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_desc2.MessageForAdding = vbLf & "Select Description Line 2 :"
                    Object_Prompt_desc2.SingleOnly = True

                    Rezultat_desc2 = Editor1.GetSelection(Object_Prompt_desc2)

                    Dim Rezultat_desc_ref As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_desc_ref As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_desc_ref.MessageForAdding = vbLf & "Select Reference Drawing :"
                    Object_Prompt_desc_ref.SingleOnly = True

                    Rezultat_desc_ref = Editor1.GetSelection(Object_Prompt_desc_ref)

                    Dim Rezultat_sta As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_sta As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_sta.MessageForAdding = vbLf & "Select Station Chainage:"
                    Object_Prompt_sta.SingleOnly = True

                    Rezultat_sta = Editor1.GetSelection(Object_Prompt_sta)
                    If Not Rezultat_sta.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If
                    Dim Rezultat_cover As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_cover As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_cover.MessageForAdding = vbLf & "Select Cover :"
                    Object_Prompt_cover.SingleOnly = True

                    Rezultat_cover = Editor1.GetSelection(Object_Prompt_cover)

                    Dim STA As Double = -1

                    If IsNothing(Rezultat_sta) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_sta.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            If IsNumeric(Replace(Text1.TextString, "+", "")) = True Then
                                STA = CDbl(Replace(Text1.TextString, "+", ""))
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            If IsNumeric(Replace(Text1.Text, "+", "")) = True Then
                                STA = CDbl(Replace(Text1.Text, "+", ""))
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "STA" Then
                                        Dim Continut As String = attref.TextString
                                        If IsNumeric(Replace(Continut, "+", "")) = True Then
                                            STA = CDbl(Replace(Continut, "+", ""))
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            If IsNumeric(Replace(Continut, "+", "")) = True Then
                                STA = CDbl(Replace(Continut, "+", ""))

                            End If
                        End If

                    End If

                    Dim Description1 As String = ""
                    If IsNothing(Rezultat_desc1) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_desc1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Description1 = Text1.TextString
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Description1 = Text1.Contents
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "DESC" Then
                                        If attref.IsMTextAttribute = False Then
                                            Dim Continut As String = attref.TextString
                                            Description1 = Continut
                                        Else
                                            Dim Continut As String = attref.MTextAttribute.Contents
                                            Description1 = Continut
                                        End If

                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            Description1 = Continut
                        End If

                    End If

                    Dim Description2 As String = ""
                    If Rezultat_desc2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_desc2.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Description2 = Text1.TextString
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Description2 = Text1.Contents
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "DESC" Then
                                        If attref.IsMTextAttribute = False Then
                                            Dim Continut As String = attref.TextString
                                            Description2 = Continut
                                        Else
                                            Dim Continut As String = attref.MTextAttribute.Contents
                                            Description2 = Continut
                                        End If

                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            Description2 = Continut
                        End If

                    End If

                    Dim Ref_ID As String = ""
                    If Rezultat_desc_ref.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_desc_ref.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Ref_ID = Text1.TextString
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Ref_ID = Text1.Contents
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "ID_NO" Then
                                        If attref.IsMTextAttribute = False Then
                                            Dim Continut As String = attref.TextString
                                            Ref_ID = Continut
                                        Else
                                            Dim Continut As String = attref.MTextAttribute.Contents
                                            Ref_ID = Continut
                                        End If

                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            Ref_ID = Continut
                        End If

                    End If

                    Dim Cover As Double = -1
                    If Rezultat_cover.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_cover.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Dim Continut As String = extrage_numar_din_text(Text1.TextString)
                            If IsNumeric(Continut) = True Then
                                Cover = CDbl(Continut)
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Dim Continut As String = extrage_numar_din_text(Text1.Text)
                            If IsNumeric(Continut) = True Then
                                Cover = CDbl(Continut)
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "ID_NO" Then
                                        If attref.IsMTextAttribute = False Then
                                            Dim Continut As String = extrage_numar_din_text(attref.TextString)
                                            If IsNumeric(Continut) = True Then
                                                Cover = CDbl(Continut)
                                            End If
                                        Else
                                            Dim Continut As String = extrage_numar_din_text(attref.MTextAttribute.Contents)
                                            If IsNumeric(Continut) = True Then
                                                Cover = CDbl(Continut)
                                            End If
                                        End If

                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = extrage_numar_din_text(attref.Tag.ToString)
                            If IsNumeric(Continut) = True Then
                                Cover = CDbl(Continut)
                            End If
                        End If

                    End If


                    If Not STA = -1 And Not Description1 = "" Then


                        W1.Range("B" & Index_row_excel).Value = Description1
                        If Not Description2 = "" Then
                            W1.Range("C" & Index_row_excel).Value = Description2
                        End If
                        If Not Ref_ID = "" Then
                            W1.Range("D" & Index_row_excel).Value = Ref_ID
                        End If
                        If Not Cover = -1 Then
                            W1.Range("I" & Index_row_excel).Value = Round(Cover, 1)
                        End If
                        W1.Range("E" & Index_row_excel).Value = STA
                        W1.Range("G" & Index_row_excel).Value = Curent_mat
                        W1.Range("J" & Index_row_excel).Value = "CM"
                        W1.Range("A" & Index_row_excel & ":J" & Index_row_excel).Interior.Color = 12040422
                        TextBox_ROW_START_XL.Text = (Index_row_excel + 1).ToString
                    End If

                    Trans1.Commit()
                End Using
            End Using

            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub

    Private Sub Button_PICK_facility_Click(sender As Object, e As EventArgs) Handles Button_facility.Click
        If IsNumeric(TextBox_ROW_START_XL.Text) = False Then
            Exit Sub
        Else
            Index_row_excel = CInt(TextBox_ROW_START_XL.Text)
        End If
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()

        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                    Editor1.SetImpliedSelection(Empty_array)




                    Dim Rezultat_Start_STA As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_start_sta As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_start_sta.MessageForAdding = vbLf & "Select the Start Chainage:"
                    Object_Prompt_start_sta.SingleOnly = True

                    Rezultat_Start_STA = Editor1.GetSelection(Object_Prompt_start_sta)
                    If Not Rezultat_Start_STA.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If



                    Dim Rezultat_End_sta As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_End_sta As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_End_sta.MessageForAdding = vbLf & "Select the End Chainage:"
                    Object_Prompt_End_sta.SingleOnly = True
                    Rezultat_End_sta = Editor1.GetSelection(Object_Prompt_End_sta)
                    If Not Rezultat_End_sta.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If

                    Dim Rezultat_sta As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_sta As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_sta.MessageForAdding = vbLf & "Select Station Chainage:"
                    Object_Prompt_sta.SingleOnly = True

                    Rezultat_sta = Editor1.GetSelection(Object_Prompt_sta)
                    If Not Rezultat_sta.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If

                    Dim Rezultat_desc_ref As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_desc_ref As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_desc_ref.MessageForAdding = vbLf & "Select Reference Drawing :"
                    Object_Prompt_desc_ref.SingleOnly = True

                    Rezultat_desc_ref = Editor1.GetSelection(Object_Prompt_desc_ref)

                    Dim Rezultat_Material As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_Material As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_Material.MessageForAdding = vbLf & "Select the Material:"
                    Object_Prompt_Material.SingleOnly = True
                    Rezultat_Material = Editor1.GetSelection(Object_Prompt_Material)


                    Dim STA As Double = -1
                    Dim BEGINSTA As Double = -1
                    Dim ENDSTA As Double = -1
                    Dim MAT As Integer = -1

                    If IsNothing(Rezultat_Start_STA) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_Start_STA.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            If IsNumeric(Replace(Text1.TextString, "+", "")) = True Then
                                BEGINSTA = CDbl(Replace(Text1.TextString, "+", ""))
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            If IsNumeric(Replace(Text1.Text, "+", "")) = True Then
                                BEGINSTA = CDbl(Replace(Text1.Text, "+", ""))
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "BEGINSTA" Then
                                        Dim Continut As String = attref.TextString
                                        If IsNumeric(Replace(Continut, "+", "")) = True Then
                                            BEGINSTA = CDbl(Replace(Continut, "+", ""))
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            If IsNumeric(Replace(Continut, "+", "")) = True Then
                                BEGINSTA = CDbl(Replace(Continut, "+", ""))

                            End If
                        End If

                    End If


                    If IsNothing(Rezultat_End_sta) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_End_sta.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            If IsNumeric(Replace(Text1.TextString, "+", "")) = True Then
                                ENDSTA = CDbl(Replace(Text1.TextString, "+", ""))
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            If IsNumeric(Replace(Text1.Text, "+", "")) = True Then
                                ENDSTA = CDbl(Replace(Text1.Text, "+", ""))
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "ENDSTA" Then
                                        Dim Continut As String = attref.TextString
                                        If IsNumeric(Replace(Continut, "+", "")) = True Then
                                            ENDSTA = CDbl(Replace(Continut, "+", ""))
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            If IsNumeric(Replace(Continut, "+", "")) = True Then
                                ENDSTA = CDbl(Replace(Continut, "+", ""))

                            End If
                        End If

                    End If

                    If IsNothing(Rezultat_sta) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_sta.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            If IsNumeric(Replace(Text1.TextString, "+", "")) = True Then
                                STA = CDbl(Replace(Text1.TextString, "+", ""))
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            If IsNumeric(Replace(Text1.Text, "+", "")) = True Then
                                STA = CDbl(Replace(Text1.Text, "+", ""))
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "STA" Then
                                        Dim Continut As String = attref.TextString
                                        If IsNumeric(Replace(Continut, "+", "")) = True Then
                                            STA = CDbl(Replace(Continut, "+", ""))
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            If IsNumeric(Replace(Continut, "+", "")) = True Then
                                STA = CDbl(Replace(Continut, "+", ""))

                            End If
                        End If

                    End If







                    Dim Ref_ID As String = ""
                    If Rezultat_desc_ref.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_desc_ref.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Ref_ID = Text1.TextString
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Ref_ID = Text1.Contents
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "ID_NO" Then
                                        If attref.IsMTextAttribute = False Then
                                            Dim Continut As String = attref.TextString
                                            Ref_ID = Continut
                                        Else
                                            Dim Continut As String = attref.MTextAttribute.Contents
                                            Ref_ID = Continut
                                        End If

                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            Ref_ID = Continut
                        End If

                    End If


                    If Rezultat_Material.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_Material.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            If IsNumeric(Text1.TextString) = True Then
                                MAT = CInt(Text1.TextString)
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            If IsNumeric(Text1.Text) = True Then
                                MAT = CInt(Text1.Text)
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "MAT" Then
                                        Dim Continut As String = attref.TextString
                                        If IsNumeric(Continut) = True Then
                                            MAT = CInt(Continut)
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            If IsNumeric(Continut) = True Then
                                MAT = CInt(Continut)
                            End If
                        End If

                    End If



                    If Not STA = -1 And Not BEGINSTA = -1 And Not ENDSTA = -1 Then
                        W1.Range("E" & Index_row_excel).Value = BEGINSTA
                        W1.Range("A" & Index_row_excel & ":J" & Index_row_excel).Interior.Color = 10921638
                        Index_row_excel = Index_row_excel + 1

                        W1.Range("B" & Index_row_excel).Value = "FACILITY"

                        If Not Ref_ID = "" Then
                            W1.Range("D" & Index_row_excel).Value = Ref_ID
                        End If
                        If Not MAT = -1 Then
                            W1.Range("G" & Index_row_excel).Value = MAT
                            Curent_mat = MAT
                        End If

                        W1.Range("J" & Index_row_excel).Value = "FACILITY"
                        W1.Range("F" & Index_row_excel).Value = (ENDSTA - BEGINSTA)
                        W1.Range("E" & Index_row_excel).Value = STA

                        W1.Range("A" & Index_row_excel & ":J" & Index_row_excel).Interior.Color = 10921638
                        Index_row_excel = Index_row_excel + 1
                        W1.Range("E" & Index_row_excel).Value = ENDSTA
                        W1.Range("A" & Index_row_excel & ":J" & Index_row_excel).Interior.Color = 10921638
                        TextBox_ROW_START_XL.Text = (Index_row_excel + 1).ToString
                    End If

                    Trans1.Commit()
                End Using
            End Using

            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub

    Private Sub Button_PICK_elbow_Click(sender As Object, e As EventArgs) Handles Button_ELBOW.Click
        If IsNumeric(TextBox_ROW_START_XL.Text) = False Then
            Exit Sub
        Else
            Index_row_excel = CInt(TextBox_ROW_START_XL.Text)
        End If
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()

        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat_desc As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_desc As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_desc.MessageForAdding = vbLf & "Select Description:"
                    Object_Prompt_desc.SingleOnly = True

                    Rezultat_desc = Editor1.GetSelection(Object_Prompt_desc)
                    If Not Rezultat_desc.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If


                    Dim Rezultat_desc_ref As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_desc_ref As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_desc_ref.MessageForAdding = vbLf & "Select Reference Drawing :"
                    Object_Prompt_desc_ref.SingleOnly = True

                    Rezultat_desc_ref = Editor1.GetSelection(Object_Prompt_desc_ref)

                    Dim Rezultat_sta As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_sta As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_sta.MessageForAdding = vbLf & "Select Station Chainage:"
                    Object_Prompt_sta.SingleOnly = True

                    Rezultat_sta = Editor1.GetSelection(Object_Prompt_sta)
                    If Not Rezultat_sta.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If


                    Dim Rezultat_len As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_len As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_len.MessageForAdding = vbLf & "Select the Length:"
                    Object_Prompt_len.SingleOnly = True

                    Rezultat_len = Editor1.GetSelection(Object_Prompt_len)
                    If Not Rezultat_len.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If


                    Dim Rezultat_Material As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_Material As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_Material.MessageForAdding = vbLf & "Select the Material:"
                    Object_Prompt_Material.SingleOnly = True
                    Rezultat_Material = Editor1.GetSelection(Object_Prompt_Material)
                    If Not Rezultat_Material.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If

                    Dim STA As Double = -1
                    Dim LEN As Double = -1
                    Dim MAT As Integer = -1



                    If IsNothing(Rezultat_sta) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_sta.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            If IsNumeric(Replace(Text1.TextString, "+", "")) = True Then
                                STA = CDbl(Replace(Text1.TextString, "+", ""))
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            If IsNumeric(Replace(Text1.Text, "+", "")) = True Then
                                STA = CDbl(Replace(Text1.Text, "+", ""))
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "STA" Then
                                        Dim Continut As String = attref.TextString
                                        If IsNumeric(Replace(Continut, "+", "")) = True Then
                                            STA = CDbl(Replace(Continut, "+", ""))
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            If IsNumeric(Replace(Continut, "+", "")) = True Then
                                STA = CDbl(Replace(Continut, "+", ""))

                            End If
                        End If

                    End If







                    Dim Ref_ID As String = ""
                    If Rezultat_desc_ref.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_desc_ref.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Ref_ID = Text1.TextString
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Ref_ID = Text1.Contents
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "ID_NO" Then
                                        If attref.IsMTextAttribute = False Then
                                            Dim Continut As String = attref.TextString
                                            Ref_ID = Continut
                                        Else
                                            Dim Continut As String = attref.MTextAttribute.Contents
                                            Ref_ID = Continut
                                        End If

                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            Ref_ID = Continut
                        End If

                    End If


                    If Rezultat_Material.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_Material.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            If IsNumeric(Text1.TextString) = True Then
                                MAT = CInt(Text1.TextString)
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            If IsNumeric(Text1.Text) = True Then
                                MAT = CInt(Text1.Text)
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "MAT" Then
                                        Dim Continut As String = attref.TextString
                                        If IsNumeric(Continut) = True Then
                                            MAT = CInt(Continut)
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            If IsNumeric(Continut) = True Then
                                MAT = CInt(Continut)
                            End If
                        End If

                    End If


                    If Rezultat_len.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_len.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            If IsNumeric(Text1.TextString) = True Then
                                LEN = Round(CDbl(Text1.TextString), 1)
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            If IsNumeric(Text1.Text) = True Then
                                LEN = Round(CDbl(Text1.Text), 1)
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "LENGTH" Then
                                        Dim Continut As String = attref.TextString
                                        If IsNumeric(Continut) = True Then
                                            LEN = Round(CDbl(Continut), 1)
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            If IsNumeric(Continut) = True Then
                                LEN = Round(CDbl(Continut), 1)
                            End If
                        End If

                    End If

                    Dim Description1 As String = ""
                    If IsNothing(Rezultat_desc) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_desc.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Description1 = Text1.TextString
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Description1 = Text1.Contents
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "DESC" Then
                                        If attref.IsMTextAttribute = False Then
                                            Dim Continut As String = attref.TextString
                                            Description1 = Continut
                                        Else
                                            Dim Continut As String = attref.MTextAttribute.Contents
                                            Description1 = Continut
                                        End If

                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            Description1 = Continut
                        End If

                    End If


                    If Not STA = -1 And Not LEN = -1 And Not MAT = -1 And Not Description1 = "" Then
                        W1.Range("E" & Index_row_excel).Value = STA - Ceiling((LEN * 10) / 2) / 10
                        W1.Range("G" & Index_row_excel).Value = MAT
                        W1.Range("A" & Index_row_excel & ":J" & Index_row_excel).Interior.Color = 15523812
                        Index_row_excel = Index_row_excel + 1

                        W1.Range("B" & Index_row_excel).Value = Description1

                        If Not Ref_ID = "" Then
                            W1.Range("D" & Index_row_excel).Value = Ref_ID
                        End If

                        W1.Range("G" & Index_row_excel).Value = MAT
                        W1.Range("F" & Index_row_excel).Value = LEN
                        W1.Range("E" & Index_row_excel).Value = STA
                        Curent_mat = MAT
                        W1.Range("A" & Index_row_excel & ":J" & Index_row_excel).Interior.Color = 15523812
                        Index_row_excel = Index_row_excel + 1
                        W1.Range("E" & Index_row_excel).Value = STA + Floor((LEN * 10) / 2) / 10
                        W1.Range("G" & Index_row_excel).Value = MAT
                        W1.Range("A" & Index_row_excel & ":J" & Index_row_excel).Interior.Color = 15523812
                        TextBox_ROW_START_XL.Text = (Index_row_excel + 1).ToString
                    End If

                    Trans1.Commit()
                End Using
            End Using

            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub

    Private Sub Button_PICK_Screw_anchors_start_Click(sender As Object, e As EventArgs) Handles Button_SA_start.Click
        If IsNumeric(TextBox_ROW_START_XL.Text) = False Then
            Exit Sub
        Else
            Index_row_excel = CInt(TextBox_ROW_START_XL.Text)
        End If
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()

        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat_Start_STA As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_start_sta As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_start_sta.MessageForAdding = vbLf & "Select the Start Chainage:"
                    Object_Prompt_start_sta.SingleOnly = True

                    Rezultat_Start_STA = Editor1.GetSelection(Object_Prompt_start_sta)
                    If Not Rezultat_Start_STA.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If



                    Dim Rezultat_End_sta As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_End_sta As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_End_sta.MessageForAdding = vbLf & "Select the End Chainage:"
                    Object_Prompt_End_sta.SingleOnly = True
                    Rezultat_End_sta = Editor1.GetSelection(Object_Prompt_End_sta)
                    If Not Rezultat_End_sta.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If


                    Dim Rezultat_Number As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_Number As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_Number.MessageForAdding = vbLf & "Select the Number of SA:"
                    Object_Prompt_Number.SingleOnly = True
                    Rezultat_Number = Editor1.GetSelection(Object_Prompt_Number)
                    If Not Rezultat_Number.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If

                    Dim Rezultat_Spacing As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_Spacing As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_Spacing.MessageForAdding = vbLf & "Select the Spacing of SA:"
                    Object_Prompt_Spacing.SingleOnly = True
                    Rezultat_Spacing = Editor1.GetSelection(Object_Prompt_Spacing)
                    If Not Rezultat_Spacing.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If

                    Dim BEGINSTA As Double = -1
                    Dim ENDSTA As Double = -1
                    Dim NO_TYPE As Double = -1
                    Dim SPACING As Double = -1


                    If IsNothing(Rezultat_Start_STA) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_Start_STA.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            If IsNumeric(Replace(Text1.TextString, "+", "")) = True Then
                                BEGINSTA = CDbl(Replace(Text1.TextString, "+", ""))
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            If IsNumeric(Replace(Text1.Text, "+", "")) = True Then
                                BEGINSTA = CDbl(Replace(Text1.Text, "+", ""))
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "BEGINSTA" Then
                                        Dim Continut As String = attref.TextString
                                        If IsNumeric(Replace(Continut, "+", "")) = True Then
                                            BEGINSTA = CDbl(Replace(Continut, "+", ""))
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            If IsNumeric(Replace(Continut, "+", "")) = True Then
                                BEGINSTA = CDbl(Replace(Continut, "+", ""))

                            End If
                        End If

                    End If




                    If IsNothing(Rezultat_End_sta) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_End_sta.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            If IsNumeric(Replace(Text1.TextString, "+", "")) = True Then
                                ENDSTA = CDbl(Replace(Text1.TextString, "+", ""))
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            If IsNumeric(Replace(Text1.Text, "+", "")) = True Then
                                ENDSTA = CDbl(Replace(Text1.Text, "+", ""))
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "ENDSTA" Then
                                        Dim Continut As String = attref.TextString
                                        If IsNumeric(Replace(Continut, "+", "")) = True Then
                                            ENDSTA = CDbl(Replace(Continut, "+", ""))
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            If IsNumeric(Replace(Continut, "+", "")) = True Then
                                ENDSTA = CDbl(Replace(Continut, "+", ""))

                            End If
                        End If

                    End If

                    If IsNothing(Rezultat_Spacing) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_Spacing.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Dim CONTINUT As String = extrage_numar_din_text(Text1.TextString)
                            If IsNumeric(CONTINUT) = True Then
                                SPACING = Round(CDbl(CONTINUT), 1)
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Dim CONTINUT As String = extrage_numar_din_text(Text1.Text)
                            If IsNumeric(CONTINUT) = True Then
                                SPACING = Round(CDbl(CONTINUT), 1)
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "SPACING" Then

                                        Dim CONTINUT As String = extrage_numar_din_text(attref.TextString)
                                        If IsNumeric(CONTINUT) = True Then
                                            SPACING = Round(CDbl(CONTINUT), 1)
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = extrage_numar_din_text(attref.Tag.ToString)
                            If IsNumeric(Continut) = True Then
                                SPACING = Round(CDbl(Continut), 1)
                            End If
                        End If

                    End If

                    If IsNothing(Rezultat_Number) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_Number.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Dim CONTINUT As String = extrage_numar_din_text(Text1.TextString)
                            If IsNumeric(CONTINUT) = True Then
                                NO_TYPE = CInt(CONTINUT)
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Dim CONTINUT As String = extrage_numar_din_text(Text1.Text)
                            If IsNumeric(CONTINUT) = True Then
                                NO_TYPE = CInt(CONTINUT)
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "NO_TYPE" Then

                                        Dim CONTINUT As String = extrage_numar_din_text(attref.TextString)
                                        If IsNumeric(CONTINUT) = True Then
                                            NO_TYPE = CInt(CONTINUT)
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = extrage_numar_din_text(attref.Tag.ToString)
                            If IsNumeric(Continut) = True Then
                                NO_TYPE = CInt(Continut)
                            End If
                        End If

                    End If

                    If Not SPACING = -1 And Not BEGINSTA = -1 And Not ENDSTA = -1 And Not NO_TYPE = -1 And ENDSTA_screw_anchors = -1 And Spacing_screw_anchors = "" Then
                        If BEGINSTA > ENDSTA Then
                            Dim tEMP As Double = BEGINSTA
                            BEGINSTA = ENDSTA
                            ENDSTA = tEMP
                        End If
                        ENDSTA_screw_anchors = ENDSTA
                        W1.Range("B" & Index_row_excel).Value = "SCREW ANCHOR START"
                        W1.Range("E" & Index_row_excel).Value = BEGINSTA
                        W1.Range("G" & Index_row_excel).Value = "SA1"
                        W1.Range("J" & Index_row_excel).Value = NO_TYPE & " SA" & " - " & SPACING & " C\C"
                        Spacing_screw_anchors = NO_TYPE & " SA" & " - " & SPACING & " C\C"
                        W1.Range("A" & Index_row_excel & ":J" & Index_row_excel).Interior.Color = 5540500
                        TextBox_ROW_START_XL.Text = (Index_row_excel + 1).ToString

                    End If
                    Trans1.Commit()
                End Using
            End Using

            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
        Button_SA_end.Visible = True
        Button_SA_start.Visible = False

    End Sub

    Private Sub Button_PICK_Screw_anchors_end_Click(sender As Object, e As EventArgs) Handles Button_SA_end.Click
        If IsNumeric(TextBox_ROW_START_XL.Text) = False Then
            Exit Sub
        Else
            Index_row_excel = CInt(TextBox_ROW_START_XL.Text)
        End If

        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()

        ascunde_butoanele_pentru_forms(Me, Colectie1)

        W1.Range("B" & Index_row_excel).Value = "SCREW ANCHOR END"
        W1.Range("E" & Index_row_excel).Value = ENDSTA_screw_anchors
        W1.Range("G" & Index_row_excel).Value = "SA2"
        W1.Range("J" & Index_row_excel).Value = Spacing_screw_anchors
        Spacing_screw_anchors = ""
        ENDSTA_screw_anchors = -1
        W1.Range("A" & Index_row_excel & ":J" & Index_row_excel).Interior.Color = 5540500
        TextBox_ROW_START_XL.Text = (Index_row_excel + 1).ToString

        afiseaza_butoanele_pentru_forms(Me, Colectie1)
        Button_SA_end.Visible = False
        Button_SA_start.Visible = True

    End Sub

    Private Sub Button_matchline_Click(sender As Object, e As EventArgs) Handles Button_matchline.Click
        If IsNumeric(TextBox_ROW_START_XL.Text) = False Then
            Exit Sub
        Else
            Index_row_excel = CInt(TextBox_ROW_START_XL.Text)
        End If
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()

        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat_Start_STA As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_start_sta As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_start_sta.MessageForAdding = vbLf & "Select the Chainage:"
                    Object_Prompt_start_sta.SingleOnly = True

                    Rezultat_Start_STA = Editor1.GetSelection(Object_Prompt_start_sta)
                    If Not Rezultat_Start_STA.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If


                    Dim Rezultat_Material As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_Material As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_Material.MessageForAdding = vbLf & "Select the Material:"
                    Object_Prompt_Material.SingleOnly = True
                    Rezultat_Material = Editor1.GetSelection(Object_Prompt_Material)
                    If Not Rezultat_Material.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If


                    Dim BEGINSTA As Double = -1
                    Dim MAT As Double = -1


                    If IsNothing(Rezultat_Material) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_Material.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            If IsNumeric(Text1.TextString) = True Then
                                MAT = Round(CDbl(Text1.TextString), 0)
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            If IsNumeric(Text1.Text) = True Then
                                MAT = Round(CDbl(Text1.Text), 0)
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "MAT" Then
                                        Dim Continut As String = attref.TextString
                                        If IsNumeric(Continut) = True Then
                                            MAT = Round(CDbl(Continut), 0)
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            If IsNumeric(Continut) = True Then
                                MAT = Round(CDbl(Continut), 0)
                            End If
                        End If

                    End If

                    If IsNothing(Rezultat_Start_STA) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_Start_STA.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            If IsNumeric(Replace(Text1.TextString, "+", "")) = True Then
                                BEGINSTA = CDbl(Replace(Text1.TextString, "+", ""))
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            If IsNumeric(Replace(Text1.Text, "+", "")) = True Then
                                BEGINSTA = CDbl(Replace(Text1.Text, "+", ""))
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "BEGINSTA" Then
                                        Dim Continut As String = attref.TextString
                                        If IsNumeric(Replace(Continut, "+", "")) = True Then
                                            BEGINSTA = CDbl(Replace(Continut, "+", ""))
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            If IsNumeric(Replace(Continut, "+", "")) = True Then
                                BEGINSTA = CDbl(Replace(Continut, "+", ""))

                            End If
                        End If

                    End If


                    If Not BEGINSTA = -1 And Not MAT = -1 Then

                        W1.Range("G" & Index_row_excel).Value = MAT
                        W1.Range("A" & Index_row_excel & ":J" & Index_row_excel).Interior.Color = 12040422
                        Index_row_excel = Index_row_excel + 1
                        Curent_mat = MAT

                        W1.Range("G" & Index_row_excel).Value = "M"
                        W1.Range("J" & Index_row_excel).Value = "END OF SHEET"
                        W1.Range("E" & Index_row_excel).Value = BEGINSTA
                        W1.Range("A" & Index_row_excel & ":J" & Index_row_excel).Interior.Color = 5287936
                        TextBox_ROW_START_XL.Text = (Index_row_excel + 1).ToString
                    End If

                    Trans1.Commit()
                End Using
            End Using

            Button_SA_start.Visible = True
            Button_SA_end.Visible = False
            Button_RW_start.Visible = True
            Button_RW_end.Visible = False
            Button_SBW_start.Visible = True
            Button_SBW_End.Visible = False

            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub

    Private Sub Alignment_w2XL_Form_Load(sender As Object, e As EventArgs) Handles Me.Load
        Button_SA_end.Visible = False
        Button_SA_start.Visible = True
    End Sub

    Private Sub Button_PICK_rw_start_Click(sender As Object, e As EventArgs) Handles Button_RW_start.Click
        If IsNumeric(TextBox_ROW_START_XL.Text) = False Then
            Exit Sub
        Else
            Index_row_excel = CInt(TextBox_ROW_START_XL.Text)
        End If
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()

        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat_Start_STA As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_start_sta As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_start_sta.MessageForAdding = vbLf & "Select the Start Chainage:"
                    Object_Prompt_start_sta.SingleOnly = True

                    Rezultat_Start_STA = Editor1.GetSelection(Object_Prompt_start_sta)
                    If Not Rezultat_Start_STA.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If



                    Dim Rezultat_End_sta As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_End_sta As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_End_sta.MessageForAdding = vbLf & "Select the End Chainage:"
                    Object_Prompt_End_sta.SingleOnly = True
                    Rezultat_End_sta = Editor1.GetSelection(Object_Prompt_End_sta)
                    If Not Rezultat_End_sta.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If


                    Dim Rezultat_Number As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_Number As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_Number.MessageForAdding = vbLf & "Select the Number of RW:"
                    Object_Prompt_Number.SingleOnly = True
                    Rezultat_Number = Editor1.GetSelection(Object_Prompt_Number)
                    If Not Rezultat_Number.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If

                    Dim Rezultat_Spacing As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_Spacing As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_Spacing.MessageForAdding = vbLf & "Select the Spacing of RW:"
                    Object_Prompt_Spacing.SingleOnly = True
                    Rezultat_Spacing = Editor1.GetSelection(Object_Prompt_Spacing)
                    If Not Rezultat_Spacing.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If

                    Dim BEGINSTA As Double = -1
                    Dim ENDSTA As Double = -1
                    Dim NO_TYPE As Double = -1
                    Dim SPACING As Double = -1


                    If IsNothing(Rezultat_Start_STA) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_Start_STA.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            If IsNumeric(Replace(Text1.TextString, "+", "")) = True Then
                                BEGINSTA = CDbl(Replace(Text1.TextString, "+", ""))
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            If IsNumeric(Replace(Text1.Text, "+", "")) = True Then
                                BEGINSTA = CDbl(Replace(Text1.Text, "+", ""))
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "BEGINSTA" Then
                                        Dim Continut As String = attref.TextString
                                        If IsNumeric(Replace(Continut, "+", "")) = True Then
                                            BEGINSTA = CDbl(Replace(Continut, "+", ""))
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            If IsNumeric(Replace(Continut, "+", "")) = True Then
                                BEGINSTA = CDbl(Replace(Continut, "+", ""))

                            End If
                        End If

                    End If




                    If IsNothing(Rezultat_End_sta) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_End_sta.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            If IsNumeric(Replace(Text1.TextString, "+", "")) = True Then
                                ENDSTA = CDbl(Replace(Text1.TextString, "+", ""))
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            If IsNumeric(Replace(Text1.Text, "+", "")) = True Then
                                ENDSTA = CDbl(Replace(Text1.Text, "+", ""))
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "ENDSTA" Then
                                        Dim Continut As String = attref.TextString
                                        If IsNumeric(Replace(Continut, "+", "")) = True Then
                                            ENDSTA = CDbl(Replace(Continut, "+", ""))
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            If IsNumeric(Replace(Continut, "+", "")) = True Then
                                ENDSTA = CDbl(Replace(Continut, "+", ""))

                            End If
                        End If

                    End If

                    If IsNothing(Rezultat_Spacing) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_Spacing.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Dim CONTINUT As String = extrage_numar_din_text(Text1.TextString)
                            If IsNumeric(CONTINUT) = True Then
                                SPACING = Round(CDbl(CONTINUT), 1)
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Dim CONTINUT As String = extrage_numar_din_text(Text1.Text)
                            If IsNumeric(CONTINUT) = True Then
                                SPACING = Round(CDbl(CONTINUT), 1)
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "SPACING" Then

                                        Dim CONTINUT As String = extrage_numar_din_text(attref.TextString)
                                        If IsNumeric(CONTINUT) = True Then
                                            SPACING = Round(CDbl(CONTINUT), 1)
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = extrage_numar_din_text(attref.Tag.ToString)
                            If IsNumeric(Continut) = True Then
                                SPACING = Round(CDbl(Continut), 1)
                            End If
                        End If

                    End If

                    If IsNothing(Rezultat_Number) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_Number.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Dim CONTINUT As String = extrage_numar_din_text(Text1.TextString)
                            If IsNumeric(CONTINUT) = True Then
                                NO_TYPE = CInt(CONTINUT)
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Dim CONTINUT As String = extrage_numar_din_text(Text1.Text)
                            If IsNumeric(CONTINUT) = True Then
                                NO_TYPE = CInt(CONTINUT)
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "NO_TYPE" Then

                                        Dim CONTINUT As String = extrage_numar_din_text(attref.TextString)
                                        If IsNumeric(CONTINUT) = True Then
                                            NO_TYPE = CInt(CONTINUT)
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = extrage_numar_din_text(attref.Tag.ToString)
                            If IsNumeric(Continut) = True Then
                                NO_TYPE = CInt(Continut)
                            End If
                        End If

                    End If

                    If Not SPACING = -1 And Not BEGINSTA = -1 And Not ENDSTA = -1 And Not NO_TYPE = -1 And ENDSTA_screw_anchors = -1 And Spacing_screw_anchors = "" Then
                        If BEGINSTA > ENDSTA Then
                            Dim tEMP As Double = BEGINSTA
                            BEGINSTA = ENDSTA
                            ENDSTA = tEMP
                        End If
                        ENDSTA_RW = ENDSTA
                        W1.Range("B" & Index_row_excel).Value = "RW START"
                        W1.Range("E" & Index_row_excel).Value = BEGINSTA
                        W1.Range("G" & Index_row_excel).Value = "RW1"
                        W1.Range("J" & Index_row_excel).Value = NO_TYPE & " RW" & " - " & SPACING & " C\C"
                        Spacing_RW = NO_TYPE & " RW" & " - " & SPACING & " C\C"
                        W1.Range("A" & Index_row_excel & ":J" & Index_row_excel).Interior.Color = 5540500
                        TextBox_ROW_START_XL.Text = (Index_row_excel + 1).ToString

                    End If
                    Trans1.Commit()
                End Using
            End Using

            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
        Button_RW_end.Visible = True
        Button_RW_start.Visible = False

    End Sub

    Private Sub Button_PICK_RW_end_Click(sender As Object, e As EventArgs) Handles Button_RW_end.Click
        If IsNumeric(TextBox_ROW_START_XL.Text) = False Then
            Exit Sub
        Else
            Index_row_excel = CInt(TextBox_ROW_START_XL.Text)
        End If

        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()

        ascunde_butoanele_pentru_forms(Me, Colectie1)

        W1.Range("B" & Index_row_excel).Value = "RW END"
        W1.Range("E" & Index_row_excel).Value = ENDSTA_RW
        W1.Range("G" & Index_row_excel).Value = "RW2"
        W1.Range("J" & Index_row_excel).Value = Spacing_RW
        Spacing_RW = ""
        ENDSTA_RW = -1
        W1.Range("A" & Index_row_excel & ":J" & Index_row_excel).Interior.Color = 5540500
        TextBox_ROW_START_XL.Text = (Index_row_excel + 1).ToString

        afiseaza_butoanele_pentru_forms(Me, Colectie1)
        Button_RW_end.Visible = False
        Button_RW_start.Visible = True

    End Sub

    Private Sub Button_PICK_SBW_start_Click(sender As Object, e As EventArgs) Handles Button_SBW_start.Click
        If IsNumeric(TextBox_ROW_START_XL.Text) = False Then
            Exit Sub
        Else
            Index_row_excel = CInt(TextBox_ROW_START_XL.Text)
        End If
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()

        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat_Start_STA As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_start_sta As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_start_sta.MessageForAdding = vbLf & "Select the Start Chainage:"
                    Object_Prompt_start_sta.SingleOnly = True

                    Rezultat_Start_STA = Editor1.GetSelection(Object_Prompt_start_sta)
                    If Not Rezultat_Start_STA.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If



                    Dim Rezultat_End_sta As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_End_sta As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_End_sta.MessageForAdding = vbLf & "Select the End Chainage:"
                    Object_Prompt_End_sta.SingleOnly = True
                    Rezultat_End_sta = Editor1.GetSelection(Object_Prompt_End_sta)
                    If Not Rezultat_End_sta.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If


                    Dim Rezultat_Number As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_Number As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_Number.MessageForAdding = vbLf & "Select the Number of SBW:"
                    Object_Prompt_Number.SingleOnly = True
                    Rezultat_Number = Editor1.GetSelection(Object_Prompt_Number)
                    If Not Rezultat_Number.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If

                    Dim Rezultat_Spacing As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_Spacing As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_Spacing.MessageForAdding = vbLf & "Select the Spacing of SBW:"
                    Object_Prompt_Spacing.SingleOnly = True
                    Rezultat_Spacing = Editor1.GetSelection(Object_Prompt_Spacing)
                    If Not Rezultat_Spacing.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If

                    Dim BEGINSTA As Double = -1
                    Dim ENDSTA As Double = -1
                    Dim NO_TYPE As Double = -1
                    Dim SPACING As Double = -1


                    If IsNothing(Rezultat_Start_STA) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_Start_STA.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            If IsNumeric(Replace(Text1.TextString, "+", "")) = True Then
                                BEGINSTA = CDbl(Replace(Text1.TextString, "+", ""))
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            If IsNumeric(Replace(Text1.Text, "+", "")) = True Then
                                BEGINSTA = CDbl(Replace(Text1.Text, "+", ""))
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "BEGINSTA" Then
                                        Dim Continut As String = attref.TextString
                                        If IsNumeric(Replace(Continut, "+", "")) = True Then
                                            BEGINSTA = CDbl(Replace(Continut, "+", ""))
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            If IsNumeric(Replace(Continut, "+", "")) = True Then
                                BEGINSTA = CDbl(Replace(Continut, "+", ""))

                            End If
                        End If

                    End If




                    If IsNothing(Rezultat_End_sta) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_End_sta.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            If IsNumeric(Replace(Text1.TextString, "+", "")) = True Then
                                ENDSTA = CDbl(Replace(Text1.TextString, "+", ""))
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            If IsNumeric(Replace(Text1.Text, "+", "")) = True Then
                                ENDSTA = CDbl(Replace(Text1.Text, "+", ""))
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "ENDSTA" Then
                                        Dim Continut As String = attref.TextString
                                        If IsNumeric(Replace(Continut, "+", "")) = True Then
                                            ENDSTA = CDbl(Replace(Continut, "+", ""))
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = attref.Tag.ToString
                            If IsNumeric(Replace(Continut, "+", "")) = True Then
                                ENDSTA = CDbl(Replace(Continut, "+", ""))

                            End If
                        End If

                    End If

                    If IsNothing(Rezultat_Spacing) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_Spacing.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Dim CONTINUT As String = extrage_numar_din_text(Text1.TextString)
                            If IsNumeric(CONTINUT) = True Then
                                SPACING = Round(CDbl(CONTINUT), 1)
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Dim CONTINUT As String = extrage_numar_din_text(Text1.Text)
                            If IsNumeric(CONTINUT) = True Then
                                SPACING = Round(CDbl(CONTINUT), 1)
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "SPACING" Then

                                        Dim CONTINUT As String = extrage_numar_din_text(attref.TextString)
                                        If IsNumeric(CONTINUT) = True Then
                                            SPACING = Round(CDbl(CONTINUT), 1)
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = extrage_numar_din_text(attref.Tag.ToString)
                            If IsNumeric(Continut) = True Then
                                SPACING = Round(CDbl(Continut), 1)
                            End If
                        End If

                    End If

                    If IsNothing(Rezultat_Number) = False Then
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_Number.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Dim CONTINUT As String = extrage_numar_din_text(Text1.TextString)
                            If IsNumeric(CONTINUT) = True Then
                                NO_TYPE = CInt(CONTINUT)
                            End If
                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim Text1 As MText = Ent1
                            Dim CONTINUT As String = extrage_numar_din_text(Text1.Text)
                            If IsNumeric(CONTINUT) = True Then
                                NO_TYPE = CInt(CONTINUT)
                            End If
                        End If


                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1
                            If Block1.AttributeCollection.Count > 0 Then
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                    If attref.Tag.ToUpper = "NO_TYPE" Then

                                        Dim CONTINUT As String = extrage_numar_din_text(attref.TextString)
                                        If IsNumeric(CONTINUT) = True Then
                                            NO_TYPE = CInt(CONTINUT)
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim attref As AttributeDefinition = Ent1
                            Dim Continut As String = extrage_numar_din_text(attref.Tag.ToString)
                            If IsNumeric(Continut) = True Then
                                NO_TYPE = CInt(Continut)
                            End If
                        End If

                    End If

                    If Not SPACING = -1 And Not BEGINSTA = -1 And Not ENDSTA = -1 And Not NO_TYPE = -1 And ENDSTA_screw_anchors = -1 And Spacing_screw_anchors = "" Then
                        If BEGINSTA > ENDSTA Then
                            Dim tEMP As Double = BEGINSTA
                            BEGINSTA = ENDSTA
                            ENDSTA = tEMP
                        End If
                        ENDSTA_SBW = ENDSTA
                        W1.Range("B" & Index_row_excel).Value = "SBW START"
                        W1.Range("E" & Index_row_excel).Value = BEGINSTA
                        W1.Range("G" & Index_row_excel).Value = "SBW1"
                        W1.Range("J" & Index_row_excel).Value = NO_TYPE & " SBW" & " - " & SPACING & " C\C"
                        Spacing_SBW = NO_TYPE & " SBW" & " - " & SPACING & " C\C"
                        W1.Range("A" & Index_row_excel & ":J" & Index_row_excel).Interior.Color = 5540500
                        TextBox_ROW_START_XL.Text = (Index_row_excel + 1).ToString

                    End If
                    Trans1.Commit()
                End Using
            End Using

            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
        Button_SBW_end.Visible = True
        Button_SBW_start.Visible = False

    End Sub

    Private Sub Button_PICK_SBW_end_Click(sender As Object, e As EventArgs) Handles Button_SBW_End.Click
        If IsNumeric(TextBox_ROW_START_XL.Text) = False Then
            Exit Sub
        Else
            Index_row_excel = CInt(TextBox_ROW_START_XL.Text)
        End If

        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()

        ascunde_butoanele_pentru_forms(Me, Colectie1)

        W1.Range("B" & Index_row_excel).Value = "SBW END"
        W1.Range("E" & Index_row_excel).Value = ENDSTA_SBW
        W1.Range("G" & Index_row_excel).Value = "SBW2"
        W1.Range("J" & Index_row_excel).Value = Spacing_SBW
        Spacing_SBW = ""
        ENDSTA_SBW = -1
        W1.Range("A" & Index_row_excel & ":J" & Index_row_excel).Interior.Color = 5540500
        TextBox_ROW_START_XL.Text = (Index_row_excel + 1).ToString

        afiseaza_butoanele_pentru_forms(Me, Colectie1)
        Button_SBW_end.Visible = False
        Button_SBW_start.Visible = True

    End Sub



End Class