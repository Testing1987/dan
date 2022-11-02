Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Public Class Chainage_operations_form

    Dim Colectie1 As New Specialized.StringCollection

    Private Sub Button_pick_chainage_Click(sender As Object, e As EventArgs) Handles Button_pick_chainage.Click
        ascunde_butoanele_pentru_forms(Me, Colectie1)
        TextBox_chainage_high.Text = ""
        TextBox_chainage_low.Text = ""
        TextBox_chainage_middle.Text = ""
        Try
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
            Editor1 = ThisDrawing.Editor

            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

            Using lock As DocumentLock = ThisDrawing.LockDocument
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select objects:"

                Object_Prompt.SingleOnly = False
                Rezultat1 = Editor1.GetSelection(Object_Prompt)

                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    If IsNothing(Rezultat1) = False Then
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                            BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                            If Rezultat1.Value.Count = 1 Then
                                For i = 1 To Rezultat1.Value.Count
                                    Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                    Obj1 = Rezultat1.Value.Item(i - 1)
                                    Dim Ent1 As Entity
                                    Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                                    If TypeOf Ent1 Is DBText Then
                                        Dim Text1 As DBText = Ent1
                                        If Text1.TextString.Contains("+") Then
                                            Dim Text_chainage As String = extrage_chainage_din_text_de_la_inceputul_textului(Text1.TextString)
                                            TextBox_chainage_middle.Text = Text_chainage
                                        End If
                                    End If

                                    If TypeOf Ent1 Is MText Then
                                        Dim mText1 As MText = Ent1
                                        If mText1.Contents.Contains("+") Then
                                            Dim Text_chainage As String = extrage_chainage_din_text_de_la_inceputul_textului(mText1.Text)
                                            TextBox_chainage_middle.Text = Text_chainage
                                        End If
                                    End If


                                    If TypeOf Ent1 Is MLeader Then
                                        Dim Mleader1 As MLeader = Ent1
                                        If Mleader1.MText.Contents.Contains("+") Then
                                            Dim Text_chainage As String = extrage_chainage_din_text_de_la_inceputul_textului(Mleader1.MText.Text)
                                            TextBox_chainage_middle.Text = Text_chainage
                                        End If
                                    End If

                                    If TypeOf Ent1 Is BlockReference Then
                                        Dim Block1 As BlockReference = Ent1
                                        If Block1.AttributeCollection.Count > 0 Then
                                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                            For Each id In attColl
                                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                                Dim Continut As String = attref.TextString
                                                If Continut.Contains("+") Then
                                                    Dim Text_chainage As String = extrage_chainage_din_text_de_la_inceputul_textului(Continut)
                                                    TextBox_chainage_middle.Text = Text_chainage
                                                End If

                                            Next
                                        End If
                                    End If
                                    If TypeOf Ent1 Is AttributeDefinition Then
                                        Dim attref As AttributeDefinition = Ent1
                                        Dim Continut As String = attref.Tag.ToString
                                        If Continut.Contains("+") Then
                                            Dim Text_chainage As String = extrage_chainage_din_text_de_la_inceputul_textului(Continut)
                                           TextBox_chainage_middle.Text = Text_chainage
                                        End If
                                    End If


                                Next
                            End If

                            If Rezultat1.Value.Count = 2 Then
                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Rezultat1.Value.Item(0)
                                Dim Ent1 As Entity
                                Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                                Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj2 = Rezultat1.Value.Item(1)
                                Dim Ent2 As Entity
                                Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)

                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.DBText And TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                                    Dim Text1 As DBText = Ent1
                                    Dim Text2 As DBText = Ent2
                                    Dim Chainage1, Chainage2 As Double
                                    If IsNumeric(Replace(Text1.TextString, "+", "")) = True Then
                                        Chainage1 = CDbl(Replace(Text1.TextString, "+", ""))
                                        If IsNumeric(Replace(Text2.TextString, "+", "")) = True Then
                                            Chainage2 = CDbl(Replace(Text2.TextString, "+", ""))
                                            TextBox_chainage_high.Text = Get_chainage_from_double(Chainage1, 1)
                                            TextBox_chainage_low.Text = Get_chainage_from_double(Chainage2, 1)
                                        End If
                                    End If
                                End If

                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.MText And TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                                    Dim mText1 As MText = Ent1
                                    If IsNumeric(mText1.Text) = True Then
                                        Dim mText2 As MText = Ent2
                                        Dim Chainage1, Chainage2 As Double
                                        If IsNumeric(Replace(mText1.Text, "+", "")) = True Then
                                            Chainage1 = CDbl(Replace(mText1.Text, "+", ""))
                                            If IsNumeric(Replace(mText2.Text, "+", "")) = True Then
                                                Chainage2 = CDbl(Replace(mText2.Text, "+", ""))
                                                TextBox_chainage_high.Text = Get_chainage_from_double(Chainage1, 1)
                                                TextBox_chainage_low.Text = Get_chainage_from_double(Chainage2, 1)
                                            End If
                                        End If
                                    End If
                                End If

                                If TypeOf Ent1 Is BlockReference And TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                                    Dim Block1 As BlockReference = Ent1
                                    If Block1.AttributeCollection.Count > 0 Then
                                        Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                        Dim Chainage1, Chainage2 As Double

                                        For Each id In attColl
                                            Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)

                                            If attref.Tag.ToUpper = "BEGINSTA" Then
                                                Dim Continut As String = attref.TextString
                                                If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                    Chainage1 = Replace(Continut, "+", "")
                                                End If
                                            End If

                                            If attref.Tag.ToUpper = "ENDSTA" Then
                                                Dim Continut As String = attref.TextString
                                                If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                    Chainage2 = Replace(Continut, "+", "")
                                                End If
                                            End If
                                        Next


                                        Dim Text2 As DBText = Ent2

                                        Dim Chainage3 As Double
                                        If IsNumeric(Replace(Text2.TextString, "+", "")) = True Then
                                            Chainage3 = CDbl(Replace(Text2.TextString, "+", ""))

                                            If Block1.Position.X < Text2.Position.X Then
                                                TextBox_chainage_high.Text = Get_chainage_from_double(Chainage2, 1)
                                                TextBox_chainage_low.Text = Get_chainage_from_double(Chainage3, 1)
                                            Else

                                                TextBox_chainage_high.Text = Get_chainage_from_double(Chainage1, 1)
                                                TextBox_chainage_low.Text = Get_chainage_from_double(Chainage3, 1)
                                            End If

                                        End If


                                    End If
                                End If

                                If TypeOf Ent2 Is BlockReference And TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                                    Dim Block1 As BlockReference = Ent2
                                    If Block1.AttributeCollection.Count > 0 Then
                                        Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                        Dim Chainage1, Chainage2 As Double

                                        For Each id In attColl
                                            Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)

                                            If attref.Tag.ToUpper = "BEGINSTA" Then
                                                Dim Continut As String = attref.TextString
                                                If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                    Chainage1 = Replace(Continut, "+", "")
                                                End If
                                            End If

                                            If attref.Tag.ToUpper = "ENDSTA" Then
                                                Dim Continut As String = attref.TextString
                                                If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                    Chainage2 = Replace(Continut, "+", "")
                                                End If
                                            End If
                                        Next


                                        Dim Text2 As DBText = Ent1

                                        Dim Chainage3 As Double
                                        If IsNumeric(Replace(Text2.TextString, "+", "")) = True Then
                                            Chainage3 = CDbl(Replace(Text2.TextString, "+", ""))

                                            If Block1.Position.X < Text2.Position.X Then

                                                TextBox_chainage_high.Text = Get_chainage_from_double(Chainage2, 1)
                                                TextBox_chainage_low.Text = Get_chainage_from_double(Chainage3, 1)
                                            Else

                                                TextBox_chainage_high.Text = Get_chainage_from_double(Chainage1, 1)
                                                TextBox_chainage_low.Text = Get_chainage_from_double(Chainage3, 1)
                                            End If


                                        End If

                                    End If
                                End If


                                If TypeOf Ent1 Is BlockReference And TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                                    Dim Block1 As BlockReference = Ent1
                                    If Block1.AttributeCollection.Count > 0 Then
                                        Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                        Dim Chainage1, Chainage2 As Double

                                        For Each id In attColl
                                            Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)

                                            If attref.Tag.ToUpper = "BEGINSTA" Then
                                                Dim Continut As String = attref.TextString
                                                If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                    Chainage1 = Replace(Continut, "+", "")
                                                End If
                                            End If

                                            If attref.Tag.ToUpper = "ENDSTA" Then
                                                Dim Continut As String = attref.TextString
                                                If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                    Chainage2 = Replace(Continut, "+", "")
                                                End If
                                            End If
                                        Next


                                        Dim MText2 As MText = Ent2

                                        Dim Chainage3 As Double
                                        If IsNumeric(Replace(MText2.Text, "+", "")) = True Then
                                            Chainage3 = CDbl(Replace(MText2.Text, "+", ""))

                                            If Block1.Position.X < MText2.Location.X Then

                                                TextBox_chainage_high.Text = Get_chainage_from_double(Chainage2, 1)
                                                TextBox_chainage_low.Text = Get_chainage_from_double(Chainage3, 1)
                                            Else

                                                TextBox_chainage_high.Text = Get_chainage_from_double(Chainage1, 1)
                                                TextBox_chainage_low.Text = Get_chainage_from_double(Chainage3, 1)
                                            End If

                                        End If

                                    End If
                                End If

                                If TypeOf Ent2 Is BlockReference And TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                                    Dim Block1 As BlockReference = Ent2
                                    If Block1.AttributeCollection.Count > 0 Then
                                        Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                        Dim Chainage1, Chainage2 As Double

                                        For Each id In attColl
                                            Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)

                                            If attref.Tag.ToUpper = "BEGINSTA" Then
                                                Dim Continut As String = attref.TextString
                                                If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                    Chainage1 = Replace(Continut, "+", "")
                                                End If
                                            End If

                                            If attref.Tag.ToUpper = "ENDSTA" Then
                                                Dim Continut As String = attref.TextString
                                                If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                    Chainage2 = Replace(Continut, "+", "")
                                                End If
                                            End If
                                        Next


                                        Dim MText2 As MText = Ent1

                                        Dim Chainage3 As Double
                                        If IsNumeric(Replace(MText2.Text, "+", "")) = True Then
                                            Chainage3 = CDbl(Replace(MText2.Text, "+", ""))

                                            If Block1.Position.X < MText2.Location.X Then

                                                TextBox_chainage_high.Text = Get_chainage_from_double(Chainage2, 1)
                                                TextBox_chainage_low.Text = Get_chainage_from_double(Chainage3, 1)
                                            Else

                                                TextBox_chainage_high.Text = Get_chainage_from_double(Chainage1, 1)
                                                TextBox_chainage_low.Text = Get_chainage_from_double(Chainage3, 1)
                                            End If


                                        End If

                                    End If
                                End If

                                If TypeOf Ent1 Is BlockReference And TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.BlockReference Then
                                    Dim Block1 As BlockReference = Ent1
                                    If Block1.AttributeCollection.Count > 0 Then
                                        Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                        Dim Chainage1, Chainage2 As Double

                                        For Each id In attColl
                                            Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)

                                            If attref.Tag.ToUpper = "BEGINSTA" Then
                                                Dim Continut As String = attref.TextString
                                                If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                    Chainage1 = Replace(Continut, "+", "")
                                                End If
                                            End If

                                            If attref.Tag.ToUpper = "ENDSTA" Then
                                                Dim Continut As String = attref.TextString
                                                If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                    Chainage2 = Replace(Continut, "+", "")
                                                End If
                                            End If
                                        Next



                                        Dim Block2 As BlockReference = Ent2
                                        If Block2.AttributeCollection.Count > 0 Then
                                            Dim attColl2 As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block2.AttributeCollection
                                            Dim Chainage3, Chainage4 As Double

                                            For Each id In attColl2
                                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)

                                                If attref.Tag.ToUpper = "BEGINSTA" Then
                                                    Dim Continut As String = attref.TextString
                                                    If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                        Chainage3 = Replace(Continut, "+", "")
                                                    End If
                                                End If

                                                If attref.Tag.ToUpper = "ENDSTA" Then
                                                    Dim Continut As String = attref.TextString
                                                    If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                        Chainage4 = Replace(Continut, "+", "")
                                                    End If
                                                End If
                                            Next


                                            If Block1.Position.X < Block2.Position.X Then

                                                TextBox_chainage_high.Text = Get_chainage_from_double(Chainage2, 1)
                                                TextBox_chainage_low.Text = Get_chainage_from_double(Chainage3, 1)
                                            Else

                                                TextBox_chainage_high.Text = Get_chainage_from_double(Chainage1, 1)
                                                TextBox_chainage_low.Text = Get_chainage_from_double(Chainage4, 1)
                                            End If

                                        End If

                                    End If
                                End If


                            End If






                            Editor1.Regen()
                            Trans1.Commit()
                        End Using

                    End If
                End If





                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            End Using




        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub

    Private Sub Button_split_Click(sender As Object, e As EventArgs) Handles Button_split.Click
        Try
            TextBox_chainage_high.Text = ""
            TextBox_chainage_low.Text = ""

            ascunde_butoanele_pentru_forms(Me, Colectie1)
            If Not TextBox_amount.Text = "" Then
                If IsNumeric(TextBox_amount.Text) = True Then
                    Dim Amount As Double = Round(CDbl(TextBox_amount.Text), 1)

                    If IsNumeric(Replace(TextBox_chainage_middle.Text, "+", "")) = True Then
                        Dim Mijloc As Double = Replace(TextBox_chainage_middle.Text, "+", "")
                        Dim Jumate1 As Double
                        Dim Jumate2 As Double
                        If Not Round(Amount / 2, 1) = Amount / 2 Then
                            Jumate1 = Floor(Amount * 10 / 2) / 10
                            Jumate2 = Ceiling(Amount * 10 / 2) / 10
                        Else
                            Jumate1 = Amount / 2
                            Jumate2 = Jumate1
                        End If
                        TextBox_chainage_low.Text = Get_chainage_from_double(Mijloc - Jumate1, 1)
                        TextBox_chainage_high.Text = Get_chainage_from_double(Mijloc + Jumate2, 1)
                    End If


                End If
            End If
        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub

    Private Sub Button_plus_Click(sender As Object, e As EventArgs) Handles Button_plus.Click
        Try
            TextBox_chainage_high.Text = ""
            TextBox_chainage_low.Text = ""

            ascunde_butoanele_pentru_forms(Me, Colectie1)
            If Not TextBox_amount.Text = "" Then
                If IsNumeric(TextBox_amount.Text) = True Then
                    Dim Amount As Double = Round(CDbl(TextBox_amount.Text), 1)

                    If IsNumeric(Replace(TextBox_chainage_middle.Text, "+", "")) = True Then
                        Dim Mijloc As Double = Replace(TextBox_chainage_middle.Text, "+", "")
                        

                        TextBox_chainage_high.Text = Get_chainage_from_double(Mijloc + Amount, 1)
                    End If


                End If
            End If
        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub
    Private Sub Button_MINUS_Click(sender As Object, e As EventArgs) Handles Button_minus.Click
        Try
            TextBox_chainage_high.Text = ""
            TextBox_chainage_low.Text = ""

            ascunde_butoanele_pentru_forms(Me, Colectie1)
            If Not TextBox_amount.Text = "" Then
                If IsNumeric(TextBox_amount.Text) = True Then
                    Dim Amount As Double = Round(CDbl(TextBox_amount.Text), 1)

                    If IsNumeric(Replace(TextBox_chainage_middle.Text, "+", "")) = True Then
                        Dim Mijloc As Double = Replace(TextBox_chainage_middle.Text, "+", "")


                        TextBox_chainage_low.Text = Get_chainage_from_double(Mijloc - Amount, 1)
                    End If


                End If
            End If
        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub

    Private Sub Button_pick_amount_Click(sender As Object, e As EventArgs) Handles Button_pick_amount.Click
        ascunde_butoanele_pentru_forms(Me, Colectie1)
        TextBox_amount.Text = ""
        Try
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
            Editor1 = ThisDrawing.Editor

            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

            Using lock As DocumentLock = ThisDrawing.LockDocument
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select objects:"

                Object_Prompt.SingleOnly = False
                Rezultat1 = Editor1.GetSelection(Object_Prompt)

                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    If IsNothing(Rezultat1) = False Then
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                            BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                            If Rezultat1.Value.Count = 1 Then
                                For i = 1 To Rezultat1.Value.Count
                                    Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                    Obj1 = Rezultat1.Value.Item(i - 1)
                                    Dim Ent1 As Entity
                                    Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                                    If TypeOf Ent1 Is DBText Then
                                        Dim Text1 As DBText = Ent1
                                        Dim String1 As String = Text1.TextString
                                        If IsNumeric(String1) = True Then
                                            TextBox_amount.Text = String1
                                        End If

                                    End If

                                    If TypeOf Ent1 Is MText Then
                                        Dim mText1 As MText = Ent1
                                        Dim String1 As String = mText1.Text
                                        If IsNumeric(String1) = True Then
                                            TextBox_amount.Text = String1
                                        End If

                                    End If


                                    If TypeOf Ent1 Is MLeader Then
                                        Dim Mleader1 As MLeader = Ent1
                                        Dim mText1 As MText = Mleader1.MText
                                        Dim String1 As String = mText1.Text
                                        If IsNumeric(String1) = True Then
                                            TextBox_amount.Text = String1
                                        End If
                                    End If

                                    If TypeOf Ent1 Is BlockReference Then
                                        Dim Block1 As BlockReference = Ent1
                                        If Block1.AttributeCollection.Count > 0 Then
                                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                            For Each id In attColl
                                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)

                                                Dim String1 As String = attref.TextString
                                                If IsNumeric(String1) = True Then
                                                    TextBox_amount.Text = String1
                                                    Exit For
                                                End If

                                            Next
                                        End If
                                    End If
                                    If TypeOf Ent1 Is AttributeDefinition Then
                                        Dim attref As AttributeDefinition = Ent1


                                        Dim String1 As String = attref.Tag.ToString
                                        If IsNumeric(String1) = True Then
                                            TextBox_amount.Text = String1
                                            Exit For
                                        End If

                                    End If


                                Next
                            End If







                            Editor1.Regen()
                            Trans1.Commit()
                        End Using

                    End If
                End If





                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            End Using




        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub
End Class