Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class Rechainage_Form
    Dim Colectie1 As New Specialized.StringCollection

    Private Sub Button_pick_Click(sender As Object, e As EventArgs) Handles Button_pick.Click
        ascunde_butoanele_pentru_forms(Me, Colectie1)

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

                            If (CheckBox_END_STA.Checked = False And CheckBox_BEG_STA.Checked = False) Then

                                TextBox_END_chainage.Text = ""
                                TextBox_BEG_chainage.Text = ""
                                TextBox_diference.Text = ""

                                If Rezultat1.Value.Count = 1 Then
                                    For i = 0 To Rezultat1.Value.Count - 1
                                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                        Obj1 = Rezultat1.Value.Item(i)
                                        Dim Ent1 As Entity
                                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                                        If TypeOf Ent1 Is DBText Then
                                            Dim Text1 As DBText = Ent1
                                            If Text1.TextString.Contains("+") Then
                                                Dim Text_chainage As String = extrage_chainage_din_text_de_la_inceputul_textului(Text1.TextString)
                                                If TextBox_BEG_chainage.Text = "" Then
                                                    TextBox_BEG_chainage.Text = Text_chainage
                                                Else
                                                    If TextBox_END_chainage.Text = "" Then
                                                        TextBox_END_chainage.Text = Text_chainage
                                                    End If
                                                End If

                                            End If
                                        End If

                                        If TypeOf Ent1 Is MText Then
                                            Dim mText1 As MText = Ent1
                                            If mText1.Contents.Contains("+") Then
                                                Dim Text_chainage As String = extrage_chainage_din_text_de_la_inceputul_textului(mText1.Text)
                                                If TextBox_BEG_chainage.Text = "" Then
                                                    TextBox_BEG_chainage.Text = Text_chainage
                                                Else
                                                    If TextBox_END_chainage.Text = "" Then
                                                        TextBox_END_chainage.Text = Text_chainage
                                                    End If
                                                End If
                                            End If
                                        End If


                                        If TypeOf Ent1 Is MLeader Then
                                            Dim Mleader1 As MLeader = Ent1
                                            If Mleader1.MText.Contents.Contains("+") Then
                                                Dim Text_chainage As String = extrage_chainage_din_text_de_la_inceputul_textului(Mleader1.MText.Text)
                                                If TextBox_BEG_chainage.Text = "" Then
                                                    TextBox_BEG_chainage.Text = Text_chainage
                                                Else
                                                    If TextBox_END_chainage.Text = "" Then
                                                        TextBox_END_chainage.Text = Text_chainage
                                                    End If
                                                End If
                                            End If
                                        End If

                                        If TypeOf Ent1 Is BlockReference Then
                                            Dim Block1 As BlockReference = Ent1
                                            If Block1.AttributeCollection.Count > 0 Then
                                                Dim Chainage_start1, Chainage_end1, Chainage_sta1 As Double
                                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                                For Each id In attColl
                                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                                    If attref.Tag.ToUpper = "BEGINSTA" Then
                                                        Dim Continut As String = attref.TextString
                                                        If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                            Chainage_start1 = CDbl(Replace(Continut, "+", ""))
                                                        End If
                                                    End If

                                                    If attref.Tag.ToUpper = "ENDSTA" Then
                                                        Dim Continut As String = attref.TextString
                                                        If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                            Chainage_end1 = CDbl(Replace(Continut, "+", ""))
                                                        End If
                                                    End If
                                                    If attref.Tag.ToUpper = "STA" Then
                                                        Dim Continut As String = attref.TextString
                                                        If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                            Chainage_sta1 = CDbl(Replace(Continut, "+", ""))
                                                        End If
                                                    End If
                                                Next
                                                If Chainage_end1 = 0 And Not Chainage_sta1 = 0 Then
                                                    Chainage_end1 = Chainage_sta1
                                                End If
                                                If Chainage_start1 = 0 And Not Chainage_sta1 = 0 Then
                                                    Chainage_start1 = Chainage_sta1
                                                End If
                                                TextBox_diference.Text = Get_String_Rounded(Abs(Chainage_start1 - Chainage_end1), 1)
                                                TextBox_BEG_chainage.Text = Get_chainage_from_double(Chainage_start1, 1)
                                                TextBox_END_chainage.Text = Get_chainage_from_double(Chainage_end1, 1)

                                            End If
                                        End If
                                        If TypeOf Ent1 Is AttributeDefinition Then
                                            Dim attref As AttributeDefinition = Ent1
                                            Dim Continut As String = attref.Tag.ToString
                                            If Continut.Contains("+") Then
                                                Dim Text_chainage As String = extrage_chainage_din_text_de_la_inceputul_textului(Continut)
                                                If TextBox_BEG_chainage.Text = "" Then
                                                    TextBox_BEG_chainage.Text = Text_chainage
                                                Else
                                                    If TextBox_END_chainage.Text = "" Then
                                                        TextBox_END_chainage.Text = Text_chainage
                                                    End If
                                                End If
                                            End If
                                        End If


                                    Next

                                End If



                                If Rezultat1.Value.Count > 1 Then
                                    Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                    Obj1 = Rezultat1.Value.Item(0)
                                    Dim Ent1 As Entity
                                    Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                                    Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                    Obj2 = Rezultat1.Value.Item(1)
                                    Dim Ent2 As Entity
                                    Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)


                                    Dim Chainage1, Chainage2 As Double
                                    Dim Chainage_start1, Chainage_end1, Chainage_sta1 As Double
                                    Dim Chainage_start2, Chainage_end2, Chainage_sta2 As Double
                                    Dim Text1 As DBText
                                    Dim Text2 As DBText
                                    Dim mText1 As MText
                                    Dim mText2 As MText
                                    Dim Block1 As BlockReference
                                    Dim Block2 As BlockReference

                                    If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                                        Text1 = Ent1
                                        If IsNumeric(Replace(Text1.TextString, "+", "")) = True Then
                                            Chainage1 = CDbl(Replace(Text1.TextString, "+", ""))
                                        End If
                                    End If

                                    If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.MText And TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                                        mText1 = Ent1
                                        If IsNumeric(mText1.Text) = True Then
                                            If IsNumeric(Replace(mText1.Text, "+", "")) = True Then
                                                Chainage1 = CDbl(Replace(mText1.Text, "+", ""))
                                            End If
                                        End If
                                    End If

                                    If TypeOf Ent1 Is BlockReference Then
                                        Block1 = Ent1
                                        If Block1.AttributeCollection.Count > 0 Then
                                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                            For Each id In attColl
                                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)

                                                If attref.Tag.ToUpper = "BEGINSTA" Then
                                                    Dim Continut As String = attref.TextString
                                                    If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                        Chainage_start1 = Replace(Continut, "+", "")
                                                    End If
                                                End If

                                                If attref.Tag.ToUpper = "ENDSTA" Then
                                                    Dim Continut As String = attref.TextString
                                                    If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                        Chainage_end1 = Replace(Continut, "+", "")
                                                    End If
                                                End If
                                                If attref.Tag.ToUpper = "STA" Then
                                                    Dim Continut As String = attref.TextString
                                                    If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                        Chainage_sta1 = Replace(Continut, "+", "")
                                                    End If
                                                End If

                                            Next



                                        End If
                                    End If

                                    If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                                        Text2 = Ent2
                                        If IsNumeric(Replace(Text2.TextString, "+", "")) = True Then
                                            Chainage2 = CDbl(Replace(Text2.TextString, "+", ""))
                                        End If
                                    End If

                                    If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.MText And TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                                        mText2 = Ent2
                                        If IsNumeric(mText2.Text) = True Then
                                            If IsNumeric(Replace(mText2.Text, "+", "")) = True Then
                                                Chainage2 = CDbl(Replace(mText2.Text, "+", ""))
                                            End If
                                        End If
                                    End If

                                    If TypeOf Ent2 Is BlockReference Then
                                        Block2 = Ent2
                                        If Block2.AttributeCollection.Count > 0 Then
                                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block2.AttributeCollection
                                            For Each id In attColl
                                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)

                                                If attref.Tag.ToUpper = "BEGINSTA" Then
                                                    Dim Continut As String = attref.TextString
                                                    If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                        Chainage_start2 = Replace(Continut, "+", "")
                                                    End If
                                                End If

                                                If attref.Tag.ToUpper = "ENDSTA" Then
                                                    Dim Continut As String = attref.TextString
                                                    If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                        Chainage_end2 = Replace(Continut, "+", "")
                                                    End If
                                                End If
                                                If attref.Tag.ToUpper = "STA" Then
                                                    Dim Continut As String = attref.TextString
                                                    If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                        Chainage_sta2 = Replace(Continut, "+", "")
                                                    End If
                                                End If

                                            Next
                                        End If
                                    End If

                                    If Not Chainage1 = 0 And Not Chainage2 = 0 Then
                                        TextBox_diference.Text = Get_String_Rounded(Abs(Chainage1 - Chainage2), 1)
                                        TextBox_BEG_chainage.Text = Get_chainage_from_double(Chainage1, 1)
                                        TextBox_END_chainage.Text = Get_chainage_from_double(Chainage2, 1)
                                    End If

                                    If (Chainage1 = 0 And Not Chainage2 = 0) Or (Chainage2 = 0 And Not Chainage1 = 0) Then
                                        Dim PozitieX As Double
                                        If IsNothing(Text1) = False Then
                                            PozitieX = Text1.Position.X
                                        End If
                                        If IsNothing(Text2) = False Then
                                            PozitieX = Text2.Position.X
                                        End If
                                        If IsNothing(mText1) = False Then
                                            PozitieX = mText1.Location.X
                                        End If
                                        If IsNothing(mText2) = False Then
                                            PozitieX = mText2.Location.X
                                        End If
                                        Dim PozitieBlockX As Double

                                        If IsNothing(Block1) = False Then
                                            PozitieBlockX = Block1.Position.X
                                        End If
                                        If IsNothing(Block2) = False Then
                                            PozitieBlockX = Block2.Position.X
                                        End If

                                        If PozitieX <= PozitieBlockX Then
                                            If Chainage1 = 0 Then
                                                If Not Chainage_start1 = 0 Then
                                                    Chainage1 = Chainage_start1
                                                    GoTo 123
                                                Else
                                                    If Not Chainage_sta1 = 0 Then
                                                        Chainage1 = Chainage_sta1
                                                        GoTo 123
                                                    Else
                                                        If Not Chainage_end1 = 0 Then
                                                            Chainage1 = Chainage_end1
                                                            GoTo 123
                                                        Else
                                                            afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                                            Exit Sub
                                                        End If
                                                    End If
                                                End If
                                            End If
                                            If Chainage2 = 0 Then
                                                If Not Chainage_start2 = 0 Then
                                                    Chainage2 = Chainage_start2
                                                    GoTo 123
                                                Else
                                                    If Not Chainage_sta2 = 0 Then
                                                        Chainage2 = Chainage_sta2
                                                        GoTo 123
                                                    Else
                                                        If Not Chainage_end2 = 0 Then
                                                            Chainage2 = Chainage_end2
                                                            GoTo 123
                                                        Else
                                                            afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                                            Exit Sub
                                                        End If
                                                    End If
                                                End If
                                            End If

                                        Else
                                            If Chainage1 = 0 Then
                                                If Not Chainage_end1 = 0 Then
                                                    Chainage1 = Chainage_end1
                                                    GoTo 123
                                                Else
                                                    If Not Chainage_sta1 = 0 Then
                                                        Chainage1 = Chainage_sta1
                                                        GoTo 123
                                                    Else
                                                        If Not Chainage_start1 = 0 Then
                                                            Chainage1 = Chainage_start1
                                                            GoTo 123
                                                        Else
                                                            afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                                            Exit Sub
                                                        End If
                                                    End If
                                                End If
                                            End If
                                            If Chainage2 = 0 Then
                                                If Not Chainage_end2 = 0 Then
                                                    Chainage2 = Chainage_end2
                                                    GoTo 123
                                                Else
                                                    If Not Chainage_sta1 = 0 Then
                                                        Chainage2 = Chainage_sta2
                                                        GoTo 123
                                                    Else
                                                        If Not Chainage_start2 = 0 Then
                                                            Chainage2 = Chainage_start2
                                                            GoTo 123
                                                        Else
                                                            afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                                            Exit Sub
                                                        End If
                                                    End If
                                                End If
                                            End If

                                        End If

123:

                                        Dim Semn As Double = 1

                                        If RadioButton_minus.Checked = True Then
                                            Semn = -1
                                        End If
                                        TextBox_diference.Text = Get_String_Rounded(Abs(Chainage1 + Semn * Chainage2), 1)
                                        TextBox_BEG_chainage.Text = Get_chainage_from_double(Chainage1, 1)
                                        TextBox_END_chainage.Text = Get_chainage_from_double(Chainage2, 1)
                                    End If



                                    If Chainage1 = 0 And Chainage2 = 0 Then
                                        Dim PozitieX1 As Double
                                        Dim PozitieX2 As Double

                                        If IsNothing(Block1) = False Then
                                            PozitieX1 = Block1.Position.X
                                        End If
                                        If IsNothing(Block2) = False Then
                                            PozitieX2 = Block2.Position.X
                                        End If

                                        If PozitieX1 <= PozitieX2 Then
                                            If Not Chainage_end1 = 0 And Not Chainage_start2 = 0 Then
                                                Chainage1 = Chainage_end1
                                                Chainage2 = Chainage_start2
                                                GoTo 124
                                            Else
                                                If Not Chainage_sta1 = 0 Then
                                                    Chainage1 = Chainage_sta1
                                                    Chainage2 = Chainage_start2
                                                    GoTo 124
                                                Else
                                                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                                    Exit Sub
                                                End If
                                            End If
                                        Else
                                            If Not Chainage_end2 = 0 And Not Chainage_start1 = 0 Then
                                                Chainage1 = Chainage_end2
                                                Chainage2 = Chainage_start1
                                                GoTo 124
                                            Else
                                                If Not Chainage_sta2 = 0 Then
                                                    Chainage1 = Chainage_sta2
                                                    Chainage2 = Chainage_start1
                                                    GoTo 124
                                                Else
                                                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                                    Exit Sub
                                                End If
                                            End If

                                        End If

124:
                                        Dim Semn As Double = 1
                                        If RadioButton_minus.Checked = True Then
                                            Semn = -1
                                        End If
                                        TextBox_diference.Text = Get_String_Rounded(Abs(Chainage1 + Semn * Chainage2), 1)
                                        TextBox_BEG_chainage.Text = Get_chainage_from_double(Chainage1, 1)
                                        TextBox_END_chainage.Text = Get_chainage_from_double(Chainage2, 1)
                                    End If


                                End If


                            End If

                            If CheckBox_BEG_STA.Checked = True And CheckBox_END_STA.Checked = False Then
                                Dim Chainage_start1 As Double
                                Dim cH_sTA As Double

                                For i = 0 To Rezultat1.Value.Count - 1
                                    Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                    Obj1 = Rezultat1.Value.Item(i)
                                    Dim Ent1 As Entity
                                    Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                                    If TypeOf Ent1 Is DBText Then
                                        Dim Text1 As DBText = Ent1
                                        If Text1.TextString.Contains("+") Then
                                            Dim Text_chainage As String = extrage_chainage_din_text_de_la_inceputul_textului(Text1.TextString)
                                            If IsNumeric(Replace(Text_chainage, "+", "")) = True Then

                                                Chainage_start1 = CDbl(Replace(Text_chainage, "+", ""))



                                            End If

                                        End If
                                    End If

                            If TypeOf Ent1 Is MText Then
                                Dim mText1 As MText = Ent1
                                If mText1.Contents.Contains("+") Then
                                    Dim Text_chainage As String = extrage_chainage_din_text_de_la_inceputul_textului(mText1.Text)
                                    If IsNumeric(Replace(Text_chainage, "+", "")) = True Then

                                                Chainage_start1 = CDbl(Replace(Text_chainage, "+", ""))


                                            End If
                                End If
                            End If


                            If TypeOf Ent1 Is MLeader Then
                                Dim Mleader1 As MLeader = Ent1
                                If Mleader1.MText.Contents.Contains("+") Then
                                    Dim Text_chainage As String = extrage_chainage_din_text_de_la_inceputul_textului(Mleader1.MText.Text)
                                    If IsNumeric(Replace(Text_chainage, "+", "")) = True Then

                                                Chainage_start1 = CDbl(Replace(Text_chainage, "+", ""))



                                            End If
                                        End If
                                    End If


                                    If TypeOf Ent1 Is BlockReference Then
                                        Dim Block1 As BlockReference = Ent1
                                        If Block1.AttributeCollection.Count > 0 Then

                                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                            For Each id In attColl
                                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                                If attref.Tag.ToUpper = "BEGINSTA" Then
                                                    Dim Text_chainage As String = attref.TextString
                                                    If IsNumeric(Replace(Text_chainage, "+", "")) = True Then
                                                        Chainage_start1 = CDbl(Replace(Text_chainage, "+", ""))
                                                    End If
                                                End If


                                                If attref.Tag.ToUpper = "STA" Then
                                                    Dim Text_chainage As String = attref.TextString
                                                    If IsNumeric(Replace(Text_chainage, "+", "")) = True Then
                                                        If IsNumeric(Replace(Text_chainage, "+", "")) = True Then
                                                            cH_sTA = CDbl(Replace(Text_chainage, "+", ""))
                                                            
                                                        End If
                                                    End If
                                                End If
                                            Next
                                        End If
                                    End If
                                Next
                                CheckBox_BEG_STA.Checked = False
                                CheckBox_diference.Checked = False
                                CheckBox_END_STA.Checked = False

                                If Not Chainage_start1 = 0 Then
                                    TextBox_BEG_chainage.Text = Get_chainage_from_double(Chainage_start1, 1)
                                    If Not cH_sTA = 0 Then
                                        TextBox_END_chainage.Text = Get_chainage_from_double(cH_sTA, 1)
                                    End If
                                Else
                                    If Not cH_sTA = 0 Then
                                        TextBox_BEG_chainage.Text = Get_chainage_from_double(cH_sTA, 1)
                                    End If
                                End If




                            End If

                    If CheckBox_END_STA.Checked = True And CheckBox_BEG_STA.Checked = False Then
                        Dim Chainage_end1 As Double
                        Dim ch_sta As Double
                        For i = 0 To Rezultat1.Value.Count - 1
                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(i)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                            If TypeOf Ent1 Is DBText Then
                                Dim Text1 As DBText = Ent1
                                If Text1.TextString.Contains("+") Then
                                    Dim Text_chainage As String = extrage_chainage_din_text_de_la_inceputul_textului(Text1.TextString)
                                    If IsNumeric(Replace(Text_chainage, "+", "")) = True Then
                                        Chainage_end1 = CDbl(Replace(Text_chainage, "+", ""))
                                    End If
                                End If
                            End If

                            If TypeOf Ent1 Is MText Then
                                Dim mText1 As MText = Ent1
                                If mText1.Contents.Contains("+") Then
                                    Dim Text_chainage As String = extrage_chainage_din_text_de_la_inceputul_textului(mText1.Text)
                                    If IsNumeric(Replace(Text_chainage, "+", "")) = True Then

                                        Chainage_end1 = CDbl(Replace(Text_chainage, "+", ""))


                                    End If
                                End If
                            End If


                            If TypeOf Ent1 Is MLeader Then
                                Dim Mleader1 As MLeader = Ent1
                                If Mleader1.MText.Contents.Contains("+") Then
                                    Dim Text_chainage As String = extrage_chainage_din_text_de_la_inceputul_textului(Mleader1.MText.Text)
                                    If IsNumeric(Replace(Text_chainage, "+", "")) = True Then

                                        Chainage_end1 = CDbl(Replace(Text_chainage, "+", ""))



                                    End If
                                End If
                            End If

                            If TypeOf Ent1 Is BlockReference Then
                                Dim Block1 As BlockReference = Ent1
                                If Block1.AttributeCollection.Count > 0 Then

                                    Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                    For Each id In attColl
                                        Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)


                                        If attref.Tag.ToUpper = "ENDSTA" Then
                                            Dim Text_chainage As String = attref.TextString
                                            If IsNumeric(Replace(Text_chainage, "+", "")) = True Then
                                                Chainage_end1 = CDbl(Replace(Text_chainage, "+", ""))
                                                    End If
                                        End If

                                        If attref.Tag.ToUpper = "STA" Then
                                            Dim Text_chainage As String = attref.TextString
                                            If IsNumeric(Replace(Text_chainage, "+", "")) = True Then
                                                        ch_sta = CDbl(Replace(Text_chainage, "+", ""))
                                                    End If
                                        End If
                                    Next

                                End If
                            End If
                                Next

                                CheckBox_BEG_STA.Checked = False
                                CheckBox_diference.Checked = False
                                CheckBox_END_STA.Checked = False

                                If Not Chainage_end1 = 0 Then
                                    TextBox_END_chainage.Text = Get_chainage_from_double(Chainage_end1, 1)
                                    If Not ch_sta = 0 Then
                                        TextBox_BEG_chainage.Text = Get_chainage_from_double(ch_sta, 1)
                                    End If
                                Else
                                    If Not ch_sta = 0 Then
                                        TextBox_END_chainage.Text = Get_chainage_from_double(ch_sta, 1)
                                    End If
                                End If

                    End If




                    If CheckBox_diference.Checked = True Then
                        Dim Text_chainage As String = ""
                        For i = 0 To Rezultat1.Value.Count - 1
                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(i)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                            If TypeOf Ent1 Is BlockReference Then
                                Dim Block1 As BlockReference = Ent1
                                If Block1.AttributeCollection.Count > 0 Then
                                    Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection

                                    For Each id In attColl
                                        Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                        If attref.Tag.ToUpper = "LENGTH" Then
                                            Text_chainage = attref.TextString

                                        End If
                                    Next

                                End If
                            End If

                        Next
                        TextBox_diference.Text = Text_chainage
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

    Private Sub Button_clear_Click(sender As Object, e As EventArgs) Handles Button_clear.Click
        TextBox_END_chainage.Text = ""
        TextBox_diference.Text = ""
        TextBox_BEG_chainage.Text = ""
        TextBox_screw_anchors_number.Text = ""
        Panel_color.BackColor = Drawing.Color.Gainsboro
        Label_updated_chain1_for_screw_anchors.Text = "Chainage 1"
        Label_updated_chain2_for_screw_anchors.Text = "Chainage 2"
    End Sub
    Private Sub Button_switch_Click(sender As Object, e As EventArgs) Handles Button_switch.Click
        Dim temp As String = TextBox_END_chainage.Text
        TextBox_END_chainage.Text = TextBox_BEG_chainage.Text
        TextBox_BEG_chainage.Text = temp
    End Sub
    Private Sub TextBox_calculate_Click(sender As Object, e As EventArgs) Handles TextBox_END_chainage.Click, TextBox_BEG_chainage.Click
        Dim Semn As Double = 1

        If RadioButton_minus.Checked = True Then
            Semn = -1
        End If

        If (CheckBox_BEG_STA.Checked = False And CheckBox_END_STA.Checked = False And CheckBox_diference.Checked = False) Then

            Dim Old_chainage As Double = 0
            If IsNumeric(Replace(TextBox_BEG_chainage.Text, "+", "")) = True Then
                Old_chainage = CDbl(Replace(TextBox_BEG_chainage.Text, "+", ""))
            End If

            Dim New_chainage As Double = 0
            If IsNumeric(Replace(TextBox_END_chainage.Text, "+", "")) = True Then
                New_chainage = CDbl(Replace(TextBox_END_chainage.Text, "+", ""))
            End If


            Dim Diferenta_chainage As Double = Abs(New_chainage + Semn * Old_chainage)
            TextBox_diference.Text = Get_String_Rounded(Diferenta_chainage, 1)
        End If

        If (CheckBox_BEG_STA.Checked = False And CheckBox_END_STA.Checked = True And CheckBox_diference.Checked = True) Then

            Dim Old_chainage As Double = 0
            Dim New_chainage As Double = 0
            If IsNumeric(Replace(TextBox_END_chainage.Text, "+", "")) = True Then
                New_chainage = CDbl(Replace(TextBox_END_chainage.Text, "+", ""))
            End If
            Dim Diferenta_chainage As Double = TextBox_diference.Text
            If IsNumeric(Diferenta_chainage) = True Then
                Old_chainage = New_chainage + Semn * Diferenta_chainage
                TextBox_BEG_chainage.Text = Get_chainage_from_double(Old_chainage, 1)
            End If
        End If
        If (CheckBox_BEG_STA.Checked = True And CheckBox_END_STA.Checked = False And CheckBox_diference.Checked = True) Then

            Dim Old_chainage As Double = 0
            If IsNumeric(Replace(TextBox_BEG_chainage.Text, "+", "")) = True Then
                Old_chainage = CDbl(Replace(TextBox_BEG_chainage.Text, "+", ""))
            End If
            Dim New_chainage As Double = 0

            Dim Diferenta_chainage As Double = TextBox_diference.Text
            If IsNumeric(Diferenta_chainage) = True Then
                New_chainage = Old_chainage + Semn * Diferenta_chainage
                TextBox_END_chainage.Text = Get_chainage_from_double(New_chainage, 1)
            End If
        End If


    End Sub

    Private Sub Button_calculate_screw_anchors_Click(sender As Object, e As EventArgs) Handles Button_calculate_screw_anchors.Click
        Dim Old_chainage As Double = 0


        If IsNumeric(Replace(TextBox_BEG_chainage.Text, "+", "")) = True Then
            Old_chainage = CDbl(Replace(TextBox_BEG_chainage.Text, "+", ""))
        End If


        Dim New_chainage As Double = 0
        If IsNumeric(Replace(TextBox_END_chainage.Text, "+", "")) = True Then
            New_chainage = CDbl(Replace(TextBox_END_chainage.Text, "+", ""))
        End If

        Dim Diferenta_chainage As Double = New_chainage - Old_chainage

        TextBox_diference.Text = Get_String_Rounded(Diferenta_chainage, 2)

        If IsNumeric(TextBox_screw_anchors_spacing.Text) = True Then
            Diferenta_chainage = Round(Abs(Diferenta_chainage), 2)
            Dim Screw_space As Double = CDbl(TextBox_screw_anchors_spacing.Text)
            Dim Multiplu As Double = Ceiling(Diferenta_chainage / Screw_space)


            Dim Number_screw As Integer = Floor(Multiplu) + 1

            TextBox_screw_anchors_number.Text = Number_screw

            If Old_chainage <= New_chainage Then
                Dim Chainage_middle = Old_chainage + Diferenta_chainage / 2
                Dim Chain1 As Double = Chainage_middle - (Multiplu * Screw_space) / 2
                Dim Chain2 As Double = Chainage_middle + (Multiplu * Screw_space) / 2
                Label_updated_chain1_for_screw_anchors.Text = Get_chainage_from_double(Chain1, 1)
                Label_updated_chain2_for_screw_anchors.Text = Get_chainage_from_double(Chain2, 1)
                If Not Round(Old_chainage, 2) = Round(Chain1, 2) Or Not Round(New_chainage, 2) = Round(Chain2, 2) Then
                    Panel_color.BackColor = Drawing.Color.Red
                Else
                    Panel_color.BackColor = Drawing.Color.Gainsboro
                End If
            Else
                Dim Chainage_middle = New_chainage + Diferenta_chainage / 2
                Dim Chain1 As Double = Chainage_middle - ((Number_screw - 1) * Screw_space) / 2
                Dim Chain2 As Double = Chainage_middle + ((Number_screw - 1) * Screw_space) / 2
                Label_updated_chain1_for_screw_anchors.Text = Get_chainage_from_double(Chain1, 1)
                Label_updated_chain2_for_screw_anchors.Text = Get_chainage_from_double(Chain2, 1)
                If Not Round(Old_chainage, 2) = Round(Chain2, 2) Or Not Round(New_chainage, 2) = Round(Chain1, 2) Then
                    Panel_color.BackColor = Drawing.Color.Red
                Else
                    Panel_color.BackColor = Drawing.Color.Gainsboro
                End If
            End If



        End If
    End Sub

    Private Sub TextBox_old_chainage_KeyDown(sender As Object, e As Windows.Forms.KeyEventArgs) Handles TextBox_BEG_chainage.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            TextBox_END_chainage.SelectAll()
            TextBox_END_chainage.Focus()
        End If
    End Sub
    Private Sub TextBox_new_chainage_KeyDown(sender As Object, e As Windows.Forms.KeyEventArgs) Handles TextBox_END_chainage.KeyDown, TextBox_AMOUNT_FOR_RECHAIN.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            TextBox_BEG_chainage.SelectAll()
            TextBox_BEG_chainage.Focus()
        End If
    End Sub


    Private Sub Button_UPDATE_LENGTH_Click(sender As Object, e As EventArgs) Handles Button_UPDATE_LENGTH.Click
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult


            Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt2.MessageForAdding = vbLf & "Select the block:"

            Object_Prompt2.SingleOnly = True

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
                                                If attref.Tag = "LENGTH" And CheckBox_diference.Checked = True Then
                                                    If Not TextBox_diference.Text = "" Then
                                                        If IsNumeric(TextBox_diference.Text) = True Then
                                                            attref.TextString = Get_String_Rounded(CDbl(TextBox_diference.Text), 1)
                                                        End If
                                                    End If
                                                End If
                                                If attref.Tag = "BEGINSTA" And CheckBox_BEG_STA.Checked = True Then
                                                    If Not TextBox_BEG_chainage.Text = "" Then
                                                        attref.TextString = TextBox_BEG_chainage.Text
                                                    End If
                                                End If
                                                If attref.Tag = "ENDSTA" And CheckBox_END_STA.Checked = True Then
                                                    If Not TextBox_END_chainage.Text = "" Then
                                                        attref.TextString = TextBox_END_chainage.Text
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



    Private Sub Button_RECHAINAGE_Click(sender As Object, e As EventArgs) Handles Button_RECHAINAGE.Click
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult


            Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt2.MessageForAdding = vbLf & "Select the block:"

            Object_Prompt2.SingleOnly = True

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

                                For i = 0 To Rezultat2.Value.Count - 1
                                    Dim Block1 As BlockReference
                                    Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                    Obj1 = Rezultat2.Value.Item(i)
                                    Dim Ent1 As Entity
                                    Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)


                                    If TypeOf Ent1 Is BlockReference Then
                                        Block1 = Ent1

                                        Dim valoareDemodificat As Double = 0
                                        If IsNumeric(TextBox_AMOUNT_FOR_RECHAIN.Text) = True Then
                                            valoareDemodificat = CDbl(TextBox_AMOUNT_FOR_RECHAIN.Text)
                                        End If
                                        If Not valoareDemodificat = 0 Then
                                            If Block1.AttributeCollection.Count > 0 Then
                                                Block1.UpgradeOpen()
                                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection

                                                For Each id In attColl
                                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForWrite)
                                                    Dim Continut As String = attref.TextString
                                                    If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                        Dim valoare1 As Double = CDbl(Replace(Continut, "+", ""))

                                                        If attref.Tag = "STA" And CheckBox_diference.Checked = True Then
                                                            attref.TextString = Get_chainage_from_double(valoare1 + valoareDemodificat, 1)

                                                        End If
                                                        If attref.Tag = "BEGINSTA" And CheckBox_BEG_STA.Checked = True Then
                                                            attref.TextString = Get_chainage_from_double(valoare1 + valoareDemodificat, 1)
                                                        End If
                                                        If attref.Tag = "ENDSTA" And CheckBox_END_STA.Checked = True Then
                                                            attref.TextString = Get_chainage_from_double(valoare1 + valoareDemodificat, 1)
                                                        End If
                                                    End If
                                                Next
                                            End If
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

    Private Sub Button_push_chainage_Click(sender As Object, e As EventArgs) Handles Button_push_chainage.Click
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult


            Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt1.MessageForAdding = vbLf & "Select the source block:"

            Object_Prompt1.SingleOnly = True

            Rezultat1 = Editor1.GetSelection(Object_Prompt1)


            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                Exit Sub
            End If



            Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult


            Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt2.MessageForAdding = vbLf & "Select the blocks to be pushed:"

            Object_Prompt2.SingleOnly = False

            Rezultat2 = Editor1.GetSelection(Object_Prompt2)


            If Rezultat2.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                Exit Sub
            End If

            Dim Begin_sta As Double
            Dim End_sta As Double

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then
                    Using lock1 As DocumentLock = ThisDrawing.LockDocument
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)

                            If Rezultat1.Value.Count > 0 Then

                                For i = 0 To Rezultat1.Value.Count - 1
                                    Dim Block1 As BlockReference
                                    Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                    Obj1 = Rezultat1.Value.Item(i)
                                    Dim Ent1 As Entity
                                    Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)


                                    If TypeOf Ent1 Is BlockReference Then
                                        Block1 = Ent1

                                        If Block1.AttributeCollection.Count > 0 Then

                                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection

                                            For Each id In attColl
                                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                                Dim Continut As String = attref.TextString
                                                If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                    If attref.Tag = "BEGINSTA" Then
                                                        Begin_sta = CDbl(Replace(Continut, "+", ""))
                                                    End If
                                                    If attref.Tag = "ENDSTA" Then
                                                        End_sta = CDbl(Replace(Continut, "+", ""))
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

            Dim Data_table_blocks As New System.Data.DataTable
            Data_table_blocks.Columns.Add("OBJID", GetType(ObjectId))
            Data_table_blocks.Columns.Add("X", GetType(Double))
            Data_table_blocks.Columns.Add("BEGINSTA", GetType(Double))
            Data_table_blocks.Columns.Add("ENDSTA", GetType(Double))
            Data_table_blocks.Columns.Add("STA", GetType(Double))
            Data_table_blocks.Columns.Add("LENGTH", GetType(Double))
            Dim Index1 As Integer = 0

            If Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat2) = False Then
                    Using lock1 As DocumentLock = ThisDrawing.LockDocument
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)

                            If Rezultat2.Value.Count > 0 Then

                                For i = 0 To Rezultat2.Value.Count - 1
                                    Dim Block1 As BlockReference
                                    Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                    Obj1 = Rezultat2.Value.Item(i)
                                    Dim Ent1 As Entity
                                    Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)


                                    If TypeOf Ent1 Is BlockReference Then
                                        Block1 = Ent1
                                        If Block1.AttributeCollection.Count > 0 Then
                                            Data_table_blocks.Rows.Add()
                                            Data_table_blocks.Rows(Index1).Item("OBJID") = Block1.ObjectId
                                            Data_table_blocks.Rows(Index1).Item("X") = Block1.Position.X
                                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                            For Each id In attColl
                                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                                Dim Continut As String = attref.TextString
                                                If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                    Dim valoare1 As Double = CDbl(Replace(Continut, "+", ""))

                                                    If attref.Tag = "STA" Then
                                                        Data_table_blocks.Rows(Index1).Item("STA") = valoare1

                                                    End If
                                                    If attref.Tag = "BEGINSTA" Then
                                                        Data_table_blocks.Rows(Index1).Item("BEGINSTA") = valoare1
                                                    End If
                                                    If attref.Tag = "ENDSTA" Then
                                                        Data_table_blocks.Rows(Index1).Item("ENDSTA") = valoare1
                                                    End If
                                                    If attref.Tag = "LENGTH" Then
                                                        Data_table_blocks.Rows(Index1).Item("LENGTH") = valoare1
                                                    End If
                                                End If
                                            Next
                                            Index1 = Index1 + 1
                                        End If
                                    End If

                                Next

                            End If


                            Trans1.Commit()

                        End Using
                    End Using
                End If
            End If

            If Data_table_blocks.Rows.Count > 0 Then

                Using lock1 As DocumentLock = ThisDrawing.LockDocument
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)

                        Dim valoareDemodificat As Double = 0



                        Dim DataView1 As New DataView(Data_table_blocks)
                        DataView1.Sort = "X"
                        If Not valoareDemodificat = 0 Then

                            For Each row1 In DataView1
                                Dim Block1 As BlockReference
                                Block1 = Trans1.GetObject(row1("OBJID"), OpenMode.ForRead)
                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                For Each id In attColl
                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForWrite)
                                    Dim Continut As String = attref.TextString
                                    If IsNumeric(Replace(Continut, "+", "")) = True Then
                                        Dim valoare1 As Double = CDbl(Replace(Continut, "+", ""))

                                        If attref.Tag = "STA" Then
                                            attref.TextString = Get_chainage_from_double(valoare1 + valoareDemodificat, 1)

                                        End If
                                        If attref.Tag = "BEGINSTA" Then
                                            attref.TextString = Get_chainage_from_double(valoare1 + valoareDemodificat, 1)
                                        End If
                                        If attref.Tag = "ENDSTA" Then
                                            attref.TextString = Get_chainage_from_double(valoare1 + valoareDemodificat, 1)
                                        End If
                                    End If
                                Next

                            Next
                        End If







                        Trans1.Commit()

                    End Using
                End Using

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

    Private Sub TextBox_BEG_chainage_TextChanged(sender As Object, e As EventArgs) Handles TextBox_BEG_chainage.TextChanged

    End Sub

    Private Sub TextBox_END_chainage_TextChanged(sender As Object, e As EventArgs) Handles TextBox_END_chainage.TextChanged

    End Sub

    Private Sub TextBox_diference_TextChanged(sender As Object, e As EventArgs) Handles TextBox_diference.TextChanged

    End Sub
End Class