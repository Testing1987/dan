Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class text2mtext_Form
    Dim Colectie1 As New Specialized.StringCollection
    Private Sub Button_change_Click(sender As Object, e As EventArgs) Handles Button_change.Click
        Try
            If IsNumeric(TextBox_text_height.Text) = False Then
                MsgBox("Please specify the text height!")
                Exit Sub
            End If




            ascunde_butoanele_pentru_forms(Me, Colectie1)
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
            Editor1 = ThisDrawing.Editor
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

            Using lock As DocumentLock = ThisDrawing.LockDocument
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
123:            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select the text objects:"
                Object_Prompt.SingleOnly = False
                Rezultat1 = Editor1.GetSelection(Object_Prompt)


                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    If IsNothing(Rezultat1) = False Then

                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim Data_table1 As New System.Data.DataTable
                            Data_table1.Columns.Add("X", GetType(Double))
                            Data_table1.Columns.Add("Y", GetType(Double))
                            Data_table1.Columns.Add("Z", GetType(Double))
                            Data_table1.Columns.Add("TEXT", GetType(String))
                            Dim Index1 As Integer = 0

                            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                            BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Dim Layer1 As String
                            Dim TextStyle1 As ObjectId
                            Dim Lineweight1 As LineWeight
                            Dim Colorindex1 As Integer
                            Dim Rotation1 As Double

                            For i = 0 To Rezultat1.Value.Count - 1
                                Obj1 = Rezultat1.Value.Item(i)
                                Dim Ent1 As Entity
                                Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                                If TypeOf Ent1 Is DBText Then
                                    Dim Text1 As DBText = Ent1
                                    Data_table1.Rows.Add()
                                    Data_table1.Rows(Index1).Item("X") = Text1.Position.X
                                    Data_table1.Rows(Index1).Item("Y") = Text1.Position.Y
                                    Data_table1.Rows(Index1).Item("Z") = Text1.Position.Z
                                    Data_table1.Rows(Index1).Item("TEXT") = Text1.TextString

                                    Index1 = Index1 + 1

                                    Layer1 = Text1.Layer
                                    textstyle1 = Text1.TextStyleId
                                    Lineweight1 = Text1.LineWeight
                                    Colorindex1 = Text1.ColorIndex
                                    Rotation1 = Text1.Rotation
                                    Text1.UpgradeOpen()
                                    Text1.Erase()
                                End If

                            Next


                            Data_table1 = Sort_data_table(Data_table1, "Y")
                            If Data_table1.Rows.Count > 0 Then
                                Dim mtext1 As New MText
                                mtext1.Layer = Layer1
                                mtext1.TextStyleId = TextStyle1
                                mtext1.LineWeight = Lineweight1
                                mtext1.ColorIndex = Colorindex1
                                mtext1.TextHeight = CDbl(TextBox_text_height.Text)
                                mtext1.Rotation = Rotation1
                                Dim continut As String



                                For i = Data_table1.Rows.Count - 1 To 0 Step -1
                                    If i = Data_table1.Rows.Count - 1 Then
                                        continut = Data_table1.Rows(i).Item("TEXT")
                                    Else
                                        continut = continut & vbCrLf & Data_table1.Rows(i).Item("TEXT")
                                    End If

                                Next

                                mtext1.Contents = continut
                                mtext1.Location = New Point3d(Data_table1.Rows(0).Item("X"), Data_table1.Rows(0).Item("Y"), Data_table1.Rows(0).Item("Z"))

                                BTrecord.AppendEntity(mtext1)
                                Trans1.AddNewlyCreatedDBObject(mtext1, True)
                            End If

                            Trans1.Commit()
                        End Using
                        GoTo 123
                    End If
                End If

                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
    End Sub
End Class