Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class Area_Form
    Dim Freeze_operations As Boolean = False
    Private Sub Button_calculate_Click(sender As Object, e As EventArgs) Handles Button_calculate.Click
        If Freeze_operations = False Then

            Try

                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                Editor1 = ThisDrawing.Editor
                Dim Empty_array() As ObjectId
                Editor1.SetImpliedSelection(Empty_array)

                Dim Nodec As Integer = 5
                If IsNumeric(TextBox_decimals.Text = True) Then
                    Nodec = Abs(CInt(TextBox_decimals.Text))
                End If

                Using lock As DocumentLock = ThisDrawing.LockDocument
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt.MessageForAdding = vbLf & "Select polylines:"

                    Object_Prompt.SingleOnly = False
                    Rezultat1 = Editor1.GetSelection(Object_Prompt)

                    If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        If IsNothing(Rezultat1) = False Then
                            Freeze_operations = True
                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Creaza_layer("NO PLOT", 40, "NO PLOT", False)

                                Dim Poly3d As Polyline3d = Nothing
                                Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                                BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                                Dim Area1 As Double = 0
                                Dim Point3d_colection As New Point3dCollection

                                For i = 0 To Rezultat1.Value.Count - 1
                                    Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                    Obj1 = Rezultat1.Value.Item(i)
                                    Dim Ent1 As Entity
                                    Ent1 = Trans1.GetObject(Obj1.ObjectId, OpenMode.ForRead)
                                    If TypeOf Ent1 Is Polyline Then
                                        Dim Poly1 As Polyline = Ent1
                                        Area1 = Area1 + Poly1.Area
                                        Point3d_colection.Add(Poly1.StartPoint)
                                    End If
                                Next




                                Dim PromptPT As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify label position:")
                                PromptPT.AllowNone = True

                                Dim Point_result As Autodesk.AutoCAD.EditorInput.PromptPointResult
                                Point_result = Editor1.GetPoint(PromptPT)
                                Dim Insert_label As Boolean = True
                                If Not Point_result.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                    Insert_label = False
                                End If


                                If Point3d_colection.Count > 0 Then
                                    Dim SuffiX1 As String = " SQFT"
                                    If RadioButton_sqft_to_AC.Checked = True Then
                                        Area1 = Area1 * 0.0000229568418910972
                                        SuffiX1 = " Ac."
                                    End If

                                    If Insert_label = True Then
                                        Dim LabelMtext As New MText
                                        LabelMtext.Layer = "NO PLOT"

                                        Dim rez As String = Round(Area1, Nodec).ToString()

                                        LabelMtext.Contents = rez
                                        TextBox_result.Text = rez
                                        LabelMtext.TextHeight = 10
                                        LabelMtext.Attachment = AttachmentPoint.MiddleCenter
                                        LabelMtext.Location = Point_result.Value
                                        BTrecord.AppendEntity(LabelMtext)
                                        Trans1.AddNewlyCreatedDBObject(LabelMtext, True)

                                        Windows.Forms.Clipboard.SetText(rez)

                                        For i = 0 To Point3d_colection.Count - 1
                                            Dim Line1 As New Line(Point3d_colection(i), Point_result.Value)
                                            Line1.Layer = "NO PLOT"
                                            BTrecord.AppendEntity(Line1)
                                            Trans1.AddNewlyCreatedDBObject(Line1, True)
                                        Next

                                    Else

                                        ThisDrawing.Editor.WriteMessage(vbLf & Round(Area1, Nodec) & SuffiX1)

                                    End If
                                End If












                                Trans1.Commit()
                            End Using
                        End If
                    End If
                End Using




                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")


            Catch ex As Exception
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try

        End If
        Freeze_operations = False
    End Sub
End Class