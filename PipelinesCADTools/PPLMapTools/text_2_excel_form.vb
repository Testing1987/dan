Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Public Class text_2_excel_form
    Dim Freeze_operations As Boolean = False
    Private Sub Button_table_to_excel_Click(sender As Object, e As EventArgs) Handles Button_table_to_excel.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                Try

                    Dim RowHeight As Double = CDbl(TextBox_row_height.Text)
                    Dim ColumnWidth As Double = CDbl(TextBox_column_width.Text)


                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                    Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = get_new_worksheet_from_Excel()


                    Using lock1 As DocumentLock = ThisDrawing.LockDocument
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)

                            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()


                            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                            Object_Prompt.MessageForAdding = vbLf & "Select text:"

                            Object_Prompt.SingleOnly = False
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Object_Prompt)




                            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                Freeze_operations = False
                                Exit Sub
                            End If


                            Dim DataTable1 As New System.Data.DataTable


                            DataTable1.Columns.Add("NUME", GetType(String))
                            DataTable1.Columns.Add("X", GetType(Double))
                            DataTable1.Columns.Add("Y", GetType(Double))
                            DataTable1.Columns.Add("ROW1", GetType(Integer))
                            DataTable1.Columns.Add("COL1", GetType(Integer))

                            Dim Idx As Integer = 0
                            For i = 0 To Rezultat1.Value.Count - 1
                                Dim Ent1 As Entity = Trans1.GetObject(Rezultat1.Value(i).ObjectId, OpenMode.ForRead)
                                If TypeOf (Ent1) Is DBText Then
                                    Dim TEXT1 As DBText = Ent1
                                    DataTable1.Rows.Add()

                                    DataTable1.Rows(Idx).Item("NUME") = TEXT1.TextString
                                    DataTable1.Rows(Idx).Item("X") = Round(TEXT1.Position.X, 0)
                                    DataTable1.Rows(Idx).Item("Y") = Round(TEXT1.Position.Y, 0)
                                    Idx = Idx + 1
                                End If
                                If TypeOf (Ent1) Is MText Then
                                    Dim MTEXT1 As MText = Ent1
                                    DataTable1.Rows.Add()

                                    DataTable1.Rows(Idx).Item("NUME") = MTEXT1.Text
                                    DataTable1.Rows(Idx).Item("X") = Round(MTEXT1.Location.X, 0)
                                    DataTable1.Rows(Idx).Item("Y") = Round(MTEXT1.Location.Y, 0)
                                    Idx = Idx + 1
                                End If

                            Next

                            If DataTable1.Rows.Count > 0 Then
                                DataTable1 = Sort_data_table_2_columns(DataTable1, "Y", " DESC,", "X", " ASC")
                                Dim x0 As Double
                                Dim y0 As Double

                                For i = 0 To DataTable1.Rows.Count - 1
                                    If i = 0 Then
                                        DataTable1.Rows(i).Item("ROW1") = 1
                                        DataTable1.Rows(i).Item("COL1") = 1
                                        x0 = DataTable1.Rows(i).Item("X")
                                        y0 = DataTable1.Rows(i).Item("Y")
                                    Else
                                        Dim X As Double
                                        Dim y As Double
                                        X = DataTable1.Rows(i).Item("X")
                                        y = DataTable1.Rows(i).Item("Y")

                                        Dim DistY As Double = y0 - y
                                        Dim Row1 As Double = Abs(DistY / RowHeight)

                                        Dim Distx As Double = X - x0
                                        Dim Col1 As Double = Abs(Distx / ColumnWidth)

                                        DataTable1.Rows(i).Item("ROW1") = Round(Row1, 0) + 1
                                        DataTable1.Rows(i).Item("COL1") = Round(Col1, 0) + 1

                                    End If


                                Next


                                For i = 0 To DataTable1.Rows.Count - 1

                                    Dim Row1 As Integer = DataTable1.Rows(i).Item("ROW1")
                                    Dim Col1 As Integer = DataTable1.Rows(i).Item("COL1")

                                    W1.Cells(Row1, Col1).value = DataTable1.Rows(i).Item("NUME")
                                Next


                                MsgBox("DONE")
                            End If



                        End Using
                    End Using




                Catch ex As System.SystemException
                    MsgBox(ex.Message & vbCrLf & ex.GetType.ToString)
                End Try
            Catch ex As Exception
                MsgBox(ex.Message & vbCrLf & ex.GetType.ToString)
            End Try

            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_Mtext_to_excel_Click(sender As Object, e As EventArgs) Handles Button_Mtext_to_excel.Click

        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                Try

                    Dim RowHeight As Double = CDbl(TextBox_row_height.Text)
                    Dim ColumnWidth As Double = CDbl(TextBox_column_width.Text)


                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument




                    Using lock1 As DocumentLock = ThisDrawing.LockDocument
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)

                            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()


                            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                            Object_Prompt.MessageForAdding = vbLf & "Select text:"

                            Object_Prompt.SingleOnly = False
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Object_Prompt)




                            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                Freeze_operations = False
                                Exit Sub
                            End If


                            Dim DataTable1 As New System.Data.DataTable


                            DataTable1.Columns.Add("TEXT", GetType(String))
                            DataTable1.Columns.Add("X", GetType(Double))
                            DataTable1.Columns.Add("Y", GetType(Double))


                            Dim Idx As Integer = 0
                            For i = 0 To Rezultat1.Value.Count - 1
                                Dim Ent1 As Entity = Trans1.GetObject(Rezultat1.Value(i).ObjectId, OpenMode.ForRead)
                                If TypeOf (Ent1) Is DBText Then
                                    Dim TEXT1 As DBText = Ent1
                                    DataTable1.Rows.Add()

                                    DataTable1.Rows(Idx).Item("TEXT") = TEXT1.TextString
                                    DataTable1.Rows(Idx).Item("X") = Round(TEXT1.Position.X, 3)
                                    DataTable1.Rows(Idx).Item("Y") = Round(TEXT1.Position.Y, 3)
                                    Idx = Idx + 1
                                End If

                                If TypeOf (Ent1) Is MText Then
                                    Dim MTEXT1 As MText = Ent1
                                    DataTable1.Rows.Add()
                                    Dim Continut As String = MTEXT1.Contents


                                    DataTable1.Rows(Idx).Item("TEXT") = Continut
                                    DataTable1.Rows(Idx).Item("X") = Round(MTEXT1.Location.X, 3)
                                    DataTable1.Rows(Idx).Item("Y") = Round(MTEXT1.Location.Y, 3)
                                    Continut = Continut.Replace("\P", "|")

                                    Dim Liniute() As String = Continut.Split("|")
                                    For j = 0 To Liniute.Length - 1
                                        If DataTable1.Columns.Contains("ROW" & (j + 1).ToString) = False Then
                                            DataTable1.Columns.Add("ROW" & (j + 1).ToString, GetType(String))
                                        End If
                                        DataTable1.Rows(Idx).Item("ROW" & (j + 1).ToString) = Liniute(j)
                                    Next


                                    Idx = Idx + 1
                                End If

                            Next
                            Add_to_clipboard_Data_table(DataTable1)



                            MsgBox("DONE")

                        End Using
                    End Using




                Catch ex As System.SystemException
                    MsgBox(ex.Message & vbCrLf & ex.GetType.ToString)
                End Try
            Catch ex As Exception
                MsgBox(ex.Message & vbCrLf & ex.GetType.ToString)
            End Try

            Freeze_operations = False
        End If
    End Sub
End Class