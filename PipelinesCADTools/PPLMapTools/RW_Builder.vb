
Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Public Class RW_Builder
    Dim Colectie1 As New Specialized.StringCollection
    Private Sub Button_draw_Click(sender As Object, e As EventArgs) Handles Button_draw.Click
        Try
            If TextBox_col_start.Text = "" Then
                MsgBox("Please specify the Start COLUMN!")
                Exit Sub
            End If

            If TextBox_col_end.Text = "" Then
                MsgBox("Please specify the End COLUMN!")
                Exit Sub
            End If

            If TextBox_col_type.Text = "" Then
                MsgBox("Please specify the Type COLUMN!")
                Exit Sub
            End If

            If TextBox_col_width.Text = "" Then
                MsgBox("Please specify the Width COLUMN!")
                Exit Sub
            End If

            If TextBox_col_offset.Text = "" Then
                MsgBox("Please specify the Centreline Offset COLUMN!")
                Exit Sub
            End If

            Dim Start1 As Integer = CInt(Val(TextBox_ROW_START.Text))
            Dim End1 As Integer = CInt(Val(TextBox_ROW_END.Text))
            If Start1 <= 0 Or End1 <= 0 Or Start1 > End1 Then
                MsgBox("Please specify the Excel Start - End Row properly!")
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
                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select centerline:"
                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
                Dim Col_Start As String = TextBox_col_start.Text.ToUpper
                Dim Col_End As String = TextBox_col_end.Text.ToUpper
                Dim Col_Type As String = TextBox_col_type.Text.ToUpper
                Dim Col_Width As String = TextBox_col_width.Text.ToUpper
                Dim Col_Offset As String = TextBox_col_offset.Text.ToUpper
                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    If IsNothing(Rezultat1) = False Then
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                            BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(0)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                            If TypeOf Ent1 Is Polyline Then

                                Dim Colectie_DBobjects1 As New DBObjectCollection
                                Dim Colectie_DBobjects2 As New DBObjectCollection

                                Dim Colectie_pentru_int1_start As New DoubleCollection
                                Dim Colectie_pentru_int1_end As New DoubleCollection

                                Dim Colectie_pentru_int2_start As New DoubleCollection
                                Dim Colectie_pentru_int2_end As New DoubleCollection

                                Dim Poly1 As Polyline = Ent1

                                For i = Start1 To End1
                                    Dim Excel_start_string As String
                                    Dim Excel_end_string As String

                                    Dim Excel_start As Double
                                    Dim Excel_end As Double

                                    Excel_start_string = W1.Range(TextBox_col_start.Text & i).Value
                                    Excel_end_string = W1.Range(TextBox_col_end.Text & i).Value
                                    Dim Xing_type As String = W1.Range(TextBox_col_type.Text & i).Value.ToString.ToUpper
                                    Dim Width_string As String = W1.Range(TextBox_col_width.Text & i).Value
                                    Dim Offset_fromCL_string As String = W1.Range(TextBox_col_offset.Text & i).Value
                                    If IsNumeric(Width_string) = True And IsNumeric(Offset_fromCL_string) Then
                                        If IsNumeric(Replace(Excel_start_string, "+", "")) = True And IsNumeric(Replace(Excel_end_string, "+", "")) Then
                                            Excel_start = CDbl(Replace(Excel_start_string, "+", ""))
                                            Excel_end = CDbl(Replace(Excel_end_string, "+", ""))

                                            Dim LatimeRW As Double = CDbl(Width_string)
                                            Dim Offset_fromCL As Double = CDbl(Offset_fromCL_string)

                                            If Excel_start >= 0 And Excel_end >= 0 Then
                                                If Excel_end < Excel_start Then
                                                    Dim temp1 As Double = Excel_start
                                                    Excel_start = Excel_end
                                                    Excel_end = temp1
                                                End If

                                                Dim Point1_start As Point3d = Poly1.GetPointAtDist(Excel_start)
                                                Dim Point1_end As Point3d = Poly1.GetPointAtDist(Excel_end)
                                                Dim Param1_start As Double = Poly1.GetParameterAtDistance(Excel_start)
                                                Dim Param1_end As Double = Poly1.GetParameterAtDistance(Excel_end)

                                                Dim Poly_pt_offset As New Polyline
                                                Poly_pt_offset.AddVertexAt(0, New Point2d(Point1_start.X, Point1_start.Y), 0, 0, 0)
                                                Dim Index1 As Integer = 1
                                                If Ceiling(Param1_start) <= Floor(Param1_end) Then
                                                    For j = Ceiling(Param1_start) To Floor(Param1_end)
                                                        Poly_pt_offset.AddVertexAt(Index1, Poly1.GetPoint2dAt(j), 0, 0, 0)
                                                        Index1 = Index1 + 1
                                                    Next
                                                End If

                                                Poly_pt_offset.AddVertexAt(Index1, New Point2d(Point1_end.X, Point1_end.Y), 0, 0, 0)

                                                If Abs(Param1_start - Poly1.GetParameterAtDistance(Round(Excel_start, 0))) <= 0.01 And Param1_start >= 1 Then
                                                    Colectie_pentru_int1_start.Add(1)
                                                Else
                                                    Colectie_pentru_int1_start.Add(0)
                                                End If

                                                If Abs(Param1_end - Poly1.GetParameterAtDistance(Round(Excel_end, 0))) <= 0.01 And Param1_end <= Poly1.NumberOfVertices - 2 Then
                                                    Colectie_pentru_int1_end.Add(1)
                                                Else
                                                    Colectie_pentru_int1_end.Add(0)
                                                End If
                                                
                                                Dim Object_colection1 As Autodesk.AutoCAD.DatabaseServices.DBObjectCollection = Poly_pt_offset.GetOffsetCurves((LatimeRW / 2) + Offset_fromCL)
                                                Dim Poly_offset1 As New Polyline
                                                Poly_offset1 = Object_colection1(0)
                                                Colectie_DBobjects1.Add(Poly_offset1)

                                                Dim Object_colection2 As Autodesk.AutoCAD.DatabaseServices.DBObjectCollection = Poly_pt_offset.GetOffsetCurves(-((LatimeRW / 2) - Offset_fromCL))
                                                Dim Poly_offset2 As New Polyline
                                                Poly_offset2 = Object_colection2(0)
                                                Colectie_DBobjects2.Add(Poly_offset2)
                                            End If
                                        End If
                                    End If
                                Next

                                If Colectie_DBobjects1.Count > 0 Then
                                    Dim Poly_offset_finala1 As New Polyline
                                    Dim Index1 As Integer = 0
                                    For i = 0 To Colectie_DBobjects1.Count - 1
                                        Dim poly_col1 As Polyline
                                        poly_col1 = Colectie_DBobjects1(i)
                                        For j = 0 To poly_col1.NumberOfVertices - 1
                                            Poly_offset_finala1.AddVertexAt(Index1, poly_col1.GetPoint2dAt(j), poly_col1.GetBulgeAt(j), 0, 0)
                                            Index1 = Index1 + 1
                                        Next
                                    Next
                                    Poly_offset_finala1.ColorIndex = 1
                                    BTrecord.AppendEntity(Poly_offset_finala1)
                                    Trans1.AddNewlyCreatedDBObject(Poly_offset_finala1, True)
                                End If

                                If Colectie_DBobjects2.Count > 0 Then
                                    Dim Poly_offset_finala2 As New Polyline
                                    Dim Index2 As Integer = 0
                                    For i = 0 To Colectie_DBobjects2.Count - 1
                                        Dim poly_col2 As Polyline
                                        poly_col2 = Colectie_DBobjects2(i)
                                        For j = 0 To poly_col2.NumberOfVertices - 1
                                            Poly_offset_finala2.AddVertexAt(Index2, poly_col2.GetPoint2dAt(j), poly_col2.GetBulgeAt(j), 0, 0)
                                            Index2 = Index2 + 1
                                        Next
                                    Next
                                    Poly_offset_finala2.ColorIndex = 2
                                    BTrecord.AppendEntity(Poly_offset_finala2)
                                    Trans1.AddNewlyCreatedDBObject(Poly_offset_finala2, True)
                                End If

                                Trans1.Commit()
                                Editor1.Regen()
                            End If
                        End Using
                    End If
                End If
            End Using

            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox("Done")
            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")


        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
    End Sub
End Class