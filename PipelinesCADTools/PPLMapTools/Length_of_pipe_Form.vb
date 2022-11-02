
Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Public Class Length_of_pipe_Form
    Dim Colectie1 As New Specialized.StringCollection
    Private Sub Button_output_Click(sender As Object, e As EventArgs) Handles Button_output.Click
        Try
            If TextBox_REF_chainage.Text = "" Then
                MsgBox("Please specify the Chainage!")
                Exit Sub
            End If

            If IsNumeric(Replace(TextBox_REF_chainage.Text, "+", "")) = False Then
                MsgBox("The format you specified the Chainage is not good!")
                Exit Sub
            End If


            If TextBox_dwg_no.Text = "" Then
                MsgBox("Please specify the Column for the drawing number!")
                Exit Sub
            End If

            If TextBox_len_of_pipe.Text = "" Then
                MsgBox("Please specify the Column for the length of pipe!")
                Exit Sub
            End If

            If TextBox_ref_ch_col.Text = "" Then
                MsgBox("Please specify the Column for the chainage reference!")
                Exit Sub
            End If

            If IsNumeric(TextBox_horizontal_Exxag.Text) = False Then
                MsgBox("Please specify the horizontal exaggeration!")
                Exit Sub
            End If
            If IsNumeric(TextBox_vertical_exag.Text) = False Then
                MsgBox("Please specify the vertical exaggeration!")
                Exit Sub
            End If

            If IsNumeric(TextBox_start_row.Text) = False Then
                MsgBox("Please specify the excel row!")
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
                Object_Prompt.MessageForAdding = vbLf & "Select the polyline representing the centerline of the pipe:"
                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)


                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    If IsNothing(Rezultat1) = False Then
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim Poly2d As Polyline
                            Dim Poly1_1 As New Polyline

                            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                            BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(0)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                            If TypeOf Ent1 Is Polyline Then
                                Poly2d = Ent1
                                Dim Horiz_exag As Double = CDbl(TextBox_horizontal_Exxag.Text)
                                Dim Vert_exag As Double = CDbl(TextBox_vertical_exag.Text)
                                Dim Point0 As New Point3d
                                Point0 = Poly2d.GetPointAtDist(0)

                                If Not Horiz_exag = 1 And Not Vert_exag = 1 Then
                                    For i = 0 To Poly2d.NumberOfVertices - 1
                                        If i = 0 Then
                                            Poly1_1.AddVertexAt(i, New Point2d(Point0.X, Point0.Y), 0, 0, 0)
                                        Else
                                            Dim Point2d_1_1 As New Point2d(Point0.X + (Poly2d.GetPoint2dAt(i).X - Point0.X) / Horiz_exag, (Point0.Y + (Poly2d.GetPoint2dAt(i).Y - Point0.Y) / Vert_exag))
                                            Poly1_1.AddVertexAt(i, Point2d_1_1, 0, 0, 0)
                                        End If
                                    Next
                                Else
                                    For i = 0 To Poly2d.NumberOfVertices - 1
                                        Poly1_1.AddVertexAt(i, Poly2d.GetPoint2dAt(i), Poly2d.GetBulgeAt(i), 0, 0)
                                    Next
                                End If



                                Poly1_1.Elevation = 0
                                ' BTrecord.AppendEntity(Poly1_1)
                                'Trans1.AddNewlyCreatedDBObject(Poly1_1, True)


                                Dim PP_ref As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select The reference point for the chainage:")
                                PP_ref.AllowNone = True

                                Dim Point_ref_result As Autodesk.AutoCAD.EditorInput.PromptPointResult
                                Point_ref_result = Editor1.GetPoint(PP_ref)
                                If Not Point_ref_result.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                    ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                    Exit Sub
                                End If

                                Dim Chainage_ref As Double = CDbl(Replace(TextBox_REF_chainage.Text, "+", ""))

                                Dim Param_ref As Double = Poly2d.GetParameterAtPoint(Poly2d.GetClosestPointTo(Point_ref_result.Value, Vector3d.ZAxis, False))
                                Dim Dist_de_la_zero As Double = Poly1_1.GetDistanceAtParameter(Param_ref)


                                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()

                                Dim Start1 As Integer
                                Start1 = CInt(TextBox_start_row.Text)
                                Dim doIt As Boolean = True

                                Creaza_layer("NO PLOT", 40, "NO PLOT", False)

                                Do Until doIt = False
                                    Dim PP_1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select a point along centerline of pipe:")
                                    PP_1.AllowNone = True

                                    Dim Point_result1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                                    Point_result1 = Editor1.GetPoint(PP_1)
                                    If Not Point_result1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                        doIt = False
                                    End If
                                    If doIt = True Then
                                        Dim Point_on_2d As New Point3d
                                        Point_on_2d = Poly2d.GetClosestPointTo(Point_result1.Value, Vector3d.ZAxis, False)
                                        Dim param1 As Double = Poly2d.GetParameterAtPoint(Point_on_2d)
                                        Dim Distanta1 As Double = Poly1_1.GetDistanceAtParameter(param1)
                                        Dim Len1 As Double = Distanta1 - Dist_de_la_zero
                                        W1.Range(TextBox_len_of_pipe.Text & Start1).Value = Round(Len1, 2)
                                        If Microsoft.VisualBasic.FileIO.FileSystem.FileExists(ThisDrawing.Name) = True Then
                                            W1.Range(TextBox_dwg_no.Text & Start1).Value = IO.Path.GetFileName(ThisDrawing.Name)
                                        End If
                                        W1.Range(TextBox_ref_ch_col.Text & Start1).Value = Get_chainage_from_double(Chainage_ref, 1)


                                        Start1 = Start1 + 1
                                        TextBox_start_row.Text = Start1.ToString
                                        Dim Mleader1 As New MLeader
                                        Mleader1 = Creaza_Mleader_nou_fara_UCS_transform(Point_on_2d, "Len = " & Round(Len1, 2), 1, 1, 1, 2, 5)
                                        Mleader1.Layer = "NO PLOT"
                                    End If

                                Loop

                            End If

                            Editor1.Regen()
                            Trans1.Commit()
                        End Using

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