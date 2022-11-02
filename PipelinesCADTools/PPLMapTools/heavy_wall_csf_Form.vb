Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Public Class heavy_wall_csf_Form
    Dim Colectie1 As New Specialized.StringCollection
    Private Sub Button_chainage_at_Start_end_on_top_Click(sender As Object, e As EventArgs) Handles Button_output_to_excel_top.Click
        Try
            If TextBox_east_intersection.Text = "" Then
                MsgBox("Please specify the East COLUMN!")
                Exit Sub
            End If
            If TextBox_north_intersection.Text = "" Then
                MsgBox("Please specify the North COLUMN!")
                Exit Sub
            End If
            If TextBox_elevation.Text = "" Then
                MsgBox("Please specify the Elevation COLUMN!")
                Exit Sub
            End If
            If TextBox_chainage.Text = "" Then
                MsgBox("Please specify the Chainage COLUMN!")
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
                Object_Prompt.MessageForAdding = vbLf & "Select 3D centerline:"
                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)


                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    If IsNothing(Rezultat1) = False Then
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim Data_table_for_singles As New System.Data.DataTable
                            Data_table_for_singles.Columns.Add("X", GetType(Double))
                            Data_table_for_singles.Columns.Add("Y", GetType(Double))
                            Data_table_for_singles.Columns.Add("Z", GetType(Double))
                            Data_table_for_singles.Columns.Add("CHAINAGE", GetType(Double))

                            Dim indexdt As Double = 0

                            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = get_new_worksheet_from_Excel()
                            Dim Col_east As String = TextBox_east_intersection.Text.ToUpper
                            Dim Col_north As String = TextBox_north_intersection.Text.ToUpper
                            Dim Col_Chain As String = TextBox_chainage.Text.ToUpper
                            Dim Col_Elevation As String = TextBox_elevation.Text.ToUpper





                            Dim Poly3d As Polyline3d

                            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                            BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(0)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                            If TypeOf Ent1 Is Polyline3d Then
                                Poly3d = Ent1
                                Dim Poly2D As New Polyline

                                Dim Index2d As Double = 0
                                For Each ObjId As ObjectId In Poly3d
                                    Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, OpenMode.ForRead)
                                    Poly2D.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)

                                    Index2d = Index2d + 1
                                Next
                                Poly2D.Elevation = 0





                                Dim Data_table_Poly As New System.Data.DataTable
                                Data_table_Poly.Columns.Add("2DPOLY", GetType(Curve))


                                Dim IndexXX As Double = 0





                                For Each ObjID In BTrecord
                                    Dim DBobject As DBObject = Trans1.GetObject(ObjID, OpenMode.ForRead)

                                    If TypeOf DBobject Is Curve Then
                                        Dim Poly1 As Curve = DBobject
                                        Dim Layer1 As LayerTableRecord = Trans1.GetObject(Poly1.LayerId, OpenMode.ForRead)
                                        If Layer1.IsFrozen = False And Layer1.IsOff = False Then
                                            Data_table_Poly.Rows.Add()
                                            Data_table_Poly.Rows(IndexXX).Item("2DPOLY") = Poly1
                                            IndexXX = IndexXX + 1
                                        End If


                                    End If

                                Next




                                If Data_table_Poly.Rows.Count > 0 Then
                                    For i = 0 To Data_table_Poly.Rows.Count - 1

                                        Dim Poly1 As Curve = Data_table_Poly.Rows(i).Item("2DPOLY")

                                        Dim Start_p1 As New Point3d(Poly1.StartPoint.X, Poly1.StartPoint.Y, 0)
                                        Dim End_p1 As New Point3d(Poly1.EndPoint.X, Poly1.EndPoint.Y, 0)
                                        Dim Start_p1P As New Point3d
                                        Start_p1P = Poly2D.GetClosestPointTo(Start_p1, Vector3d.ZAxis, False)
                                        Dim End_p1P As New Point3d
                                        End_p1P = Poly2D.GetClosestPointTo(End_p1, Vector3d.ZAxis, False)
                                        Dim Col_int As New Point3dCollection
                                        If Start_p1.GetVectorTo(Start_p1P).Length < 0.1 And End_p1.GetVectorTo(End_p1P).Length < 0.1 Then
                                            Col_int.Add(Start_p1P)
                                            Col_int.Add(End_p1P)
                                        End If



                                        If Col_int.Count > 0 Then
                                            For j = 0 To Col_int.Count - 1
                                                Dim point_on_2d As New Point3d
                                                point_on_2d = Col_int(j)
                                                Dim Point_on_3d As New Point3d
                                                Dim Param1 As Double = Poly2D.GetParameterAtPoint(point_on_2d)
                                                Point_on_3d = Poly3d.GetPointAtParameter(Param1)


                                                Data_table_for_singles.Rows.Add()
                                                Data_table_for_singles.Rows(indexdt).Item("X") = Round(Point_on_3d.X, 3)
                                                Data_table_for_singles.Rows(indexdt).Item("Y") = Round(Point_on_3d.Y, 3)
                                                Data_table_for_singles.Rows(indexdt).Item("Z") = Round(Point_on_3d.Z, 3)
                                                Data_table_for_singles.Rows(indexdt).Item("CHAINAGE") = Round(Poly3d.GetDistAtPoint(Point_on_3d), 3)
                                                indexdt = indexdt + 1
                                            Next

                                        End If
                                    Next
                                End If



                            End If


                            If Data_table_for_singles.Rows.Count > 0 Then

                                Dim start1 As Integer = 2
                                W1.Range(Col_east & "1").Value = "East"
                                W1.Range(Col_north & "1").Value = "North"
                                W1.Range(Col_Elevation & "1").Value = "Elevation"
                                W1.Range(Col_Chain & "1").Value = "Chainage"


                                For k = 0 To Data_table_for_singles.Rows.Count - 1

                                    Dim new_leader As New MLeader
                                    new_leader = Creaza_Mleader_nou_fara_UCS_transform(New Point3d(Data_table_for_singles.Rows(k).Item("X"), Data_table_for_singles.Rows(k).Item("Y"), Data_table_for_singles.Rows(k).Item("Z")),
                                                                                       "East = " & Data_table_for_singles.Rows(k).Item("X") & vbCrLf &
                                                                                       "North = " & Data_table_for_singles.Rows(k).Item("Y") & vbCrLf &
                                                                                       "Elev = " & Data_table_for_singles.Rows(k).Item("Z") & vbCrLf &
                                                                                       "Chainage = " & Data_table_for_singles.Rows(k).Item("CHAINAGE") _
                                                                                       , 1, 0.2, 0.5, 2, 7)


                                    W1.Range(Col_east & start1).Value = Data_table_for_singles.Rows(k).Item("X")
                                    W1.Range(Col_north & start1).Value = Data_table_for_singles.Rows(k).Item("Y")
                                    W1.Range(Col_Elevation & start1).Value = Data_table_for_singles.Rows(k).Item("Z")
                                    W1.Range(Col_Chain & start1).Value = Data_table_for_singles.Rows(k).Item("CHAINAGE")


                                    start1 = start1 + 1

                                Next














                            End If

                            Editor1.Regen()
                            Trans1.Commit()
                        End Using

                    End If
                End If
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                MsgBox("Done")
                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_output_to_excel_parallel_Click(sender As Object, e As EventArgs) Handles Button_output_to_excel_parallel_middle.Click
        Try
            If TextBox_east_intersection.Text = "" Then
                MsgBox("Please specify the East COLUMN!")
                Exit Sub
            End If
            If TextBox_north_intersection.Text = "" Then
                MsgBox("Please specify the North COLUMN!")
                Exit Sub
            End If
            If TextBox_elevation.Text = "" Then
                MsgBox("Please specify the Elevation COLUMN!")
                Exit Sub
            End If
            If TextBox_chainage.Text = "" Then
                MsgBox("Please specify the Chainage COLUMN!")
                Exit Sub
            End If
            If TextBox_CSF_Chainage.Text = "" Then
                MsgBox("Please specify the CSF Chainage COLUMN!")
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
                Object_Prompt.MessageForAdding = vbLf & "Select 3D centerline:"
                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)


                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    If IsNothing(Rezultat1) = False Then
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction


                            Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                            Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                            Object_Prompt2.MessageForAdding = vbLf & "Select a line or polyline representing the heavy wall:"
                            Object_Prompt2.SingleOnly = True
                            Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                            If Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                If IsNothing(Rezultat2) = False Then

                                    Dim Data_table_for_singles As New System.Data.DataTable
                                    Data_table_for_singles.Columns.Add("X", GetType(Double))
                                    Data_table_for_singles.Columns.Add("Y", GetType(Double))
                                    Data_table_for_singles.Columns.Add("Z", GetType(Double))
                                    Data_table_for_singles.Columns.Add("CHAINAGE", GetType(Double))
                                    Data_table_for_singles.Columns.Add("CSFCHAINAGE", GetType(Double))


                                    Dim indexdt As Double = 0

                                    Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = get_new_worksheet_from_Excel()
                                    Dim Col_east As String = TextBox_east_intersection.Text.ToUpper
                                    Dim Col_north As String = TextBox_north_intersection.Text.ToUpper
                                    Dim Col_Chain As String = TextBox_chainage.Text.ToUpper
                                    Dim Col_Elevation As String = TextBox_elevation.Text.ToUpper
                                    Dim Col_CSF_CHAIN As String = TextBox_CSF_Chainage.Text.ToUpper




                                    Dim Poly3d As Polyline3d
                                    Dim Liniesample1 As Line
                                    Dim Polysample1 As Polyline

                                    Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                    BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                                    Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                    Obj1 = Rezultat1.Value.Item(0)
                                    Dim Ent1 As Entity
                                    Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                                    Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                    Obj2 = Rezultat2.Value.Item(0)
                                    Dim Ent2 As Entity
                                    Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)


                                    If TypeOf Ent1 Is Polyline3d And (TypeOf Ent2 Is Line Or TypeOf Ent2 Is Polyline) Then

                                        Poly3d = Ent1

                                        Dim Layersample As String
                                        If TypeOf Ent2 Is Line Then
                                            Liniesample1 = Ent2
                                            Layersample = Liniesample1.Layer
                                        Else
                                            Polysample1 = Ent2
                                            Layersample = Polysample1.Layer
                                        End If




                                        Dim Poly2D As New Polyline

                                        Dim Index2d As Double = 0
                                        For Each ObjId As ObjectId In Poly3d
                                            Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, OpenMode.ForRead)
                                            Poly2D.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)

                                            Index2d = Index2d + 1
                                        Next
                                        Poly2D.Elevation = 0


                                        Dim ChainageCSF_colection As New DBObjectCollection
                                        Dim CSF_colection As New DBObjectCollection


                                        Dim Data_table_Poly As New System.Data.DataTable
                                        Data_table_Poly.Columns.Add("2DPOLY", GetType(Curve))


                                        Dim IndexXX As Double = 0


                                        For Each ObjID In BTrecord
                                            Dim DBobject As DBObject = Trans1.GetObject(ObjID, OpenMode.ForRead)

                                            If TypeOf DBobject Is Curve Then
                                                Dim Poly1 As Curve = DBobject
                                                If Poly1.Layer = Layersample Then
                                                    Data_table_Poly.Rows.Add()
                                                    Data_table_Poly.Rows(IndexXX).Item("2DPOLY") = Poly1
                                                    IndexXX = IndexXX + 1
                                                End If

                                            End If
                                            If TypeOf DBobject Is DBText Then
                                                Dim CsfCh1 As DBText = DBobject
                                                If CsfCh1.TextString.Contains("+") = True Then
                                                    ChainageCSF_colection.Add(CsfCh1)
                                                End If
                                                If CsfCh1.TextString.Contains("CSF") = True Then
                                                    CSF_colection.Add(CsfCh1)
                                                End If
                                            End If

                                        Next




                                        If Data_table_Poly.Rows.Count > 0 Then
                                            For i = 0 To Data_table_Poly.Rows.Count - 1

                                                Dim Poly1 As Curve = Data_table_Poly.Rows(i).Item("2DPOLY")

                                                Dim Start_p1 As New Point3d(Poly1.StartPoint.X, Poly1.StartPoint.Y, 0)
                                                Dim End_p1 As New Point3d(Poly1.EndPoint.X, Poly1.EndPoint.Y, 0)
                                                Dim Middle_p1 As New Point3d((Poly1.StartPoint.X + Poly1.EndPoint.X) / 2, (Poly1.StartPoint.Y + Poly1.EndPoint.Y) / 2, 0)
                                                Dim Length1 As Double = Start_p1.GetVectorTo(End_p1).Length

                                                Dim Middle_p1P As New Point3d
                                                Middle_p1P = Poly2D.GetClosestPointTo(Middle_p1, Vector3d.ZAxis, False)
                                                Dim Param_middle As Double = Poly2D.GetParameterAtPoint(Middle_p1P)
                                                Dim Len_middle As Double = Poly3d.GetDistanceAtParameter(Param_middle)


                                                If Len_middle - Length1 / 2 >= 0 And Len_middle - Length1 / 2 <= Poly3d.Length Then
                                                    Dim Start_p1P As New Point3d
                                                    Start_p1P = Poly3d.GetPointAtDist(Len_middle - Length1 / 2)
                                                    Data_table_for_singles.Rows.Add()
                                                    Data_table_for_singles.Rows(indexdt).Item("X") = Round(Start_p1P.X, 3)
                                                    Data_table_for_singles.Rows(indexdt).Item("Y") = Round(Start_p1P.Y, 3)
                                                    Data_table_for_singles.Rows(indexdt).Item("Z") = Round(Start_p1P.Z, 3)
                                                    Data_table_for_singles.Rows(indexdt).Item("CHAINAGE") = Round(Poly3d.GetDistAtPoint(Start_p1P), 3)
                                                    indexdt = indexdt + 1

                                                    Dim End_p1P As New Point3d
                                                    End_p1P = Poly3d.GetPointAtDist(Len_middle + Length1 / 2)
                                                    Data_table_for_singles.Rows.Add()
                                                    Data_table_for_singles.Rows(indexdt).Item("X") = Round(End_p1P.X, 3)
                                                    Data_table_for_singles.Rows(indexdt).Item("Y") = Round(End_p1P.Y, 3)
                                                    Data_table_for_singles.Rows(indexdt).Item("Z") = Round(End_p1P.Z, 3)
                                                    Data_table_for_singles.Rows(indexdt).Item("CHAINAGE") = Round(Poly3d.GetDistAtPoint(End_p1P), 3)
                                                    indexdt = indexdt + 1


                                                    If CSF_colection.Count > 0 And ChainageCSF_colection.Count > 0 Then
                                                        Dim DistMin1 As Double = 50
                                                        Dim Chain_index1 As Double
                                                        Dim Pt_index1 As New Point3d

                                                        Dim DistMin2 As Double = 50
                                                        Dim Chain_index2 As Double
                                                        Dim Pt_index2 As New Point3d

                                                        For r = 0 To ChainageCSF_colection.Count - 1
                                                            Dim DbText1 As DBText = ChainageCSF_colection(r)

                                                            Dim Dist1 As Double
                                                            Dist1 = New Point3d(Start_p1P.X, Start_p1P.Y, 0).GetVectorTo(New Point3d(DbText1.Position.X, DbText1.Position.Y, 0)).Length
                                                            If Dist1 < DistMin1 And IsNumeric(Replace(DbText1.TextString, "+", "")) = True Then
                                                                DistMin1 = Dist1
                                                                Chain_index1 = CDbl(Replace(DbText1.TextString, "+", ""))
                                                                Pt_index1 = DbText1.Position
                                                            End If
                                                            Dim Dist2 As Double
                                                            Dist2 = New Point3d(End_p1P.X, End_p1P.Y, 0).GetVectorTo(New Point3d(DbText1.Position.X, DbText1.Position.Y, 0)).Length
                                                            If Dist2 < DistMin2 And IsNumeric(Replace(DbText1.TextString, "+", "")) = True Then
                                                                DistMin2 = Dist2
                                                                Chain_index2 = CDbl(Replace(DbText1.TextString, "+", ""))
                                                                Pt_index2 = DbText1.Position
                                                            End If

                                                        Next

                                                        Dim Csf1 As Double
                                                        Dim Csf2 As Double

                                                        For r = 0 To CSF_colection.Count - 1
                                                            Dim DbText1 As DBText = CSF_colection(r)
                                                            Dim String1 As String = DbText1.TextString
                                                            String1 = extrage_numar_din_text_de_la_sfarsitul_textului(String1)
                                                            If IsNumeric(String1) = True And Abs(Pt_index1.X - DbText1.Position.X) < 0.1 And Abs(Pt_index1.Y - DbText1.Position.Y) < 0.1 Then
                                                                Csf1 = CDbl(String1)
                                                            End If
                                                            If IsNumeric(String1) = True And Abs(Pt_index2.X - DbText1.Position.X) < 0.1 And Abs(Pt_index2.Y - DbText1.Position.Y) < 0.1 Then
                                                                Csf2 = CDbl(String1)
                                                            End If
                                                        Next

                                                        Dim pT_1 As New Point3d
                                                        pT_1 = Poly2D.GetClosestPointTo(Pt_index1, Vector3d.ZAxis, False)
                                                        Dim Param_1 As Double = Poly2D.GetParameterAtPoint(pT_1)
                                                        Dim Len_1 As Double = Poly3d.GetDistanceAtParameter(Param_1)


                                                        Data_table_for_singles.Rows(indexdt - 2).Item("CSFCHAINAGE") = Chain_index1 + (Poly3d.GetDistAtPoint(Start_p1P) - Len_1) / Csf1

                                                        Dim pT_2 As New Point3d
                                                        pT_2 = Poly2D.GetClosestPointTo(Pt_index2, Vector3d.ZAxis, False)
                                                        Dim Param_2 As Double = Poly2D.GetParameterAtPoint(pT_2)
                                                        Dim Len_2 As Double = Poly3d.GetDistanceAtParameter(Param_2)


                                                        Data_table_for_singles.Rows(indexdt - 1).Item("CSFCHAINAGE") = Chain_index2 + (Poly3d.GetDistAtPoint(End_p1P) - Len_2) / Csf2

                                                    End If
                                                End If





                                            Next
                                        End If



                                    End If


                                    If Data_table_for_singles.Rows.Count > 0 Then

                                        Dim start1 As Integer = 2
                                        W1.Range(Col_east & "1").Value = "East"
                                        W1.Range(Col_north & "1").Value = "North"
                                        W1.Range(Col_Elevation & "1").Value = "Elevation"
                                        W1.Range(Col_Chain & "1").Value = "Grid Chainage"
                                        W1.Range(Col_CSF_CHAIN & "1").Value = "CSF Chainage"

                                        For k = 0 To Data_table_for_singles.Rows.Count - 1
                                            Dim csf_CH As Double = 0
                                            If IsDBNull(Data_table_for_singles.Rows(k).Item("CSFCHAINAGE")) = False Then
                                                csf_CH = Round(Data_table_for_singles.Rows(k).Item("CSFCHAINAGE"), 3)
                                            End If

                                            Dim new_leader As New MLeader
                                            new_leader = Creaza_Mleader_nou_fara_UCS_transform(New Point3d(Data_table_for_singles.Rows(k).Item("X"), Data_table_for_singles.Rows(k).Item("Y"), Data_table_for_singles.Rows(k).Item("Z")),
                                                                                               "East = " & Data_table_for_singles.Rows(k).Item("X") & vbCrLf &
                                                                                               "North = " & Data_table_for_singles.Rows(k).Item("Y") & vbCrLf &
                                                                                               "Elev = " & Data_table_for_singles.Rows(k).Item("Z") & vbCrLf &
                                                                                               "Grid Chainage = " & Data_table_for_singles.Rows(k).Item("CHAINAGE") & vbCrLf &
                                                                                               "CSF Chainage = " & csf_CH _
                                                                                               , 1, 0.2, 0.5, 2, 7)

                                            W1.Range(Col_east & start1).Value = Data_table_for_singles.Rows(k).Item("X")
                                            W1.Range(Col_north & start1).Value = Data_table_for_singles.Rows(k).Item("Y")
                                            W1.Range(Col_Elevation & start1).Value = Data_table_for_singles.Rows(k).Item("Z")
                                            W1.Range(Col_Chain & start1).Value = Data_table_for_singles.Rows(k).Item("CHAINAGE")
                                            W1.Range(Col_CSF_CHAIN & start1).Value = csf_CH
                                            start1 = start1 + 1
                                        Next

                                    End If

                                    Editor1.Regen()
                                    Trans1.Commit()
                                End If
                            End If
                        End Using

                    End If
                End If
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                MsgBox("Done")
                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
    End Sub



    Private Sub Button_output_to_excel_parallel_ends_Click(sender As Object, e As EventArgs) Handles Button_output_to_excel_parallel_ends.Click
        Try
            If TextBox_east_intersection.Text = "" Then
                MsgBox("Please specify the East COLUMN!")
                Exit Sub
            End If
            If TextBox_north_intersection.Text = "" Then
                MsgBox("Please specify the North COLUMN!")
                Exit Sub
            End If
            If TextBox_elevation.Text = "" Then
                MsgBox("Please specify the Elevation COLUMN!")
                Exit Sub
            End If
            If TextBox_chainage.Text = "" Then
                MsgBox("Please specify the Chainage COLUMN!")
                Exit Sub
            End If
            If TextBox_CSF_Chainage.Text = "" Then
                MsgBox("Please specify the CSF Chainage COLUMN!")
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
                Object_Prompt.MessageForAdding = vbLf & "Select 3D centerline:"
                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)


                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    If IsNothing(Rezultat1) = False Then
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction


                            Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                            Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                            Object_Prompt2.MessageForAdding = vbLf & "Select a line or polyline representing the heavy wall:"
                            Object_Prompt2.SingleOnly = True
                            Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                            If Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                If IsNothing(Rezultat2) = False Then

                                    Dim Data_table_for_singles As New System.Data.DataTable
                                    Data_table_for_singles.Columns.Add("X", GetType(Double))
                                    Data_table_for_singles.Columns.Add("Y", GetType(Double))
                                    Data_table_for_singles.Columns.Add("Z", GetType(Double))
                                    Data_table_for_singles.Columns.Add("CHAINAGE", GetType(Double))
                                    Data_table_for_singles.Columns.Add("CSFCHAINAGE", GetType(Double))


                                    Dim indexdt As Double = 0

                                    Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = get_new_worksheet_from_Excel()
                                    Dim Col_east As String = TextBox_east_intersection.Text.ToUpper
                                    Dim Col_north As String = TextBox_north_intersection.Text.ToUpper
                                    Dim Col_Chain As String = TextBox_chainage.Text.ToUpper
                                    Dim Col_Elevation As String = TextBox_elevation.Text.ToUpper
                                    Dim Col_CSF_CHAIN As String = TextBox_CSF_Chainage.Text.ToUpper




                                    Dim Poly3d As Polyline3d
                                    Dim Liniesample1 As Line
                                    Dim Polysample1 As Polyline

                                    Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                    BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                                    Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                    Obj1 = Rezultat1.Value.Item(0)
                                    Dim Ent1 As Entity
                                    Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                                    Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                    Obj2 = Rezultat2.Value.Item(0)
                                    Dim Ent2 As Entity
                                    Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)


                                    If TypeOf Ent1 Is Polyline3d And (TypeOf Ent2 Is Line Or TypeOf Ent2 Is Polyline) Then

                                        Poly3d = Ent1

                                        Dim Layersample As String
                                        If TypeOf Ent2 Is Line Then
                                            Liniesample1 = Ent2
                                            Layersample = Liniesample1.Layer
                                        Else
                                            Polysample1 = Ent2
                                            Layersample = Polysample1.Layer
                                        End If




                                        Dim Poly2D As New Polyline

                                        Dim Index2d As Double = 0
                                        For Each ObjId As ObjectId In Poly3d
                                            Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, OpenMode.ForRead)
                                            Poly2D.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)

                                            Index2d = Index2d + 1
                                        Next
                                        Poly2D.Elevation = 0


                                        Dim ChainageCSF_colection As New DBObjectCollection
                                        Dim CSF_colection As New DBObjectCollection


                                        Dim Data_table_Poly As New System.Data.DataTable
                                        Data_table_Poly.Columns.Add("2DPOLY", GetType(Curve))


                                        Dim IndexXX As Double = 0


                                        For Each ObjID In BTrecord
                                            Dim DBobject As DBObject = Trans1.GetObject(ObjID, OpenMode.ForRead)

                                            If TypeOf DBobject Is Curve Then
                                                Dim Poly1 As Curve = DBobject
                                                If Poly1.Layer = Layersample Then
                                                    Data_table_Poly.Rows.Add()
                                                    Data_table_Poly.Rows(IndexXX).Item("2DPOLY") = Poly1
                                                    IndexXX = IndexXX + 1
                                                End If

                                            End If
                                            If TypeOf DBobject Is DBText Then
                                                Dim CsfCh1 As DBText = DBobject
                                                If CsfCh1.TextString.Contains("+") = True Then
                                                    ChainageCSF_colection.Add(CsfCh1)
                                                End If
                                                If CsfCh1.TextString.Contains("CSF") = True Then
                                                    CSF_colection.Add(CsfCh1)
                                                End If
                                            End If

                                        Next




                                        If Data_table_Poly.Rows.Count > 0 Then
                                            For i = 0 To Data_table_Poly.Rows.Count - 1

                                                Dim Poly1 As Curve = Data_table_Poly.Rows(i).Item("2DPOLY")

                                                Dim Start_p1 As New Point3d(Poly1.StartPoint.X, Poly1.StartPoint.Y, 0)
                                                Dim End_p1 As New Point3d(Poly1.EndPoint.X, Poly1.EndPoint.Y, 0)


                                                Dim Start_p1P As New Point3d
                                                Dim Param_Start As Double = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(Start_p1, Vector3d.ZAxis, False))

                                                Start_p1P = Poly3d.GetPointAtParameter(Param_Start)
                                                Data_table_for_singles.Rows.Add()
                                                Data_table_for_singles.Rows(indexdt).Item("X") = Round(Start_p1P.X, 3)
                                                Data_table_for_singles.Rows(indexdt).Item("Y") = Round(Start_p1P.Y, 3)
                                                Data_table_for_singles.Rows(indexdt).Item("Z") = Round(Start_p1P.Z, 3)
                                                Data_table_for_singles.Rows(indexdt).Item("CHAINAGE") = Round(Poly3d.GetDistAtPoint(Start_p1P), 3)
                                                indexdt = indexdt + 1

                                                Dim End_p1P As New Point3d
                                                Dim Param_End As Double = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(End_p1, Vector3d.ZAxis, False))
                                                End_p1P = Poly3d.GetPointAtParameter(Param_End)
                                                Data_table_for_singles.Rows.Add()
                                                Data_table_for_singles.Rows(indexdt).Item("X") = Round(End_p1P.X, 3)
                                                Data_table_for_singles.Rows(indexdt).Item("Y") = Round(End_p1P.Y, 3)
                                                Data_table_for_singles.Rows(indexdt).Item("Z") = Round(End_p1P.Z, 3)
                                                Data_table_for_singles.Rows(indexdt).Item("CHAINAGE") = Round(Poly3d.GetDistAtPoint(End_p1P), 3)
                                                indexdt = indexdt + 1


                                                If CSF_colection.Count > 0 And ChainageCSF_colection.Count > 0 Then
                                                    Dim DistMin1 As Double = 50
                                                    Dim Chain_index1 As Double
                                                    Dim Pt_index1 As New Point3d

                                                    Dim DistMin2 As Double = 50
                                                    Dim Chain_index2 As Double
                                                    Dim Pt_index2 As New Point3d

                                                    For r = 0 To ChainageCSF_colection.Count - 1
                                                        Dim DbText1 As DBText = ChainageCSF_colection(r)

                                                        Dim Dist1 As Double
                                                        Dist1 = New Point3d(Start_p1P.X, Start_p1P.Y, 0).GetVectorTo(New Point3d(DbText1.Position.X, DbText1.Position.Y, 0)).Length
                                                        If Dist1 < DistMin1 And IsNumeric(Replace(DbText1.TextString, "+", "")) = True Then
                                                            DistMin1 = Dist1
                                                            Chain_index1 = CDbl(Replace(DbText1.TextString, "+", ""))
                                                            Pt_index1 = DbText1.Position
                                                        End If
                                                        Dim Dist2 As Double
                                                        Dist2 = New Point3d(End_p1P.X, End_p1P.Y, 0).GetVectorTo(New Point3d(DbText1.Position.X, DbText1.Position.Y, 0)).Length
                                                        If Dist2 < DistMin2 And IsNumeric(Replace(DbText1.TextString, "+", "")) = True Then
                                                            DistMin2 = Dist2
                                                            Chain_index2 = CDbl(Replace(DbText1.TextString, "+", ""))
                                                            Pt_index2 = DbText1.Position
                                                        End If

                                                    Next

                                                    Dim Csf1 As Double
                                                    Dim Csf2 As Double

                                                    For r = 0 To CSF_colection.Count - 1
                                                        Dim DbText1 As DBText = CSF_colection(r)
                                                        Dim String1 As String = DbText1.TextString
                                                        String1 = extrage_numar_din_text_de_la_sfarsitul_textului(String1)
                                                        If IsNumeric(String1) = True And Abs(Pt_index1.X - DbText1.Position.X) < 0.1 And Abs(Pt_index1.Y - DbText1.Position.Y) < 0.1 Then
                                                            Csf1 = CDbl(String1)
                                                        End If
                                                        If IsNumeric(String1) = True And Abs(Pt_index2.X - DbText1.Position.X) < 0.1 And Abs(Pt_index2.Y - DbText1.Position.Y) < 0.1 Then
                                                            Csf2 = CDbl(String1)
                                                        End If
                                                    Next

                                                    Dim pT_1 As New Point3d
                                                    pT_1 = Poly2D.GetClosestPointTo(Pt_index1, Vector3d.ZAxis, False)
                                                    Dim Param_1 As Double = Poly2D.GetParameterAtPoint(pT_1)
                                                    Dim Len_1 As Double = Poly3d.GetDistanceAtParameter(Param_1)


                                                    Data_table_for_singles.Rows(indexdt - 2).Item("CSFCHAINAGE") = Chain_index1 + (Poly3d.GetDistAtPoint(Start_p1P) - Len_1) / Csf1

                                                    Dim pT_2 As New Point3d
                                                    pT_2 = Poly2D.GetClosestPointTo(Pt_index2, Vector3d.ZAxis, False)
                                                    Dim Param_2 As Double = Poly2D.GetParameterAtPoint(pT_2)
                                                    Dim Len_2 As Double = Poly3d.GetDistanceAtParameter(Param_2)


                                                    Data_table_for_singles.Rows(indexdt - 1).Item("CSFCHAINAGE") = Chain_index2 + (Poly3d.GetDistAtPoint(End_p1P) - Len_2) / Csf2

                                                End If






                                            Next
                                        End If



                                    End If


                                    If Data_table_for_singles.Rows.Count > 0 Then

                                        Dim start1 As Integer = 2
                                        W1.Range(Col_east & "1").Value = "East"
                                        W1.Range(Col_north & "1").Value = "North"
                                        W1.Range(Col_Elevation & "1").Value = "Elevation"
                                        W1.Range(Col_Chain & "1").Value = "Grid Chainage"
                                        W1.Range(Col_CSF_CHAIN & "1").Value = "CSF Chainage"

                                        For k = 0 To Data_table_for_singles.Rows.Count - 1
                                            Dim csf_CH As Double = 0
                                            If IsDBNull(Data_table_for_singles.Rows(k).Item("CSFCHAINAGE")) = False Then
                                                csf_CH = Round(Data_table_for_singles.Rows(k).Item("CSFCHAINAGE"), 3)
                                            End If

                                            Dim new_leader As New MLeader
                                            new_leader = Creaza_Mleader_nou_fara_UCS_transform(New Point3d(Data_table_for_singles.Rows(k).Item("X"), Data_table_for_singles.Rows(k).Item("Y"), Data_table_for_singles.Rows(k).Item("Z")),
                                                                                               "East = " & Data_table_for_singles.Rows(k).Item("X") & vbCrLf &
                                                                                               "North = " & Data_table_for_singles.Rows(k).Item("Y") & vbCrLf &
                                                                                               "Elev = " & Data_table_for_singles.Rows(k).Item("Z") & vbCrLf &
                                                                                               "Grid Chainage = " & Data_table_for_singles.Rows(k).Item("CHAINAGE") & vbCrLf &
                                                                                               "CSF Chainage = " & csf_CH _
                                                                                               , 1, 0.2, 0.5, 2, 7)

                                            W1.Range(Col_east & start1).Value = Data_table_for_singles.Rows(k).Item("X")
                                            W1.Range(Col_north & start1).Value = Data_table_for_singles.Rows(k).Item("Y")
                                            W1.Range(Col_Elevation & start1).Value = Data_table_for_singles.Rows(k).Item("Z")
                                            W1.Range(Col_Chain & start1).Value = Data_table_for_singles.Rows(k).Item("CHAINAGE")
                                            W1.Range(Col_CSF_CHAIN & start1).Value = csf_CH
                                            start1 = start1 + 1
                                        Next

                                    End If

                                    Editor1.Regen()
                                    Trans1.Commit()
                                End If
                            End If
                        End Using

                    End If
                End If
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                MsgBox("Done")
                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
    End Sub







    Private Sub Button_start_end_without_CSF_Click(sender As Object, e As EventArgs) Handles Button_start_end_without_CSF.Click
        Try
            If TextBox_east_intersection.Text = "" Then
                MsgBox("Please specify the East COLUMN!")
                Exit Sub
            End If
            If TextBox_north_intersection.Text = "" Then
                MsgBox("Please specify the North COLUMN!")
                Exit Sub
            End If
            If TextBox_elevation.Text = "" Then
                MsgBox("Please specify the Elevation COLUMN!")
                Exit Sub
            End If
            If TextBox_chainage.Text = "" Then
                MsgBox("Please specify the Chainage COLUMN!")
                Exit Sub
            End If
            If TextBox_CSF_Chainage.Text = "" Then
                MsgBox("Please specify the CSF Chainage COLUMN!")
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


                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    If IsNothing(Rezultat1) = False Then
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction


                            Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                            Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                            Object_Prompt2.MessageForAdding = vbLf & "Select a sample segment:"
                            Object_Prompt2.SingleOnly = True
                            Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                            If Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                If IsNothing(Rezultat2) = False Then

                                    Dim Data_table_for_singles As New System.Data.DataTable
                                    Data_table_for_singles.Columns.Add("X", GetType(Double))
                                    Data_table_for_singles.Columns.Add("Y", GetType(Double))
                                    Data_table_for_singles.Columns.Add("Z", GetType(Double))
                                    Data_table_for_singles.Columns.Add("STATION", GetType(Double))
                                    Data_table_for_singles.Columns.Add("INDEX", GetType(Double))


                                    Dim indexdt As Double = 0

                                    Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = get_new_worksheet_from_Excel()
                                    Dim Col_east As String = TextBox_east_intersection.Text.ToUpper
                                    Dim Col_north As String = TextBox_north_intersection.Text.ToUpper
                                    Dim Col_station As String = TextBox_chainage.Text.ToUpper
                                    Dim Col_Elevation As String = TextBox_elevation.Text.ToUpper





                                    Dim Poly3d As Polyline3d
                                    Dim Liniesample1 As Line
                                    Dim Polysample1 As Polyline

                                    Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                    BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                                    Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                    Obj1 = Rezultat1.Value.Item(0)
                                    Dim Ent1 As Entity
                                    Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                                    Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                    Obj2 = Rezultat2.Value.Item(0)
                                    Dim Ent2 As Entity
                                    Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)


                                    If TypeOf Ent1 Is Polyline3d And (TypeOf Ent2 Is Line Or TypeOf Ent2 Is Polyline) Then

                                        Poly3d = Ent1

                                        Dim Layersample As String
                                        If TypeOf Ent2 Is Line Then
                                            Liniesample1 = Ent2
                                            Layersample = Liniesample1.Layer
                                        Else
                                            Polysample1 = Ent2
                                            Layersample = Polysample1.Layer
                                        End If




                                        Dim Poly2D As New Polyline

                                        Dim Index2d As Double = 0
                                        For Each ObjId As ObjectId In Poly3d
                                            Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, OpenMode.ForRead)
                                            Poly2D.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)

                                            Index2d = Index2d + 1
                                        Next
                                        Poly2D.Elevation = 0


                                        Dim Data_table_Poly As New System.Data.DataTable
                                        Data_table_Poly.Columns.Add("2DPOLY", GetType(Curve))


                                        Dim IndexXX As Double = 0


                                        For Each ObjID In BTrecord
                                            Dim DBobject As DBObject = Trans1.GetObject(ObjID, OpenMode.ForRead)

                                            If TypeOf DBobject Is Curve Then
                                                Dim Poly1 As Curve = DBobject
                                                If Poly1.Layer = Layersample Then
                                                    Data_table_Poly.Rows.Add()
                                                    Data_table_Poly.Rows(IndexXX).Item("2DPOLY") = Poly1
                                                    IndexXX = IndexXX + 1
                                                End If

                                            End If
                                        Next



                                        If Data_table_Poly.Rows.Count > 0 Then
                                            For i = 0 To Data_table_Poly.Rows.Count - 1

                                                Dim Poly1 As Curve = Data_table_Poly.Rows(i).Item("2DPOLY")

                                                Dim Start_p1 As New Point3d(Poly1.StartPoint.X, Poly1.StartPoint.Y, 0)
                                                Dim End_p1 As New Point3d(Poly1.EndPoint.X, Poly1.EndPoint.Y, 0)


                                                Dim Start_p1P As New Point3d
                                                Dim Param_Start As Double = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(Start_p1, Vector3d.ZAxis, False))

                                                Start_p1P = Poly3d.GetPointAtParameter(Param_Start)
                                                Data_table_for_singles.Rows.Add()
                                                Data_table_for_singles.Rows(indexdt).Item("X") = Round(Start_p1P.X, 3)
                                                Data_table_for_singles.Rows(indexdt).Item("Y") = Round(Start_p1P.Y, 3)
                                                Data_table_for_singles.Rows(indexdt).Item("Z") = Round(Start_p1P.Z, 3)
                                                Data_table_for_singles.Rows(indexdt).Item("STATION") = Round(Poly3d.GetDistAtPoint(Start_p1P), 3)
                                                Data_table_for_singles.Rows(indexdt).Item("INDEX") = i + 1
                                                indexdt = indexdt + 1

                                                Dim End_p1P As New Point3d
                                                Dim Param_End As Double = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(End_p1, Vector3d.ZAxis, False))
                                                End_p1P = Poly3d.GetPointAtParameter(Param_End)
                                                Data_table_for_singles.Rows.Add()
                                                Data_table_for_singles.Rows(indexdt).Item("X") = Round(End_p1P.X, 3)
                                                Data_table_for_singles.Rows(indexdt).Item("Y") = Round(End_p1P.Y, 3)
                                                Data_table_for_singles.Rows(indexdt).Item("Z") = Round(End_p1P.Z, 3)
                                                Data_table_for_singles.Rows(indexdt).Item("STATION") = Round(Poly3d.GetDistAtPoint(End_p1P), 3)
                                                Data_table_for_singles.Rows(indexdt).Item("INDEX") = i + 1
                                                indexdt = indexdt + 1









                                            Next
                                        End If



                                    End If

                                    If TypeOf Ent1 Is Polyline And (TypeOf Ent2 Is Line Or TypeOf Ent2 Is Polyline) Then

                                        Dim Layersample As String
                                        If TypeOf Ent2 Is Line Then
                                            Liniesample1 = Ent2
                                            Layersample = Liniesample1.Layer
                                        Else
                                            Polysample1 = Ent2
                                            Layersample = Polysample1.Layer
                                        End If




                                        Dim Poly2D As Polyline = Ent1



                                        Dim Data_table_Poly As New System.Data.DataTable
                                        Data_table_Poly.Columns.Add("2DPOLY", GetType(Curve))

                                        Dim IndexXX As Double = 0

                                        For Each ObjID In BTrecord
                                            Dim DBobject As DBObject = Trans1.GetObject(ObjID, OpenMode.ForRead)

                                            If TypeOf DBobject Is Curve Then
                                                Dim Poly1 As Curve = DBobject
                                                If Poly1.Layer = Layersample Then
                                                    Data_table_Poly.Rows.Add()
                                                    Data_table_Poly.Rows(IndexXX).Item("2DPOLY") = Poly1
                                                    IndexXX = IndexXX + 1
                                                End If

                                            End If
                                        Next


                                        If Data_table_Poly.Rows.Count > 0 Then
                                            For i = 0 To Data_table_Poly.Rows.Count - 1

                                                Dim Poly1 As Curve = Data_table_Poly.Rows(i).Item("2DPOLY")

                                                Dim Start_p1 As New Point3d(Poly1.StartPoint.X, Poly1.StartPoint.Y, 0)
                                                Dim End_p1 As New Point3d(Poly1.EndPoint.X, Poly1.EndPoint.Y, 0)


                                                Dim Start_p1P As New Point3d
                                                Start_p1P = Poly2D.GetClosestPointTo(Start_p1, Vector3d.ZAxis, False)
                                                Data_table_for_singles.Rows.Add()
                                                Data_table_for_singles.Rows(indexdt).Item("X") = Round(Start_p1P.X, 3)
                                                Data_table_for_singles.Rows(indexdt).Item("Y") = Round(Start_p1P.Y, 3)
                                                Data_table_for_singles.Rows(indexdt).Item("Z") = Round(Start_p1P.Z, 3)
                                                Data_table_for_singles.Rows(indexdt).Item("STATION") = Round(Poly2D.GetDistAtPoint(Start_p1P), 3)
                                                Data_table_for_singles.Rows(indexdt).Item("INDEX") = i + 1
                                                indexdt = indexdt + 1

                                                Dim End_p1P As New Point3d

                                                End_p1P = Poly2D.GetClosestPointTo(End_p1, Vector3d.ZAxis, False)
                                                Data_table_for_singles.Rows.Add()
                                                Data_table_for_singles.Rows(indexdt).Item("X") = Round(End_p1P.X, 3)
                                                Data_table_for_singles.Rows(indexdt).Item("Y") = Round(End_p1P.Y, 3)
                                                Data_table_for_singles.Rows(indexdt).Item("Z") = Round(End_p1P.Z, 3)
                                                Data_table_for_singles.Rows(indexdt).Item("STATION") = Round(Poly2D.GetDistAtPoint(End_p1P), 3)
                                                Data_table_for_singles.Rows(indexdt).Item("INDEX") = i + 1
                                                indexdt = indexdt + 1









                                            Next
                                        End If



                                    End If


                                    If Data_table_for_singles.Rows.Count > 0 Then

                                        Dim start1 As Integer = 2
                                        W1.Range(Col_east & "1").Value = "East"
                                        W1.Range(Col_north & "1").Value = "North"
                                        W1.Range(Col_Elevation & "1").Value = "Elevation"
                                        W1.Range(Col_station & "1").Value = "STATION"

                                        Dim Last_col As Integer = W1.Range(Col_east & "1").Column

                                        If W1.Range(Col_north & "1").Column > Last_col Then
                                            Last_col = W1.Range(Col_north & "1").Column
                                        End If

                                        If W1.Range(Col_Elevation & "1").Column > Last_col Then
                                            Last_col = W1.Range(Col_Elevation & "1").Column
                                        End If

                                        If W1.Range(Col_station & "1").Column > Last_col Then
                                            Last_col = W1.Range(Col_station & "1").Column
                                        End If

                                        Last_col = Last_col + 1
                                        W1.Cells(1, Last_col).Value = "INDEX"

                                        For k = 0 To Data_table_for_singles.Rows.Count - 1

                                            Dim new_leader As New MLeader
                                            new_leader = Creaza_Mleader_nou_fara_UCS_transform(New Point3d(Data_table_for_singles.Rows(k).Item("X"), Data_table_for_singles.Rows(k).Item("Y"), Data_table_for_singles.Rows(k).Item("Z")),
                                                                                               "East = " & Data_table_for_singles.Rows(k).Item("X") & vbCrLf &
                                                                                               "North = " & Data_table_for_singles.Rows(k).Item("Y") & vbCrLf &
                                                                                               "Elev = " & Data_table_for_singles.Rows(k).Item("Z") & vbCrLf &
                                                                                               "Station = " & Data_table_for_singles.Rows(k).Item("STATION"), 1, 0.2, 0.5, 2, 7)



                                            W1.Range(Col_east & start1).Value = Data_table_for_singles.Rows(k).Item("X")
                                            W1.Range(Col_north & start1).Value = Data_table_for_singles.Rows(k).Item("Y")
                                            W1.Range(Col_Elevation & start1).Value = Data_table_for_singles.Rows(k).Item("Z")
                                            W1.Range(Col_station & start1).Value = Data_table_for_singles.Rows(k).Item("STATION")
                                            W1.Cells(start1, Last_col).Value = Data_table_for_singles.Rows(k).Item("INDEX")
                                            start1 = start1 + 1
                                        Next

                                    End If

                                    Editor1.Regen()
                                    Trans1.Commit()
                                End If
                            End If
                        End Using

                    End If
                End If
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                MsgBox("Done")
                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
    End Sub



    Private Sub Button_CSF_TO_AUTOCAD_Click(sender As Object, e As EventArgs) Handles Button_CSF_TO_AUTOCAD.Click
        Try
            If TextBox_east_intersection.Text = "" Then
                MsgBox("Please specify the East COLUMN!")
                Exit Sub
            End If
            If TextBox_north_intersection.Text = "" Then
                MsgBox("Please specify the North COLUMN!")
                Exit Sub
            End If
            If TextBox_elevation.Text = "" Then
                MsgBox("Please specify the Elevation COLUMN!")
                Exit Sub
            End If

            If TextBox_CSF_Chainage.Text = "" Then
                MsgBox("Please specify the CSF Chainage COLUMN!")
                Exit Sub
            End If

            If IsNumeric(TextBox_start.Text) = False Then
                MsgBox("Please specify the Start Row")
                Exit Sub
            End If
            If IsNumeric(TextBox_end.Text) = False Then
                MsgBox("Please specify the End Row")
                Exit Sub
            End If
            If CInt(TextBox_end.Text) < CInt(TextBox_start.Text) Then
                MsgBox("Start row smaller than end row")
                Exit Sub
            End If
            If CInt(TextBox_end.Text) < 1 Then
                MsgBox("End row smaller than 1")
                Exit Sub
            End If
            If CInt(TextBox_start.Text) < 1 Then
                MsgBox("Start row smaller than 1")
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
                Object_Prompt.MessageForAdding = vbLf & "Select the centerline:"
                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
                Dim Col_east As String = TextBox_east_intersection.Text.ToUpper
                Dim Col_north As String = TextBox_north_intersection.Text.ToUpper
                Dim Col_Elevation As String = TextBox_elevation.Text.ToUpper
                Dim Col_CSF_CHAIN As String = TextBox_CSF_Chainage.Text.ToUpper
                Dim start1 As Integer = CInt(TextBox_start.Text)
                Dim end1 As Integer = CInt(TextBox_end.Text)
                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    If IsNothing(Rezultat1) = False Then
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim Data_table_3D_POLY As New System.Data.DataTable
                            Data_table_3D_POLY.Columns.Add("X", GetType(Double))
                            Data_table_3D_POLY.Columns.Add("Y", GetType(Double))
                            Data_table_3D_POLY.Columns.Add("Z", GetType(Double))
                            Dim indexdt As Double = 0
                            Dim Poly3d As Polyline3d
                            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                            BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(0)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                            If TypeOf Ent1 Is Polyline3d Then
                                Poly3d = Ent1
                                Dim ChainageCSF_colection As New DBObjectCollection
                                Dim CSF_colection As New DBObjectCollection
                                Dim Poly2D As New Polyline
                                Dim Index2d As Double = 0
                                For Each ObjId As ObjectId In Poly3d
                                    Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, OpenMode.ForRead)
                                    Poly2D.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)
                                    Index2d = Index2d + 1
                                Next
                                For Each ObjID In BTrecord
                                    Dim DBobject As DBObject = Trans1.GetObject(ObjID, OpenMode.ForRead)
                                    If TypeOf DBobject Is DBText Then
                                        Dim CsfCh1 As DBText = DBobject
                                        If CsfCh1.TextString.Contains("+") = True Then
                                            ChainageCSF_colection.Add(CsfCh1)
                                        End If
                                        If CsfCh1.TextString.Contains("CSF") = True Then
                                            CSF_colection.Add(CsfCh1)
                                        End If
                                    End If
                                Next
                                If CSF_colection.Count > 0 And ChainageCSF_colection.Count > 0 Then
                                    For i = start1 To end1
                                        Dim ChainageCSF_txt As String
                                        ChainageCSF_txt = W1.Range(Col_CSF_CHAIN & i).Value
                                        If IsNumeric(Replace(ChainageCSF_txt, "+", "")) = True Then
                                            Dim ChainageCSF As Double
                                            ChainageCSF = CDbl(Replace(ChainageCSF_txt, "+", ""))
                                            Dim DistMin1 As Double = 50
                                            Dim Chain_index1 As Double
                                            Dim Pt_index1 As New Point3d
                                            For r = 0 To ChainageCSF_colection.Count - 1
                                                Dim DbText1 As DBText = ChainageCSF_colection(r)
                                                If IsNumeric(Replace(DbText1.TextString, "+", "")) = True Then
                                                    Dim Dist1 As Double = CDbl(Replace(DbText1.TextString, "+", ""))
                                                    Dist1 = Dist1 - ChainageCSF
                                                    If Abs(Dist1) < Abs(DistMin1) Then
                                                        DistMin1 = Dist1
                                                        Chain_index1 = CDbl(Replace(DbText1.TextString, "+", ""))
                                                        Pt_index1 = DbText1.Position
                                                    End If
                                                End If
                                            Next

                                            Dim Csf1 As Double
                                            For r = 0 To CSF_colection.Count - 1
                                                Dim DbText1 As DBText = CSF_colection(r)
                                                Dim String1 As String = DbText1.TextString
                                                String1 = extrage_numar_din_text_de_la_sfarsitul_textului(String1)
                                                If IsNumeric(String1) = True And Abs(Pt_index1.X - DbText1.Position.X) < 0.1 And Abs(Pt_index1.Y - DbText1.Position.Y) < 0.1 Then
                                                    Csf1 = CDbl(String1)
                                                End If
                                            Next
                                            Dim pT_1 As New Point3d
                                            pT_1 = Poly2D.GetClosestPointTo(Pt_index1, Vector3d.ZAxis, False)
                                            Dim Param_1 As Double = Poly2D.GetParameterAtPoint(pT_1)
                                            Dim Len_1 As Double = Poly3d.GetDistanceAtParameter(Param_1)
                                            Dim Pt_at_Chainage_CSF As New Point3d
                                            Pt_at_Chainage_CSF = Poly3d.GetPointAtDist(Len_1 - DistMin1 * Csf1)


                                            W1.Range(Col_east & i).Value = Round(Pt_at_Chainage_CSF.X, 3)
                                            W1.Range(Col_north & i).Value = Round(Pt_at_Chainage_CSF.Y, 3)
                                            W1.Range(Col_Elevation & i).Value = Round(Pt_at_Chainage_CSF.Z, 3)
                                            Dim new_leader As New MLeader
                                            new_leader = Creaza_Mleader_nou_fara_UCS_transform(Pt_at_Chainage_CSF,
                                                                                               "East = " & Round(Pt_at_Chainage_CSF.X, 3) & vbCrLf &
                                                                                               "North = " & Round(Pt_at_Chainage_CSF.Y, 3) & vbCrLf &
                                                                                               "Elev = " & Round(Pt_at_Chainage_CSF.Z, 3) & vbCrLf &
                                                                                               "CSF Chainage = " & Get_chainage_from_double(ChainageCSF, 3) _
                                                                                               , 1, 0.2, 0.5, 2, 7)
                                        End If
                                    Next
                                    Editor1.Regen()
                                    Trans1.Commit()
                                Else
                                    Dim Pt_at_Chainage As New Point3d
                                    For i = start1 To end1
                                        Dim Chainage_txt As String
                                        Chainage_txt = W1.Range(Col_CSF_CHAIN & i).Value
                                        If IsNumeric(Replace(Chainage_txt, "+", "")) = True Then
                                            Dim Chainage As Double
                                            Chainage = CDbl(Replace(Chainage_txt, "+", ""))

                                            Pt_at_Chainage = Poly3d.GetPointAtDist(Chainage)
                                            W1.Range(Col_east & i).Value = Round(Pt_at_Chainage.X, 3)
                                            W1.Range(Col_north & i).Value = Round(Pt_at_Chainage.Y, 3)
                                            W1.Range(Col_Elevation & i).Value = Round(Pt_at_Chainage.Z, 3)
                                            Dim new_leader As New MLeader
                                            new_leader = Creaza_Mleader_nou_fara_UCS_transform(Pt_at_Chainage,
                                                                                               "East = " & Round(Pt_at_Chainage.X, 3) & vbCrLf &
                                                                                               "North = " & Round(Pt_at_Chainage.Y, 3) & vbCrLf &
                                                                                               "Elev = " & Round(Pt_at_Chainage.Z, 3) & vbCrLf &
                                                                                               "Chainage = " & Get_chainage_from_double(Chainage, 3) _
                                                                                               , 1, 0.2, 0.5, 2, 7)
                                        End If

                                    Next



                                    Editor1.Regen()
                                    Trans1.Commit()
                                End If
                            End If
                            If TypeOf Ent1 Is Polyline Then
                                Dim PolyLW As Polyline
                                PolyLW = Ent1


                                Dim Pt_at_Chainage As New Point3d
                                For i = start1 To end1
                                    Dim Chainage_txt As String
                                    Chainage_txt = W1.Range(Col_CSF_CHAIN & i).Value
                                    If IsNumeric(Replace(Chainage_txt, "+", "")) = True Then
                                        Dim Chainage As Double
                                        Chainage = CDbl(Replace(Chainage_txt, "+", ""))

                                        Pt_at_Chainage = PolyLW.GetPointAtDist(Chainage)
                                        W1.Range(Col_east & i).Value = Round(Pt_at_Chainage.X, 3)
                                        W1.Range(Col_north & i).Value = Round(Pt_at_Chainage.Y, 3)
                                        W1.Range(Col_Elevation & i).Value = Round(Pt_at_Chainage.Z, 3)
                                        Dim new_leader As New MLeader
                                        new_leader = Creaza_Mleader_nou_fara_UCS_transform(Pt_at_Chainage,
                                                                                           "East = " & Round(Pt_at_Chainage.X, 3) & vbCrLf &
                                                                                           "North = " & Round(Pt_at_Chainage.Y, 3) & vbCrLf &
                                                                                           "Elev = " & Round(Pt_at_Chainage.Z, 3) & vbCrLf &
                                                                                           "Chainage = " & Get_chainage_from_double(Chainage, 3) _
                                                                                           , 1, 0.2, 0.5, 2, 7)
                                    End If

                                Next



                                Editor1.Regen()
                                Trans1.Commit()

                            End If

                        End Using
                    End If
                End If
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                MsgBox("Done")
                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_read_sta_wr2xl_xy_Click(sender As Object, e As EventArgs) Handles Button_read_sta_wr_2xl.Click
        Try
            If TextBox_east_intersection.Text = "" Then
                MsgBox("Please specify the East COLUMN!")
                Exit Sub
            End If
            If TextBox_north_intersection.Text = "" Then
                MsgBox("Please specify the North COLUMN!")
                Exit Sub
            End If
            If TextBox_elevation.Text = "" Then
                MsgBox("Please specify the Elevation COLUMN!")
                Exit Sub
            End If


            If IsNumeric(TextBox_start.Text) = False Then
                MsgBox("Please specify the Start Row")
                Exit Sub
            End If
            If IsNumeric(TextBox_end.Text) = False Then
                MsgBox("Please specify the End Row")
                Exit Sub
            End If
            If CInt(TextBox_end.Text) < CInt(TextBox_start.Text) Then
                MsgBox("Start row smaller than end row")
                Exit Sub
            End If
            If CInt(TextBox_end.Text) < 1 Then
                MsgBox("End row smaller than 1")
                Exit Sub
            End If
            If CInt(TextBox_start.Text) < 1 Then
                MsgBox("Start row smaller than 1")
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
                Object_Prompt.MessageForAdding = vbLf & "Select the polyline"
                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
                Dim Col_east As String = TextBox_east_intersection.Text.ToUpper
                Dim Col_north As String = TextBox_north_intersection.Text.ToUpper
                Dim Col_Elevation As String = TextBox_elevation.Text.ToUpper
                Dim Col_sta As String = TextBox_chainage.Text.ToUpper
                Dim start1 As Integer = CInt(TextBox_start.Text)
                Dim end1 As Integer = CInt(TextBox_end.Text)
                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    If IsNothing(Rezultat1) = False Then
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim Data_table_poly As New System.Data.DataTable
                            Data_table_poly.Columns.Add("X", GetType(Double))
                            Data_table_poly.Columns.Add("Y", GetType(Double))
                            Data_table_poly.Columns.Add("Z", GetType(Double))
                            Dim indexdt As Double = 0
                            Dim Poly1 As Polyline
                            Dim Poly3D As Polyline
                            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                            BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(0)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                            If TypeOf Ent1 Is Polyline Then
                                Poly1 = Ent1


                                For i = start1 To end1
                                    Dim Stationtxt As String
                                    Stationtxt = W1.Range(Col_sta & i).Value
                                    If IsNumeric(Replace(Stationtxt, "+", "")) = True Then
                                        Dim station1 As Double
                                        station1 = CDbl(Replace(Stationtxt, "+", ""))
                                        If station1 <= Poly1.Length Then
                                            Dim pt_on_curve = Poly1.GetPointAtDist(station1)
                                            W1.Range(Col_east & i).Value = Round(pt_on_curve.X, 2)
                                            W1.Range(Col_north & i).Value = Round(pt_on_curve.Y, 2)
                                        End If




                                    End If
                                Next
                                Editor1.Regen()
                                Trans1.Commit()

                            End If
                            If TypeOf Ent1 Is Polyline3d Then
                                Poly3D = Ent1
                                For i = start1 To end1
                                    Dim Stationtxt As String
                                    Stationtxt = W1.Range(Col_sta & i).Value
                                    If IsNumeric(Replace(Stationtxt, "+", "")) = True Then
                                        Dim station1 As Double
                                        station1 = CDbl(Replace(Stationtxt, "+", ""))
                                        If station1 <= Poly3D.Length Then
                                            Dim pt_on_curve = Poly3D.GetPointAtDist(station1)
                                            W1.Range(Col_east & i).Value = Round(pt_on_curve.X, 2)
                                            W1.Range(Col_north & i).Value = Round(pt_on_curve.Y, 2)
                                            W1.Range(Col_Elevation & i).Value = Round(pt_on_curve.Z, 2)
                                        End If




                                    End If

                                Next



                                Editor1.Regen()
                                Trans1.Commit()

                            End If

                        End Using
                    End If
                End If
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                MsgBox("Done")
                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_rerouteMP_Click(sender As Object, e As EventArgs) Handles Button_rerouteMP.Click
        Try
            If TextBox_east_intersection.Text = "" Then
                MsgBox("Please specify the East COLUMN!")
                Exit Sub
            End If
            If TextBox_north_intersection.Text = "" Then
                MsgBox("Please specify the North COLUMN!")
                Exit Sub
            End If
            If TextBox_elevation.Text = "" Then
                MsgBox("Please specify the Elevation COLUMN!")
                Exit Sub
            End If


            If IsNumeric(TextBox_start.Text) = False Then
                MsgBox("Please specify the Start Row")
                Exit Sub
            End If

            If CInt(TextBox_start.Text) < 1 Then
                MsgBox("Start row smaller than 1")
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
                Object_Prompt.MessageForAdding = vbLf & "Select the polyline"
                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
                Dim Col_east As String = TextBox_east_intersection.Text.ToUpper
                Dim Col_north As String = TextBox_north_intersection.Text.ToUpper
                Dim Col_Elevation As String = TextBox_elevation.Text.ToUpper
                Dim Col_sta As String = TextBox_chainage.Text.ToUpper
                Dim start1 As Integer = CInt(TextBox_start.Text)

                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    If IsNothing(Rezultat1) = False Then
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim Data_table_poly As New System.Data.DataTable
                            Data_table_poly.Columns.Add("X", GetType(Double))
                            Data_table_poly.Columns.Add("Y", GetType(Double))
                            Data_table_poly.Columns.Add("Z", GetType(Double))
                            Dim indexdt As Double = 0
                            Dim Poly1 As Polyline
                            Dim Poly3D As Polyline
                            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                            BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(0)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                            If TypeOf Ent1 Is Polyline Then
                                Poly1 = Ent1

123:
                                Dim Result_point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                                Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Please pick the reroute start:")

                                PP1.AllowNone = False
                                Result_point1 = Editor1.GetPoint(PP1)
                                If Result_point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                    Trans1.Commit()
                                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                    Exit Sub
                                End If


                                Dim Result_point2 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                                Dim PP2 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Please pick the reroute end:")

                                PP2.AllowNone = False
                                Result_point2 = Editor1.GetPoint(PP2)
                                If Result_point2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                    Trans1.Commit()
                                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                    Exit Sub
                                End If

                                Dim Sta_start As Double = -1

                                Dim rez_sta As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify the start station:")
                                rez_sta.AllowNone = False
                                Dim Rezultat22 As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(rez_sta)
                                If Rezultat22.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                    Sta_start = Rezultat22.Value
                                End If


                                If Sta_start > -1 Then
                                    Dim pt1 As New Point3d()
                                    pt1 = Poly1.GetClosestPointTo(Result_point1.Value, Vector3d.ZAxis, False)

                                    Dim pt2 As New Point3d()
                                    pt2 = Poly1.GetClosestPointTo(Result_point2.Value, Vector3d.ZAxis, False)

                                    Dim l_start = Poly1.GetDistAtPoint(pt1)
                                    Dim l_end = Poly1.GetDistAtPoint(pt2)

                                    Dim reroute_length = l_end - l_start

                                    Dim interval = 528

                                    Dim startMP As Double = Sta_start / interval

                                    Dim Mp = Ceiling(startMP)
                                    Dim Diferenta1 As Double = Mp * interval - Sta_start


                                    Dim l1 As Double = l_start + Diferenta1

                                    Do While (l1 <= l_end)
                                        W1.Range(Col_sta & start1).Value = Mp / 10
                                        Dim pt_on_curve = Poly1.GetPointAtDist(l1)
                                        W1.Range(Col_east & start1).Value = Round(pt_on_curve.X, 4)
                                        W1.Range(Col_north & start1).Value = Round(pt_on_curve.Y, 4)
                                        start1 = start1 + 1
                                        Mp = Mp + 1
                                        l1 = l1 + interval


                                        Dim DBP As New DBPoint(pt_on_curve)
                                        BTrecord.AppendEntity(DBP)
                                        Trans1.AddNewlyCreatedDBObject(DBP, True)

                                    Loop






                                End If



                            End If
                            GoTo 123

                        End Using
                    End If
                End If
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                MsgBox("Done")
                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_pt_to_CSF_Click(sender As Object, e As EventArgs) Handles Button_pt_to_CSF.Click
        Try
            If TextBox_east_intersection.Text = "" Then
                MsgBox("Please specify the East COLUMN!")
                Exit Sub
            End If
            If TextBox_north_intersection.Text = "" Then
                MsgBox("Please specify the North COLUMN!")
                Exit Sub
            End If

            If TextBox_CSF_Chainage.Text = "" Then
                MsgBox("Please specify the CSF Chainage COLUMN!")
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
                Object_Prompt.MessageForAdding = vbLf & "Select 3D centerline:"
                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
                Dim Col_east As String = TextBox_east_intersection.Text.ToUpper
                Dim Col_north As String = TextBox_north_intersection.Text.ToUpper
                Dim Col_CSF_CHAIN As String = TextBox_CSF_Chainage.Text.ToUpper
                Dim start1 As Integer = CInt(TextBox_start.Text)
                Dim end1 As Integer = CInt(TextBox_end.Text)
                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    If IsNothing(Rezultat1) = False Then
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim Data_table_3D_POLY As New System.Data.DataTable
                            Data_table_3D_POLY.Columns.Add("X", GetType(Double))
                            Data_table_3D_POLY.Columns.Add("Y", GetType(Double))
                            Data_table_3D_POLY.Columns.Add("Z", GetType(Double))
                            Dim indexdt As Double = 0
                            Dim Poly3d As Polyline3d
                            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                            BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(0)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                            If TypeOf Ent1 Is Polyline3d Then
                                Poly3d = Ent1
                                Dim ChainageCSF_colection As New DBObjectCollection
                                Dim CSF_colection As New DBObjectCollection
                                Dim Poly2D As New Polyline
                                Dim Index2d As Double = 0
                                For Each ObjId As ObjectId In Poly3d
                                    Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, OpenMode.ForRead)
                                    Poly2D.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)
                                    Index2d = Index2d + 1
                                Next
                                Poly2D.Elevation = 0
                                For Each ObjID In BTrecord
                                    Dim DBobject As DBObject = Trans1.GetObject(ObjID, OpenMode.ForRead)
                                    If TypeOf DBobject Is DBText Then
                                        Dim CsfCh1 As DBText = DBobject
                                        If CsfCh1.TextString.Contains("+") = True Then
                                            ChainageCSF_colection.Add(CsfCh1)
                                        End If
                                        If CsfCh1.TextString.Contains("CSF") = True Then
                                            CSF_colection.Add(CsfCh1)
                                        End If
                                    End If
                                Next
                                If CSF_colection.Count > 0 And ChainageCSF_colection.Count > 0 Then
                                    For i = start1 To end1

                                        Dim X As Double
                                        Dim Xstring As String = W1.Range(Col_east & i).Value

                                        Dim Y As Double
                                        Dim Ystring As String = W1.Range(Col_north & i).Value

                                        If IsNumeric(Xstring) = True And IsNumeric(Ystring) = True Then
                                            X = CDbl(Xstring)
                                            Y = CDbl(Ystring)
                                            Dim Point_on_poly2d As New Point3d
                                            Dim Point_on_poly3d As New Point3d
                                            Point_on_poly2d = Poly2D.GetClosestPointTo(New Point3d(X, Y, 0), Vector3d.ZAxis, False)
                                            Point_on_poly3d = Poly3d.GetPointAtParameter(Poly2D.GetParameterAtPoint(Point_on_poly2d))
                                            Dim ChainageCSF As Double = Get_chainage_with_CSF_from_dbtext(Poly3d, Point_on_poly3d, CSF_colection, ChainageCSF_colection)

                                            W1.Range(Col_CSF_CHAIN & i).Value = Round(ChainageCSF, 3)
                                            Dim new_leader As New MLeader
                                            new_leader = Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly3d,
                                                                                               "East = " & Round(Point_on_poly3d.X, 3) & vbCrLf &
                                                                                               "North = " & Round(Point_on_poly3d.Y, 3) & vbCrLf &
                                                                                               "Elev = " & Round(Point_on_poly3d.Z, 3) & vbCrLf &
                                                                                               "CSF Chainage = " & Get_chainage_from_double(ChainageCSF, 3) _
                                                                                               , 1, 0.2, 0.5, 5, 5)
                                        End If

                                    Next

                                End If


                            End If

                            Editor1.Regen()
                            Trans1.Commit()
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

    Private Sub Button_2dgrid_Click(sender As Object, e As EventArgs) Handles Button_2dgrid.Click

        Try
            If TextBox_east_intersection.Text = "" Then
                MsgBox("Please specify the East COLUMN!")
                Exit Sub
            End If
            If TextBox_north_intersection.Text = "" Then
                MsgBox("Please specify the North COLUMN!")
                Exit Sub
            End If
            If TextBox_elevation.Text = "" Then
                MsgBox("Please specify the Elevation COLUMN!")
                Exit Sub
            End If

            If TextBox_CSF_Chainage.Text = "" Then
                MsgBox("Please specify the CSF Chainage COLUMN!")
                Exit Sub
            End If

            If IsNumeric(TextBox_start.Text) = False Then
                MsgBox("Please specify the Start Row")
                Exit Sub
            End If
            If IsNumeric(TextBox_end.Text) = False Then
                MsgBox("Please specify the End Row")
                Exit Sub
            End If
            If CInt(TextBox_end.Text) < CInt(TextBox_start.Text) Then
                MsgBox("Start row smaller than end row")
                Exit Sub
            End If
            If CInt(TextBox_end.Text) < 1 Then
                MsgBox("End row smaller than 1")
                Exit Sub
            End If
            If CInt(TextBox_start.Text) < 1 Then
                MsgBox("Start row smaller than 1")
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
                Object_Prompt.MessageForAdding = vbLf & "Select 3D centerline:"
                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
                Dim Col_east As String = TextBox_east_intersection.Text.ToUpper
                Dim Col_north As String = TextBox_north_intersection.Text.ToUpper
                Dim Col_Elevation As String = TextBox_elevation.Text.ToUpper
                Dim Col_CSF_CHAIN As String = TextBox_CSF_Chainage.Text.ToUpper
                Dim Col_Grid_2D_chainage As String = TextBox_chainage.Text.ToUpper
                Dim start1 As Integer = CInt(TextBox_start.Text)
                Dim end1 As Integer = CInt(TextBox_end.Text)
                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    If IsNothing(Rezultat1) = False Then
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim Poly3d As Polyline3d
                            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                            BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(0)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                            If TypeOf Ent1 Is Polyline3d Then
                                Poly3d = Ent1
                                Dim ChainageCSF_colection As New DBObjectCollection
                                Dim CSF_colection As New DBObjectCollection
                                Dim Poly2D As New Polyline
                                Dim Index2d As Double = 0
                                For Each ObjId As ObjectId In Poly3d
                                    Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, OpenMode.ForRead)
                                    Poly2D.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)
                                    Index2d = Index2d + 1
                                Next
                                For Each ObjID In BTrecord
                                    Dim DBobject As DBObject = Trans1.GetObject(ObjID, OpenMode.ForRead)
                                    If TypeOf DBobject Is DBText Then
                                        Dim CsfCh1 As DBText = DBobject
                                        If CsfCh1.TextString.Contains("+") = True Then
                                            ChainageCSF_colection.Add(CsfCh1)
                                        End If
                                        If CsfCh1.TextString.Contains("CSF") = True Then
                                            CSF_colection.Add(CsfCh1)
                                        End If
                                    End If
                                Next
                                If CSF_colection.Count > 0 And ChainageCSF_colection.Count > 0 Then
                                    For i = start1 To end1
                                        Dim Chainage2Dgrid_txt As String
                                        Chainage2Dgrid_txt = W1.Range(Col_Grid_2D_chainage & i).Value
                                        If IsNumeric(Replace(Chainage2Dgrid_txt, "+", "")) = True Then
                                            Dim Chainage2D_grid As Double
                                            Chainage2D_grid = CDbl(Replace(Chainage2Dgrid_txt, "+", ""))
                                            Dim Param_grid As Double = Poly2D.GetParameterAtDistance(Chainage2D_grid)
                                            Dim Pt_at_Chainage_CSF As New Point3d
                                            Pt_at_Chainage_CSF = Poly3d.GetPointAtParameter(Param_grid)
                                            Dim ChainageCSF As Double = Get_chainage_with_CSF_from_dbtext(Poly3d, Pt_at_Chainage_CSF, CSF_colection, ChainageCSF_colection)


                                            W1.Range(Col_east & i).Value = Round(Pt_at_Chainage_CSF.X, 3)
                                            W1.Range(Col_north & i).Value = Round(Pt_at_Chainage_CSF.Y, 3)
                                            W1.Range(Col_Elevation & i).Value = Round(Pt_at_Chainage_CSF.Z, 3)
                                            W1.Range(Col_CSF_CHAIN & i).Value = Round(ChainageCSF, 3)

                                            Dim new_leader As New MLeader
                                            new_leader = Creaza_Mleader_nou_fara_UCS_transform(Pt_at_Chainage_CSF,
                                                                                               "East = " & Round(Pt_at_Chainage_CSF.X, 3) & vbCrLf &
                                                                                               "North = " & Round(Pt_at_Chainage_CSF.Y, 3) & vbCrLf &
                                                                                               "Elev = " & Round(Pt_at_Chainage_CSF.Z, 3) & vbCrLf &
                                                                                               "CSF Chainage = " & Get_chainage_from_double(ChainageCSF, 3) _
                                                                                               , 1, 0.2, 0.5, 2, 7)
                                        End If
                                    Next
                                    Editor1.Regen()
                                    Trans1.Commit()
                                End If
                            End If
                        End Using
                    End If
                End If
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                MsgBox("Done")
                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Sub Button_3d_2d_Click(sender As Object, e As EventArgs) Handles Button_3d_2d.Click

        Try
            If TextBox_east_intersection.Text = "" Then
                MsgBox("Please specify the East COLUMN!")
                Exit Sub
            End If
            If TextBox_north_intersection.Text = "" Then
                MsgBox("Please specify the North COLUMN!")
                Exit Sub
            End If

            If TextBox_CSF_Chainage.Text = "" Then
                MsgBox("Please specify the station COLUMN!")
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
                Object_Prompt.MessageForAdding = vbLf & "Select 3D centerline:"
                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
                Dim Col_east As String = TextBox_east_intersection.Text.ToUpper
                Dim Col_north As String = TextBox_north_intersection.Text.ToUpper
                Dim Col_CSF_CHAIN As String = TextBox_CSF_Chainage.Text.ToUpper
                Dim start1 As Integer = CInt(TextBox_start.Text)
                Dim end1 As Integer = CInt(TextBox_end.Text)
                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    If IsNothing(Rezultat1) = False Then
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim Data_table_3D_POLY As New System.Data.DataTable
                            Data_table_3D_POLY.Columns.Add("X", GetType(Double))
                            Data_table_3D_POLY.Columns.Add("Y", GetType(Double))
                            Data_table_3D_POLY.Columns.Add("Z", GetType(Double))
                            Dim indexdt As Double = 0
                            Dim Poly3d As Polyline3d
                            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                            BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(0)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                            If TypeOf Ent1 Is Polyline3d Then
                                Poly3d = Ent1
                                Dim Poly2D As New Polyline
                                Dim Index2d As Double = 0
                                For Each ObjId As ObjectId In Poly3d
                                    Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, OpenMode.ForRead)
                                    Poly2D.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)
                                    Index2d = Index2d + 1
                                Next
                                Poly2D.Elevation = 0


                                For i = start1 To end1

                                    Dim X As Double
                                    Dim Xstring As String = W1.Range(Col_east & i).Value

                                    Dim Y As Double
                                    Dim Ystring As String = W1.Range(Col_north & i).Value

                                    If IsNumeric(Xstring) = True And IsNumeric(Ystring) = True Then
                                        X = CDbl(Xstring)
                                        Y = CDbl(Ystring)
                                        Dim Point_on_poly2d As New Point3d
                                        Dim Point_on_poly3d As New Point3d
                                        Point_on_poly2d = Poly2D.GetClosestPointTo(New Point3d(X, Y, 0), Vector3d.ZAxis, False)
                                        Point_on_poly3d = Poly3d.GetPointAtParameter(Poly2D.GetParameterAtPoint(Point_on_poly2d))
                                        Dim ChainageCSF As Double = Poly3d.GetDistAtPoint(Point_on_poly3d)

                                        W1.Range(Col_CSF_CHAIN & i).Value = Round(ChainageCSF, 3)
                                        If Not TextBox_elevation.Text = "" Then
                                            W1.Range(TextBox_elevation.Text & i).Value = Round(Point_on_poly3d.Z, 3)
                                        End If
                                    End If

                                Next


                            Else
                                MsgBox("No 3d Polyline selected")


                            End If

                            Editor1.Regen()
                            Trans1.Commit()
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

    Private Sub Button_reroute_projection_Click(sender As Object, e As EventArgs) Handles Button_reroute_projection.Click
        Dim jj As Integer
        Dim ii As Integer
        Try


            If IsNumeric(TextBox_start.Text) = False Then
                MsgBox("Please specify the Start Row")
                Exit Sub
            End If
            If IsNumeric(TextBox_end.Text) = False Then
                MsgBox("Please specify the End Row")
                Exit Sub
            End If
            If CInt(TextBox_end.Text) < CInt(TextBox_start.Text) Then
                MsgBox("Start row smaller than end row")
                Exit Sub
            End If
            If CInt(TextBox_end.Text) < 1 Then
                MsgBox("End row smaller than 1")
                Exit Sub
            End If
            If CInt(TextBox_start.Text) < 1 Then
                MsgBox("Start row smaller than 1")
                Exit Sub
            End If


            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
            Editor1 = ThisDrawing.Editor
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

            Using lock As DocumentLock = ThisDrawing.LockDocument
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()


                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
                Dim Col_x As String = TextBox_X.Text.ToUpper
                Dim Col_y As String = TextBox_Y.Text.ToUpper
                Dim Col_z As String = TextBox_Z.Text.ToUpper
                Dim Col_csf_sta As String = TextBox_CSF_STA.Text.ToUpper



                Dim start1 As Integer = CInt(TextBox_start.Text)
                Dim end1 As Integer = CInt(TextBox_end.Text)


                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select the 2D lightweight polyline:"
                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)

                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    If IsNothing(Rezultat1) = False Then
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim dt_from_excel As New System.Data.DataTable

                            dt_from_excel.Columns.Add("X", GetType(Double))
                            dt_from_excel.Columns.Add("Y", GetType(Double))
                            dt_from_excel.Columns.Add("Z", GetType(Double))
                            dt_from_excel.Columns.Add("CSF", GetType(Double))
                            dt_from_excel.Columns.Add("STA", GetType(Double))


                            Dim dt_poly2d As New System.Data.DataTable

                            dt_poly2d.Columns.Add("X", GetType(Double))
                            dt_poly2d.Columns.Add("Y", GetType(Double))
                            dt_poly2d.Columns.Add("STA", GetType(Double))


                            Dim Poly1 As Polyline

                            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)

                            Dim Ent1 As Entity = Trans1.GetObject(Rezultat1.Value.Item(0).ObjectId, OpenMode.ForRead)

                            If TypeOf Ent1 Is Polyline Then
                                Poly1 = Ent1

                                For i = 0 To Poly1.NumberOfVertices - 1
                                    Dim pt1 As Point3d = Poly1.GetPointAtParameter(i)
                                    dt_poly2d.Rows.Add()
                                    dt_poly2d.Rows(i).Item("X") = pt1.X
                                    dt_poly2d.Rows(i).Item("Y") = pt1.Y
                                    dt_poly2d.Rows(i).Item("STA") = Poly1.GetDistanceAtParameter(i)

                                Next


                                For i = start1 To end1

                                    Dim csf_string As String = W1.Range(Col_csf_sta & i).Value2
                                    If IsNumeric(csf_string.Replace("+", "")) = False Then
                                        MsgBox("non numeric value on cell " & Col_csf_sta & i)
                                        Exit Sub
                                    End If

                                    Dim x_string As String = W1.Range(Col_x & i).Value2
                                    If IsNumeric(x_string) = False Then
                                        MsgBox("non numeric value on cell " & Col_x & i)
                                        Exit Sub
                                    End If


                                    Dim y_string As String = W1.Range(Col_y & i).Value2
                                    If IsNumeric(y_string) = False Then
                                        MsgBox("non numeric value on cell " & Col_y & i)
                                        Exit Sub
                                    End If

                                    Dim z_string As String = W1.Range(Col_z & i).Value2
                                    If IsNumeric(z_string) = False Then
                                        MsgBox("non numeric value on cell " & Col_z & i)
                                        Exit Sub
                                    End If


                                    Dim csf As Double = csf_string.Replace("+", "")
                                    Dim x As Double = Convert.ToDouble(x_string)
                                    Dim y As Double = Convert.ToDouble(y_string)
                                    Dim z As Double = Convert.ToDouble(z_string)


                                    Dim pt2 As New Point3d(x, y, Poly1.Elevation)

                                    Dim point_on_poly As New Point3d()
                                    point_on_poly = Poly1.GetClosestPointTo(pt2, Vector3d.ZAxis, False)

                                    dt_from_excel.Rows.Add()
                                    dt_from_excel.Rows(dt_from_excel.Rows.Count - 1).Item("X") = point_on_poly.X
                                    dt_from_excel.Rows(dt_from_excel.Rows.Count - 1).Item("Y") = point_on_poly.Y
                                    dt_from_excel.Rows(dt_from_excel.Rows.Count - 1).Item("Z") = z
                                    dt_from_excel.Rows(dt_from_excel.Rows.Count - 1).Item("CSF") = csf
                                    Dim sta As Double = Poly1.GetDistAtPoint(point_on_poly)

                                    dt_from_excel.Rows(dt_from_excel.Rows.Count - 1).Item("STA") = sta



                                Next


                            End If

                            If dt_from_excel.Rows.Count > 0 Then
                                For i = 0 To dt_poly2d.Rows.Count - 1

                                    ii = i
                                    Dim x1 As Double = dt_poly2d.Rows(i).Item("X")
                                    Dim y1 As Double = dt_poly2d.Rows(i).Item("Y")
                                    Dim sta1 As Double = dt_poly2d.Rows(i).Item("STA")


                                    Dim pt2 As New Point3d(x1, y1, Poly1.Elevation)
                                    Dim point_on_poly As New Point3d()
                                    point_on_poly = Poly1.GetClosestPointTo(pt2, Vector3d.ZAxis, False)

                                    For j = 1 To dt_from_excel.Rows.Count - 1
                                        jj = j

                                        Dim sta0 As Double = 0
                                        sta0 = dt_from_excel.Rows(j - 1).Item("STA")



                                        Dim sta2 As Double = 0
                                        sta2 = dt_from_excel.Rows(j).Item("STA")



                                        If sta1 > sta0 And sta1 < sta2 Then
                                            Dim x0 As Double = dt_from_excel.Rows(j - 1).Item("X")
                                            Dim x2 As Double = dt_from_excel.Rows(j).Item("X")
                                            Dim y0 As Double = dt_from_excel.Rows(j - 1).Item("Y")
                                            Dim y2 As Double = dt_from_excel.Rows(j).Item("Y")
                                            Dim z0 As Double = dt_from_excel.Rows(j - 1).Item("Z")
                                            Dim z2 As Double = dt_from_excel.Rows(j).Item("Z")
                                            Dim csf0 As Double = dt_from_excel.Rows(j - 1).Item("CSF")
                                            Dim csf2 As Double = dt_from_excel.Rows(j).Item("CSF")



                                            Dim pt3 As New Point3d(x1, y1, Poly1.Elevation)

                                            Dim pt_on_poly = New Point3d()
                                            pt_on_poly = Poly1.GetClosestPointTo(pt3, Vector3d.YAxis, False)
                                            Dim param1 As Double = Poly1.GetParameterAtPoint(pt_on_poly)


                                            Dim row1 As System.Data.DataRow
                                            row1 = dt_from_excel.NewRow()
                                            row1("X") = x1
                                            row1("Y") = y1

                                            row1("STA") = sta1


                                            Dim difsta As Double = sta2 - sta0
                                            Dim difelev As Double = z2 - z0
                                            Dim coeficient1 As Double = difelev / difsta

                                            row1("Z") = z0 + coeficient1 * (sta1 - sta0)
                                            Dim difcsf As Double = csf2 - csf0

                                            Dim coeficient2 As Double = difcsf / difsta
                                            row1("CSF") = csf0 + coeficient2 * (sta1 - sta0)





                                            dt_from_excel.Rows.InsertAt(row1, j)

                                            Exit For
                                        End If
                                    Next
                                Next
                                Transfer_datatable_to_new_excel_spreadsheet(dt_from_excel)
                                Transfer_datatable_to_new_excel_spreadsheet(dt_poly2d)
                            End If


                        End Using
                    End If
                End If




                MsgBox("Done")
                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            End Using

        Catch ex As System.Exception

            MsgBox(ii & vbCrLf & jj & vbCrLf & ex.Message)
        End Try
    End Sub

    Private Sub Button_match_csf_Click(sender As Object, e As EventArgs) Handles Button_read_write_csf.Click
        Dim jj As Integer
        Dim ii As Integer
        Try


            If IsNumeric(TextBox_start.Text) = False Then
                MsgBox("Please specify the Start Row")
                Exit Sub
            End If
            If IsNumeric(TextBox_end.Text) = False Then
                MsgBox("Please specify the End Row")
                Exit Sub
            End If
            If CInt(TextBox_end.Text) < CInt(TextBox_start.Text) Then
                MsgBox("Start row smaller than end row")
                Exit Sub
            End If
            If CInt(TextBox_end.Text) < 1 Then
                MsgBox("End row smaller than 1")
                Exit Sub
            End If
            If CInt(TextBox_start.Text) < 1 Then
                MsgBox("Start row smaller than 1")
                Exit Sub
            End If

            If IsNumeric(TextBox_s1.Text) = False Then
                MsgBox("Please specify the Start Row")
                Exit Sub
            End If
            If IsNumeric(TextBox_e1.Text) = False Then
                MsgBox("Please specify the End Row")
                Exit Sub
            End If
            If CInt(TextBox_e1.Text) < CInt(TextBox_s1.Text) Then
                MsgBox("Start row smaller than end row")
                Exit Sub
            End If
            If CInt(TextBox_e1.Text) < 1 Then
                MsgBox("End row smaller than 1")
                Exit Sub
            End If
            If CInt(TextBox_s1.Text) < 1 Then
                MsgBox("Start row smaller than 1")
                Exit Sub
            End If


            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
            Editor1 = ThisDrawing.Editor
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

            Using lock As DocumentLock = ThisDrawing.LockDocument
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                Dim Workbook1 As Microsoft.Office.Interop.Excel.Workbook = Get_active_workbook_from_Excel()
                Dim W3 As Microsoft.Office.Interop.Excel.Worksheet = Workbook1.Sheets(CInt(TextBoxs1.Text))

                Dim W4 As Microsoft.Office.Interop.Excel.Worksheet = Workbook1.Sheets(CInt(TextBoxs.Text))
                Dim Col_x4 As String = TextBox_X.Text.ToUpper
                Dim Col_y4 As String = TextBox_Y.Text.ToUpper
                Dim Col_z4 As String = TextBox_Z.Text.ToUpper
                Dim Col_csf_sta4 As String = TextBox_CSF_STA.Text.ToUpper

                Dim start4 As Integer = CInt(TextBox_s1.Text)
                Dim end4 As Integer = CInt(TextBox_e1.Text)


                Dim start3 As Integer = CInt(TextBox_start.Text)
                Dim end3 As Integer = CInt(TextBox_end.Text)
                Dim Col_x3 As String = TextBoxx1.Text.ToUpper
                Dim Col_y3 As String = TextBoxy1.Text.ToUpper
                Dim Col_z3 As String = TextBoxz1.Text.ToUpper
                Dim Col_csf_sta31 As String = TextBoxcsf1.Text.ToUpper
                Dim Col_csf_sta32 As String = TextBoxCSF2.Text.ToUpper



                Dim start0 As Integer = start3



                For i = start4 To end4




                    Dim x_string As String = W4.Range(Col_x4 & i).Value2
                    If IsNumeric(x_string) = False Then
                        MsgBox("non numeric value on cell " & Col_x4 & i)
                        Exit Sub
                    End If


                    Dim y_string As String = W4.Range(Col_y4 & i).Value2
                    If IsNumeric(y_string) = False Then
                        MsgBox("non numeric value on cell " & Col_y4 & i)
                        Exit Sub
                    End If

                    Dim z_string As String = W4.Range(Col_z4 & i).Value2
                    If IsNumeric(z_string) = False Then
                        MsgBox("non numeric value on cell " & Col_z4 & i)
                        Exit Sub
                    End If



                    Dim x4 As Double = Convert.ToDouble(x_string)
                    Dim y4 As Double = Convert.ToDouble(y_string)
                    Dim z4 As Double = Convert.ToDouble(z_string)

                    For j = start0 To end3




                        Dim x_string3 As String = W3.Range(Col_x3 & j).Value2
                        If IsNumeric(x_string3) = False Then
                            MsgBox("non numeric value on cell " & Col_x3 & j)
                            Exit Sub
                        End If


                        Dim y_string3 As String = W3.Range(Col_y3 & j).Value2
                        If IsNumeric(y_string3) = False Then
                            MsgBox("non numeric value on cell " & Col_y3 & j)
                            Exit Sub
                        End If

                        Dim z_string3 As String = W3.Range(Col_z3 & j).Value2
                        If IsNumeric(z_string3) = False Then
                            MsgBox("non numeric value on cell " & Col_z3 & j)
                            Exit Sub
                        End If



                        Dim x3 As Double = Convert.ToDouble(x_string3)
                        Dim y3 As Double = Convert.ToDouble(y_string3)
                        Dim z3 As Double = Convert.ToDouble(z_string3)

                        If Math.Round(x3, 3) = Math.Round(x4, 3) Then
                            If Math.Round(y3, 3) = Math.Round(y4, 3) Then
                                If Math.Round(z3, 3) = Math.Round(z4, 3) Then
                                    Dim csf31 As String = W3.Range(Col_csf_sta32 & j).Value2
                                    Dim csf32 As String = W3.Range(Col_csf_sta31 & j).Value2

                                    If Not (csf31 = "") Then

                                        W4.Range(Col_csf_sta4 & i).Value2 = csf31
                                    Else
                                        W4.Range(Col_csf_sta4 & i).Value2 = csf32
                                    End If
                                    start0 = j + 1
                                    j = end3

                                End If
                            End If
                        End If


                    Next




                Next





                MsgBox("Done")
                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            End Using

        Catch ex As System.Exception

            MsgBox(ii & vbCrLf & jj & vbCrLf & ex.Message)
        End Try
    End Sub

    Private Sub Button_gen_station_Click(sender As Object, e As EventArgs) Handles Button_gen_station.Click
        Try
            If TextBox_east_intersection.Text = "" Then
                MsgBox("Please specify the East COLUMN!")
                Exit Sub
            End If
            If TextBox_north_intersection.Text = "" Then
                MsgBox("Please specify the North COLUMN!")
                Exit Sub
            End If

            If TextBox_CSF_Chainage.Text = "" Then
                MsgBox("Please specify the station COLUMN!")
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
                Object_Prompt.MessageForAdding = vbLf & "Select  centerline:"
                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
                Dim Col_east As String = TextBox_east_intersection.Text.ToUpper
                Dim Col_north As String = TextBox_north_intersection.Text.ToUpper
                Dim Col_station As String = TextBox_CSF_Chainage.Text.ToUpper
                Dim start1 As Integer = CInt(TextBox_start.Text)
                Dim end1 As Integer = CInt(TextBox_end.Text)
                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    If IsNothing(Rezultat1) = False Then
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                            BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)

                            Dim Ent1 As Entity
                            Ent1 = Trans1.GetObject(Rezultat1.Value.Item(0).ObjectId, OpenMode.ForRead)
                            If TypeOf Ent1 Is Polyline3d Then
                                Dim Poly3d As Polyline3d = Ent1
                                Dim Poly2D As New Polyline
                                Dim Index2d As Double = 0
                                For Each ObjId As ObjectId In Poly3d
                                    Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, OpenMode.ForRead)
                                    Poly2D.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)
                                    Index2d = Index2d + 1
                                Next
                                Poly2D.Elevation = 0

                                For i = start1 To end1
                                    Dim X As Double
                                    Dim Xstring As String = W1.Range(Col_east & i).Value
                                    Dim Y As Double
                                    Dim Ystring As String = W1.Range(Col_north & i).Value
                                    If IsNumeric(Xstring) = True And IsNumeric(Ystring) = True Then
                                        X = CDbl(Xstring)
                                        Y = CDbl(Ystring)
                                        Dim Point_on_poly2d As New Point3d
                                        Dim param1 As Double
                                        Point_on_poly2d = Poly2D.GetClosestPointTo(New Point3d(X, Y, 0), Vector3d.ZAxis, False)
                                        param1 = Poly2D.GetParameterAtPoint(Point_on_poly2d)
                                        Dim station As Double = Poly3d.GetDistanceAtParameter(param1)

                                        W1.Range(Col_station & i).Value = Round(station, 3)
                                    End If
                                Next
                            End If

                            If TypeOf Ent1 Is Polyline Then
                                Dim Poly2d As Polyline = Ent1
                                For i = start1 To end1
                                    Dim X As Double
                                    Dim Xstring As String = W1.Range(Col_east & i).Value
                                    Dim Y As Double
                                    Dim Ystring As String = W1.Range(Col_north & i).Value
                                    If IsNumeric(Xstring) = True And IsNumeric(Ystring) = True Then
                                        X = CDbl(Xstring)
                                        Y = CDbl(Ystring)
                                        Dim Point_on_poly2d As New Point3d
                                        Dim param1 As Double
                                        Point_on_poly2d = Poly2d.GetClosestPointTo(New Point3d(X, Y, 0), Vector3d.ZAxis, False)
                                        param1 = Poly2d.GetParameterAtPoint(Point_on_poly2d)
                                        Dim station As Double = Poly2d.GetDistanceAtParameter(param1)
                                        W1.Range(Col_station & i).Value = Round(station, 3)
                                    End If
                                Next
                            End If

                            Editor1.Regen()
                            Trans1.Commit()
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