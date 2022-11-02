Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class New_chainage_calc_from_excel
    Dim Colectie1 As New Specialized.StringCollection
    Private Sub Button_recalculate_chainage_on_a_new_route_Click(sender As Object, e As EventArgs) Handles Button_recalculate_chainage_on_a_new_route.Click
        Try
            If TextBox_old_east.Text = "" Then
                MsgBox("Please specify the Old East COLUMN!")
                Exit Sub
            End If
            If TextBox_new_east.Text = "" Then
                MsgBox("Please specify the New East COLUMN!")
                Exit Sub
            End If
            If TextBox_old_north.Text = "" Then
                MsgBox("Please specify the Old North COLUMN!")
                Exit Sub
            End If
            If TextBox_new_north.Text = "" Then
                MsgBox("Please specify the New North COLUMN!")
                Exit Sub
            End If

            If TextBox_old_elevation.Text = "" Then
                MsgBox("Please specify the Old Elevation COLUMN!")
                Exit Sub
            End If
            If TextBox_new_elevation.Text = "" Then
                MsgBox("Please specify the New Elevation COLUMN!")
                Exit Sub
            End If


            If TextBox_old_ch.Text = "" Then
                MsgBox("Please specify the Old Chainage COLUMN!")
                Exit Sub
            End If
            If TextBox_new_ch.Text = "" Then
                MsgBox("Please specify the New Chainage COLUMN!")
                Exit Sub
            End If

            If TextBox_row_Start1.Text = "" Then
                MsgBox("Please specify the Start ROW!")
                Exit Sub
            End If
            If TextBox_row_end1.Text = "" Then
                MsgBox("Please specify the end ROW!")
                Exit Sub
            End If

            If IsNumeric(TextBox_row_Start1.Text) = False Then
                With TextBox_row_Start1
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify start row")
                Exit Sub
            End If

            If IsNumeric(TextBox_row_end1.Text) = False Then
                With TextBox_row_end1
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify end row")
                Exit Sub
            End If

            If Val(TextBox_row_Start1.Text) < 1 Then
                With TextBox_row_Start1
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Start row can't be smaller than 1")

                Exit Sub
            End If

            If Val(TextBox_row_end1.Text) < 1 Then
                With TextBox_row_end1
                    .Text = ""
                    .Focus()
                End With
                MsgBox("End row can't be smaller than 1")

                Exit Sub
            End If
            Dim start1 As Integer = CInt(TextBox_row_Start1.Text)
            Dim end1 As Integer = CInt(TextBox_row_end1.Text)
            If end1 < start1 Then
                MsgBox("End row can't be smaller than start row")

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
                Object_Prompt.MessageForAdding = vbLf & "Select 3D old polyline:"

                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)

                Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt2.MessageForAdding = vbLf & "Select 3D  new polyline:"

                Object_Prompt2.SingleOnly = True
                Rezultat2 = Editor1.GetSelection(Object_Prompt2)


                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK And Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    If IsNothing(Rezultat1) = False And IsNothing(Rezultat2) = False Then
                        Dim Data_table1 As New System.Data.DataTable
                        Data_table1.Columns.Add("TEXT325", GetType(DBText))
                        Dim Index1 As Double = 0

                        Dim Data_table2 As New System.Data.DataTable
                        Data_table2.Columns.Add("TEXT0", GetType(DBText))
                        Dim Index2 As Double = 0

                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                            Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj2 = Rezultat2.Value.Item(0)
                            Dim Ent2 As Entity
                            Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)
                            If TypeOf Ent2 Is Polyline3d Then
                                Dim Poly3d_NEW As Polyline3d = Ent2



                                For Each ObjID In BTrecord
                                    Dim DBobject As DBObject = Trans1.GetObject(ObjID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                    If TypeOf DBobject Is DBText Then
                                        Dim Text1 As DBText = DBobject
                                        If Text1.Layer = Poly3d_NEW.Layer Then
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
                            End If
                        End Using



                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                            BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)



                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(0)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                            Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj2 = Rezultat2.Value.Item(0)
                            Dim Ent2 As Entity
                            Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)

                            If TypeOf Ent1 Is Polyline3d And TypeOf Ent2 Is Polyline3d Then
                                Dim Poly3d_OLD As Polyline3d = Ent1
                                Dim Poly3d_NEW As Polyline3d = Ent2

                                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
                                Dim Col_OLD_east As String = TextBox_old_east.Text.ToUpper
                                Dim Col_OLD_north As String = TextBox_old_north.Text.ToUpper
                                Dim Col_OLD_elevation As String = TextBox_old_elevation.Text.ToUpper
                                Dim Col_OLD_Chain As String = TextBox_old_ch.Text.ToUpper

                                Dim Col_new_east As String = TextBox_new_east.Text.ToUpper
                                Dim Col_new_north As String = TextBox_new_north.Text.ToUpper
                                Dim Col_new_elevation As String = TextBox_new_elevation.Text.ToUpper
                                Dim Col_new_Chain As String = TextBox_new_ch.Text.ToUpper

                                For i = start1 To end1
                                    Dim Old_chainage As Double
                                    Dim Excel_old_chainage As String = W1.Range(Col_OLD_Chain & i).Value
                                    If IsNumeric(Replace(Excel_old_chainage, "+", "")) = True Then
                                        Old_chainage = CDbl(Replace(Excel_old_chainage, "+", ""))
                                        Dim Old_point As New Point3d
                                        Old_point = Poly3d_OLD.GetPointAtDist(Old_chainage)
                                        W1.Range(Col_OLD_east & i).Value = Round(Old_point.X, 2)
                                        W1.Range(Col_OLD_north & i).Value = Round(Old_point.Y, 2)
                                        W1.Range(Col_OLD_elevation & i).Value = Round(Old_point.Z, 2)

                                        Dim New_point As New Point3d
                                        New_point = Poly3d_NEW.GetClosestPointTo(Old_point, Vector3d.ZAxis, False)
                                        W1.Range(Col_new_east & i).Value = Round(New_point.X, 2)
                                        W1.Range(Col_new_north & i).Value = Round(New_point.Y, 2)
                                        W1.Range(Col_new_elevation & i).Value = Round(New_point.Z, 2)

                                        Dim Parameter_picked As Double = Round(Poly3d_NEW.GetParameterAtPoint(New_point), 3)

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
                                            Point_CHAINAGE = Poly3d_NEW.GetPointAtParameter(Parameter_start)
                                            Distanta_pana_la_Vertex = Point_CHAINAGE.GetVectorTo(New_point).Length

                                            For J = 0 To Data_table1.Rows.Count - 1
                                                Dim Text1 As DBText = Data_table1.Rows(J).Item("TEXT325")
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
                                                Point_CHAINAGE1 = Poly3d_NEW.GetPointAtParameter(Parameter_start)
                                                Dim Point_CHAINAGE2 As New Point3d
                                                Point_CHAINAGE2 = Poly3d_NEW.GetPointAtParameter(Parameter_end)

                                                For J = 0 To Data_table2.Rows.Count - 1
                                                    Dim Text1 As DBText = Data_table2.Rows(J).Item("TEXT0")
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
                                        If Not CSF1 + CSF2 = 0 And Not CSF1 = 0 And Not CSF2 = 0 Then
                                            New_ch = Chainage_on_vertex + Distanta_pana_la_Vertex / ((CSF1 + CSF2) / 2)
                                        Else
                                            New_ch = Chainage_on_vertex + Distanta_pana_la_Vertex
                                        End If

                                        W1.Range(Col_new_Chain & i).Value = Round(New_ch, 2)

                                        Dim Old_leader As New MLeader
                                        Old_leader = Creaza_Mleader_nou_fara_UCS_transform(Old_point, "East = " & Round(Old_point.X, 2) & vbCrLf & _
                                                                                           "North = " & Round(Old_point.Y, 2) & vbCrLf & _
                                                                                           "Elev = " & Round(Old_point.Z, 2) & vbCrLf & _
                                                                                           "Chainage = " & Old_chainage _
                                                                                           , 1, 0.2, 0.5, 2, 7)
                                        Old_leader.Layer = Poly3d_OLD.Layer
                                        Dim new_leader As New MLeader
                                        new_leader = Creaza_Mleader_nou_fara_UCS_transform(New_point, "East = " & Round(New_point.X, 2) & vbCrLf & _
                                                                                           "North = " & Round(New_point.Y, 2) & vbCrLf & _
                                                                                           "Elev = " & Round(New_point.Z, 2) & vbCrLf & _
                                                                                           "Chainage = " & Round(New_ch, 2) _
                                                                                           , 1, 0.2, 0.5, 2, 14)
                                        new_leader.Layer = Poly3d_NEW.Layer

                                    End If
                                Next
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

    Private Sub Button_recalc_chainage_in_excel_based_on_new_poly_Click(sender As Object, e As EventArgs) Handles Button_recalc_chainage_in_excel_based_on_new_poly.Click
        Try
            If TextBox_east.Text = "" Then
                MsgBox("Please specify the East COLUMN!")
                Exit Sub
            End If
            If TextBox_north.Text = "" Then
                MsgBox("Please specify the North COLUMN!")
                Exit Sub
            End If
            If TextBox_row_start.Text = "" Then
                MsgBox("Please specify the Start ROW!")
                Exit Sub
            End If
            If TextBox_row_end.Text = "" Then
                MsgBox("Please specify the end ROW!")
                Exit Sub
            End If
            If IsNumeric(TextBox_row_start.Text) = False Then
                With TextBox_row_start
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify start row")
                Exit Sub
            End If
            If IsNumeric(TextBox_row_end.Text) = False Then
                With TextBox_row_end
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify end row")
                Exit Sub
            End If
            If Val(TextBox_row_start.Text) < 1 Then
                With TextBox_row_start
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Start row can't be smaller than 1")

                Exit Sub
            End If

            If Val(TextBox_row_end.Text) < 1 Then
                With TextBox_row_end
                    .Text = ""
                    .Focus()
                End With
                MsgBox("End row can't be smaller than 1")

                Exit Sub
            End If
            Dim start1 As Integer = CInt(TextBox_row_start.Text)
            Dim end1 As Integer = CInt(TextBox_row_end.Text)
            If end1 < start1 Then
                MsgBox("End row can't be smaller than start row")
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
                Object_Prompt.MessageForAdding = vbLf & "Select 3d polyline:"
                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)

                Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt2.MessageForAdding = vbLf & "Select 2d polyline:"
                Object_Prompt2.SingleOnly = True
                Rezultat2 = Editor1.GetSelection(Object_Prompt2)

                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK And Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    If IsNothing(Rezultat1) = False And IsNothing(Rezultat2) = False Then
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                            BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)
                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(0)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                            Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj2 = Rezultat2.Value.Item(0)
                            Dim Ent2 As Entity
                            Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)
                            If TypeOf Ent1 Is Polyline3d Then
                                Dim Poly3d As Polyline3d = Ent1
                                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
                                Dim Col_east As String = TextBox_east.Text.ToUpper
                                Dim Col_north As String = TextBox_north.Text.ToUpper
                                Dim Col_Chain As String = TextBox_recalc_chainage.Text.ToUpper
                                For i = start1 To end1
                                    Dim East, North As Double
                                    Dim CellEast As String
                                    CellEast = W1.Range(Col_east & i).Value
                                    Dim CellNorth As String
                                    CellNorth = W1.Range(Col_north & i).Value
                                    If IsNumeric(CellEast) = True And IsNumeric(CellNorth) = True Then
                                        East = CDbl(CellEast)
                                        North = CDbl(CellNorth)
                                        If Not East = 0 Or Not North = 0 Then
                                            Dim Point1 As New Point3d(East, North, 0)
                                            Dim Point_on_poly As New Point3d
                                            Point_on_poly = Poly3d.GetClosestPointTo(Point1, Vector3d.ZAxis, False)
                                            Dim new_chainage As Double = Poly3d.GetDistAtPoint(Point_on_poly)
                                            If TypeOf Ent2 Is Polyline Then
                                                Dim Poly2 As Polyline = Ent2
                                                Dim Point_langa_poly As New Point3d(Point1.X, Point1.Y, 0)
                                                Dim Point_pe_poly2d As New Point3d
                                                Point_pe_poly2d = Poly2.GetClosestPointTo(Point_langa_poly, Vector3d.ZAxis, False)
                                                Dim Distanta As Double = Point_langa_poly.GetVectorTo(New Point3d(Point_pe_poly2d.X, Point_pe_poly2d.Y, 0)).Length
                                                W1.Range(Chr(Asc(Col_Chain) + 1) & i).Value = Round(Distanta, 2)
                                            End If
                                            W1.Range(Col_Chain & i).Value = Get_chainage_from_double(new_chainage, 1)
                                            Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly, "East = " & CellEast & vbCrLf & "North = " & CellNorth, 1, 0.2, 0.5, 2, 2)
                                            Creaza_Mleader_nou_fara_UCS_transform(Point1, "East = " & CellEast & vbCrLf & "North = " & CellNorth, 1, 0.2, 0.5, 2, 2)
                                        End If
                                    End If
                                Next
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

    Private Sub Button_chainage_at_intersection_Click(sender As Object, e As EventArgs) Handles Button_chainage_at_intersection.Click
        Try
            If TextBox_east_intersection.Text = "" Then
                MsgBox("Please specify the East COLUMN!")
                Exit Sub
            End If
            If TextBox_north_intersection.Text = "" Then
                MsgBox("Please specify the North COLUMN!")
                Exit Sub
            End If
            If TextBox_elevation_intersection.Text = "" Then
                MsgBox("Please specify the Elevation COLUMN!")
                Exit Sub
            End If
            If TextBox_chainage_intersection.Text = "" Then
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
                Object_Prompt.MessageForAdding = vbLf & "Select 3d polyline:"
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

                            Dim Data_table_for_doubles As New System.Data.DataTable
                            Data_table_for_doubles.Columns.Add("X", GetType(Double))
                            Data_table_for_doubles.Columns.Add("Y", GetType(Double))
                            Data_table_for_doubles.Columns.Add("Z", GetType(Double))
                            Data_table_for_doubles.Columns.Add("CHAINAGE", GetType(Double))

                            Dim indexdtd As Double = 0

                            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = get_new_worksheet_from_Excel()
                            Dim Col_east As String = TextBox_east_intersection.Text.ToUpper
                            Dim Col_north As String = TextBox_north_intersection.Text.ToUpper
                            Dim Col_Chain As String = TextBox_chainage_intersection.Text.ToUpper
                            Dim Col_Elevation As String = TextBox_elevation_intersection.Text.ToUpper
                           



                            Dim Data_table1 As New System.Data.DataTable
                            Data_table1.Columns.Add("TEXT325", GetType(DBText))
                            Dim Index1 As Double = 0

                            Dim Data_table2 As New System.Data.DataTable
                            Data_table2.Columns.Add("TEXT0", GetType(DBText))
                            Dim Index2 As Double = 0

                            Dim Poly3d As Polyline3d

                            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                            BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)
                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(0)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                            If TypeOf Ent1 Is Polyline3d Then
                                Poly3d = Ent1
                                Dim Poly2D As New Polyline

                                Dim Index2d As Double = 0
                                For Each ObjId As ObjectId In Poly3d
                                    Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                    Poly2D.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)

                                    Index2d = Index2d + 1
                                Next
                                Poly2D.Elevation = 0





                                Dim Data_table_Poly As New System.Data.DataTable
                                Data_table_Poly.Columns.Add("2DPOLY", GetType(Polyline))
                                Data_table_Poly.Columns.Add("FEATID", GetType(String))
                                Data_table_Poly.Columns.Add("NAME", GetType(String))
                                Dim Data_table_LINE As New System.Data.DataTable
                                Data_table_LINE.Columns.Add("LINE", GetType(Line))

                                Dim IndexXX As Double = 0
                                Dim IndexXXL As Double = 0
                                Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                                Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables



                                For Each ObjID In BTrecord
                                    Dim DBobject As DBObject = Trans1.GetObject(ObjID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                    If TypeOf DBobject Is Polyline Then
                                        Dim Poly1 As Polyline = DBobject
                                        'If Poly1.Closed = True Then
                                        Data_table_Poly.Rows.Add()
                                        Data_table_Poly.Rows(IndexXX).Item("2DPOLY") = Poly1
                                        IndexXX = IndexXX + 1
                                        'End If
                                    End If

                                    If TypeOf DBobject Is Line Then
                                        Dim Line1 As Line = DBobject
                                        Data_table_LINE.Rows.Add()
                                        Data_table_LINE.Rows(IndexXXL).Item("LINE") = Line1
                                        IndexXXL = IndexXXL + 1
                                        'End If
                                    End If

                                    If TypeOf DBobject Is DBText Then
                                        Dim Text1 As DBText = DBobject
                                        If Text1.Layer = Poly3d.Layer Then
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




                                If Data_table_Poly.Rows.Count > 0 Then
                                    For i = 0 To Data_table_Poly.Rows.Count - 1

                                        Dim Poly1 As Polyline = Data_table_Poly.Rows(i).Item("2DPOLY")
                                        Dim Col_int As New Point3dCollection
                                        Poly1.IntersectWith(Poly2D, Intersect.OnBothOperands, Col_int, IntPtr.Zero, IntPtr.Zero)

                                        If Col_int.Count > 0 Then
                                            For j = 0 To Col_int.Count - 1
                                                Dim point_on_2d As New Point3d
                                                point_on_2d = Col_int(j)
                                                Dim Point_on_3d As New Point3d
                                                Dim Param1 As Double = Poly2D.GetParameterAtPoint(point_on_2d)
                                                Point_on_3d = Poly3d.GetPointAtParameter(Param1)




                                                


                                                Dim Add_value As Boolean = True
                                                Dim Nr_val_data_table As Double = Data_table_for_singles.Rows.Count

                                                If Nr_val_data_table > 0 Then
                                                    For k = 0 To Nr_val_data_table - 1
                                                        If Data_table_for_singles.Rows(k).Item("X") = Round(point_on_2d.X, 2) And Data_table_for_singles.Rows(k).Item("Y") = Round(point_on_2d.Y, 2) Then
                                                            Add_value = False
                                                            Exit For
                                                        End If
                                                    Next
                                                End If


                                                If Add_value = True Then
                                                    Data_table_for_singles.Rows.Add()
                                                    Data_table_for_singles.Rows(indexdt).Item("X") = Round(Point_on_3d.X, 2)
                                                    Data_table_for_singles.Rows(indexdt).Item("Y") = Round(Point_on_3d.Y, 2)
                                                    Data_table_for_singles.Rows(indexdt).Item("Z") = Round(Point_on_3d.Z, 2)
                                                    Data_table_for_singles.Rows(indexdt).Item("CHAINAGE") = Round(Get_chainage_with_CSF(Poly3d, Point_on_3d, Data_table2, Data_table1), 2)
                                                    indexdt = indexdt + 1
                                                End If

                                                Data_table_for_doubles.Rows.Add()
                                                Data_table_for_doubles.Rows(indexdtd).Item("X") = Round(Point_on_3d.X, 2)
                                                Data_table_for_doubles.Rows(indexdtd).Item("Y") = Round(Point_on_3d.Y, 2)
                                                Data_table_for_doubles.Rows(indexdtd).Item("Z") = Round(Point_on_3d.Z, 2)
                                                Data_table_for_doubles.Rows(indexdtd).Item("CHAINAGE") = Round(Get_chainage_with_CSF(Poly3d, Point_on_3d, Data_table2, Data_table1), 2)


                                                Dim Id1 As ObjectId = Poly1.ObjectId
                                                Dim Records1 As Autodesk.Gis.Map.ObjectData.Records
                                                If Not Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False).Count = 0 Then
                                                    Records1 = Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                                    Dim Record1 As Autodesk.Gis.Map.ObjectData.Record
                                                    For Each Record1 In Records1
                                                        Dim Tabla1 As Autodesk.Gis.Map.ObjectData.Table
                                                        Tabla1 = Tables1(Record1.TableName)

                                                        Dim Field_defs1 As Autodesk.Gis.Map.ObjectData.FieldDefinitions
                                                        Field_defs1 = Tabla1.FieldDefinitions
                                                        For ii = 0 To Record1.Count - 1
                                                            Dim Field_def1 As Autodesk.Gis.Map.ObjectData.FieldDefinition
                                                            Field_def1 = Field_defs1(ii)
                                                            Dim Nume_field As String = Field_def1.Name.ToUpper
                                                            Dim Valoare_field As String = Record1(ii).StrValue

                                                            If Data_table_for_doubles.Columns.Contains(Nume_field) = False Then
                                                                Data_table_for_doubles.Columns.Add(Nume_field, GetType(String))
                                                            End If
                                                            Data_table_for_doubles.Rows(indexdtd).Item(Nume_field) = Valoare_field

                                                        Next

                                                    Next

                                                End If

                                                indexdtd = indexdtd + 1
                                            Next

                                        End If
                                    Next
                                End If

                                If Data_table_LINE.Rows.Count > 0 Then
                                    For i = 0 To Data_table_LINE.Rows.Count - 1

                                        Dim LINE1 As Line = Data_table_LINE.Rows(i).Item("LINE")
                                        Dim Col_int As New Point3dCollection
                                        LINE1.IntersectWith(Poly2D, Intersect.OnBothOperands, Col_int, IntPtr.Zero, IntPtr.Zero)
                                        If Col_int.Count > 0 Then


                                            For j = 0 To Col_int.Count - 1
                                                Dim point_on_2d As New Point3d
                                                point_on_2d = Col_int(j)
                                                Dim Point_on_3d As New Point3d
                                                Dim Param1 As Double = Poly2D.GetParameterAtPoint(point_on_2d)
                                                Point_on_3d = Poly3d.GetPointAtParameter(Param1)





                                                Dim Add_value As Boolean = True
                                                Dim Nr_val_data_table As Double = Data_table_for_singles.Rows.Count

                                                If Nr_val_data_table > 0 Then
                                                    For k = 0 To Nr_val_data_table - 1
                                                        If Data_table_for_singles.Rows(k).Item("X") = Round(point_on_2d.X, 2) And Data_table_for_singles.Rows(k).Item("Y") = Round(point_on_2d.Y, 2) Then
                                                            Add_value = False
                                                            Exit For
                                                        End If
                                                    Next
                                                End If


                                                If Add_value = True Then
                                                    Data_table_for_singles.Rows.Add()
                                                    Data_table_for_singles.Rows(indexdt).Item("X") = Round(Point_on_3d.X, 2)
                                                    Data_table_for_singles.Rows(indexdt).Item("Y") = Round(Point_on_3d.Y, 2)
                                                    Data_table_for_singles.Rows(indexdt).Item("Z") = Round(Point_on_3d.Z, 2)
                                                    Data_table_for_singles.Rows(indexdt).Item("CHAINAGE") = Round(Get_chainage_with_CSF(Poly3d, Point_on_3d, Data_table2, Data_table1), 2)
                                                    indexdt = indexdt + 1
                                                End If
                                                Data_table_for_doubles.Rows.Add()
                                                Data_table_for_doubles.Rows(indexdtd).Item("X") = Round(Point_on_3d.X, 2)
                                                Data_table_for_doubles.Rows(indexdtd).Item("Y") = Round(Point_on_3d.Y, 2)
                                                Data_table_for_doubles.Rows(indexdtd).Item("Z") = Round(Point_on_3d.Z, 2)
                                                Data_table_for_doubles.Rows(indexdtd).Item("CHAINAGE") = Round(Get_chainage_with_CSF(Poly3d, Point_on_3d, Data_table2, Data_table1), 2)


                                                Dim Id1 As ObjectId = LINE1.ObjectId
                                                Dim Records1 As Autodesk.Gis.Map.ObjectData.Records
                                                If Not Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False).Count = 0 Then
                                                    Records1 = Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                                    Dim Record1 As Autodesk.Gis.Map.ObjectData.Record
                                                    For Each Record1 In Records1
                                                        Dim Tabla1 As Autodesk.Gis.Map.ObjectData.Table
                                                        Tabla1 = Tables1(Record1.TableName)

                                                        Dim Field_defs1 As Autodesk.Gis.Map.ObjectData.FieldDefinitions
                                                        Field_defs1 = Tabla1.FieldDefinitions
                                                        For ii = 0 To Record1.Count - 1
                                                            Dim Field_def1 As Autodesk.Gis.Map.ObjectData.FieldDefinition
                                                            Field_def1 = Field_defs1(ii)
                                                            Dim Nume_field As String = Field_def1.Name.ToUpper
                                                            Dim Valoare_field As String = Record1(ii).StrValue

                                                            If Data_table_for_doubles.Columns.Contains(Nume_field) = False Then
                                                                Data_table_for_doubles.Columns.Add(Nume_field, GetType(String))
                                                            End If
                                                            Data_table_for_doubles.Rows(indexdtd).Item(Nume_field) = Valoare_field

                                                        Next

                                                    Next

                                                End If

                                                indexdtd = indexdtd + 1
                                            Next





                                        End If
                                    Next
                                End If

                            End If


                            If Data_table_for_singles.Rows.Count > 0 Then


                                For k = 0 To Data_table_for_singles.Rows.Count - 1
                                    Dim New_chainage As Double = Round(Data_table_for_singles.Rows(k).Item("CHAINAGE"), 2)
                                    Dim new_leader As New MLeader
                                    new_leader = Creaza_Mleader_nou_fara_UCS_transform(New Point3d(Data_table_for_singles.Rows(k).Item("X"), Data_table_for_singles.Rows(k).Item("Y"), Data_table_for_singles.Rows(k).Item("Z")), _
                                                                                       "East = " & Data_table_for_singles.Rows(k).Item("X") & vbCrLf & _
                                                                                       "North = " & Data_table_for_singles.Rows(k).Item("Y") & vbCrLf & _
                                                                                       "Elev = " & Data_table_for_singles.Rows(k).Item("Z") & vbCrLf & _
                                                                                       "Chainage = " & Round(New_chainage, 2) _
                                                                                       , 1, 0.2, 0.5, 2, 7)
                                Next



                                W1.Range(Col_east & "1").Value = "East"
                                W1.Range(Col_north & "1").Value = "North"
                                W1.Range(Col_Elevation & "1").Value = "Elevation"
                                W1.Range(Col_Chain & "1").Value = "Chainage"
                                If Data_table_for_doubles.Columns.Count > 4 Then
                                    Dim litera_next As Integer = Asc(Col_Chain) + 1
                                    Dim PREFIX As String = ""
                                    For k = 4 To Data_table_for_doubles.Columns.Count - 1
                                        W1.Range(PREFIX & Chr(litera_next) & "1").Value = Data_table_for_doubles.Columns(k).ColumnName
                                        If Chr(litera_next) = "Z" Then
                                            If PREFIX = "" Then
                                                PREFIX = "A"
                                            Else
                                                PREFIX = Chr(Asc(PREFIX) + 1)
                                            End If

                                            litera_next = Asc("A")
                                        Else
                                            litera_next = litera_next + 1
                                        End If

                                    Next
                                End If

                                Dim start1 As Integer = 2

                                For k = 0 To Data_table_for_doubles.Rows.Count - 1
                                    W1.Range(Col_east & start1).Value = Data_table_for_doubles.Rows(k).Item("X")
                                    W1.Range(Col_north & start1).Value = Data_table_for_doubles.Rows(k).Item("Y")
                                    W1.Range(Col_Elevation & start1).Value = Data_table_for_doubles.Rows(k).Item("Z")
                                    Dim New_chainage As Double = Round(Data_table_for_doubles.Rows(k).Item("CHAINAGE"), 2)
                                    W1.Range(Col_Chain & start1).Value = New_chainage
                                    If Data_table_for_doubles.Columns.Count > 4 Then
                                        Dim litera_next As Integer = Asc(Col_Chain) + 1
                                        Dim PREFIX As String = ""
                                        For ks = 4 To Data_table_for_doubles.Columns.Count - 1
                                            If IsDBNull(Data_table_for_doubles.Rows(k).Item(ks)) = False Then
                                                W1.Range(PREFIX & Chr(litera_next) & start1).Value = Data_table_for_doubles.Rows(k).Item(ks)
                                                If Chr(litera_next) = "Z" Then
                                                    If PREFIX = "" Then
                                                        PREFIX = "A"
                                                    Else
                                                        PREFIX = Chr(Asc(PREFIX) + 1)
                                                    End If
                                                    litera_next = Asc("A")
                                                Else
                                                    litera_next = litera_next + 1
                                                End If
                                            End If

                                        Next
                                    End If

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



    Private Sub Button_chainage_point_objects_Click(sender As Object, e As EventArgs) Handles Button_chainage_point_objects.Click

        Try
            If TextBox_east_point.Text = "" Then
                MsgBox("Please specify the East COLUMN!")
                Exit Sub
            End If
            If TextBox_north_point.Text = "" Then
                MsgBox("Please specify the North COLUMN!")
                Exit Sub
            End If
            If TextBox_elevation_Point.Text = "" Then
                MsgBox("Please specify the Elevation COLUMN!")
                Exit Sub
            End If
            If TextBox_chainage_point.Text = "" Then
                MsgBox("Please specify the Chainage COLUMN!")
                Exit Sub
            End If

            Dim start1 As Integer = 2

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
                Object_Prompt.MessageForAdding = vbLf & "Select 3d polyline:"
                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)


                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    If IsNothing(Rezultat1) = False Then
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                            BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)
                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(0)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                            If TypeOf Ent1 Is Polyline3d Then
                                Dim Poly3d As Polyline3d = Ent1
                                Dim Poly2D As New Polyline

                                Dim Index2d As Double = 0
                                For Each ObjId As ObjectId In Poly3d
                                    Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                    Poly2D.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)
                                    Index2d = Index2d + 1
                                Next


                                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = get_new_worksheet_from_Excel()
                                Dim Col_east As String = TextBox_east_intersection.Text.ToUpper
                                Dim Col_north As String = TextBox_north_intersection.Text.ToUpper
                                Dim Col_Chain As String = TextBox_chainage_intersection.Text.ToUpper
                                Dim Col_Elevation As String = TextBox_elevation_intersection.Text.ToUpper
                               

                                W1.Range(Col_east & "1").Value = "East"
                                W1.Range(Col_north & "1").Value = "North"
                                W1.Range(Col_Elevation & "1").Value = "Elevation"
                                W1.Range(Col_Chain & "1").Value = "Chainage"
                             



                                Dim Data_table1 As New System.Data.DataTable
                                Data_table1.Columns.Add("TEXT325", GetType(DBText))
                                Dim Index1 As Double = 0

                                Dim Data_table2 As New System.Data.DataTable
                                Data_table2.Columns.Add("TEXT0", GetType(DBText))
                                Dim Index2 As Double = 0


                                Dim Data_table_Points As New System.Data.DataTable
                                Data_table_Points.Columns.Add("X", GetType(Double))
                                Data_table_Points.Columns.Add("Y", GetType(Double))

                                Dim IndexXX As Double = 0

                                For Each ObjID In BTrecord
                                    Dim DBobject As DBObject = Trans1.GetObject(ObjID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                    If TypeOf DBobject Is DBPoint Then
                                        Dim Point1 As DBPoint = DBobject

                                        Data_table_Points.Rows.Add()
                                        Data_table_Points.Rows(IndexXX).Item("X") = Point1.Position.X
                                        Data_table_Points.Rows(IndexXX).Item("Y") = Point1.Position.Y
                                        IndexXX = IndexXX + 1

                                    End If
                                    If TypeOf DBobject Is DBText Then
                                        Dim Text1 As DBText = DBobject
                                        If Text1.Layer = Poly3d.Layer Then
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






                                For i = 0 To Data_table_Points.Rows.Count - 1

                                    Dim Point_zero As New Point3d(Data_table_Points(i).Item("X"), Data_table_Points(i).Item("Y"), 0)

                                    Dim Point_on_3d As New Point3d
                                    Point_on_3d = Poly3d.GetClosestPointTo(Point_zero, Vector3d.ZAxis, False)
                                    W1.Range(Col_east & start1).Value = Round(Point_on_3d.X, 2)
                                    W1.Range(Col_north & start1).Value = Round(Point_on_3d.Y, 2)
                                    W1.Range(Col_Elevation & start1).Value = Round(Point_on_3d.Z, 2)
                                    Dim New_chainage As Double = Round(Get_chainage_with_CSF(Poly3d, Point_on_3d, Data_table2, Data_table1), 1)
                                    W1.Range(Col_Chain & start1).Value = New_chainage
                                    Dim new_leader As New MLeader
                                    new_leader = Creaza_Mleader_nou_fara_UCS_transform(Point_on_3d, "East = " & Round(Point_on_3d.X, 2) & vbCrLf & _
                                                                                       "North = " & Round(Point_on_3d.Y, 2) & vbCrLf & _
                                                                                       "Elev = " & Round(Point_on_3d.Z, 2) & vbCrLf & _
                                                                                       "Chainage = " & Round(New_chainage, 2) _
                                                                                       , 1, 0.2, 0.5, 2, 7)
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

End Class