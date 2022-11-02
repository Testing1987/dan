Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class Graph_converter
    Dim Data_table_station_equation As System.Data.DataTable
    Dim Freeze_operations As Boolean = False
    Dim Poly_graph As Polyline
    Dim Point_min As New Point3d
    Dim Data_table_split As System.Data.DataTable

    Public Function Get_equation_value(ByVal Station_measured As Double) As Double
        Dim Valoare As Double = 0
        If IsNothing(Data_table_station_equation) = False Then
            If Data_table_station_equation.Rows.Count > 0 Then
                For i = 0 To Data_table_station_equation.Rows.Count - 1
                    If IsDBNull(Data_table_station_equation.Rows(i).Item("STATION_BACK")) = False And IsDBNull(Data_table_station_equation.Rows(i).Item("STATION_AHEAD")) = False Then
                        Dim Station_back As Double = Data_table_station_equation.Rows(i).Item("STATION_BACK")
                        Dim Station_ahead As Double = Data_table_station_equation.Rows(i).Item("STATION_AHEAD")

                        If Station_measured + Valoare < Station_back Then
                            Exit For
                        End If

                        Valoare = Valoare + Station_ahead - Station_back

                    End If
                Next
            End If


        End If


        Return Valoare
    End Function


    Private Sub Button_load_equations_from_excel_Click(sender As Object, e As EventArgs) Handles Button_load_equations_from_excel.Click
        If Freeze_operations = False Then
            Freeze_operations = True

            Try

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Dim Start1 As Integer = 0
                Dim End1 As Integer = 0
                If IsNumeric(TextBox_Row_Start_eq.Text) = True Then
                    Start1 = CInt(TextBox_Row_Start_eq.Text)
                End If
                If IsNumeric(TextBox_Row_End_eq.Text) = True Then
                    End1 = CInt(TextBox_Row_End_eq.Text)
                End If

                If End1 = 0 Or Start1 = 0 Then
                    Freeze_operations = False
                    Exit Sub
                End If
                If End1 < Start1 Then
                    Freeze_operations = False
                    Exit Sub
                End If
                Dim Column_Sta_Back As String = ""
                Column_Sta_Back = TextBox_col_station_back.Text.ToUpper
                Dim Column_sta_ahead As String = ""
                Column_sta_ahead = TextBox_col_statation_ahead.Text.ToUpper

                Data_table_station_equation = New System.Data.DataTable
                Data_table_station_equation.Columns.Add("STATION_BACK", GetType(Double))
                Data_table_station_equation.Columns.Add("STATION_AHEAD", GetType(Double))


                Dim Index_data_table As Double



                For i = Start1 To End1
                    Dim Station_back As String = Replace(W1.Range(Column_Sta_Back & i).Value2, "+", "")
                    Dim Station_ahead As String = Replace(W1.Range(Column_sta_ahead & i).Value2, "+", "")
                    If IsNumeric(Station_ahead) = True And IsNumeric(Station_back) = True Then

                        Data_table_station_equation.Rows.Add()
                        Data_table_station_equation.Rows(Index_data_table).Item("STATION_BACK") = CDbl(Station_back)
                        Data_table_station_equation.Rows(Index_data_table).Item("STATION_AHEAD") = CDbl(Station_ahead)
                        Index_data_table = Index_data_table + 1

                    Else
                        MsgBox("non numerical values on row " & i)
                        W1.Rows(i).select()
                        Freeze_operations = False
                        Exit Sub

                    End If
                Next


                Data_table_station_equation = Sort_data_table(Data_table_station_equation, "STATION_BACK")

                'MsgBox(Data_table_Centerline.Rows.Count)



            Catch ex As Exception
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try
        End If
        Freeze_operations = False
    End Sub

    Private Sub Button_load_parameters_Click(sender As Object, e As EventArgs) Handles Button_load_parameters.Click
        If Freeze_operations = False Then
            Freeze_operations = True

            If IsNumeric(TextBox_L_elevation.Text) = False Then
                MsgBox("Non numeric lowest elevation")
                Freeze_operations = False
                Exit Sub
            End If

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
            ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Dim Empty_array() As ObjectId
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Try
                Using Lock1 As Autodesk.AutoCAD.ApplicationServices.DocumentLock = ThisDrawing.LockDocument



                    Dim Result_point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                    Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Please the bottom left point: (STATION = 0+00  AND ELEVATION = " & TextBox_L_elevation.Text & ")")

                    PP1.AllowNone = False
                    Result_point1 = Editor1.GetPoint(PP1)
                    If Result_point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                        Freeze_operations = False
                        Exit Sub
                    End If


                    Point_min = Result_point1.Value


                    Dim Rezultat_profile_poly As Autodesk.AutoCAD.EditorInput.PromptEntityResult

                    Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select graph polyline:")

                    Object_Prompt1.SetRejectMessage(vbLf & "Please select a polyline")

                    Object_Prompt1.AddAllowedClass(GetType(Polyline), True)

                    Rezultat_profile_poly = Editor1.GetEntity(Object_Prompt1)


                    If Not Rezultat_profile_poly.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        MsgBox("NO centerline")
                        Editor1.WriteMessage(vbLf & "Command:")
                        Freeze_operations = False
                        Editor1.SetImpliedSelection(Empty_array)
                        Exit Sub
                    End If




                    Dim Rezultat_labels As Autodesk.AutoCAD.EditorInput.PromptEntityResult

                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select a graph label:")

                    Object_Prompt2.SetRejectMessage(vbLf & "Please select a text or an mtext object")
                    Object_Prompt2.AddAllowedClass(GetType(DBText), True)
                    Object_Prompt2.AddAllowedClass(GetType(MText), True)

                    Rezultat_labels = Editor1.GetEntity(Object_Prompt2)


                    If Not Rezultat_labels.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        MsgBox("NO centerline")
                        Editor1.WriteMessage(vbLf & "Command:")
                        Freeze_operations = False
                        Editor1.SetImpliedSelection(Empty_array)
                        Exit Sub
                    End If


                    Dim Rezultat_grid_lines As Autodesk.AutoCAD.EditorInput.PromptEntityResult

                    Dim Object_Prompt21 As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select a graph line:")

                    Object_Prompt21.SetRejectMessage(vbLf & "Please select a line or polyline")
                    Object_Prompt21.AddAllowedClass(GetType(Line), True)
                    Object_Prompt21.AddAllowedClass(GetType(Polyline), True)

                    Rezultat_grid_lines = Editor1.GetEntity(Object_Prompt21)


                    If Not Rezultat_grid_lines.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        MsgBox("NO centerline")
                        Editor1.WriteMessage(vbLf & "Command:")
                        Freeze_operations = False
                        Editor1.SetImpliedSelection(Empty_array)
                        Exit Sub
                    End If




                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                        Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord

                        BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                        Dim Text1 As DBText = TryCast(Trans1.GetObject(Rezultat_labels.ObjectId, OpenMode.ForRead), DBText)
                        If IsNothing(Text1) = False Then
                            If ComboBox_text_styles.Items.Contains(Text1.TextStyleName) = False Then
                                ComboBox_text_styles.Items.Add(Text1.TextStyleName)
                            End If
                            ComboBox_text_styles.SelectedIndex = ComboBox_text_styles.Items.IndexOf(Text1.TextStyleName)
                            TextBox_text_height.Text = Text1.Height
                            If ComboBox_layer_text.Items.Contains(Text1.Layer) = False Then
                                ComboBox_layer_text.Items.Add(Text1.Layer)
                            End If
                            ComboBox_layer_text.SelectedIndex = ComboBox_layer_text.Items.IndexOf(Text1.Layer)
                        End If

                        Dim Mtext1 As MText = TryCast(Trans1.GetObject(Rezultat_labels.ObjectId, OpenMode.ForRead), MText)
                        If IsNothing(Mtext1) = False Then
                            If ComboBox_text_styles.Items.Contains(Mtext1.TextStyleName) = False Then
                                ComboBox_text_styles.Items.Add(Mtext1.TextStyleName)
                            End If
                            ComboBox_text_styles.SelectedIndex = ComboBox_text_styles.Items.IndexOf(Mtext1.TextStyleName)
                            TextBox_text_height.Text = Mtext1.TextHeight
                            If ComboBox_layer_text.Items.Contains(Mtext1.Layer) = False Then
                                ComboBox_layer_text.Items.Add(Mtext1.Layer)
                            End If
                            ComboBox_layer_text.SelectedIndex = ComboBox_layer_text.Items.IndexOf(Mtext1.Layer)
                        End If

                        Poly_graph = TryCast(Trans1.GetObject(Rezultat_profile_poly.ObjectId, OpenMode.ForRead), Polyline)
                        If IsNothing(Poly_graph) = False Then
                            If ComboBox_layer_profile_polyline.Items.Contains(Poly_graph.Layer) = False Then
                                ComboBox_layer_profile_polyline.Items.Add(Poly_graph.Layer)
                            End If
                            ComboBox_layer_profile_polyline.SelectedIndex = ComboBox_layer_profile_polyline.Items.IndexOf(Poly_graph.Layer)

                        End If


                        Dim Poly_grid As Polyline = TryCast(Trans1.GetObject(Rezultat_grid_lines.ObjectId, OpenMode.ForRead), Polyline)
                        If IsNothing(Poly_grid) = False Then
                            If ComboBox_layer_grid_lines.Items.Contains(Poly_grid.Layer) = False Then
                                ComboBox_layer_grid_lines.Items.Add(Poly_grid.Layer)
                            End If
                            ComboBox_layer_grid_lines.SelectedIndex = ComboBox_layer_grid_lines.Items.IndexOf(Poly_grid.Layer)

                        End If

                        Dim Line_grid As Line = TryCast(Trans1.GetObject(Rezultat_grid_lines.ObjectId, OpenMode.ForRead), Line)
                        If IsNothing(Line_grid) = False Then
                            If ComboBox_layer_grid_lines.Items.Contains(Line_grid.Layer) = False Then
                                ComboBox_layer_grid_lines.Items.Add(Line_grid.Layer)
                            End If
                            ComboBox_layer_grid_lines.SelectedIndex = ComboBox_layer_grid_lines.Items.IndexOf(Line_grid.Layer)

                        End If

                    End Using
                End Using
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_draw_NEW_GRAPH_Click(sender As Object, e As EventArgs) Handles Button_draw_NEW_GRAPH.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
            ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Dim Empty_array() As ObjectId
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Try
                Using Lock1 As Autodesk.AutoCAD.ApplicationServices.DocumentLock = ThisDrawing.LockDocument

                    Dim Result_point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                    Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Please pick the grid location:")

                    PP1.AllowNone = False
                    Result_point1 = Editor1.GetPoint(PP1)
                    If Result_point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                        Freeze_operations = False
                        Exit Sub
                    End If
                    If IsNothing(Poly_graph) = True Then
                        MsgBox("Please load the existing parameters (no polyline loaded)")
                        Freeze_operations = False
                        Exit Sub

                    End If

                    If IsNothing(Data_table_station_equation) = True Then
                        MsgBox("Please load Station Equations values")
                        Freeze_operations = False
                        Exit Sub
                    End If

                    If Data_table_station_equation.Rows.Count = 0 Then
                        MsgBox("Please load Station Equations values")
                        Freeze_operations = False
                        Exit Sub
                    End If
                    Dim no_plot As String = "NO PLOT"
                    Creaza_layer(no_plot, 40, no_plot, False)

                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction


                        Dim Text_style_table As Autodesk.AutoCAD.DatabaseServices.TextStyleTable = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)


                        If Text_style_table.Has(ComboBox_text_styles.Text) = False Then
                            MsgBox("Not valid text style")
                            Freeze_operations = False
                            Exit Sub
                        End If

                        Dim Text_style_ID As Autodesk.AutoCAD.DatabaseServices.ObjectId = Text_style_table.Item(ComboBox_text_styles.Text)

                        Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord

                        BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                        Dim Grid_Z_Max, Grid_Z_Min As Double
                        If IsNumeric(TextBox_L_elevation.Text) = True Then
                            Grid_Z_Min = CDbl(TextBox_L_elevation.Text)
                        Else
                            MsgBox("Non numeric lowest elevation")
                            Freeze_operations = False
                            Exit Sub
                        End If

                        If IsNumeric(TextBox_H_Elevation.Text) = True Then
                            Grid_Z_Max = CDbl(TextBox_H_Elevation.Text)
                        Else
                            MsgBox("Non numeric highest elevation")
                            Freeze_operations = False
                            Exit Sub
                        End If

                        Dim Chainage_Graph_start As Double
                        Dim Chainage_Graph_end As Double

                        If IsNumeric(TextBox_Minimum_chainage.Text) = True Then
                            Chainage_Graph_start = CDbl(TextBox_Minimum_chainage.Text)
                        Else
                            MsgBox("Non numeric Chainage start Graph")
                            Freeze_operations = False
                            Exit Sub
                        End If

                        If IsNumeric(TextBox_Maximum_chainage.Text) = True Then
                            Chainage_Graph_end = CDbl(TextBox_Maximum_chainage.Text)
                        Else
                            MsgBox("Non numeric Chainage end Graph")
                            Freeze_operations = False
                            Exit Sub
                        End If

                        Dim Nr_hor_Lines As Double
                        Dim Horiz_increment, Vert_increment As Double
                        Dim hSCALE, vSCALE, Vertical_exag As Double
                        Dim Nr_vert_Lines, Chainage_poly_start As Double

                        Chainage_poly_start = 0


                        If IsNumeric(TextBox_Hincr.Text) = True Then
                            Horiz_increment = CDbl(TextBox_Hincr.Text)
                        Else
                            MsgBox("Non numeric horizontal increment")
                            Freeze_operations = False
                            Exit Sub
                        End If

                        If IsNumeric(TextBox_Vincr.Text) = True Then
                            Vert_increment = CDbl(TextBox_Vincr.Text)
                        Else
                            MsgBox("Non numeric vertical increment")
                            Freeze_operations = False
                            Exit Sub
                        End If

                        If Horiz_increment < Abs(Chainage_Graph_end) Then
                            Nr_vert_Lines = Ceiling(Abs(Chainage_Graph_end) / Horiz_increment)
                        Else
                            Nr_vert_Lines = 1
                        End If



                        Nr_hor_Lines = Ceiling((Grid_Z_Max - Grid_Z_Min) / Vert_increment)
                        If IsNumeric(TextBox_text_height.Text) = False Then
                            MsgBox("Non numeric text height")
                            Freeze_operations = False
                            Exit Sub
                        End If
                        Dim Text_height As Double = CDbl(TextBox_text_height.Text)

                        If IsNumeric(TextBox_Hscale.Text) = True Then
                            hSCALE = CDbl(TextBox_Hscale.Text)
                        Else
                            MsgBox("Non numeric horizontal scale")
                            Freeze_operations = False
                            Exit Sub
                        End If

                        If IsNumeric(TextBox_Vscale.Text) = True Then
                            vSCALE = CDbl(TextBox_Vscale.Text)
                        Else
                            MsgBox("Non numeric vertical scale")
                            Freeze_operations = False
                            Exit Sub
                        End If

                        Dim Printing_scale As Double = 1
                        Dim Viewport_scale As Double = 1
                        Dim Factor_print_viewport As Double
                        Factor_print_viewport = Printing_scale / Viewport_scale



                        Vertical_exag = hSCALE / vSCALE




                        Dim Pt1(0 To 1) As Double
                        Dim Pt2(0 To 1) As Double
                        Dim Pt3(0 To 1) As Double
                        Dim Pt4(0 To 1) As Double
                        Dim Pt5(0 To 1) As Double
                        Dim Pt6(0 To 1) As Double
                        Dim Pt7(0 To 1) As Double
                        Dim Pt8(0 To 1) As Double
                        Dim Pt9(0 To 1) As Double
                        Dim Pt10(0 To 1) As Double
                        Dim Pt11(0 To 1) As Double
                        Dim Pt12(0 To 1) As Double
                        Dim Pt13(0 To 1) As Double
                        Dim Pt14(0 To 1) As Double
                        Dim Pt15(0 To 1) As Double
                        Dim Pt16(0 To 1) As Double
                        Dim Pt17(0 To 1) As Double
                        Dim Pt18(0 To 1) As Double
                        Dim Pt19(0 To 1) As Double
                        Dim Pt20(0 To 1) As Double

                        Pt1(0) = Result_point1.Value.X
                        Pt1(1) = Result_point1.Value.Y

                        Dim Ytop As Double = Result_point1.Value.Y + Nr_hor_Lines * Vert_increment * (Factor_print_viewport / hSCALE) * Vertical_exag
                        Dim X As Double = Result_point1.Value.X
                        Dim Y As Double = Result_point1.Value.Y

                        Dim Stat_prev As Double = 0
                        Dim row1 As Integer = 0

                        Dim Vert_zero_line As New Autodesk.AutoCAD.DatabaseServices.Line
                        Vert_zero_line.Layer = ComboBox_layer_grid_lines.Text
                        Vert_zero_line.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(X, Y, 0)
                        Vert_zero_line.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(X, Ytop, 0)
                        BTrecord.AppendEntity(Vert_zero_line)
                        Dim BackSt0 As Double = Data_table_station_equation.Rows(0).Item("STATION_BACK")
                        Dim AheadSt0 As Double = Data_table_station_equation.Rows(0).Item("STATION_AHEAD")

                        If BackSt0 = 0 Then
                            Stat_prev = AheadSt0
                            row1 = 1
                        End If


                        Dim MText0 As New Autodesk.AutoCAD.DatabaseServices.MText
                        MText0.Contents = Get_chainage_feet_from_double(Stat_prev, 0)
                        MText0.Layer = ComboBox_layer_text.Text
                        MText0.TextStyleId = Text_style_ID
                        MText0.TextHeight = Text_height
                        MText0.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.TopCenter
                        MText0.Location = New Autodesk.AutoCAD.Geometry.Point3d(X, Y - 1.2 * Text_height, 0)
                        BTrecord.AppendEntity(MText0)
                        Trans1.AddNewlyCreatedDBObject(MText0, True)

                        Dim DataTable1 As New System.Data.DataTable
                        DataTable1.Columns.Add("STA1", GetType(Double))
                        DataTable1.Columns.Add("STA2", GetType(Double))
                        DataTable1.Columns.Add("STATION_AHEAD", GetType(Double))
                        Dim idx As Integer = 0
                        Dim Stat_prev1 As Double = 0
                        For i = row1 To Data_table_station_equation.Rows.Count - 1
                            Dim BackSt As Double = Data_table_station_equation.Rows(i).Item("STATION_BACK")
                            Dim AheadSt As Double = Data_table_station_equation.Rows(i).Item("STATION_AHEAD")
                            DataTable1.Rows.Add()
                            DataTable1.Rows(idx).Item(0) = Stat_prev1
                            DataTable1.Rows(idx).Item(1) = BackSt
                            DataTable1.Rows(idx).Item(2) = AheadSt
                            idx = idx + 1
                            Stat_prev1 = AheadSt
                        Next

                        Dim Last_label As Double = 0

                        For i = 0 To DataTable1.Rows.Count - 1
                            Dim Sta1 As Double = DataTable1.Rows(i).Item(0)
                            Dim Sta2 As Double = DataTable1.Rows(i).Item(1)
                            Dim Sta3 As Double = DataTable1.Rows(i).Item(2)
                            Last_label = Sta3
                            Dim Dist_NP As Double = (Sta2 - Sta1) * (Factor_print_viewport / hSCALE)


                            Dim Vert_np_line As New Autodesk.AutoCAD.DatabaseServices.Line
                            Vert_np_line.Layer = no_plot
                            Vert_np_line.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(X + Dist_NP, Y - 3 * Text_height, 0)
                            Vert_np_line.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(X + Dist_NP, Ytop, 0)

                            BTrecord.AppendEntity(Vert_np_line)
                            Trans1.AddNewlyCreatedDBObject(Vert_np_line, True)

                            Dim MText_np As New Autodesk.AutoCAD.DatabaseServices.MText
                            MText_np.Contents = "BACK = " & Sta2 & vbCrLf & "AHEAD = " & Sta3
                            MText_np.Layer = no_plot
                            MText_np.TextStyleId = Text_style_ID
                            MText_np.TextHeight = Text_height / 2
                            MText_np.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.TopCenter
                            MText_np.Location = New Autodesk.AutoCAD.Geometry.Point3d(X + Dist_NP, Y - 3 * Text_height - 1.2 * Text_height, 0)
                            BTrecord.AppendEntity(MText_np)
                            Trans1.AddNewlyCreatedDBObject(MText_np, True)




                            If Sta2 - Sta1 > Horiz_increment Then



                                Dim Label_next As Double = Ceiling((Sta1 + Horiz_increment) / Horiz_increment) * Horiz_increment - Horiz_increment
                                Dim Dist1 As Double = (Label_next - Sta1) * (Factor_print_viewport / hSCALE)
                                Dim X1 As Double = X + Dist1

                                For j = 0 To Floor((Sta2 - Sta1) / Horiz_increment)
                                    If X1 < X + Dist_NP Then
                                        Dim Vert_line As New Autodesk.AutoCAD.DatabaseServices.Line
                                        Vert_line.Layer = ComboBox_layer_grid_lines.Text
                                        Vert_line.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(X1, Y, 0)
                                        Vert_line.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(X1, Ytop, 0)
                                        BTrecord.AppendEntity(Vert_line)
                                        Trans1.AddNewlyCreatedDBObject(Vert_line, True)

                                        Dim MText_Station As New Autodesk.AutoCAD.DatabaseServices.MText
                                        MText_Station.Contents = Get_chainage_feet_from_double(Label_next, 0)
                                        MText_Station.Layer = ComboBox_layer_text.Text
                                        MText_Station.TextStyleId = Text_style_ID
                                        MText_Station.TextHeight = Text_height
                                        MText_Station.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.TopCenter
                                        MText_Station.Location = New Autodesk.AutoCAD.Geometry.Point3d(X1, Y - 1.2 * Text_height, 0)
                                        BTrecord.AppendEntity(MText_Station)
                                        Trans1.AddNewlyCreatedDBObject(MText_Station, True)
                                        Label_next = Label_next + Horiz_increment
                                        X1 = X1 + Horiz_increment * (Factor_print_viewport / hSCALE)
                                    Else
                                        Exit For
                                    End If



                                Next
                            End If

                            X = X + Dist_NP
                        Next

                        Dim Xend As Double = Result_point1.Value.X + Nr_vert_Lines * Horiz_increment * (Factor_print_viewport / hSCALE)

                        Dim Vert_line_end As New Autodesk.AutoCAD.DatabaseServices.Line
                        Vert_line_end.Layer = ComboBox_layer_grid_lines.Text
                        Vert_line_end.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xend, Y, 0)
                        Vert_line_end.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xend, Ytop, 0)
                        BTrecord.AppendEntity(Vert_line_end)
                        Trans1.AddNewlyCreatedDBObject(Vert_line_end, True)

                        If Xend > X Then
                            Dim Nr1 As Integer = Floor((Xend - X) / (Horiz_increment * (Factor_print_viewport / hSCALE)))
                            If Nr1 > 0 Then
                                Dim Label_next As Double = Ceiling((Last_label + Horiz_increment) / Horiz_increment) * Horiz_increment - Horiz_increment
                                Dim Dist1 As Double = (Label_next - Last_label) * (Factor_print_viewport / hSCALE)
                                Dim X1 As Double = X + Dist1
                                For i = 0 To Nr1
                                    If X1 < Xend Then
                                        Dim Vert_line As New Autodesk.AutoCAD.DatabaseServices.Line
                                        Vert_line.Layer = ComboBox_layer_grid_lines.Text
                                        Vert_line.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(X1, Y, 0)
                                        Vert_line.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(X1, Ytop, 0)
                                        BTrecord.AppendEntity(Vert_line)
                                        Trans1.AddNewlyCreatedDBObject(Vert_line, True)

                                        Dim MText_Station As New Autodesk.AutoCAD.DatabaseServices.MText
                                        MText_Station.Contents = Get_chainage_feet_from_double(Label_next, 0)
                                        MText_Station.Layer = ComboBox_layer_text.Text
                                        MText_Station.TextStyleId = Text_style_ID
                                        MText_Station.TextHeight = Text_height
                                        MText_Station.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.TopCenter
                                        MText_Station.Location = New Autodesk.AutoCAD.Geometry.Point3d(X1, Y - 1.2 * Text_height, 0)
                                        BTrecord.AppendEntity(MText_Station)
                                        Trans1.AddNewlyCreatedDBObject(MText_Station, True)
                                        Label_next = Label_next + Horiz_increment
                                        X1 = X1 + Horiz_increment * (Factor_print_viewport / hSCALE)
                                    Else
                                        Exit For
                                    End If


                                Next
                            End If

                        End If







                        Pt3(1) = Pt1(1)

                        Pt3(0) = Pt1(0)

                        Pt4(0) = Pt1(0) + Nr_vert_Lines * Horiz_increment * (Factor_print_viewport / hSCALE)
                        Pt4(1) = Pt1(1)

                        Pt6(0) = Pt3(0) ' - 10 * (Factor_print_viewport / 1000)
                        Pt6(1) = Pt1(1) '

                        Pt7(0) = Pt4(0) ' + 10 * (Factor_print_viewport / 1000)
                        Pt7(1) = Pt1(1) '


                        Pt8(0) = Pt1(0)
                        Pt8(1) = Pt1(1) - 1.2 * Text_height

                        Pt9(1) = Pt8(1)

                        Pt10(0) = Pt6(0) - 0.5 * Text_height
                        Pt10(1) = Pt6(1)

                        Pt12(0) = Pt7(0) + 0.5 * Text_height
                        Pt12(1) = Pt7(1)

                        Pt11(0) = Pt10(0)
                        Pt13(0) = Pt12(0)














                        Dim Horiz_Zero_line As New Autodesk.AutoCAD.DatabaseServices.Line
                        Horiz_Zero_line.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Pt3(0), Pt3(1), 0)
                        Horiz_Zero_line.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Pt4(0), Pt4(1), 0)
                        Horiz_Zero_line.Layer = ComboBox_layer_grid_lines.Text
                        BTrecord.AppendEntity(Horiz_Zero_line)
                        Trans1.AddNewlyCreatedDBObject(Horiz_Zero_line, True)

                        For i = 1 To Nr_hor_Lines
                            Dim Xh1, Yh1, Xh2, Yh2 As Double
                            Xh1 = Pt3(0)
                            Xh2 = Pt4(0)

                            Yh1 = Pt3(1) + i * (Factor_print_viewport / hSCALE) * Vertical_exag * Vert_increment
                            Yh2 = Yh1
                            Dim h_off_line As New Autodesk.AutoCAD.DatabaseServices.Line
                            h_off_line.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xh1, Yh1, 0)
                            h_off_line.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xh2, Yh2, 0)
                            h_off_line.Layer = ComboBox_layer_grid_lines.Text
                            BTrecord.AppendEntity(h_off_line)
                            Trans1.AddNewlyCreatedDBObject(h_off_line, True)
                        Next












                        Dim Vstring As String
                        Vstring = Round(Grid_Z_Min, 0).ToString

                        Dim Mtext_ch_Ver_left As New Autodesk.AutoCAD.DatabaseServices.MText
                        Mtext_ch_Ver_left.Contents = Vstring
                        Mtext_ch_Ver_left.Layer = ComboBox_layer_text.Text
                        Mtext_ch_Ver_left.TextStyleId = Text_style_ID
                        Mtext_ch_Ver_left.TextHeight = Text_height

                        Mtext_ch_Ver_left.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.MiddleRight

                        Mtext_ch_Ver_left.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt10(0), Pt10(1), 0)
                        BTrecord.AppendEntity(Mtext_ch_Ver_left)
                        Trans1.AddNewlyCreatedDBObject(Mtext_ch_Ver_left, True)

                        Dim Mtext_ch_Ver_right As New Autodesk.AutoCAD.DatabaseServices.MText
                        Mtext_ch_Ver_right.Contents = Vstring
                        Mtext_ch_Ver_right.Layer = ComboBox_layer_text.Text
                        Mtext_ch_Ver_right.TextStyleId = Text_style_ID
                        Mtext_ch_Ver_right.TextHeight = Text_height
                        Mtext_ch_Ver_right.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.MiddleLeft

                        Mtext_ch_Ver_right.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt12(0), Pt12(1), 0)
                        BTrecord.AppendEntity(Mtext_ch_Ver_right)
                        Trans1.AddNewlyCreatedDBObject(Mtext_ch_Ver_right, True)

                        For i = 1 To Nr_hor_Lines

                            Pt11(1) = Pt10(1) + i * Vert_increment * (Factor_print_viewport / hSCALE) * Vertical_exag

                            Pt13(1) = Pt12(1) + i * Vert_increment * (Factor_print_viewport / hSCALE) * Vertical_exag



                            Vstring = Round((i * Vert_increment) + Grid_Z_Min, 0).ToString

                            Dim Mtext_ch_Ver_left1 As New Autodesk.AutoCAD.DatabaseServices.MText
                            Mtext_ch_Ver_left1.Contents = Vstring
                            Mtext_ch_Ver_left1.Layer = ComboBox_layer_text.Text
                            Mtext_ch_Ver_left1.TextStyleId = Text_style_ID
                            Mtext_ch_Ver_left1.TextHeight = Text_height
                            Mtext_ch_Ver_left1.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.MiddleRight

                            Mtext_ch_Ver_left1.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt11(0), Pt11(1), 0)
                            BTrecord.AppendEntity(Mtext_ch_Ver_left1)
                            Trans1.AddNewlyCreatedDBObject(Mtext_ch_Ver_left1, True)


                            Dim Mtext_ch_Ver_right1 As New Autodesk.AutoCAD.DatabaseServices.MText
                            Mtext_ch_Ver_right1.Contents = Vstring
                            Mtext_ch_Ver_right1.Layer = ComboBox_layer_text.Text
                            Mtext_ch_Ver_right1.TextStyleId = Text_style_ID
                            Mtext_ch_Ver_right1.TextHeight = Text_height
                            Mtext_ch_Ver_right1.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.MiddleLeft
                            Mtext_ch_Ver_right1.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt13(0), Pt13(1), 0)
                            BTrecord.AppendEntity(Mtext_ch_Ver_right1)
                            Trans1.AddNewlyCreatedDBObject(Mtext_ch_Ver_right1, True)


                        Next

                        Dim Profile_x_point(0 To Poly_graph.NumberOfVertices - 1) As Double
                        Dim Profile_y_point(0 To Poly_graph.NumberOfVertices - 1) As Double

                        Dim Poly_copy As New Autodesk.AutoCAD.DatabaseServices.Polyline()



                        Profile_x_point(0) = Result_point1.Value.X + (Chainage_poly_start) * (Factor_print_viewport / hSCALE)
                        Profile_y_point(0) = Result_point1.Value.Y + Poly_graph.StartPoint.Y - Point_min.Y



                        Poly_copy.AddVertexAt(0, New Autodesk.AutoCAD.Geometry.Point2d(Profile_x_point(0), Profile_y_point(0)), 0, 0, 0)
                        'here You can put CSF!
                        For i = 1 To Poly_graph.NumberOfVertices - 1
                            Profile_x_point(i) = Profile_x_point(i - 1) + Poly_graph.GetPoint3dAt(i).X - Poly_graph.GetPoint3dAt(i - 1).X
                            Profile_y_point(i) = Profile_y_point(i - 1) + Poly_graph.GetPoint3dAt(i).Y - Poly_graph.GetPoint3dAt(i - 1).Y
                            Poly_copy.AddVertexAt(i, New Autodesk.AutoCAD.Geometry.Point2d(Profile_x_point(i), Profile_y_point(i)), 0, 0, 0)
                        Next

                        Poly_copy.Layer = ComboBox_layer_profile_polyline.Text
                        Poly_copy.Linetype = Poly_graph.Linetype
                        Poly_copy.LineWeight = Poly_graph.LineWeight
                        Poly_copy.ColorIndex = Poly_graph.ColorIndex

                        BTrecord.AppendEntity(Poly_copy)
                        Trans1.AddNewlyCreatedDBObject(Poly_copy, True)
                        Trans1.Commit()
                    End Using
                End Using

                Editor1.Regen()

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_blocks_to_excel_Click(sender As Object, e As EventArgs) Handles Button_blocks_to_excel.Click
        Try




            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            If Freeze_operations = False Then

                Freeze_operations = True


                Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument

                    Dim Table_data1 As New System.Data.DataTable

                    Table_data1.Columns.Add("BLOCK_NAME", GetType(String))
                    Table_data1.Columns.Add("STATION_MEASURED", GetType(Double))
                    Table_data1.Columns.Add("STATION_EQ", GetType(Double))
                   



                    ' Dim k As Double = 1
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        

                        Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)


                        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor

                        Editor1 = ThisDrawing.Editor
                        Dim Empty_array() As ObjectId
                        Editor1.SetImpliedSelection(Empty_array)

                        Dim UCS_CURENT As Matrix3d = Editor1.CurrentUserCoordinateSystem


                        '****************************************************************************************



                        Dim Rezultat_hline As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                        Dim Object_promptH As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_promptH.MessageForAdding = vbLf & "Select a known Vertical line and the label for it (STATION):"

                        Object_promptH.SingleOnly = False
                        Rezultat_hline = Editor1.GetSelection(Object_promptH)


                        Dim Rezultat_hlineSCALE As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify Horizontal Exaggeration:")
                        Rezultat_hlineSCALE.DefaultValue = 1
                        Rezultat_hlineSCALE.AllowNone = True
                        Dim Rezultat_hline44 As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(Rezultat_hlineSCALE)

                        Dim H_EXAG As Double = Rezultat_hline44.Value
                        If H_EXAG = 0 Then H_EXAG = 1

                        If Rezultat_hline.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Editor1.SetImpliedSelection(Empty_array)
                            Freeze_operations = False
                            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                            Freeze_operations = False
                            Exit Sub
                        End If

                        If Rezultat_hline.Value.Count <> 2 Then
                            MsgBox("Your selection contains " & Rezultat_hline.Value.Count & " objects")
                            Editor1.SetImpliedSelection(Empty_array)
                            Freeze_operations = False
                            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                            Freeze_operations = False
                            Exit Sub
                        End If

                        Dim Known_station As Double = -100000



                        Dim mText_cunoscut_chainage As Autodesk.AutoCAD.DatabaseServices.MText
                        Dim Text_cunoscut_chainage As Autodesk.AutoCAD.DatabaseServices.DBText


                        Dim Obj2_chainage As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj2_chainage = Rezultat_hline.Value.Item(0)
                        Dim Ent2_chainage As Entity
                        Ent2_chainage = Obj2_chainage.ObjectId.GetObject(OpenMode.ForRead)
                        Dim Obj3_chainage As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj3_chainage = Rezultat_hline.Value.Item(1)
                        Dim Ent3_chainage As Entity
                        Ent3_chainage = Obj3_chainage.ObjectId.GetObject(OpenMode.ForRead)

                        If TypeOf Ent2_chainage Is Autodesk.AutoCAD.DatabaseServices.MText Then
                            mText_cunoscut_chainage = Ent2_chainage
                            If IsNumeric(Replace(mText_cunoscut_chainage.Text, "+", "")) = True Then Known_station = CDbl(Replace(mText_cunoscut_chainage.Text, "+", ""))
                        End If

                        If TypeOf Ent3_chainage Is Autodesk.AutoCAD.DatabaseServices.MText Then
                            mText_cunoscut_chainage = Ent3_chainage
                            If IsNumeric(Replace(mText_cunoscut_chainage.Text, "+", "")) = True Then Known_station = CDbl(Replace(mText_cunoscut_chainage.Text, "+", ""))
                        End If

                        If TypeOf Ent2_chainage Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                            Text_cunoscut_chainage = Ent2_chainage
                            If IsNumeric(Replace(Text_cunoscut_chainage.TextString, "+", "")) = True Then Known_station = CDbl(Replace(Text_cunoscut_chainage.TextString, "+", ""))
                        End If

                        If TypeOf Ent3_chainage Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                            Text_cunoscut_chainage = Ent3_chainage
                            If IsNumeric(Replace(Text_cunoscut_chainage.TextString, "+", "")) = True Then Known_station = CDbl(Replace(Text_cunoscut_chainage.TextString, "+", ""))
                        End If

                        If Known_station = -100000 Then
                            MsgBox("Chainage datum not numeric")
                            Editor1.SetImpliedSelection(Empty_array)
                            Freeze_operations = False
                            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                            Freeze_operations = False
                            Exit Sub
                        End If


                        Dim Linia_cunoscuta As Autodesk.AutoCAD.DatabaseServices.Line
                        Dim x0_sta1, x0_sta2 As Double
                        Dim y0_sta1, y0_sta2 As Double

                        If TypeOf Ent2_chainage Is Autodesk.AutoCAD.DatabaseServices.Line Then
                            Linia_cunoscuta = Ent2_chainage
                            x0_sta1 = Linia_cunoscuta.StartPoint.X 'TransformBy(UCS_CURENT).X
                            x0_sta2 = Linia_cunoscuta.EndPoint.X 'TransformBy(UCS_CURENT).X
                            y0_sta1 = Linia_cunoscuta.StartPoint.Y 'TransformBy(UCS_CURENT).Y
                            y0_sta2 = Linia_cunoscuta.EndPoint.Y ' TransformBy(UCS_CURENT).Y
                            If Abs(x0_sta1 - x0_sta2) > 0.001 Then
                                MsgBox("Vertical line you selected is not vertical")
                                Editor1.SetImpliedSelection(Empty_array)
                                Freeze_operations = False
                                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                Freeze_operations = False
                                Exit Sub
                            End If

                        End If

                        If TypeOf Ent3_chainage Is Autodesk.AutoCAD.DatabaseServices.Line Then
                            Linia_cunoscuta = Ent3_chainage
                            x0_sta1 = Linia_cunoscuta.StartPoint.X 'TransformBy(UCS_CURENT).X
                            x0_sta2 = Linia_cunoscuta.EndPoint.X 'TransformBy(UCS_CURENT).X
                            y0_sta1 = Linia_cunoscuta.StartPoint.Y 'TransformBy(UCS_CURENT).Y
                            y0_sta2 = Linia_cunoscuta.EndPoint.Y 'TransformBy(UCS_CURENT).Y
                            If Abs(x0_sta1 - x0_sta2) > 0.001 Then
                                MsgBox("Vertical line you selected is not vertical")
                                Editor1.SetImpliedSelection(Empty_array)
                                Freeze_operations = False
                                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                Freeze_operations = False
                                Exit Sub
                            End If

                        End If

                        If TypeOf Ent2_chainage Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                            Dim Poly123 As Polyline = Ent2_chainage
                            x0_sta1 = Poly123.StartPoint.X 'TransformBy(UCS_CURENT).X
                            x0_sta2 = Poly123.EndPoint.X 'TransformBy(UCS_CURENT).X
                            y0_sta1 = Poly123.StartPoint.Y 'TransformBy(UCS_CURENT).Y
                            y0_sta2 = Poly123.EndPoint.Y ' TransformBy(UCS_CURENT).Y
                            If Abs(x0_sta1 - x0_sta2) > 0.001 Then
                                MsgBox("Vertical line you selected is not vertical")
                                Editor1.SetImpliedSelection(Empty_array)
                                Freeze_operations = False
                                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                Freeze_operations = False
                                Exit Sub
                            End If

                        End If

                        If TypeOf Ent3_chainage Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                            Dim Poly123 As Polyline = Ent3_chainage
                            x0_sta1 = Poly123.StartPoint.X 'TransformBy(UCS_CURENT).X
                            x0_sta2 = Poly123.EndPoint.X 'TransformBy(UCS_CURENT).X
                            y0_sta1 = Poly123.StartPoint.Y 'TransformBy(UCS_CURENT).Y
                            y0_sta2 = Poly123.EndPoint.Y ' TransformBy(UCS_CURENT).Y
                            If Abs(x0_sta1 - x0_sta2) > 0.001 Then
                                MsgBox("Vertical line you selected is not vertical")
                                Editor1.SetImpliedSelection(Empty_array)
                                Freeze_operations = False
                                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                Freeze_operations = False
                                Exit Sub
                            End If

                        End If

                       

                        Dim Elevatia_cunoscuta As Double = -100000
                        Dim V_EXAG As Double = 1
                        Dim Y_elev_cunoscut As Double = 0



                        Dim Rezultat_vert As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                        Dim Object_prompt_vert As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_prompt_vert.MessageForAdding = vbLf & "Select a known horizontal line (ELEVATION) and the label for it:"

                        Object_prompt_vert.SingleOnly = False
                        Rezultat_vert = Editor1.GetSelection(Object_prompt_vert)


                        Dim Rezultat_vertSCALE As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify Vertical Exaggeration:")
                        Rezultat_vertSCALE.DefaultValue = 1
                        Rezultat_vertSCALE.AllowNone = True
                        Dim Rezultat_vscale As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(Rezultat_vertSCALE)

                        V_EXAG = Rezultat_vscale.Value
                        If V_EXAG = 0 Then V_EXAG = 1


                        Dim Obj1v As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1v = Rezultat_vert.Value.Item(0)
                        Dim Ent1v As Entity
                        Ent1v = Obj1v.ObjectId.GetObject(OpenMode.ForRead)

                        Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj2 = Rezultat_vert.Value.Item(1)
                        Dim Ent2 As Entity
                        Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)

                        Dim mText_cunoscut As Autodesk.AutoCAD.DatabaseServices.MText
                        Dim Text_cunoscut As Autodesk.AutoCAD.DatabaseServices.DBText

                        If TypeOf Ent1v Is Autodesk.AutoCAD.DatabaseServices.MText Then
                            mText_cunoscut = Ent1v
                            If IsNumeric(Replace(mText_cunoscut.Text, "'", "")) = True Then Elevatia_cunoscuta = CDbl(Replace(mText_cunoscut.Text, "'", ""))
                        End If

                        If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                            mText_cunoscut = Ent2
                            If IsNumeric(Replace(mText_cunoscut.Text, "'", "")) = True Then Elevatia_cunoscuta = CDbl(Replace(mText_cunoscut.Text, "'", ""))
                        End If

                        If TypeOf Ent1v Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                            Text_cunoscut = Ent1v
                            If IsNumeric(Replace(Text_cunoscut.TextString, "'", "")) = True Then Elevatia_cunoscuta = CDbl(Replace(Text_cunoscut.TextString, "'", ""))
                        End If

                        If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                            Text_cunoscut = Ent2
                            If IsNumeric(Replace(Text_cunoscut.TextString, "'", "")) = True Then Elevatia_cunoscuta = CDbl(Replace(Text_cunoscut.TextString, "'", ""))
                        End If

                        Dim Linia_Elev_cunoscuta As Autodesk.AutoCAD.DatabaseServices.Line
                        Dim polylinia_elev_cunoscuta As Autodesk.AutoCAD.DatabaseServices.Polyline

                        Dim x0e1, y0e1, x0e2, y0e2 As Double

                        If TypeOf Ent1v Is Autodesk.AutoCAD.DatabaseServices.Line Then
                            Linia_Elev_cunoscuta = Ent1v
                            x0e1 = Linia_Elev_cunoscuta.StartPoint.X
                            y0e1 = Linia_Elev_cunoscuta.StartPoint.Y
                            x0e2 = Linia_Elev_cunoscuta.EndPoint.X
                            y0e2 = Linia_Elev_cunoscuta.EndPoint.Y
                            If Abs(y0e1 - y0e2) > 0.001 Then
                                Editor1.SetImpliedSelection(Empty_array)
                                MsgBox("Segment not horizontal")
                                Freeze_operations = False
                                Exit Sub
                            End If
                            Y_elev_cunoscut = y0e1

                        End If


                        If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                            Linia_Elev_cunoscuta = Ent2
                            x0e1 = Linia_Elev_cunoscuta.StartPoint.X
                            y0e1 = Linia_Elev_cunoscuta.StartPoint.Y
                            x0e2 = Linia_Elev_cunoscuta.EndPoint.X
                            y0e2 = Linia_Elev_cunoscuta.EndPoint.Y
                            If Abs(y0e1 - y0e2) > 0.001 Then
                                Editor1.SetImpliedSelection(Empty_array)
                                Freeze_operations = False
                                Exit Sub
                            End If
                            Y_elev_cunoscut = y0e1

                        End If

                        If TypeOf Ent1v Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                            polylinia_elev_cunoscuta = Ent1v

                            x0e1 = polylinia_elev_cunoscuta.StartPoint.X
                            y0e1 = polylinia_elev_cunoscuta.StartPoint.Y
                            x0e2 = polylinia_elev_cunoscuta.EndPoint.X
                            y0e2 = polylinia_elev_cunoscuta.EndPoint.Y
                            If Abs(y0e1 - y0e2) > 0.001 Then
                                Editor1.SetImpliedSelection(Empty_array)
                                MsgBox("Segment not horizontal")
                                Freeze_operations = False
                                Exit Sub
                            End If
                            Y_elev_cunoscut = y0e1

                        End If

                        If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                            polylinia_elev_cunoscuta = Ent2

                            x0e1 = polylinia_elev_cunoscuta.StartPoint.X
                            y0e1 = polylinia_elev_cunoscuta.StartPoint.Y
                            x0e2 = polylinia_elev_cunoscuta.EndPoint.X
                            y0e2 = polylinia_elev_cunoscuta.EndPoint.Y
                            If Abs(y0e1 - y0e2) > 0.001 Then
                                Editor1.SetImpliedSelection(Empty_array)
                                MsgBox("Segment not horizontal")
                                Freeze_operations = False
                                Exit Sub
                            End If
                            Y_elev_cunoscut = y0e1

                        End If

                        Dim Rezultat_blocks As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                        Dim Object_Prompt_blocks As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_Prompt_blocks.MessageForAdding = vbLf & "Select the blocks you want to transfer to excel:"
                        Object_Prompt_blocks.SingleOnly = False
                        Rezultat_blocks = Editor1.GetSelection(Object_Prompt_blocks)

                        Dim Index1 As Integer = 0

                        If Rezultat_blocks.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            For i = 0 To Rezultat_blocks.Value.Count - 1
                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Rezultat_blocks.Value.Item(i)
                                Dim Ent1 As Entity
                                Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)


                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.BlockReference Then
                                    Dim Block1 As BlockReference = Ent1
                                    Dim X As Double = Block1.Position.X
                                    Dim Measured_station As Double = Known_station - (x0_sta1 - X) * H_EXAG
                                    Dim Station_with_eq As Double = Measured_station + Get_equation_value(Measured_station)
                                    Dim Block_name As String

                                    Dim BlockTrec As BlockTableRecord = Nothing
                                    If Block1.IsDynamicBlock = True Then
                                        BlockTrec = Trans1.GetObject(Block1.DynamicBlockTableRecord, OpenMode.ForRead)
                                        Block_name = BlockTrec.Name
                                    Else
                                        BlockTrec = Trans1.GetObject(Block1.BlockTableRecord, OpenMode.ForRead)
                                        Block_name = BlockTrec.Name
                                    End If
                                    Table_data1.Rows.Add()
                                    Table_data1.Rows(Index1).Item("BLOCK_NAME") = Block_name
                                    Table_data1.Rows(Index1).Item("STATION_MEASURED") = Measured_station
                                    Table_data1.Rows(Index1).Item("STATION_EQ") = Station_with_eq

                                    Index1 = Index1 + 1
                                End If



                            Next
                        End If




                        

                        Trans1.Commit()
                        ' asta e de la tranzactie
                    End Using

                    If IsNothing(Table_data1) = False Then
                        If Table_data1.Rows.Count > 0 Then
                            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_NEW_worksheet_from_Excel()
                            Transfer_data_to_Excel(W1, Table_data1)

                        End If
                    End If






                    Freeze_operations = False



                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    ' asta e de la lock
                End Using
                Freeze_operations = False
            End If
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

        Catch ex As Exception
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

            Freeze_operations = False
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Function Transfer_data_to_Excel(ByVal w1 As Microsoft.Office.Interop.Excel.Worksheet, ByVal Data_table1 As System.Data.DataTable)
        If IsNothing(Data_table1) = False Then
            If Data_table1.Rows.Count > 0 Then


                Dim NrR As Integer = Data_table1.Rows.Count
                Dim NrC As Integer = Data_table1.Columns.Count

                Dim values(NrR, NrC - 1) As Object


                For i = 0 To NrR - 1
                    For j = 0 To NrC - 1
                        If IsDBNull(Data_table1.Rows(i).Item(j)) = False Then
                            values(i + 1, j) = Data_table1.Rows(i).Item(j)
                        End If
                    Next
                Next
                For j = 0 To NrC - 1
                    values(0, j) = Data_table1.Columns(j).ColumnName
                Next


                Dim range1 As Microsoft.Office.Interop.Excel.Range = w1.Range(w1.Cells(1, 1), w1.Cells(Data_table1.Rows.Count + 1, Data_table1.Columns.Count))
                range1.Cells.NumberFormat = "@"
                range1.Value2 = values


            End If
        End If
    End Function


    Private Sub Button_split_load_from_excel_Click(sender As Object, e As EventArgs) Handles Button_split_load_from_excel.Click
        If Freeze_operations = False Then
            Freeze_operations = True

            Try

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Dim Start1 As Integer = 0
                Dim End1 As Integer = 0
                If IsNumeric(TextBox_split_row_start.Text) = True Then
                    Start1 = CInt(TextBox_split_row_start.Text)
                End If
                If IsNumeric(TextBox_split_row_end.Text) = True Then
                    End1 = CInt(TextBox_split_row_end.Text)
                End If

                If End1 = 0 Or Start1 = 0 Then
                    Freeze_operations = False
                    Exit Sub
                End If
                If End1 < Start1 Then
                    Freeze_operations = False
                    Exit Sub
                End If
                Dim Column_Sta1 As String = ""
                Column_Sta1 = TextBox_split_sta1.Text.ToUpper
                Dim Column_sta2 As String = ""
                Column_sta2 = TextBox_split_sta2.Text.ToUpper
                Dim Column_lbl As String = ""
                Column_lbl = TextBox_split_label.Text.ToUpper

                Data_table_split = New System.Data.DataTable
                Data_table_split.Columns.Add("STA1", GetType(Double))
                Data_table_split.Columns.Add("STA2", GetType(Double))
                Data_table_split.Columns.Add("LBL", GetType(String))



                For i = Start1 To End1
                    Dim Station1 As String = Replace(W1.Range(Column_Sta1 & i).Value2, "+", "")
                    Dim Station2 As String = Replace(W1.Range(Column_sta2 & i).Value2, "+", "")
                    Dim label As String = W1.Range(Column_lbl & i).Value2
                    If IsNumeric(Station2) = True And IsNumeric(Station1) = True Then

                        Data_table_split.Rows.Add()
                        Data_table_split.Rows(Data_table_split.Rows.Count - 1).Item("STA1") = CDbl(Station1)
                        Data_table_split.Rows(Data_table_split.Rows.Count - 1).Item("STA2") = CDbl(Station2)
                        Data_table_split.Rows(Data_table_split.Rows.Count - 1).Item("LBL") = label

                    Else
                        MsgBox("non numerical values on row " & i)
                        W1.Rows(i).select()
                        Freeze_operations = False
                        Exit Sub

                    End If
                Next

                Transfer_datatable_to_new_excel_spreadsheet(Data_table_split)

                'MsgBox(Data_table_Centerline.Rows.Count)



            Catch ex As Exception
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try
        End If
        Freeze_operations = False
    End Sub

    Private Sub Button_split_graph_Click(sender As Object, e As EventArgs) Handles Button_split_graph.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Dim last_good_index As Integer = 0

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
            ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Dim Empty_array() As ObjectId
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Try
                Using Lock1 As Autodesk.AutoCAD.ApplicationServices.DocumentLock = ThisDrawing.LockDocument

                    Dim Result_point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                    Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Please pick the grid location:")

                    PP1.AllowNone = False
                    Result_point1 = Editor1.GetPoint(PP1)
                    If Result_point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                        Freeze_operations = False
                        Exit Sub
                    End If
                    If IsNothing(Poly_graph) = True Then
                        MsgBox("Please load the existing parameters (no polyline loaded)")
                        Freeze_operations = False
                        Exit Sub

                    End If

                    If IsNothing(Data_table_split) = True Then
                        MsgBox("Please load split values")
                        Freeze_operations = False
                        Exit Sub
                    End If

                    If Data_table_split.Rows.Count = 0 Then
                        MsgBox("Please load split values")
                        Freeze_operations = False
                        Exit Sub
                    End If

                    Dim no_plot As String = "NO PLOT"
                    Creaza_layer(no_plot, 40, no_plot, False)

                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction


                        Dim Text_style_table As Autodesk.AutoCAD.DatabaseServices.TextStyleTable = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)


                        If Text_style_table.Has(ComboBox_text_styles.Text) = False Then
                            MsgBox("Not valid text style")
                            Freeze_operations = False
                            Exit Sub
                        End If

                        Dim Text_style_ID As Autodesk.AutoCAD.DatabaseServices.ObjectId = Text_style_table.Item(ComboBox_text_styles.Text)

                        Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord

                        BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                        Dim Grid_Z_Max, Grid_Z_Min As Double
                        If IsNumeric(TextBox_L_elevation.Text) = True Then
                            Grid_Z_Min = CDbl(TextBox_L_elevation.Text)
                        Else
                            MsgBox("Non numeric lowest elevation")
                            Freeze_operations = False
                            Exit Sub
                        End If

                        If IsNumeric(TextBox_H_Elevation.Text) = True Then
                            Grid_Z_Max = CDbl(TextBox_H_Elevation.Text)
                        Else
                            MsgBox("Non numeric highest elevation")
                            Freeze_operations = False
                            Exit Sub
                        End If



                        Dim X0 As Double = Result_point1.Value.X
                        Dim Y0 As Double = Result_point1.Value.Y

                        Dim Horiz_increment, Vert_increment As Double

                        If IsNumeric(TextBox_Hincr.Text) = True Then
                            Horiz_increment = CDbl(TextBox_Hincr.Text)
                        Else
                            MsgBox("Non numeric horizontal increment")
                            Freeze_operations = False
                            Exit Sub
                        End If

                        If IsNumeric(TextBox_Vincr.Text) = True Then
                            Vert_increment = CDbl(TextBox_Vincr.Text)
                        Else
                            MsgBox("Non numeric vertical increment")
                            Freeze_operations = False
                            Exit Sub
                        End If


                        If IsNumeric(TextBox_text_height.Text) = False Then
                            MsgBox("Non numeric text height")
                            Freeze_operations = False
                            Exit Sub
                        End If
                        Dim Text_height As Double = CDbl(TextBox_text_height.Text)

                        Dim hSCALE, vSCALE, Vertical_exag As Double

                        If IsNumeric(TextBox_Hscale.Text) = True Then
                            hSCALE = CDbl(TextBox_Hscale.Text)
                        Else
                            MsgBox("Non numeric horizontal scale")
                            Freeze_operations = False
                            Exit Sub
                        End If

                        If IsNumeric(TextBox_Vscale.Text) = True Then
                            vSCALE = CDbl(TextBox_Vscale.Text)
                        Else
                            MsgBox("Non numeric vertical scale")
                            Freeze_operations = False
                            Exit Sub
                        End If

                        Vertical_exag = hSCALE / vSCALE

                        Dim Printing_scale As Double = 1
                        Dim Viewport_scale As Double = 1
                        Dim Factor_print_viewport As Double
                        Factor_print_viewport = Printing_scale / Viewport_scale

                        Dim Spacing As Double = 1.5 * (Grid_Z_Max - Grid_Z_Min) * (Factor_print_viewport / hSCALE) * Vertical_exag



                        Dim Nr_hor_Lines As Double
                        Nr_hor_Lines = Ceiling((Grid_Z_Max - Grid_Z_Min) / Vert_increment)

                        Y0 = Y0 + Spacing



                        For k = 0 To Data_table_split.Rows.Count - 1
                            last_good_index = k

                            Y0 = Y0 - Spacing

                            Dim Sta1 As Double = Data_table_split.Rows(k).Item("STA1")
                            Dim Sta2 As Double = Data_table_split.Rows(k).Item("STA2")
                            Dim Label1 As String = Data_table_split.Rows(k).Item("LBL")

                            If Sta1 > Sta2 Then
                                Dim t As Double = Sta1
                                Sta1 = Sta2
                                Sta2 = t
                            End If


                            Dim Chainage_Graph_start As Double = 0
                            Dim Chainage_Graph_end As Double = Sta2 - Sta1
                            Dim Chainage_poly_start As Double = 0

                            Dim Pt1(0 To 1) As Double
                            Dim Pt2(0 To 1) As Double
                            Dim Pt3(0 To 1) As Double
                            Dim Pt4(0 To 1) As Double
                            Dim Pt5(0 To 1) As Double
                            Dim Pt6(0 To 1) As Double
                            Dim Pt7(0 To 1) As Double
                            Dim Pt8(0 To 1) As Double
                            Dim Pt9(0 To 1) As Double
                            Dim Pt10(0 To 1) As Double
                            Dim Pt11(0 To 1) As Double
                            Dim Pt12(0 To 1) As Double
                            Dim Pt13(0 To 1) As Double
                            Dim Pt14(0 To 1) As Double
                            Dim Pt15(0 To 1) As Double
                            Dim Pt16(0 To 1) As Double
                            Dim Pt17(0 To 1) As Double
                            Dim Pt18(0 To 1) As Double
                            Dim Pt19(0 To 1) As Double
                            Dim Pt20(0 To 1) As Double


                            Pt1(0) = X0
                            Pt1(1) = Y0

                            Dim Ytop As Double = Y0 + Nr_hor_Lines * Vert_increment * (Factor_print_viewport / hSCALE) * Vertical_exag

                            Dim Label_graph As New Autodesk.AutoCAD.DatabaseServices.MText
                            Label_graph.Contents = Label1
                            Label_graph.Layer = no_plot
                            Label_graph.TextHeight = 4 * Text_height
                            Label_graph.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.MiddleRight
                            Label_graph.Location = New Point3d(X0 - 5 * Text_height, (Ytop + Y0) / 2, 0)
                            BTrecord.AppendEntity(Label_graph)
                            Trans1.AddNewlyCreatedDBObject(Label_graph, True)


                            Dim Stat_prev As Double = 0
                            Dim row1 As Integer = 0

                            Dim Vert_zero_line As New Autodesk.AutoCAD.DatabaseServices.Line
                            Vert_zero_line.Layer = ComboBox_layer_grid_lines.Text
                            Vert_zero_line.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(X0, Y0, 0)
                            Vert_zero_line.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(X0, Ytop, 0)
                            BTrecord.AppendEntity(Vert_zero_line)



                            Dim MText0 As New Autodesk.AutoCAD.DatabaseServices.MText
                            MText0.Contents = Get_chainage_feet_from_double(Stat_prev, 0)
                            MText0.Layer = ComboBox_layer_text.Text
                            MText0.TextStyleId = Text_style_ID
                            MText0.TextHeight = Text_height
                            MText0.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.TopCenter
                            MText0.Location = New Autodesk.AutoCAD.Geometry.Point3d(X0, Y0 - 1.2 * Text_height, 0)
                            BTrecord.AppendEntity(MText0)
                            Trans1.AddNewlyCreatedDBObject(MText0, True)



                            Dim First_label As Double = 0


                            Dim Xend As Double = X0 + Chainage_Graph_end * (Factor_print_viewport / hSCALE)


                            Dim Vert_line_end As New Autodesk.AutoCAD.DatabaseServices.Line
                            Vert_line_end.Layer = ComboBox_layer_grid_lines.Text
                            Vert_line_end.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xend, Y0, 0)
                            Vert_line_end.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xend, Ytop, 0)
                            BTrecord.AppendEntity(Vert_line_end)
                            Trans1.AddNewlyCreatedDBObject(Vert_line_end, True)

                            If Xend > X0 Then
                                Dim Nr1 As Integer = Floor((Xend - X0) / (Horiz_increment * (Factor_print_viewport / hSCALE)))
                                If Nr1 > 0 Then
                                    Dim Label_next As Double = Ceiling((First_label + Horiz_increment) / Horiz_increment) * Horiz_increment - Horiz_increment
                                    Dim Dist1 As Double = (Label_next - First_label) * (Factor_print_viewport / hSCALE)
                                    Dim X1 As Double = X0 + Dist1
                                    For i = 0 To Nr1
                                        If X1 < Xend Then
                                            Dim Vert_line As New Autodesk.AutoCAD.DatabaseServices.Line
                                            Vert_line.Layer = ComboBox_layer_grid_lines.Text
                                            Vert_line.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(X1, Y0, 0)
                                            Vert_line.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(X1, Ytop, 0)
                                            BTrecord.AppendEntity(Vert_line)
                                            Trans1.AddNewlyCreatedDBObject(Vert_line, True)

                                            Dim MText_Station As New Autodesk.AutoCAD.DatabaseServices.MText
                                            MText_Station.Contents = Get_chainage_feet_from_double(Label_next, 0)
                                            MText_Station.Layer = ComboBox_layer_text.Text
                                            MText_Station.TextStyleId = Text_style_ID
                                            MText_Station.TextHeight = Text_height
                                            MText_Station.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.TopCenter
                                            MText_Station.Location = New Autodesk.AutoCAD.Geometry.Point3d(X1, Y0 - 1.2 * Text_height, 0)
                                            BTrecord.AppendEntity(MText_Station)
                                            Trans1.AddNewlyCreatedDBObject(MText_Station, True)
                                            Label_next = Label_next + Horiz_increment
                                            X1 = X1 + Horiz_increment * (Factor_print_viewport / hSCALE)
                                        Else
                                            Exit For
                                        End If


                                    Next
                                End If

                            End If







                            Pt3(1) = Pt1(1)

                            Pt3(0) = Pt1(0)

                            Pt4(0) = Pt1(0) + Chainage_Graph_end * (Factor_print_viewport / hSCALE)
                            Pt4(1) = Pt1(1)

                            Pt6(0) = Pt3(0) ' - 10 * (Factor_print_viewport / 1000)
                            Pt6(1) = Pt1(1) '

                            Pt7(0) = Pt4(0) ' + 10 * (Factor_print_viewport / 1000)
                            Pt7(1) = Pt1(1) '


                            Pt8(0) = Pt1(0)
                            Pt8(1) = Pt1(1) - 1.2 * Text_height

                            Pt9(1) = Pt8(1)

                            Pt10(0) = Pt6(0) - 0.5 * Text_height
                            Pt10(1) = Pt6(1)

                            Pt12(0) = Pt7(0) + 0.5 * Text_height
                            Pt12(1) = Pt7(1)

                            Pt11(0) = Pt10(0)
                            Pt13(0) = Pt12(0)














                            Dim Horiz_Zero_line As New Autodesk.AutoCAD.DatabaseServices.Line
                            Horiz_Zero_line.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Pt3(0), Pt3(1), 0)
                            Horiz_Zero_line.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Pt4(0), Pt4(1), 0)
                            Horiz_Zero_line.Layer = ComboBox_layer_grid_lines.Text
                            BTrecord.AppendEntity(Horiz_Zero_line)
                            Trans1.AddNewlyCreatedDBObject(Horiz_Zero_line, True)

                            For i = 1 To Nr_hor_Lines
                                Dim Xh1, Yh1, Xh2, Yh2 As Double
                                Xh1 = Pt3(0)
                                Xh2 = Pt4(0)

                                Yh1 = Pt3(1) + i * (Factor_print_viewport / hSCALE) * Vertical_exag * Vert_increment
                                Yh2 = Yh1
                                Dim h_off_line As New Autodesk.AutoCAD.DatabaseServices.Line
                                h_off_line.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xh1, Yh1, 0)
                                h_off_line.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xh2, Yh2, 0)
                                h_off_line.Layer = ComboBox_layer_grid_lines.Text
                                BTrecord.AppendEntity(h_off_line)
                                Trans1.AddNewlyCreatedDBObject(h_off_line, True)
                            Next












                            Dim Vstring As String
                            Vstring = Round(Grid_Z_Min, 0).ToString

                            Dim Mtext_ch_Ver_left As New Autodesk.AutoCAD.DatabaseServices.MText
                            Mtext_ch_Ver_left.Contents = Vstring
                            Mtext_ch_Ver_left.Layer = ComboBox_layer_text.Text
                            Mtext_ch_Ver_left.TextStyleId = Text_style_ID
                            Mtext_ch_Ver_left.TextHeight = Text_height

                            Mtext_ch_Ver_left.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.MiddleRight

                            Mtext_ch_Ver_left.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt10(0), Pt10(1), 0)
                            BTrecord.AppendEntity(Mtext_ch_Ver_left)
                            Trans1.AddNewlyCreatedDBObject(Mtext_ch_Ver_left, True)

                            Dim Mtext_ch_Ver_right As New Autodesk.AutoCAD.DatabaseServices.MText
                            Mtext_ch_Ver_right.Contents = Vstring
                            Mtext_ch_Ver_right.Layer = ComboBox_layer_text.Text
                            Mtext_ch_Ver_right.TextStyleId = Text_style_ID
                            Mtext_ch_Ver_right.TextHeight = Text_height
                            Mtext_ch_Ver_right.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.MiddleLeft

                            Mtext_ch_Ver_right.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt12(0), Pt12(1), 0)
                            BTrecord.AppendEntity(Mtext_ch_Ver_right)
                            Trans1.AddNewlyCreatedDBObject(Mtext_ch_Ver_right, True)

                            For i = 1 To Nr_hor_Lines

                                Pt11(1) = Pt10(1) + i * Vert_increment * (Factor_print_viewport / hSCALE) * Vertical_exag

                                Pt13(1) = Pt12(1) + i * Vert_increment * (Factor_print_viewport / hSCALE) * Vertical_exag



                                Vstring = Round((i * Vert_increment) + Grid_Z_Min, 0).ToString

                                Dim Mtext_ch_Ver_left1 As New Autodesk.AutoCAD.DatabaseServices.MText
                                Mtext_ch_Ver_left1.Contents = Vstring
                                Mtext_ch_Ver_left1.Layer = ComboBox_layer_text.Text
                                Mtext_ch_Ver_left1.TextStyleId = Text_style_ID
                                Mtext_ch_Ver_left1.TextHeight = Text_height
                                Mtext_ch_Ver_left1.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.MiddleRight

                                Mtext_ch_Ver_left1.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt11(0), Pt11(1), 0)
                                BTrecord.AppendEntity(Mtext_ch_Ver_left1)
                                Trans1.AddNewlyCreatedDBObject(Mtext_ch_Ver_left1, True)


                                Dim Mtext_ch_Ver_right1 As New Autodesk.AutoCAD.DatabaseServices.MText
                                Mtext_ch_Ver_right1.Contents = Vstring
                                Mtext_ch_Ver_right1.Layer = ComboBox_layer_text.Text
                                Mtext_ch_Ver_right1.TextStyleId = Text_style_ID
                                Mtext_ch_Ver_right1.TextHeight = Text_height
                                Mtext_ch_Ver_right1.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.MiddleLeft
                                Mtext_ch_Ver_right1.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt13(0), Pt13(1), 0)
                                BTrecord.AppendEntity(Mtext_ch_Ver_right1)
                                Trans1.AddNewlyCreatedDBObject(Mtext_ch_Ver_right1, True)


                            Next

                            Dim Profile_x_point(0 To Poly_graph.NumberOfVertices - 1) As Double
                            Dim Profile_y_point(0 To Poly_graph.NumberOfVertices - 1) As Double

                            Dim Poly_copy As New Autodesk.AutoCAD.DatabaseServices.Polyline()



                            Profile_x_point(0) = X0 + (Chainage_poly_start) * (Factor_print_viewport / hSCALE)
                            Profile_y_point(0) = Y0 + Poly_graph.StartPoint.Y - Point_min.Y



                            Poly_copy.AddVertexAt(0, New Autodesk.AutoCAD.Geometry.Point2d(Profile_x_point(0), Profile_y_point(0)), 0, 0, 0)
                            'here You can put CSF!
                            For i = 1 To Poly_graph.NumberOfVertices - 1
                                Profile_x_point(i) = Profile_x_point(i - 1) + Poly_graph.GetPoint3dAt(i).X - Poly_graph.GetPoint3dAt(i - 1).X
                                Profile_y_point(i) = Profile_y_point(i - 1) + Poly_graph.GetPoint3dAt(i).Y - Poly_graph.GetPoint3dAt(i - 1).Y
                                Poly_copy.AddVertexAt(i, New Autodesk.AutoCAD.Geometry.Point2d(Profile_x_point(i), Profile_y_point(i)), 0, 0, 0)
                            Next

                            Poly_copy.Layer = ComboBox_layer_profile_polyline.Text
                            Poly_copy.Linetype = Poly_graph.Linetype
                            Poly_copy.LineWeight = Poly_graph.LineWeight
                            Poly_copy.ColorIndex = Poly_graph.ColorIndex

                            Poly_copy.TransformBy(Matrix3d.Displacement(New Point3d(X0, Y0, 0).GetVectorTo(New Point3d(X0 - Sta1, Y0, 0))))

                            Dim Vert_line_start As New Autodesk.AutoCAD.DatabaseServices.Line

                            Vert_line_start.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(X0, Y0, 0)
                            Vert_line_start.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(X0, Ytop, 0)

                            Dim Col_start As New Point3dCollection
                            Col_start = Intersect_on_both_operands(Vert_line_start, Poly_copy)

                            Dim Col_end As New Point3dCollection
                            Col_end = Intersect_on_both_operands(Vert_line_end, Poly_copy)

                            Dim index1 As Integer = -1
                            Dim index2 As Integer = -1

                            For i = Poly_graph.NumberOfVertices - 1 To 0 Step -1
                                Dim X1 As Double = Poly_copy.GetPointAtParameter(i).X
                                If X1 > Xend Then
                                    If Poly_copy.NumberOfVertices >= 2 Then
                                        Poly_copy.RemoveVertexAt(i)
                                    Else
                                        index1 = i
                                    End If

                                End If
                                If X1 < X0 Then
                                    If Poly_copy.NumberOfVertices >= 2 Then
                                        Poly_copy.RemoveVertexAt(i)
                                    Else
                                        index2 = i
                                    End If

                                End If
                            Next

                            Poly_copy.AddVertexAt(Poly_copy.NumberOfVertices, New Autodesk.AutoCAD.Geometry.Point2d(Col_end(0).X, Col_end(0).Y), 0, 0, 0)
                            Poly_copy.AddVertexAt(0, New Autodesk.AutoCAD.Geometry.Point2d(Col_start(0).X, Col_start(0).Y), 0, 0, 0)

                            If index2 > -1 Then
                                Poly_copy.RemoveVertexAt(1)
                            End If

                            If index1 > -1 Then
                                Poly_copy.RemoveVertexAt(Poly_copy.NumberOfVertices - 2)
                            End If

                            BTrecord.AppendEntity(Poly_copy)
                            Trans1.AddNewlyCreatedDBObject(Poly_copy, True)








                        Next
                        Trans1.Commit()
                    End Using
                End Using

                Editor1.Regen()

            Catch ex As Exception
                MsgBox(ex.Message & vbCrLf & "last index = " & last_good_index)
            End Try
            Freeze_operations = False
        End If
    End Sub
End Class