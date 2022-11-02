Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry


Public Class Profiler_from_excel_Form
    Dim Profile_table As System.Data.DataTable
    Dim Profile_table2 As System.Data.DataTable
    Dim Diferenta_randuri
    Dim start1 As Double
    Dim end1 As Double
    Dim Linia_chainage_zero As Line
    Dim Polylinie_profil As Polyline
    Dim Polylinie_profil2 As Polyline
    Dim Deseneaza_poly_profil2 As Boolean = False
    Dim Am_folosit_chainage As Boolean
    Dim Am_folosit_chainage2 As Boolean
    Dim Lowest_elevation_from_load As Double
    Dim Highest_elevation_from_load As Double
    Dim Max_chainage_from_load As Double
    Dim Min_chainage_from_load As Double

    Private Sub Profiler_from_excel_Form_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Panel_chainage.Visible = False
        Panel_chainage2.Visible = False
        Button_Transfer_to_acad.Visible = False
        Panel_BLOCKS.Visible = False

        Incarca_existing_LINETYPES_to_combobox(ComboBox_LINETYPE)
        Incarca_existing_textstyles_to_combobox(ComboBox_text_styles)
        Incarca_existing_layers_to_combobox(ComboBox_layer_grid_lines)
        Incarca_existing_layers_to_combobox(ComboBox_LAYER_TEXT_AND_BLOCKS)
        Incarca_existing_layers_to_combobox(ComboBox_LAYER_PROFILE_POLYLINE)

        TextBox_color_index_grid_lines.Text = "9"

        If ComboBox_text_styles.Items.Contains("ROMANS") = True Then
            ComboBox_text_styles.Text = "ROMANS"
        End If
        If ComboBox_LINETYPE.Items.Contains("TCDASH2") = True Then
            ComboBox_LINETYPE.Text = "TCDASH2"
        End If
        If ComboBox_layer_grid_lines.Items.Contains("GRID") = True Then
            ComboBox_layer_grid_lines.Text = "GRID"
        End If
        If ComboBox_LAYER_TEXT_AND_BLOCKS.Items.Contains("TEXT") = True Then
            ComboBox_LAYER_TEXT_AND_BLOCKS.Text = "TEXT"
        End If
        If ComboBox_LAYER_PROFILE_POLYLINE.Items.Contains("PGRADE") = True Then
            ComboBox_LAYER_PROFILE_POLYLINE.Text = "PGRADE"
        End If


        Me.Width = 380
        Incarca_existing_Blocks_with_attributes_to_combobox(ComboBox_blocks)


    End Sub
    Private Sub Button_aquisition_Click(sender As System.Object, e As System.EventArgs) Handles Button_aquisition.Click
        Button_load_values.Visible = True
        Button_Transfer_to_acad.Visible = False
        TextBox_row_start.ReadOnly = False
        TextBox_row_end.ReadOnly = False
        TextBox_CHAINAGE.ReadOnly = False
        TextBox_chainage2.ReadOnly = False
        TextBox_point_name.ReadOnly = False
        TextBox_East.ReadOnly = False
        TextBox_East2.ReadOnly = False
        TextBox_NORTH.ReadOnly = False
        TextBox_NORTH2.ReadOnly = False
        TextBox_elevation.ReadOnly = False
        TextBox_elevation2.ReadOnly = False
        TextBox_description_extra.ReadOnly = False
        TextBox_Description.ReadOnly = False
        Incarca_existing_Blocks_with_attributes_to_combobox(ComboBox_blocks)
        Me.Width = 380
    End Sub

    Private Sub Label_graph_lowest_el_Click(sender As Object, e As System.EventArgs) Handles Label_graph_lowest_el.Click
        TextBox_Minimum_chainage.Text = Min_chainage_from_load
        TextBox_L_elevation.Text = Lowest_elevation_from_load
        TextBox_H_Elevation.Text = Highest_elevation_from_load
        TextBox_Maximum_chainage.Text = Max_chainage_from_load
        TextBox_row_0.Text = start1

    End Sub


    Private Sub Button_switch_EN_KP_Click(sender As Object, e As System.EventArgs) Handles Button_switch_EN_KP.Click
        If Button_load_values.Visible = True Then
            Select Case Panel_chainage.Visible
                Case True
                    Panel_chainage.Visible = False
                    Panel_E_N.Visible = True
                Case False
                    Panel_chainage.Visible = True
                    Panel_E_N.Visible = False
            End Select
        End If
    End Sub
    Private Sub Button_switch_EN_KP2_Click(sender As Object, e As System.EventArgs) Handles Button_switch_EN_KP2.Click
        If Button_load_values.Visible = True Then
            Select Case Panel_chainage2.Visible
                Case True
                    Panel_chainage2.Visible = False
                    Panel_E_N2.Visible = True
                Case False
                    Panel_chainage2.Visible = True
                    Panel_E_N2.Visible = False
            End Select
        End If
    End Sub


    Private Sub Panel_labels_for_blocks_Click(sender As Object, e As System.EventArgs)
        If ComboBox_blocks.Items.Count > 0 Then ComboBox_blocks.SelectedIndex = 0
        Incarca_existing_Blocks_with_attributes_to_combobox(ComboBox_blocks)
    End Sub

    Private Sub Panel_formating_Click(sender As Object, e As System.EventArgs) Handles Panel_formating.Click
        Incarca_existing_textstyles_to_combobox(ComboBox_text_styles)
        Incarca_existing_LINETYPES_to_combobox(ComboBox_LINETYPE)
        Incarca_existing_layers_to_combobox(ComboBox_layer_grid_lines)
        Incarca_existing_layers_to_combobox(ComboBox_LAYER_TEXT_AND_BLOCKS)
        Incarca_existing_layers_to_combobox(ComboBox_LAYER_PROFILE_POLYLINE)

        If ComboBox_text_styles.Items.Contains("ROMANS") = True Then
            ComboBox_text_styles.Text = "ROMANS"
        End If
        If ComboBox_LINETYPE.Items.Contains("TCDASH2") = True Then
            ComboBox_LINETYPE.Text = "TCDASH2"
        End If
        If ComboBox_layer_grid_lines.Items.Contains("GRID") = True Then
            ComboBox_layer_grid_lines.Text = "GRID"
        End If
        If ComboBox_LAYER_TEXT_AND_BLOCKS.Items.Contains("TEXT") = True Then
            ComboBox_LAYER_TEXT_AND_BLOCKS.Text = "TEXT"
        End If
        If ComboBox_LAYER_PROFILE_POLYLINE.Items.Contains("PGRADE") = True Then
            ComboBox_LAYER_PROFILE_POLYLINE.Text = "PGRADE"
        End If

    End Sub




    Private Sub Button_load_values_Click(sender As Object, e As System.EventArgs) Handles Button_load_values.Click


        Profile_table = New System.Data.DataTable
        Profile_table2 = New System.Data.DataTable

        If Val(TextBox_row_start.Text) < 1 Or IsNumeric(TextBox_row_start.Text) = False Then
            With TextBox_row_start
                .Text = ""
                .Focus()
            End With
            Exit Sub
        End If

        If Val(TextBox_row_end.Text) < 1 Or IsNumeric(TextBox_row_end.Text) = False Then
            With TextBox_row_end
                .Text = ""
                .Focus()
            End With
            Exit Sub
        End If

        If Val(TextBox_row_end.Text) < Val(TextBox_row_start.Text) Then
            With TextBox_row_end
                .Text = ""
                .Focus()
            End With
            Exit Sub
        End If

        Button_load_values.Visible = False
        Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Dim Col_East As Integer
            Dim Col_North As Integer

            If Panel_E_N.Visible = True Then
                Col_East = Stabileste_coloanele(TextBox_East.Text)
                Col_North = Stabileste_coloanele(TextBox_NORTH.Text)
                Am_folosit_chainage = False
            Else
                Am_folosit_chainage = True
            End If



            Dim Col_Elevation As Integer = Stabileste_coloanele(TextBox_elevation.Text)

            Dim Col_Chainage As Integer


            Col_Chainage = Stabileste_coloanele(TextBox_CHAINAGE.Text)


            Dim Col_Description As Integer = Stabileste_coloanele(TextBox_Description.Text)
            Dim Col_Description_extra As Integer = Stabileste_coloanele(TextBox_description_extra.Text)

            Dim Col_PN As Integer = Stabileste_coloanele(TextBox_point_name.Text)

            start1 = CDbl(TextBox_row_start.Text)
            end1 = CDbl(TextBox_row_end.Text)
            Diferenta_randuri = start1
            TextBox_row_start.ReadOnly = True
            TextBox_row_end.ReadOnly = True
            TextBox_CHAINAGE.ReadOnly = True
            TextBox_chainage2.ReadOnly = True
            TextBox_point_name.ReadOnly = True
            TextBox_East.ReadOnly = True
            TextBox_East2.ReadOnly = True
            TextBox_NORTH.ReadOnly = True
            TextBox_NORTH2.ReadOnly = True
            TextBox_elevation.ReadOnly = True
            TextBox_elevation2.ReadOnly = True
            TextBox_description_extra.ReadOnly = True
            TextBox_Description.ReadOnly = True

            Profile_table = Load_Excel_to_data_table(start1, end1, Col_PN, Col_East, Col_North, Col_Elevation, Col_Chainage, Col_Description, Col_Description_extra)


            If Profile_table.Rows.Count > 0 Then



                Dim Lowest_elev As Double
                Dim Highest_elev As Double
                Dim Total_chainage As Double
                Dim Total_chainage2 As Double

                If Am_folosit_chainage = False Then
                    Profile_table.Rows.Item(0).Item("Chainage") = 0
                End If



                If IsDBNull(Profile_table.Rows.Item(0).Item("Elevation")) = False Then
                    Lowest_elev = Profile_table.Rows.Item(0).Item("Elevation")
                    Highest_elev = Profile_table.Rows.Item(0).Item("Elevation")
                Else
                    Button_load_values.Visible = True
                    TextBox_row_start.ReadOnly = False
                    TextBox_row_end.ReadOnly = False
                    MsgBox("Elevation not specified on row " & Diferenta_randuri)
                    TextBox_L_elevation.Text = ""
                    TextBox_H_Elevation.Text = ""
                    TextBox_Maximum_chainage.Text = ""
                    TextBox_row_0.Text = ""

                    TextBox_row_start.ReadOnly = False
                    TextBox_row_end.ReadOnly = False
                    TextBox_CHAINAGE.ReadOnly = False
                    TextBox_chainage2.ReadOnly = False
                    TextBox_point_name.ReadOnly = False
                    TextBox_East.ReadOnly = False
                    TextBox_East2.ReadOnly = False
                    TextBox_NORTH.ReadOnly = False
                    TextBox_NORTH2.ReadOnly = False
                    TextBox_elevation.ReadOnly = False
                    TextBox_elevation2.ReadOnly = False
                    TextBox_description_extra.ReadOnly = False
                    TextBox_Description.ReadOnly = False
                    Exit Sub

                End If



                For i = 1 To Profile_table.Rows.Count - 1

                    If Am_folosit_chainage = False Then
                        If IsDBNull(Profile_table.Rows.Item(i).Item("East")) = False And IsDBNull(Profile_table.Rows.Item(i).Item("North")) = False Then
                            Dim x1, y1, x2, y2 As Double
                            If IsDBNull(Profile_table.Rows.Item(i - 1).Item("East")) = False Then x1 = Profile_table.Rows.Item(i - 1).Item("East")
                            If IsDBNull(Profile_table.Rows.Item(i - 1).Item("North")) = False Then y1 = Profile_table.Rows.Item(i - 1).Item("North")
                            x2 = Profile_table.Rows.Item(i).Item("East")
                            y2 = Profile_table.Rows.Item(i).Item("North")
                            Total_chainage = Total_chainage + ((x1 - x2) ^ 2 + (y1 - y2) ^ 2) ^ 0.5
                            Profile_table.Rows.Item(i).Item("Chainage") = Total_chainage
                        Else
                            Button_load_values.Visible = True
                            TextBox_row_start.ReadOnly = False
                            TextBox_row_end.ReadOnly = False
                            MsgBox("X/Y not specified on row " & i + Diferenta_randuri)
                            TextBox_L_elevation.Text = ""
                            TextBox_H_Elevation.Text = ""
                            TextBox_Maximum_chainage.Text = ""
                            TextBox_row_0.Text = ""
                            TextBox_row_start.ReadOnly = False
                            TextBox_row_end.ReadOnly = False
                            TextBox_CHAINAGE.ReadOnly = False
                            TextBox_chainage2.ReadOnly = False
                            TextBox_point_name.ReadOnly = False
                            TextBox_East.ReadOnly = False
                            TextBox_East2.ReadOnly = False
                            TextBox_NORTH.ReadOnly = False
                            TextBox_NORTH2.ReadOnly = False
                            TextBox_elevation.ReadOnly = False
                            TextBox_elevation2.ReadOnly = False
                            TextBox_description_extra.ReadOnly = False
                            TextBox_Description.ReadOnly = False
                            Exit Sub
                        End If
                    End If



                    If IsDBNull(Profile_table.Rows.Item(i).Item("Elevation")) = False Then
                        Dim z2 As Double = Profile_table.Rows.Item(i).Item("Elevation")
                        If z2 < Lowest_elev Then Lowest_elev = z2
                        If z2 > Highest_elev Then Highest_elev = z2
                    Else
                        Button_load_values.Visible = True
                        TextBox_row_start.ReadOnly = False
                        TextBox_row_end.ReadOnly = False
                        MsgBox("Z not specified on row " & i + Diferenta_randuri)
                        TextBox_L_elevation.Text = ""
                        TextBox_H_Elevation.Text = ""
                        TextBox_Maximum_chainage.Text = ""
                        TextBox_row_0.Text = ""
                        TextBox_row_start.ReadOnly = False
                        TextBox_row_end.ReadOnly = False
                        TextBox_CHAINAGE.ReadOnly = False
                        TextBox_chainage2.ReadOnly = False
                        TextBox_point_name.ReadOnly = False
                        TextBox_East.ReadOnly = False
                        TextBox_East2.ReadOnly = False
                        TextBox_NORTH.ReadOnly = False
                        TextBox_NORTH2.ReadOnly = False
                        TextBox_elevation.ReadOnly = False
                        TextBox_elevation2.ReadOnly = False
                        TextBox_description_extra.ReadOnly = False
                        TextBox_Description.ReadOnly = False
                        Exit Sub
                    End If

                    If IsDBNull(Profile_table.Rows.Item(i).Item("Chainage")) = True Then
                        Button_load_values.Visible = True
                        TextBox_row_start.ReadOnly = False
                        TextBox_row_end.ReadOnly = False
                        MsgBox("Chainage not specified on row " & i + Diferenta_randuri)
                        TextBox_L_elevation.Text = ""
                        TextBox_H_Elevation.Text = ""
                        TextBox_Maximum_chainage.Text = ""
                        TextBox_row_0.Text = ""
                        TextBox_row_start.ReadOnly = False
                        TextBox_row_end.ReadOnly = False
                        TextBox_CHAINAGE.ReadOnly = False
                        TextBox_chainage2.ReadOnly = False
                        TextBox_point_name.ReadOnly = False
                        TextBox_East.ReadOnly = False
                        TextBox_East2.ReadOnly = False
                        TextBox_NORTH.ReadOnly = False
                        TextBox_NORTH2.ReadOnly = False
                        TextBox_elevation.ReadOnly = False
                        TextBox_elevation2.ReadOnly = False
                        TextBox_description_extra.ReadOnly = False
                        TextBox_Description.ReadOnly = False
                        Exit Sub
                    End If

                Next



                If Am_folosit_chainage = True Then
                    Total_chainage = Profile_table.Rows.Item(Profile_table.Rows.Count - 1).Item("Chainage")
                    TextBox_Minimum_chainage.Text = Profile_table.Rows.Item(0).Item("Chainage")
                Else
                    TextBox_Minimum_chainage.Text = 0
                End If

                TextBox_L_elevation.Text = Round(Lowest_elev, 0) - 1
                TextBox_H_Elevation.Text = Round(Highest_elev, 0) + 1
                TextBox_Maximum_chainage.Text = Round(Total_chainage, 0)
                TextBox_row_0.Text = start1

                Lowest_elevation_from_load = TextBox_L_elevation.Text
                Highest_elevation_from_load = TextBox_H_Elevation.Text
                Max_chainage_from_load = TextBox_Maximum_chainage.Text
                Min_chainage_from_load = TextBox_Minimum_chainage.Text

                '***************************GRAPH2


                Dim Col_East2 As Integer
                Dim Col_North2 As Integer

                If Panel_E_N2.Visible = True Then
                    Col_East2 = Stabileste_coloanele(TextBox_East2.Text)
                    Col_North2 = Stabileste_coloanele(TextBox_NORTH2.Text)
                    Am_folosit_chainage2 = False
                Else
                    Am_folosit_chainage2 = True
                End If

                Dim Col_Elevation2 As Integer = Stabileste_coloanele(TextBox_elevation2.Text)

                Dim Col_Chainage2 As Integer
                Col_Chainage2 = Stabileste_coloanele(TextBox_chainage2.Text)

                If Col_Elevation2 > 0 Then
                    If Am_folosit_chainage2 = False Then
                        If Col_East2 > 0 Then
                            If Col_North2 > 0 Then
                                Deseneaza_poly_profil2 = True
                            End If
                        End If
                    Else
                        If Col_Chainage2 > 0 Then
                            Deseneaza_poly_profil2 = True
                        End If
                    End If
                End If

                If Deseneaza_poly_profil2 = True Then
                    Dim Col_Description2 As Integer
                    Dim Col_Description_extra2 As Integer
                    Dim Col_PN2 As Integer
                    Profile_table2 = Load_Excel_to_data_table(start1, end1, Col_PN2, Col_East2, Col_North2, Col_Elevation2, Col_Chainage2, Col_Description2, Col_Description_extra2)

                    If Am_folosit_chainage2 = False Then
                        Profile_table2.Rows.Item(0).Item("Chainage") = 0
                    End If

                    For i = 1 To Profile_table2.Rows.Count - 1

                        If Am_folosit_chainage2 = False Then
                            If IsDBNull(Profile_table2.Rows.Item(i).Item("East")) = False And IsDBNull(Profile_table2.Rows.Item(i).Item("North")) = False Then
                                Dim x1, y1, x2, y2 As Double
                                If IsDBNull(Profile_table2.Rows.Item(i - 1).Item("East")) = False Then x1 = Profile_table2.Rows.Item(i - 1).Item("East")
                                If IsDBNull(Profile_table2.Rows.Item(i - 1).Item("North")) = False Then y1 = Profile_table2.Rows.Item(i - 1).Item("North")
                                x2 = Profile_table2.Rows.Item(i).Item("East")
                                y2 = Profile_table2.Rows.Item(i).Item("North")
                                Total_chainage2 = Total_chainage2 + ((x1 - x2) ^ 2 + (y1 - y2) ^ 2) ^ 0.5
                                Profile_table2.Rows.Item(i).Item("Chainage") = Total_chainage2
                            Else
                                Button_load_values.Visible = True
                                TextBox_row_start.ReadOnly = False
                                TextBox_row_end.ReadOnly = False
                                MsgBox("X/Y not specified on row " & i + Diferenta_randuri)
                                TextBox_L_elevation.Text = ""
                                TextBox_H_Elevation.Text = ""
                                TextBox_Maximum_chainage.Text = ""
                                TextBox_row_0.Text = ""
                                TextBox_row_start.ReadOnly = False
                                TextBox_row_end.ReadOnly = False
                                TextBox_CHAINAGE.ReadOnly = False
                                TextBox_chainage2.ReadOnly = False
                                TextBox_point_name.ReadOnly = False
                                TextBox_East.ReadOnly = False
                                TextBox_East2.ReadOnly = False
                                TextBox_NORTH.ReadOnly = False
                                TextBox_NORTH2.ReadOnly = False
                                TextBox_elevation.ReadOnly = False
                                TextBox_elevation2.ReadOnly = False
                                TextBox_description_extra.ReadOnly = False
                                TextBox_Description.ReadOnly = False
                                Exit Sub
                            End If
                        End If



                        If IsDBNull(Profile_table2.Rows.Item(i).Item("Elevation")) = False Then
                            Dim z2 As Double = Profile_table2.Rows.Item(i).Item("Elevation")
                            If z2 < Lowest_elev Then Lowest_elev = z2
                            If z2 > Highest_elev Then Highest_elev = z2
                        Else
                            Button_load_values.Visible = True
                            TextBox_row_start.ReadOnly = False
                            TextBox_row_end.ReadOnly = False
                            MsgBox("Z not specified on row " & i + Diferenta_randuri)
                            TextBox_L_elevation.Text = ""
                            TextBox_H_Elevation.Text = ""
                            TextBox_Maximum_chainage.Text = ""
                            TextBox_row_0.Text = ""
                            TextBox_row_start.ReadOnly = False
                            TextBox_row_end.ReadOnly = False
                            TextBox_CHAINAGE.ReadOnly = False
                            TextBox_chainage2.ReadOnly = False
                            TextBox_point_name.ReadOnly = False
                            TextBox_East.ReadOnly = False
                            TextBox_East2.ReadOnly = False
                            TextBox_NORTH.ReadOnly = False
                            TextBox_NORTH2.ReadOnly = False
                            TextBox_elevation.ReadOnly = False
                            TextBox_elevation2.ReadOnly = False
                            TextBox_description_extra.ReadOnly = False
                            TextBox_Description.ReadOnly = False
                            Exit Sub
                        End If

                        If IsDBNull(Profile_table2.Rows.Item(i).Item("Chainage")) = True Then
                            Button_load_values.Visible = True
                            TextBox_row_start.ReadOnly = False
                            TextBox_row_end.ReadOnly = False
                            MsgBox("Chainage not specified on row " & i + Diferenta_randuri)
                            TextBox_L_elevation.Text = ""
                            TextBox_H_Elevation.Text = ""
                            TextBox_Maximum_chainage.Text = ""
                            TextBox_row_0.Text = ""
                            TextBox_row_start.ReadOnly = False
                            TextBox_row_end.ReadOnly = False
                            TextBox_CHAINAGE.ReadOnly = False
                            TextBox_chainage2.ReadOnly = False
                            TextBox_point_name.ReadOnly = False
                            TextBox_East.ReadOnly = False
                            TextBox_East2.ReadOnly = False
                            TextBox_NORTH.ReadOnly = False
                            TextBox_NORTH2.ReadOnly = False
                            TextBox_elevation.ReadOnly = False
                            TextBox_elevation2.ReadOnly = False
                            TextBox_description_extra.ReadOnly = False
                            TextBox_Description.ReadOnly = False
                            Exit Sub
                        End If

                    Next

                End If

            End If

        End Using

        Button_Transfer_to_acad.Visible = True
        If Panel_E_N.Visible = True Then Panel_DRAW_POLYLINES.Visible = True
        Me.Width = 1000
    End Sub

    Public Function Load_Excel_prin_copy_paste_to_data_table(ByVal Start1 As Double, ByVal End1 As Double, ByVal Col_pn As Integer, ByVal Col_East As Integer, ByVal Col_North As Integer, ByVal Col_Elevation As Integer, _
                                                             ByVal Col_chainage As Integer, ByVal Col_Description As Integer, ByVal Col_Description_extra As Integer) As System.Data.DataTable

        Try
            Dim Table_data1 As New System.Data.DataTable
            Table_data1.Columns.Add("PN", GetType(String))
            Table_data1.Columns.Add("East", GetType(String))
            Table_data1.Columns.Add("North", GetType(String))
            Table_data1.Columns.Add("Elevation", GetType(String))
            Table_data1.Columns.Add("Chainage", GetType(String))
            Table_data1.Columns.Add("Description", GetType(String))
            Table_data1.Columns.Add("Description_extra", GetType(String))
            Table_data1.Columns.Add("Chainage_Recalculated", GetType(String))

            Dim Path_to_desktop As String = Environment.GetFolderPath(System.Environment.SpecialFolder.DesktopDirectory)
            Dim Nume_fisier_csv As String = Path_to_desktop & "\profiler_excel.csv"


            Dim Excel1 As Microsoft.Office.Interop.Excel.Application
            Excel1 = CType(System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)

            Excel1.DisplayAlerts = False
            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet

            W1 = Get_active_worksheet_from_Excel()

            W1.Rows(Start1 & ":" & End1).Copy()
            Excel1.Workbooks.Add()
            Excel1.ActiveSheet.Paste()
            W1 = Get_active_worksheet_from_Excel()
            W1.Cells.MergeCells = False
            Dim Sterge As Boolean = True

            For i = 1 To 100
                If i > Col_pn And _
                      i > Col_East And _
                      i > Col_North And _
                      i > Col_Elevation And _
                      i > Col_chainage And _
                      i > Col_Description And _
                      i > Col_Description_extra Then
                    Sterge = False
                Else
                    Sterge = True
                End If

                If i = Col_pn Or _
                      i = Col_East Or _
                      i = Col_North Or _
                      i = Col_Elevation Or _
                      i = Col_chainage Or _
                      i = Col_Description Or _
                      i = Col_Description_extra Then
                    Sterge = False
                End If

                If Sterge = True Then
                    If Not i = Col_pn And _
                        Not i = Col_East And _
                        Not i = Col_North And _
                        Not i = Col_Elevation And _
                        Not i = Col_chainage And _
                        Not i = Col_Description And _
                        Not i = Col_Description_extra  Then

                        If i < Col_pn Then
                            Col_pn = Col_pn - 1
                        End If
                        If i < Col_East Then
                            Col_East = Col_East - 1
                        End If
                        If i < Col_North Then
                            Col_North = Col_North - 1
                        End If
                        If i < Col_Elevation Then
                            Col_Elevation = Col_Elevation - 1
                        End If
                        If i < Col_chainage Then
                            Col_chainage = Col_chainage - 1
                        End If
                        If i < Col_Description Then
                            Col_Description = Col_Description - 1
                        End If
                        If i < Col_Description_extra Then
                            Col_Description_extra = Col_Description_extra - 1
                        End If

                        W1.Columns(i).delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftToLeft)
                        i = i - 1
                    End If
                End If
            Next


            W1.SaveAs(Nume_fisier_csv, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV)
            Excel1.ActiveWindow.Close()

            If System.IO.File.Exists(Nume_fisier_csv) = True Then

                Using Reader1 As New System.IO.StreamReader(Nume_fisier_csv)
                    Dim Line1 As String
                    'Dim j As Integer = 0

                    Dim Index1 As Integer = 1

                    While Reader1.Peek > 0
                        'Dim Sec1 As Integer = Environment.TickCount


                        Line1 = Reader1.ReadLine

                        Table_data1.Rows.Add()

                        If InStr(Line1, ",") > 0 Then


                            Dim Bucati_linie() As String
                            Bucati_linie = Split(Line1, ",")

                            If Col_pn > 0 And Bucati_linie.Length >= Col_pn Then
                                If Not Bucati_linie(Col_pn - 1) = "" Then
                                    Table_data1.Rows.Item(Index1 - 1).Item("PN") = Bucati_linie(Col_pn - 1)
                                End If
                            End If

                            If Col_East > 0 And Bucati_linie.Length >= Col_East Then
                                If IsNumeric(Bucati_linie(Col_East - 1)) = True Then
                                    Table_data1.Rows.Item(Index1 - 1).Item("East") = Bucati_linie(Col_East - 1)
                                End If
                            End If

                            If Col_North > 0 And Bucati_linie.Length >= Col_North Then
                                If IsNumeric(Bucati_linie(Col_North - 1)) = True Then
                                    Table_data1.Rows.Item(Index1 - 1).Item("North") = Bucati_linie(Col_North - 1)
                                End If
                            End If

                            If Col_Elevation > 0 And Bucati_linie.Length >= Col_Elevation Then
                                If IsNumeric(Bucati_linie(Col_Elevation - 1)) = True Then
                                    Table_data1.Rows.Item(Index1 - 1).Item("Elevation") = Bucati_linie(Col_Elevation - 1)
                                End If
                            End If


                            If Col_chainage > 0 And Bucati_linie.Length >= Col_chainage Then
                                Dim String_ch As String = Replace(Bucati_linie(Col_chainage - 1), "+", "")
                                If IsNumeric(String_ch) = True Then
                                    Table_data1.Rows.Item(Index1 - 1).Item("Chainage") = String_ch
                                End If
                            End If

                            If Col_Description > 0 And Bucati_linie.Length >= Col_Description Then
                                If Not Bucati_linie(Col_Description - 1) = "" Then
                                    Table_data1.Rows.Item(Index1 - 1).Item("Description") = Bucati_linie(Col_Description - 1)
                                End If
                            End If

                            If Col_Description_extra > 0 And Bucati_linie.Length >= Col_Description_extra Then
                                If Not Bucati_linie(Col_Description_extra - 1) = "" Then
                                    Table_data1.Rows.Item(Index1 - 1).Item("Description_extra") = Bucati_linie(Col_Description_extra - 1)
                                End If
                            End If
                            Index1 = Index1 + 1

                            ' asta e de la If j >= start1 And j <= end1
                        End If



                        ' ASTA E DE LA If InStr(Line1, ",") > 0 

                        'Dim Sec2 As Integer = Environment.TickCount

                        'MsgBox(Sec2 - Sec1)
                        'asta e de la reader.peek >0
                    End While

                    'asta e de la reader
                End Using

                System.IO.File.Delete(Nume_fisier_csv)
                Excel1.DisplayAlerts = True
                'asta e de la fisierul exista

                'MsgBox(Col_chainage & vbCrLf & Col_Elevation)
                Return Table_data1

            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    Public Function Load_Excel_to_data_table(ByVal Start1 As Double, ByVal End1 As Double, ByVal Col_pn As Integer, ByVal Col_East As Integer, ByVal Col_North As Integer, ByVal Col_Elevation As Integer, _
                                                             ByVal Col_chainage As Integer, ByVal Col_Description As Integer, ByVal Col_Description_extra As Integer) As System.Data.DataTable

        Try

            Dim Table_data1 As New System.Data.DataTable

            Table_data1.Columns.Add("PN", GetType(String))
            Table_data1.Columns.Add("East", GetType(String))
            Table_data1.Columns.Add("North", GetType(String))
            Table_data1.Columns.Add("Elevation", GetType(String))
            Table_data1.Columns.Add("Chainage", GetType(String))
            Table_data1.Columns.Add("Description", GetType(String))
            Table_data1.Columns.Add("Description_extra", GetType(String))
            Table_data1.Columns.Add("Chainage_Recalculated", GetType(String))


            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet

            W1 = Get_active_worksheet_from_Excel()


            Dim Index_data_table As Integer = 0

            For i = Start1 To End1

                Table_data1.Rows.Add()




                If Col_PN > 0 Then
                    Table_data1.Rows.Item(Index_data_table).Item("PN") = W1.Cells(i, Col_PN).value2
                End If

                If Col_East > 0 Then
                    Dim Cell_value As String = W1.Cells(i, Col_East).value2

                    If IsNumeric(Cell_value) = True Then
                        Table_data1.Rows.Item(Index_data_table).Item("East") = Cell_value
                    End If
                End If

                If Col_North > 0 Then
                    Dim Cell_value As String = W1.Cells(i, Col_North).value2

                    If IsNumeric(Cell_value) = True Then
                        Table_data1.Rows.Item(Index_data_table).Item("North") = Cell_value
                    End If
                End If

                If Col_Elevation > 0 Then
                    Dim Cell_value As String = W1.Cells(i, Col_Elevation).value2

                    If IsNumeric(Cell_value) = True Then
                        Table_data1.Rows.Item(Index_data_table).Item("Elevation") = Cell_value
                    End If
                Else
                    Table_data1.Rows.Item(Index_data_table).Item("Elevation") = "0"
                End If


                If Col_chainage > 0 Then
                    Dim Cell_value As String = Replace(W1.Cells(i, Col_chainage).value2, "+", "")
                    If IsNumeric(Cell_value) = True Then
                        Table_data1.Rows.Item(Index_data_table).Item("Chainage") = Cell_value
                    End If
                End If

                If Col_Description > 0 Then
                    Dim Cell_value As String = W1.Cells(i, Col_Description).value2

                    If Not Cell_value = "" Then
                        Table_data1.Rows.Item(Index_data_table).Item("Description") = Cell_value
                    End If
                End If

                If Col_Description_extra > 0 Then
                    Dim Cell_value As String = W1.Cells(i, Col_Description_extra).value2
                    If Not Cell_value = "" Then
                        Table_data1.Rows.Item(Index_data_table).Item("Description_extra") = Cell_value
                    End If
                End If



               


                Index_data_table = Index_data_table + 1

            Next

            Return Table_data1


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function



    Private Sub Button_Transfer_to_acad_Click(sender As Object, e As System.EventArgs) Handles Button_Transfer_to_acad.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
        ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                Dim Grid_Z_Max, Grid_Z_Min As Double
                If IsNumeric(TextBox_L_elevation.Text) = True Then
                    Grid_Z_Min = CDbl(TextBox_L_elevation.Text)
                Else
                    MsgBox("Non numeric lowest elevation")
                    TextBox_row_start.ReadOnly = False
                    TextBox_row_end.ReadOnly = False
                    Exit Sub
                End If

                If IsNumeric(TextBox_H_Elevation.Text) = True Then
                    Grid_Z_Max = CDbl(TextBox_H_Elevation.Text)
                Else
                    MsgBox("Non numeric highest elevation")
                    TextBox_row_start.ReadOnly = False
                    TextBox_row_end.ReadOnly = False
                    Exit Sub
                End If

                Dim CSF As Double = 1

                Dim Printing_scale, Viewport_scale As Double
                If IsNumeric(TextBox_printing_scale.Text) = True Then
                    Printing_scale = CDbl(TextBox_printing_scale.Text)
                Else
                    MsgBox("Non numeric printing scale")
                    TextBox_row_start.ReadOnly = False
                    TextBox_row_end.ReadOnly = False
                    Exit Sub
                End If

                If IsNumeric(TextBox_viewport_scale.Text) = True Then
                    Viewport_scale = CDbl(TextBox_viewport_scale.Text)
                Else
                    MsgBox("Non numeric viewport scale")
                    TextBox_row_start.ReadOnly = False
                    TextBox_row_end.ReadOnly = False
                    Exit Sub
                End If

                Dim Color_index_grid As Integer

                If IsNumeric(TextBox_color_index_grid_lines.Text) = True Then
                    Color_index_grid = CInt(TextBox_color_index_grid_lines.Text)
                Else
                    MsgBox("Non numeric grid color")
                    TextBox_row_start.ReadOnly = False
                    TextBox_row_end.ReadOnly = False
                    Exit Sub
                End If
                If ComboBox_LAYER_TEXT_AND_BLOCKS.Text = "" Or ComboBox_layer_grid_lines.Text = "" Or ComboBox_LAYER_PROFILE_POLYLINE.Text = "" Or TextBox_Layer_NO_PLOT.Text = "" Then
                    MsgBox("Layer not specified")
                    TextBox_row_start.ReadOnly = False
                    TextBox_row_end.ReadOnly = False
                    Exit Sub
                End If


                Dim Total_chainage_grid As Double
                If IsNumeric(TextBox_Maximum_chainage.Text) = True And IsNumeric(TextBox_Minimum_chainage.Text) = True Then
                    Total_chainage_grid = CDbl(TextBox_Maximum_chainage.Text) - CDbl(TextBox_Minimum_chainage.Text)
                Else
                    MsgBox("Non numeric total Chainage")
                    TextBox_row_start.ReadOnly = False
                    TextBox_row_end.ReadOnly = False
                    Exit Sub
                End If


                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction



                    Dim Text_style_table As Autodesk.AutoCAD.DatabaseServices.TextStyleTable = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                    Dim Text_style_ID As Autodesk.AutoCAD.DatabaseServices.ObjectId = Text_style_table.Item(ComboBox_text_styles.Text)


                    Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)



1234:

                    Dim Row_zero As Double

                    If IsNumeric(TextBox_row_0.Text) = True Then
                        Row_zero = CDbl(TextBox_row_0.Text)
                    End If

                    If Row_zero < 1 Then
                        MsgBox("Non valid zero chainage position")
                        TextBox_row_start.ReadOnly = False
                        TextBox_row_end.ReadOnly = False
                        Exit Sub
                    End If

                    If Row_zero < start1 Or Row_zero > end1 Then
                        MsgBox("Non valid zero chainage position")
                        TextBox_row_start.ReadOnly = False
                        TextBox_row_end.ReadOnly = False
                        Exit Sub
                    End If


                    Dim Delta_Chainage As Double = 0
                    Dim Delta_Chainage2 As Double = 0
                    Dim Nou_zero As Boolean = False
                    If Row_zero > start1 Then
                        Delta_Chainage = Profile_table.Rows.Item(Row_zero - start1).Item("Chainage")
                        If Deseneaza_poly_profil2 = True Then
                            Delta_Chainage2 = Profile_table2.Rows.Item(Row_zero - start1).Item("Chainage")
                        End If
                        Nou_zero = True
                    End If

                    Dim Poly2 As New Autodesk.AutoCAD.DatabaseServices.Polyline
                    Dim Poly3 As New Autodesk.AutoCAD.DatabaseServices.Polyline3d
                    Dim Colectie_puncte_3d As New Point3dCollection

                    If Panel_DRAW_POLYLINES.Visible = True Then
                        If CheckBox_draw_3Dpoly.Checked = True Then
                            BTrecord.AppendEntity(Poly3)
                            Trans1.AddNewlyCreatedDBObject(Poly3, True)
                        End If
                    End If


                    Creaza_layer(TextBox_Layer_NO_PLOT.Text, 40, "No Plot", False)

                    Dim Leader_text_start, Leader_text_end As String

                    If Am_folosit_chainage = False Then
                        For i = 0 To Profile_table.Rows.Count - 1
                            If IsDBNull(Profile_table.Rows.Item(i).Item("East")) = False And IsDBNull(Profile_table.Rows.Item(i).Item("North")) = False And IsDBNull(Profile_table.Rows.Item(i).Item("Chainage")) = False Then
                                Profile_table.Rows.Item(i).Item("Chainage_Recalculated") = Profile_table.Rows.Item(i).Item("Chainage") - Delta_Chainage

                                If Panel_DRAW_POLYLINES.Visible = True Then

                                    Dim Descriptie_punct As String = ""

                                    If IsDBNull(Profile_table.Rows.Item(i).Item("Description")) = False Then
                                        Descriptie_punct = Profile_table.Rows.Item(i).Item("Description")
                                    End If
                                    If IsDBNull(Profile_table.Rows.Item(i).Item("Elevation")) = False Then
                                        If Descriptie_punct = "" Then
                                            Descriptie_punct = "ELEV = " & Profile_table.Rows.Item(i).Item("Elevation")
                                        Else
                                            Descriptie_punct = Descriptie_punct & vbCrLf & "ELEV = " & Profile_table.Rows.Item(i).Item("Elevation")
                                        End If
                                    End If

                                    If IsDBNull(Profile_table.Rows.Item(i).Item("PN")) = False Then
                                        If Descriptie_punct = "" Then
                                            Descriptie_punct = Profile_table.Rows.Item(i).Item("PN")
                                        Else
                                            Descriptie_punct = Descriptie_punct & vbCrLf & Profile_table.Rows.Item(i).Item("PN")
                                        End If
                                    End If

                                    If i = 0 Then Leader_text_start = Descriptie_punct
                                    If i = Profile_table.Rows.Count - 1 Then Leader_text_end = Descriptie_punct


                                    If CheckBox_draw_2Dpoly.Checked = True Then
                                        Poly2.AddVertexAt(i, New Point2d(Profile_table.Rows.Item(i).Item("East"), Profile_table.Rows.Item(i).Item("North")), 0, 0, 0)


                                    End If


                                    If CheckBox_draw_3Dpoly.Checked = True Then
                                        Dim Vertex_poly_3d As New PolylineVertex3d(New Point3d(Profile_table.Rows.Item(i).Item("East"), Profile_table.Rows.Item(i).Item("North"), Profile_table.Rows.Item(i).Item("Elevation")))
                                        Poly3.AppendVertex(Vertex_poly_3d)
                                        Trans1.AddNewlyCreatedDBObject(Vertex_poly_3d, True)

                                    End If
                                End If
                            End If

                            '***** PROFILE 2
                            If Deseneaza_poly_profil2 = True Then
                                If Am_folosit_chainage2 = False Then
                                    If IsDBNull(Profile_table2.Rows.Item(i).Item("East")) = False And IsDBNull(Profile_table2.Rows.Item(i).Item("North")) = False And IsDBNull(Profile_table2.Rows.Item(i).Item("Chainage")) = False Then
                                        Profile_table2.Rows.Item(i).Item("Chainage_Recalculated") = Profile_table2.Rows.Item(i).Item("Chainage") - Delta_Chainage2
                                    End If
                                Else
                                    If IsDBNull(Profile_table2.Rows.Item(i).Item("Chainage")) = False Then
                                        Profile_table2.Rows.Item(i).Item("Chainage_Recalculated") = Profile_table2.Rows.Item(i).Item("Chainage")
                                    End If
                                End If
                            End If
                        Next
                    End If

                    If Am_folosit_chainage = True Then
                        Nou_zero = True
                        For i = 0 To Profile_table.Rows.Count - 1
                            If IsDBNull(Profile_table.Rows.Item(i).Item("Chainage")) = False Then
                                Profile_table.Rows.Item(i).Item("Chainage_Recalculated") = Profile_table.Rows.Item(i).Item("Chainage")
                            End If

                            '***** PROFILE 2
                            If Deseneaza_poly_profil2 = True Then
                                If Am_folosit_chainage2 = True Then
                                    If IsDBNull(Profile_table2.Rows.Item(i).Item("Chainage")) = False Then
                                        Profile_table2.Rows.Item(i).Item("Chainage_Recalculated") = Profile_table2.Rows.Item(i).Item("Chainage")
                                    End If
                                Else
                                    If IsDBNull(Profile_table2.Rows.Item(i).Item("East")) = False And IsDBNull(Profile_table2.Rows.Item(i).Item("North")) = False And IsDBNull(Profile_table2.Rows.Item(i).Item("Chainage")) = False Then
                                        Profile_table2.Rows.Item(i).Item("Chainage_Recalculated") = Profile_table2.Rows.Item(i).Item("Chainage") - Delta_Chainage2
                                    End If
                                End If
                            End If
                        Next
                    End If

                    If Panel_DRAW_POLYLINES.Visible = True And Am_folosit_chainage = False Then
                        If CheckBox_draw_2Dpoly.Checked = True Then
                            BTrecord.AppendEntity(Poly2)
                            Trans1.AddNewlyCreatedDBObject(Poly2, True)
                        End If
                    End If


                    Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                    Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Please pick the grid location:")

                    PP1.AllowNone = False
                    Point1 = Editor1.GetPoint(PP1)
                    If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                        Exit Sub
                    End If



                    Dim Xg0, Yg0 As Double
                    Xg0 = Point1.Value.X
                    Yg0 = Point1.Value.Y


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



                    Dim Nr_hor_Lines As Double
                    Dim Horiz_increment, Vert_increment As Double
                    Dim hSCALE, vSCALE, Vertical_exag As Double


                    If IsNumeric(TextBox_Hincr.Text) = True Then
                        Horiz_increment = CDbl(TextBox_Hincr.Text)
                    Else
                        MsgBox("Non numeric horizontal increment")
                        Exit Sub
                    End If

                    If IsNumeric(TextBox_Vincr.Text) = True Then
                        Vert_increment = CDbl(TextBox_Vincr.Text)
                    Else
                        MsgBox("Non numeric vertical increment")
                        Exit Sub
                    End If
                    Dim Nr_vert_Lines_stanga, Nr_vert_Lines_dreapta, Chainage_poly_start As Double
                    Dim Chainage_Graph_start, Chainage_Graph_end As Double

                    If IsNumeric(TextBox_Minimum_chainage.Text) = True Then
                        Chainage_Graph_start = CDbl(TextBox_Minimum_chainage.Text)
                    Else
                        MsgBox("Non numeric Chainage start Graph")
                        Exit Sub
                    End If

                    If IsNumeric(TextBox_Maximum_chainage.Text) = True Then
                        Chainage_Graph_end = CDbl(TextBox_Maximum_chainage.Text)
                    Else
                        MsgBox("Non numeric Chainage end Graph")
                        Exit Sub
                    End If

                    Chainage_poly_start = Profile_table.Rows.Item(0).Item("Chainage_Recalculated") / CSF



                    If Horiz_increment < Abs(Chainage_Graph_end) Then
                        Nr_vert_Lines_dreapta = Ceiling(Abs(Chainage_Graph_end) / Horiz_increment)
                    Else
                        Nr_vert_Lines_dreapta = 1
                    End If


                    If Horiz_increment < Abs(Chainage_Graph_start) Then
                        Nr_vert_Lines_stanga = Ceiling(Abs(Chainage_Graph_start) / Horiz_increment)
                    Else
                        If Not Chainage_Graph_start = 0 Then
                            Nr_vert_Lines_stanga = 1
                        Else
                            Nr_vert_Lines_stanga = 0
                        End If

                    End If


                    Nr_hor_Lines = Ceiling((Grid_Z_Max - Grid_Z_Min) / Vert_increment)


                    Pt1(0) = Xg0
                    Pt1(1) = Yg0



                    Dim Factor_print_viewport As Double
                    Factor_print_viewport = Printing_scale / Viewport_scale

                    If IsNumeric(TextBox_text_height.Text) = False Then
                        MsgBox("Non numeric text height")
                        Exit Sub
                    End If
                    Dim Text_height As Double = CDbl(TextBox_text_height.Text)

                    If IsNumeric(TextBox_Hscale.Text) = True Then
                        hSCALE = CDbl(TextBox_Hscale.Text)
                    Else
                        MsgBox("Non numeric horizontal scale")
                        Exit Sub
                    End If

                    If IsNumeric(TextBox_Vscale.Text) = True Then
                        vSCALE = CDbl(TextBox_Vscale.Text)
                    Else
                        MsgBox("Non numeric vertical scale")
                        Exit Sub
                    End If
                    Vertical_exag = hSCALE / vSCALE

                    Pt2(0) = Xg0
                    Pt2(1) = Yg0 + Nr_hor_Lines * Vert_increment * (Factor_print_viewport / hSCALE) * Vertical_exag

                    '**** extraX, extraY

                    Dim Extra_y As Double = 0
                    If Vert_increment = 50 Then
                        Extra_y = 0.5 * Vert_increment * (Factor_print_viewport / hSCALE) * Vertical_exag
                    End If

                    Dim Extra_x As Double = 0
                    If CheckBox_Hydrostatic_style.Checked = True Then
                        Extra_x = Horiz_increment * 1.11
                    End If


                    Dim Vert_Zero_line As New Autodesk.AutoCAD.DatabaseServices.Line
                    Vert_Zero_line.Layer = ComboBox_layer_grid_lines.Text
                    Vert_Zero_line.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Pt1(0), Pt1(1), 0)

                    If CheckBox_No_vertical_lines.Checked = False Then
                        Vert_Zero_line.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Pt2(0), Pt2(1), 0)



                        If Nou_zero = True Then
                            Vert_Zero_line.ColorIndex = Color_index_grid
                            Vert_Zero_line.Linetype = ComboBox_LINETYPE.Text
                            Vert_Zero_line.LineWeight = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight000
                        End If

                        If TextBox_Minimum_chainage.Text = "0" Then
                            Vert_Zero_line.Linetype = "CONTINUOUS"
                        End If


                        If Chainage_Graph_start < 0 Then
                            Vert_Zero_line.ColorIndex = Color_index_grid
                            Vert_Zero_line.Linetype = ComboBox_LINETYPE.Text
                            Vert_Zero_line.LineWeight = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight000
                        End If
                    Else

                        If Nr_vert_Lines_stanga = 0 Then
                            Vert_Zero_line.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Pt2(0), Pt2(1), 0)
                        End If


                        Vert_Zero_line.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Pt2(0), Yg0 - Text_height, 0)
                        Vert_Zero_line.Linetype = "CONTINUOUS"
                        If TextBox_Minimum_chainage.Text = "0" Then
                            Dim Vert_Zero_line_stanga As New Autodesk.AutoCAD.DatabaseServices.Line
                            Vert_Zero_line_stanga.Layer = ComboBox_layer_grid_lines.Text
                            Vert_Zero_line_stanga.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Pt1(0), Pt1(1), 0)
                            Vert_Zero_line_stanga.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Pt2(0), Pt2(1), 0)
                            Vert_Zero_line_stanga.Linetype = "CONTINUOUS"
                            BTrecord.AppendEntity(Vert_Zero_line_stanga)
                            Trans1.AddNewlyCreatedDBObject(Vert_Zero_line_stanga, True)
                        End If

                    End If

                    If CheckBox_Hydrostatic_style.Checked = True Then

                        Vert_Zero_line.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xg0, Yg0 - Extra_y, 0)
                        Vert_Zero_line.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xg0, Yg0 - Extra_y - 2 * Text_height, 0)
                        Vert_Zero_line.Linetype = "CONTINUOUS"

                        If TextBox_Minimum_chainage.Text = "0" Then
                            Dim Vert_Zero_line_stanga As New Autodesk.AutoCAD.DatabaseServices.Line
                            Vert_Zero_line_stanga.Layer = ComboBox_layer_grid_lines.Text
                            Vert_Zero_line_stanga.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Pt1(0) - Extra_x, Pt1(1) - Extra_y, 0)
                            Vert_Zero_line_stanga.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Pt2(0) - Extra_x, Pt2(1), 0)
                            Vert_Zero_line_stanga.Linetype = "CONTINUOUS"
                            Vert_Zero_line_stanga.ColorIndex = 5
                            BTrecord.AppendEntity(Vert_Zero_line_stanga)
                            Trans1.AddNewlyCreatedDBObject(Vert_Zero_line_stanga, True)
                        End If


                    End If


                    BTrecord.AppendEntity(Vert_Zero_line)
                    Trans1.AddNewlyCreatedDBObject(Vert_Zero_line, True)

                    Linia_chainage_zero = Vert_Zero_line



                    For i = 1 To Nr_vert_Lines_dreapta
                        Dim Xv1, Yv1, Xv2, Yv2 As Double

                        Xv1 = Pt1(0) + i * Horiz_increment * (Factor_print_viewport / hSCALE)
                        Yv1 = Pt1(1)
                        Xv2 = Pt1(0) + i * Horiz_increment * (Factor_print_viewport / hSCALE)
                        Yv2 = Pt2(1)

                        Dim V_off_line_dreapta As New Autodesk.AutoCAD.DatabaseServices.Line
                        V_off_line_dreapta.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xv1, Yv1, 0)
                        V_off_line_dreapta.Layer = ComboBox_layer_grid_lines.Text

                        If CheckBox_No_vertical_lines.Checked = False Then
                            V_off_line_dreapta.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xv2, Yv2, 0)

                            If i < Nr_vert_Lines_dreapta Then
                                V_off_line_dreapta.ColorIndex = Color_index_grid
                                V_off_line_dreapta.Linetype = ComboBox_LINETYPE.Text
                                V_off_line_dreapta.LineWeight = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight000
                            Else
                                V_off_line_dreapta.Linetype = "CONTINUOUS"
                            End If
                        Else
                            If i < Nr_vert_Lines_dreapta Then
                                V_off_line_dreapta.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xv2, Yg0 - CDbl(TextBox_text_height.Text), 0)
                            Else
                                V_off_line_dreapta.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xv2, Yv2, 0)
                            End If
                            V_off_line_dreapta.Linetype = "CONTINUOUS"
                        End If

                        '  AICI BAGI EXTRA LINIE

                        If CheckBox_Hydrostatic_style.Checked = True Then
                            V_off_line_dreapta.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xv1, Yv1 - Extra_y, 0)
                            V_off_line_dreapta.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xv2, Yg0 - Extra_y - 2 * CDbl(TextBox_text_height.Text), 0)
                            If i = Nr_vert_Lines_dreapta Then
                                Dim Vertical_linie_dreapta As New Autodesk.AutoCAD.DatabaseServices.Line
                                Vertical_linie_dreapta.Layer = ComboBox_layer_grid_lines.Text
                                Vertical_linie_dreapta.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xv1 + Extra_x, Yv1 - Extra_y, 0)
                                Vertical_linie_dreapta.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xv2 + Extra_x, Yv2, 0)
                                Vertical_linie_dreapta.Linetype = "CONTINUOUS"
                                Vertical_linie_dreapta.ColorIndex = 5
                                BTrecord.AppendEntity(Vertical_linie_dreapta)
                                Trans1.AddNewlyCreatedDBObject(Vertical_linie_dreapta, True)
                            End If
                            V_off_line_dreapta.Linetype = "CONTINUOUS"
                        End If



                        BTrecord.AppendEntity(V_off_line_dreapta)
                        Trans1.AddNewlyCreatedDBObject(V_off_line_dreapta, True)

                    Next


                    If Nr_vert_Lines_stanga > 0 Then
                        For i = 1 To Nr_vert_Lines_stanga
                            Dim Xv1, Yv1, Xv2, Yv2 As Double
                            Xv1 = Pt1(0) - i * Horiz_increment * (Factor_print_viewport / hSCALE)
                            Xv2 = Pt1(0) - i * Horiz_increment * (Factor_print_viewport / hSCALE)


                            Yv1 = Pt1(1)
                            Yv2 = Pt2(1)
                            Dim V_off_line_stanga As New Autodesk.AutoCAD.DatabaseServices.Line
                            V_off_line_stanga.Layer = ComboBox_layer_grid_lines.Text

                            V_off_line_stanga.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xv1, Yv1, 0)

                            If CheckBox_No_vertical_lines.Checked = False Then
                                V_off_line_stanga.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xv2, Yv2, 0)
                                If i < Nr_vert_Lines_stanga Then
                                    V_off_line_stanga.ColorIndex = Color_index_grid
                                    V_off_line_stanga.Linetype = ComboBox_LINETYPE.Text
                                    V_off_line_stanga.LineWeight = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight000
                                Else
                                    V_off_line_stanga.Linetype = "CONTINUOUS"
                                End If
                            Else
                                If i < Nr_vert_Lines_stanga Then
                                    V_off_line_stanga.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xv2, Yg0 - CDbl(TextBox_text_height.Text), 0)
                                Else
                                    V_off_line_stanga.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xv2, Yv2, 0)
                                End If
                                V_off_line_stanga.Linetype = "CONTINUOUS"
                            End If

                            If CheckBox_Hydrostatic_style.Checked = True Then
                                V_off_line_stanga.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xv1, Yv1 - Extra_y, 0)
                                V_off_line_stanga.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xv2, Yg0 - Extra_y - 2 * CDbl(TextBox_text_height.Text), 0)
                                If i = Nr_vert_Lines_stanga Then

                                    Dim Vertical_linie_stanga As New Autodesk.AutoCAD.DatabaseServices.Line
                                    Vertical_linie_stanga.Layer = ComboBox_layer_grid_lines.Text
                                    Vertical_linie_stanga.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xv1 - Extra_x, Yv1 - Extra_y, 0)
                                    Vertical_linie_stanga.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xv2 - Extra_x, Yv2, 0)
                                    Vertical_linie_stanga.Linetype = "CONTINUOUS"
                                    Vertical_linie_stanga.ColorIndex = 5
                                    BTrecord.AppendEntity(Vertical_linie_stanga)
                                    Trans1.AddNewlyCreatedDBObject(Vertical_linie_stanga, True)
                                End If
                                V_off_line_stanga.Linetype = "CONTINUOUS"
                            End If




                            BTrecord.AppendEntity(V_off_line_stanga)
                            Trans1.AddNewlyCreatedDBObject(V_off_line_stanga, True)

                        Next

                    End If




                    If Nr_vert_Lines_stanga > 0 Then
                        Pt3(0) = Pt1(0) - Nr_vert_Lines_stanga * Horiz_increment * (Factor_print_viewport / hSCALE)

                    Else
                        Pt3(0) = Pt1(0)
                    End If

                    Pt3(1) = Pt1(1)

                    Pt4(0) = Pt1(0) + Nr_vert_Lines_dreapta * Horiz_increment * (Factor_print_viewport / hSCALE)
                    Pt4(1) = Pt1(1)


                    Dim Horiz_Zero_line As New Autodesk.AutoCAD.DatabaseServices.Line
                    Horiz_Zero_line.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Pt3(0), Pt3(1), 0)
                    Horiz_Zero_line.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Pt4(0), Pt4(1), 0)
                    Horiz_Zero_line.Layer = ComboBox_layer_grid_lines.Text

                    If CheckBox_Hydrostatic_style.Checked = True Then
                        Horiz_Zero_line.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Pt3(0) - Extra_x, Pt3(1), 0)
                        Horiz_Zero_line.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Pt4(0) + Extra_x, Pt4(1), 0)
                    Else
                        Horiz_Zero_line.Linetype = "CONTINUOUS"
                    End If

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

                        h_off_line.ColorIndex = Color_index_grid
                        h_off_line.Linetype = ComboBox_LINETYPE.Text
                        h_off_line.LineWeight = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight000
                        h_off_line.Layer = ComboBox_layer_grid_lines.Text
                        If CheckBox_Hydrostatic_style.Checked = True Then
                            If i = Nr_hor_Lines Then
                                h_off_line.Linetype = "CONTINUOUS"
                                h_off_line.ColorIndex = 5
                            End If
                            h_off_line.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xh1 - Extra_x, Yh1, 0)
                            h_off_line.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xh2 + Extra_x, Yh2, 0)
                        End If
                        BTrecord.AppendEntity(h_off_line)
                        Trans1.AddNewlyCreatedDBObject(h_off_line, True)




                    Next

                    If CheckBox_Hydrostatic_style.Checked = True Then
                        If Vert_increment = 50 Then
                            For i = 0 To Nr_hor_Lines ' desenez linia de jos de aia am pus de la zero
                                Dim Xh1, Yh1, Xh2, Yh2 As Double
                                Xh1 = Pt3(0)
                                Xh2 = Pt4(0)
                                Yh1 = Pt3(1) + i * (Factor_print_viewport / hSCALE) * Vertical_exag * Vert_increment - Extra_y
                                Yh2 = Yh1
                                Dim h_off_line As New Autodesk.AutoCAD.DatabaseServices.Line
                                h_off_line.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xh1 - Extra_x, Yh1, 0)
                                h_off_line.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xh2 + Extra_x, Yh2, 0)
                                h_off_line.ColorIndex = Color_index_grid
                                h_off_line.Linetype = ComboBox_LINETYPE.Text
                                h_off_line.LineWeight = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight000
                                h_off_line.Layer = ComboBox_layer_grid_lines.Text
                                BTrecord.AppendEntity(h_off_line)
                                Trans1.AddNewlyCreatedDBObject(h_off_line, True)

                                For j = 5 To 20 Step 5
                                    Dim Liniuta_stanga As New Autodesk.AutoCAD.DatabaseServices.Line
                                    Liniuta_stanga.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xh1 - Extra_x, Yh1 + j * Vertical_exag * (Factor_print_viewport / hSCALE), 0)
                                    Liniuta_stanga.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xh1 - Extra_x + 2 * Text_height, Yh1 + j * Vertical_exag * (Factor_print_viewport / hSCALE), 0)
                                    Liniuta_stanga.ColorIndex = Color_index_grid
                                    Liniuta_stanga.Linetype = "CONTINUOUS"
                                    Liniuta_stanga.LineWeight = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight000
                                    Liniuta_stanga.Layer = ComboBox_layer_grid_lines.Text
                                    BTrecord.AppendEntity(Liniuta_stanga)
                                    Trans1.AddNewlyCreatedDBObject(Liniuta_stanga, True)

                                    Dim Liniuta_dreapta As New Autodesk.AutoCAD.DatabaseServices.Line
                                    Liniuta_dreapta.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xh2 + Extra_x, Yh1 + j * Vertical_exag * (Factor_print_viewport / hSCALE), 0)
                                    Liniuta_dreapta.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xh2 + Extra_x - 2 * Text_height, Yh1 + j * Vertical_exag * (Factor_print_viewport / hSCALE), 0)
                                    Liniuta_dreapta.ColorIndex = Color_index_grid
                                    Liniuta_dreapta.Linetype = "CONTINUOUS"
                                    Liniuta_dreapta.LineWeight = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight000
                                    Liniuta_dreapta.Layer = ComboBox_layer_grid_lines.Text
                                    BTrecord.AppendEntity(Liniuta_dreapta)
                                    Trans1.AddNewlyCreatedDBObject(Liniuta_dreapta, True)

                                Next

                                For j = 30 To 45 Step 5
                                    If i < Nr_hor_Lines Then
                                        Dim Liniuta_stanga As New Autodesk.AutoCAD.DatabaseServices.Line
                                        Liniuta_stanga.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xh1 - Extra_x, Yh1 + j * Vertical_exag * (Factor_print_viewport / hSCALE), 0)
                                        Liniuta_stanga.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xh1 - Extra_x + 2 * Text_height, Yh1 + j * Vertical_exag * (Factor_print_viewport / hSCALE), 0)
                                        Liniuta_stanga.ColorIndex = Color_index_grid
                                        Liniuta_stanga.Linetype = "CONTINUOUS"
                                        Liniuta_stanga.LineWeight = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight000
                                        Liniuta_stanga.Layer = ComboBox_layer_grid_lines.Text
                                        BTrecord.AppendEntity(Liniuta_stanga)
                                        Trans1.AddNewlyCreatedDBObject(Liniuta_stanga, True)

                                        Dim Liniuta_dreapta As New Autodesk.AutoCAD.DatabaseServices.Line
                                        Liniuta_dreapta.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xh2 + Extra_x, Yh1 + j * Vertical_exag * (Factor_print_viewport / hSCALE), 0)
                                        Liniuta_dreapta.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xh2 + Extra_x - 2 * Text_height, Yh1 + j * Vertical_exag * (Factor_print_viewport / hSCALE), 0)
                                        Liniuta_dreapta.ColorIndex = Color_index_grid
                                        Liniuta_dreapta.Linetype = "CONTINUOUS"
                                        Liniuta_dreapta.LineWeight = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight000
                                        Liniuta_dreapta.Layer = ComboBox_layer_grid_lines.Text
                                        BTrecord.AppendEntity(Liniuta_dreapta)
                                        Trans1.AddNewlyCreatedDBObject(Liniuta_dreapta, True)

                                    End If
                                Next

                            Next
                        End If
                    End If


                    If CheckBox_No_vertical_lines.Checked = True Then
                        For i = 1 To Nr_hor_Lines
                            Dim Xh1, Yh1, Xh2, Yh2 As Double
                            Xh1 = Pt3(0)
                            Yh1 = Pt3(1) + i * (Factor_print_viewport / hSCALE) * Vertical_exag * Vert_increment

                            Xh2 = Pt3(0) + Text_height * 4
                            Yh2 = Yh1


                            Dim h_off_line_stanga As New Autodesk.AutoCAD.DatabaseServices.Line
                            h_off_line_stanga.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xh1, Yh1, 0)
                            h_off_line_stanga.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xh2, Yh2, 0)

                            h_off_line_stanga.Linetype = "CONTINUOUS"
                            h_off_line_stanga.Layer = ComboBox_layer_grid_lines.Text
                            BTrecord.AppendEntity(h_off_line_stanga)
                            Trans1.AddNewlyCreatedDBObject(h_off_line_stanga, True)

                            Xh1 = Pt3(0)
                            Yh1 = Pt3(1) + (i - 1) * (Factor_print_viewport / hSCALE) * Vertical_exag * Vert_increment + 0.5 * (Factor_print_viewport / hSCALE) * Vertical_exag * Vert_increment
                            Xh2 = Pt3(0) + Text_height * 2
                            Yh2 = Yh1

                            Dim h_off_line_stanga_mica As New Autodesk.AutoCAD.DatabaseServices.Line
                            h_off_line_stanga_mica.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xh1, Yh1, 0)
                            h_off_line_stanga_mica.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xh2, Yh2, 0)

                            h_off_line_stanga_mica.Linetype = "CONTINUOUS"
                            h_off_line_stanga_mica.Layer = ComboBox_layer_grid_lines.Text
                            BTrecord.AppendEntity(h_off_line_stanga_mica)
                            Trans1.AddNewlyCreatedDBObject(h_off_line_stanga_mica, True)

                            Xh1 = Pt4(0) - Text_height * 4
                            Yh1 = Pt3(1) + i * (Factor_print_viewport / hSCALE) * Vertical_exag * Vert_increment

                            Xh2 = Pt4(0)
                            Yh2 = Yh1

                            Dim h_off_line_dreapta As New Autodesk.AutoCAD.DatabaseServices.Line
                            h_off_line_dreapta.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xh1, Yh1, 0)
                            h_off_line_dreapta.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xh2, Yh2, 0)

                            h_off_line_dreapta.Linetype = "CONTINUOUS"
                            h_off_line_dreapta.Layer = ComboBox_layer_grid_lines.Text
                            BTrecord.AppendEntity(h_off_line_dreapta)
                            Trans1.AddNewlyCreatedDBObject(h_off_line_dreapta, True)

                            Xh1 = Pt4(0) - Text_height * 2
                            Yh1 = Pt3(1) + (i - 1) * (Factor_print_viewport / hSCALE) * Vertical_exag * Vert_increment + 0.5 * (Factor_print_viewport / hSCALE) * Vertical_exag * Vert_increment
                            Xh2 = Pt4(0)
                            Yh2 = Yh1

                            Dim h_off_line_dreapta_mica As New Autodesk.AutoCAD.DatabaseServices.Line
                            h_off_line_dreapta_mica.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xh1, Yh1, 0)
                            h_off_line_dreapta_mica.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(Xh2, Yh2, 0)

                            h_off_line_dreapta_mica.Linetype = "CONTINUOUS"
                            h_off_line_dreapta_mica.Layer = ComboBox_layer_grid_lines.Text
                            BTrecord.AppendEntity(h_off_line_dreapta_mica)
                            Trans1.AddNewlyCreatedDBObject(h_off_line_dreapta_mica, True)

                        Next
                    End If




                    Pt6(0) = Pt3(0) ' - 10 * (Factor_print_viewport / 1000)
                    Pt6(1) = Pt1(1) '

                    Pt7(0) = Pt4(0) ' + 10 * (Factor_print_viewport / 1000)
                    Pt7(1) = Pt1(1) '


                    Pt8(0) = Pt1(0)
                    Pt8(1) = Pt1(1) - 1.2 * Text_height

                    If CheckBox_No_vertical_lines.Checked = True Then
                        Pt8(1) = Yg0 - 1.5 * Text_height
                    End If
                    If CheckBox_Hydrostatic_style.Checked = True Then
                        Pt8(1) = Yg0 - 3.5 * Text_height - Extra_y
                    End If


                    Dim MText_ch_Hor As New Autodesk.AutoCAD.DatabaseServices.MText
                    If CheckBox_US_style.Checked = False Then
                        MText_ch_Hor.Contents = "0+000"
                    Else
                        MText_ch_Hor.Contents = "0+00"
                    End If

                    MText_ch_Hor.Layer = ComboBox_LAYER_TEXT_AND_BLOCKS.Text

                    MText_ch_Hor.TextStyleId = Text_style_ID
                    MText_ch_Hor.TextHeight = Text_height
                    MText_ch_Hor.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.TopCenter
                    MText_ch_Hor.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt8(0), Pt8(1), 0)
                    BTrecord.AppendEntity(MText_ch_Hor)
                    Trans1.AddNewlyCreatedDBObject(MText_ch_Hor, True)

                    For i = 1 To Nr_vert_Lines_dreapta

                        Pt9(0) = Pt8(0) + i * Horiz_increment * (Factor_print_viewport / hSCALE)
                        Pt9(1) = Pt8(1)

                        Dim Text_ch_Hor_h_dreapta As String

                        If CheckBox_US_style.Checked = False Then
                            Text_ch_Hor_h_dreapta = Get_chainage_from_double(i * Horiz_increment, 0)
                        Else
                            Text_ch_Hor_h_dreapta = Get_chainage_feet_from_double(i * Horiz_increment, 0)
                        End If




                        Dim MText_ch_Hor_h_dreapta As New Autodesk.AutoCAD.DatabaseServices.MText
                        MText_ch_Hor_h_dreapta.Contents = Text_ch_Hor_h_dreapta
                        MText_ch_Hor_h_dreapta.Layer = ComboBox_LAYER_TEXT_AND_BLOCKS.Text
                        MText_ch_Hor_h_dreapta.TextStyleId = Text_style_ID
                        MText_ch_Hor_h_dreapta.TextHeight = Text_height
                        MText_ch_Hor_h_dreapta.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.TopCenter
                        MText_ch_Hor_h_dreapta.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt9(0), Pt9(1), 0)
                        BTrecord.AppendEntity(MText_ch_Hor_h_dreapta)
                        Trans1.AddNewlyCreatedDBObject(MText_ch_Hor_h_dreapta, True)


                    Next

                    If Nr_vert_Lines_stanga > 0 Then

                        For i = 1 To Nr_vert_Lines_stanga

                            Pt9(0) = Pt8(0) - i * Horiz_increment * (Factor_print_viewport / hSCALE)

                            Pt9(1) = Pt8(1)

                            Dim Text_ch_Hor_h_stanga As String

                            If CheckBox_US_style.Checked = False Then
                                Text_ch_Hor_h_stanga = Get_chainage_from_double(-i * Horiz_increment, 0)
                            Else
                                Text_ch_Hor_h_stanga = Get_chainage_feet_from_double(-i * Horiz_increment, 0)
                            End If


                            Dim MText_ch_Hor_h_stanga As New Autodesk.AutoCAD.DatabaseServices.MText
                            MText_ch_Hor_h_stanga.Contents = Text_ch_Hor_h_stanga
                            MText_ch_Hor_h_stanga.Layer = ComboBox_LAYER_TEXT_AND_BLOCKS.Text
                            MText_ch_Hor_h_stanga.TextStyleId = Text_style_ID
                            MText_ch_Hor_h_stanga.TextHeight = Text_height
                            MText_ch_Hor_h_stanga.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.TopCenter
                            MText_ch_Hor_h_stanga.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt9(0), Pt9(1), 0)
                            BTrecord.AppendEntity(MText_ch_Hor_h_stanga)
                            Trans1.AddNewlyCreatedDBObject(MText_ch_Hor_h_stanga, True)

                        Next

                    End If






                    Pt10(0) = Pt6(0) - 0.5 * Text_height
                    Pt10(1) = Pt6(1)

                    Pt12(0) = Pt7(0) + 0.5 * Text_height
                    Pt12(1) = Pt7(1)

                    If CheckBox_No_vertical_lines.Checked = True Then
                        Pt10(0) = Pt6(0) + 0.7 * Text_height
                        Pt10(1) = Pt6(1) + 0.5 * Text_height

                        Pt12(0) = Pt7(0) - 0.7 * Text_height
                        Pt12(1) = Pt7(1) + 0.5 * Text_height
                    End If

                    If CheckBox_Hydrostatic_style.Checked = True Then
                        Pt10(0) = Pt6(0) + 3.4 * Text_height - Extra_x
                        Pt10(1) = Pt6(1) + 0.7 * Text_height

                        Pt12(0) = Pt7(0) - 3.4 * Text_height + Extra_x
                        Pt12(1) = Pt7(1) + 0.7 * Text_height
                    End If

                    Dim Vstring As String
                    Vstring = Round(Grid_Z_Min, 0).ToString

                    If CheckBox_US_style.Checked = True Then
                        Vstring = Vstring & "'"
                    End If


                    Dim Mtext_ch_Ver_left As New Autodesk.AutoCAD.DatabaseServices.MText
                    Mtext_ch_Ver_left.Contents = Vstring
                    Mtext_ch_Ver_left.Layer = ComboBox_LAYER_TEXT_AND_BLOCKS.Text
                    Mtext_ch_Ver_left.TextStyleId = Text_style_ID
                    Mtext_ch_Ver_left.TextHeight = Text_height

                    Mtext_ch_Ver_left.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.MiddleRight

                    If CheckBox_No_vertical_lines.Checked = True Or CheckBox_Hydrostatic_style.Checked = True Then
                        Mtext_ch_Ver_left.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.BottomLeft
                    End If

                    Mtext_ch_Ver_left.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt10(0), Pt10(1), 0)
                    BTrecord.AppendEntity(Mtext_ch_Ver_left)
                    Trans1.AddNewlyCreatedDBObject(Mtext_ch_Ver_left, True)

                    Dim Mtext_ch_Ver_right As New Autodesk.AutoCAD.DatabaseServices.MText
                    Mtext_ch_Ver_right.Contents = Vstring
                    Mtext_ch_Ver_right.Layer = ComboBox_LAYER_TEXT_AND_BLOCKS.Text
                    Mtext_ch_Ver_right.TextStyleId = Text_style_ID
                    Mtext_ch_Ver_right.TextHeight = Text_height
                    Mtext_ch_Ver_right.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.MiddleLeft

                    If CheckBox_No_vertical_lines.Checked = True Or CheckBox_Hydrostatic_style.Checked = True Then
                        Mtext_ch_Ver_right.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.BottomRight
                    End If

                    Mtext_ch_Ver_right.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt12(0), Pt12(1), 0)
                    BTrecord.AppendEntity(Mtext_ch_Ver_right)
                    Trans1.AddNewlyCreatedDBObject(Mtext_ch_Ver_right, True)

                    For i = 1 To Nr_hor_Lines
                        Pt11(0) = Pt10(0)
                        Pt11(1) = Pt10(1) + i * Vert_increment * (Factor_print_viewport / hSCALE) * Vertical_exag
                        Pt13(0) = Pt12(0)
                        Pt13(1) = Pt12(1) + i * Vert_increment * (Factor_print_viewport / hSCALE) * Vertical_exag



                        Vstring = Round((i * Vert_increment) + Grid_Z_Min, 0).ToString

                        If CheckBox_US_style.Checked = True Then
                            Vstring = Vstring & "'"
                        End If

                        Dim Mtext_ch_Ver_left1 As New Autodesk.AutoCAD.DatabaseServices.MText
                        Mtext_ch_Ver_left1.Contents = Vstring
                        Mtext_ch_Ver_left1.Layer = ComboBox_LAYER_TEXT_AND_BLOCKS.Text
                        Mtext_ch_Ver_left1.TextStyleId = Text_style_ID
                        Mtext_ch_Ver_left1.TextHeight = Text_height
                        Mtext_ch_Ver_left1.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.MiddleRight

                        If CheckBox_No_vertical_lines.Checked = True Or CheckBox_Hydrostatic_style.Checked = True Then
                            Mtext_ch_Ver_left1.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.BottomLeft
                        End If

                        Mtext_ch_Ver_left1.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt11(0), Pt11(1), 0)


                        If CheckBox_No_vertical_lines.Checked = False Then
                            BTrecord.AppendEntity(Mtext_ch_Ver_left1)
                            Trans1.AddNewlyCreatedDBObject(Mtext_ch_Ver_left1, True)
                        Else
                            If i < Nr_hor_Lines Then

                                BTrecord.AppendEntity(Mtext_ch_Ver_left1)
                                Trans1.AddNewlyCreatedDBObject(Mtext_ch_Ver_left1, True)

                            End If

                        End If


                        Dim Mtext_ch_Ver_right1 As New Autodesk.AutoCAD.DatabaseServices.MText
                        Mtext_ch_Ver_right1.Contents = Vstring
                        Mtext_ch_Ver_right1.Layer = ComboBox_LAYER_TEXT_AND_BLOCKS.Text
                        Mtext_ch_Ver_right1.TextStyleId = Text_style_ID
                        Mtext_ch_Ver_right1.TextHeight = Text_height
                        Mtext_ch_Ver_right1.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.MiddleLeft

                        If CheckBox_No_vertical_lines.Checked = True Or CheckBox_Hydrostatic_style.Checked = True Then
                            Mtext_ch_Ver_right1.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.BottomRight
                        End If
                        Mtext_ch_Ver_right1.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt13(0), Pt13(1), 0)


                        If CheckBox_No_vertical_lines.Checked = False Then
                            BTrecord.AppendEntity(Mtext_ch_Ver_right1)
                            Trans1.AddNewlyCreatedDBObject(Mtext_ch_Ver_right1, True)
                        Else
                            If i < Nr_hor_Lines Then
                                BTrecord.AppendEntity(Mtext_ch_Ver_right1)
                                Trans1.AddNewlyCreatedDBObject(Mtext_ch_Ver_right1, True)
                            End If
                        End If
                    Next

                    Pt14(0) = (Pt3(0) + Pt4(0)) / 2
                    Pt14(1) = (Pt3(1) + Pt4(1)) / 2 - 6.5 * Text_height - Extra_y
                    Pt15(0) = (Pt3(0) + Pt4(0)) / 2
                    Pt15(1) = (Pt3(1) + Pt4(1)) / 2 - 10.5 * Text_height - Extra_y




                    Dim Titlu2 As New Autodesk.AutoCAD.DatabaseServices.MText
                    Titlu2.Contents = "EXAGGERATION - HORIZ= " & (1000 / hSCALE) & "  VERT=" & (1000 / vSCALE)
                    Titlu2.Layer = ComboBox_LAYER_TEXT_AND_BLOCKS.Text
                    Titlu2.TextStyleId = Text_style_ID
                    Titlu2.TextHeight = 4 * Text_height
                    Titlu2.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.TopCenter
                    Titlu2.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt15(0), Pt15(1), 0)
                    BTrecord.AppendEntity(Titlu2)
                    Trans1.AddNewlyCreatedDBObject(Titlu2, True)
                    Pt19(0) = Pt1(0) - Nr_vert_Lines_stanga * Horiz_increment * (Factor_print_viewport / hSCALE)


                    Pt19(1) = Pt9(1) - 1.5 * Text_height
                    Pt20(0) = Pt1(0) + Nr_vert_Lines_dreapta * Horiz_increment * (Factor_print_viewport / hSCALE)
                    Pt20(1) = Pt9(1) - 1.5 * Text_height


                    If CheckBox_Hydrostatic_style.Checked = False Then
                        Dim North_South1 As New Autodesk.AutoCAD.DatabaseServices.MText
                        North_South1.Contents = "Specify Direction"
                        North_South1.Layer = ComboBox_LAYER_TEXT_AND_BLOCKS.Text
                        North_South1.TextStyleId = Text_style_ID
                        North_South1.TextHeight = Text_height
                        North_South1.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.TopCenter
                        North_South1.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt19(0), Pt19(1), 0)
                        BTrecord.AppendEntity(North_South1)
                        Trans1.AddNewlyCreatedDBObject(North_South1, True)

                        Dim North_South2 As New Autodesk.AutoCAD.DatabaseServices.MText
                        North_South2.Contents = "Specify Direction"
                        North_South2.Layer = ComboBox_LAYER_TEXT_AND_BLOCKS.Text
                        North_South2.TextStyleId = Text_style_ID
                        North_South2.TextHeight = Text_height
                        North_South2.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.TopCenter
                        North_South2.Location = New Autodesk.AutoCAD.Geometry.Point3d(Pt20(0), Pt20(1), 0)
                        BTrecord.AppendEntity(North_South2)
                        Trans1.AddNewlyCreatedDBObject(North_South2, True)
                    End If




                    Dim Profile_x_point(0 To Profile_table.Rows.Count - 1) As Double
                    Dim Profile_y_point(0 To Profile_table.Rows.Count - 1) As Double

                    Dim Poly1 As New Autodesk.AutoCAD.DatabaseServices.Polyline()


                    Profile_x_point(0) = Pt1(0) + (Chainage_poly_start) * (Factor_print_viewport / hSCALE)
                    Profile_y_point(0) = Yg0 + (Profile_table.Rows.Item(0).Item("Elevation") - Grid_Z_Min) * (Factor_print_viewport / hSCALE) * Vertical_exag



                    Poly1.AddVertexAt(0, New Autodesk.AutoCAD.Geometry.Point2d(Profile_x_point(0), Profile_y_point(0)), 0, 0, 0)
                    'here You can put CSF!
                    For i = 1 To Profile_table.Rows.Count - 1
                        Dim Grid_Chainage As Double = Profile_table.Rows.Item(i).Item("Chainage_Recalculated") - Profile_table.Rows.Item(i - 1).Item("Chainage_Recalculated")
                        Profile_x_point(i) = Profile_x_point(i - 1) + Grid_Chainage * (Factor_print_viewport / hSCALE) / CSF
                        Profile_y_point(i) = Yg0 + (Profile_table.Rows.Item(i).Item("Elevation") - Grid_Z_Min) * (Factor_print_viewport / hSCALE) * Vertical_exag
                        Poly1.AddVertexAt(i, New Autodesk.AutoCAD.Geometry.Point2d(Profile_x_point(i), Profile_y_point(i)), 0, 0, 0)
                    Next

                    Poly1.Layer = ComboBox_LAYER_PROFILE_POLYLINE.Text
                    BTrecord.AppendEntity(Poly1)
                    Trans1.AddNewlyCreatedDBObject(Poly1, True)
                    Polylinie_profil = Poly1





                    If CheckBox_add_description_on_graph.Checked = True Then
                        creaza_linie_pt_pozitie_si_insereaza_block_daca_e_selectat_Blockul()
                    End If
                    If Deseneaza_poly_profil2 = True Then
                        Dim Poly_prof2 As New Autodesk.AutoCAD.DatabaseServices.Polyline()
                        Profile_x_point(0) = Pt1(0) + (Profile_table2.Rows.Item(0).Item("Chainage_Recalculated") / CSF) * (Factor_print_viewport / hSCALE)
                        Profile_y_point(0) = Yg0 + (Profile_table2.Rows.Item(0).Item("Elevation") - Grid_Z_Min) * (Factor_print_viewport / hSCALE) * Vertical_exag
                        Poly_prof2.AddVertexAt(0, New Autodesk.AutoCAD.Geometry.Point2d(Profile_x_point(0), Profile_y_point(0)), 0, 0, 0)

                        For i = 1 To Profile_table2.Rows.Count - 1
                            Dim Grid_Chainage As Double = Profile_table2.Rows.Item(i).Item("Chainage_Recalculated") - Profile_table2.Rows.Item(i - 1).Item("Chainage_Recalculated")
                            Profile_x_point(i) = Profile_x_point(i - 1) + Grid_Chainage * (Factor_print_viewport / hSCALE) / CSF
                            Profile_y_point(i) = Yg0 + (Profile_table2.Rows.Item(i).Item("Elevation") - Grid_Z_Min) * (Factor_print_viewport / hSCALE) * Vertical_exag
                            Poly_prof2.AddVertexAt(i, New Autodesk.AutoCAD.Geometry.Point2d(Profile_x_point(i), Profile_y_point(i)), 0, 0, 0)
                        Next
                        Poly_prof2.Layer = ComboBox_LAYER_PROFILE_POLYLINE.Text
                        Poly_prof2.ColorIndex = 131
                        BTrecord.AppendEntity(Poly_prof2)
                        Trans1.AddNewlyCreatedDBObject(Poly_prof2, True)
                    End If

                    Editor1.Regen()
                    Trans1.Commit()
                End Using ' asta e de la transaction

            End Using ' ASTA E DE LA LOCK1

            CheckBox_draw_2Dpoly.Checked = False
            CheckBox_draw_3Dpoly.Checked = False


        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub


    Private Sub ComboBox_blocks_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ComboBox_blocks.SelectedIndexChanged
        Try
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Using LOCK1 As DocumentLock = ThisDrawing.LockDocument
                If ComboBox_atrib_description_1.Items.Count > 0 Then ComboBox_atrib_description_1.Items.Clear()
                If ComboBox_atrib_chainage.Items.Count > 0 Then ComboBox_atrib_chainage.Items.Clear()
                If ComboBox_atrib_description_2.Items.Count > 0 Then ComboBox_atrib_description_2.Items.Clear()
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim Block_table As Autodesk.AutoCAD.DatabaseServices.BlockTable = Trans1.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    If Block_table.Has(ComboBox_blocks.Text) = True Then
                        Dim Block1 As BlockTableRecord = Trans1.GetObject(Block_table.Item(ComboBox_blocks.Text), OpenMode.ForRead)
                        If Block1.HasAttributeDefinitions = True Then
                            For Each Id1 As ObjectId In Block1
                                Dim ent As Entity = TryCast(Trans1.GetObject(Id1, OpenMode.ForRead), Entity)
                                If ent IsNot Nothing Then
                                    Dim attDefinition1 As AttributeDefinition = TryCast(ent, AttributeDefinition)
                                    If attDefinition1 IsNot Nothing Then
                                        ComboBox_atrib_description_1.Items.Add(attDefinition1.Tag)
                                        ComboBox_atrib_chainage.Items.Add(attDefinition1.Tag)
                                        ComboBox_atrib_description_2.Items.Add(attDefinition1.Tag)
                                    End If
                                End If
                            Next
                        End If
                    End If
                End Using ' asta e de la trans1
            End Using
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub eXAMPLE_UpdateAttributesInDatabase(db As Database, blockName As String, attbName As String, attbValue As String)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Autodesk.AutoCAD.EditorInput.Editor = doc.Editor
        ' Get the IDs of the spaces we want to process and simply call a function to process each
        Dim ModelSpaceId As ObjectId
        Dim PaperSpaceId As ObjectId
        Dim Transactn As Transaction = db.TransactionManager.StartTransaction()
        Using Transactn
            Dim bt As BlockTable = DirectCast(Transactn.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
            ModelSpaceId = bt(BlockTableRecord.ModelSpace)
            PaperSpaceId = bt(BlockTableRecord.PaperSpace)
            Transactn.Commit()
        End Using
        eXAMPLE_UpdateAttributesInBlock(ModelSpaceId, blockName, attbName, attbValue)
        eXAMPLE_UpdateAttributesInBlock(PaperSpaceId, blockName, attbName, attbValue)
        ed.Regen()
    End Sub
    '============================================================
    Private Sub eXAMPLE_UpdateAttributesInBlock(btRecordId As ObjectId, blockName As String, attbName As String, attbValue As String)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Autodesk.AutoCAD.EditorInput.Editor = doc.Editor
        Dim tr As Transaction = doc.TransactionManager.StartTransaction()
        Using tr
            Dim btRecord As BlockTableRecord = DirectCast(tr.GetObject(btRecordId, OpenMode.ForRead), BlockTableRecord)
            For Each entId As ObjectId In btRecord
                Dim ent As Entity = TryCast(tr.GetObject(entId, OpenMode.ForRead), Entity)
                If ent IsNot Nothing Then
                    Dim br As BlockReference = TryCast(ent, BlockReference)
                    If br IsNot Nothing Then
                        Dim bd As BlockTableRecord = DirectCast(tr.GetObject(br.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)  'to see whether it's a block with the name we're after
                        If bd.Name.ToUpper() = blockName Then               ' Check each of the attributes...
                            For Each attReferenceId As ObjectId In br.AttributeCollection
                                Dim obj As DBObject = tr.GetObject(attReferenceId, OpenMode.ForRead)
                                Dim attReference As AttributeReference = TryCast(obj, AttributeReference)
                                If attReference IsNot Nothing Then                    'to see whether it has the tag we're after
                                    If attReference.Tag.ToUpper() = attbName Then     ' If so, update the value and increment the counter
                                        attReference.UpgradeOpen()
                                        attReference.TextString = attbValue
                                        attReference.DowngradeOpen()
                                    End If
                                End If
                            Next
                        End If
                        eXAMPLE_UpdateAttributesInBlock(br.BlockTableRecord, blockName, attbName, attbValue)    ' Recurse for nested blocks
                    End If
                End If
            Next
            tr.Commit()
        End Using
    End Sub

    Private Sub creaza_linie_pt_pozitie_si_insereaza_block_daca_e_selectat_Blockul()
        Try


            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1

                ' Dim k As Double = 1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
                    Dim start1 As Double = CDbl(TextBox_row_start.Text)
                    Dim end1 As Double = CDbl(TextBox_row_end.Text)
                    Dim Diferenta_start As Double = start1 - 1


                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim H_EXAG As Double = 1000 / CDbl(TextBox_Hscale.Text)
                    If H_EXAG = 0 Then H_EXAG = 1

                    Dim Chainage_cunoscuta As Double = 0





                    Dim Linia_cunoscuta As Autodesk.AutoCAD.DatabaseServices.Line
                    Dim x01, x02 As Double
                    Dim y01, y02 As Double

                    Dim UCS_Curent As Matrix3d = Editor1.CurrentUserCoordinateSystem


                    Linia_cunoscuta = Linia_chainage_zero
                    x01 = Linia_cunoscuta.StartPoint.TransformBy(UCS_Curent).X
                    x02 = Linia_cunoscuta.EndPoint.TransformBy(UCS_Curent).X
                    y01 = Linia_cunoscuta.StartPoint.TransformBy(UCS_Curent).Y
                    y02 = Linia_cunoscuta.EndPoint.TransformBy(UCS_Curent).Y

                    Creaza_layer(TextBox_Layer_NO_PLOT.Text, 40, "No Plot", False)

                    For i = 0 To Profile_table.Rows.Count - 1


                        If IsDBNull(Profile_table.Rows.Item(i).Item("Description")) = False Then
                            Dim Descriptie1 As String = Profile_table.Rows.Item(i).Item("Description")
                            Dim Descriptie2 As String = ""
                            If IsDBNull(Profile_table.Rows.Item(i).Item("Description_extra")) = False Then Descriptie2 = Profile_table.Rows.Item(i).Item("Description_extra")

                            If Not Replace(Descriptie1, " ", "") = "" Then
                                If IsDBNull(Profile_table.Rows.Item(i).Item("Chainage_Recalculated")) = False Then
                                    If IsNumeric(Profile_table.Rows.Item(i).Item("Chainage_Recalculated")) = True Then
                                        Dim Valoare_chainage_la_punct As Double = CDbl(Profile_table.Rows.Item(i).Item("Chainage_Recalculated"))
                                        Dim Block_name_string As String = ComboBox_blocks.Text
                                        Dim Chainage_string As String = Get_chainage_from_double(Profile_table.Rows.Item(i).Item("Chainage_Recalculated"), 1)


                                        'de aici sunt bLOCURI

                                        Dim x, y As Double
                                        x = x01 + (Valoare_chainage_la_punct - Chainage_cunoscuta) * H_EXAG
                                        y = (y01 + y02) / 2

                                        Dim Linie1 As New Line
                                        Linie1.StartPoint = New Point3d(x, y01, 0)
                                        Linie1.EndPoint = New Point3d(x, y02, 0)
                                        Linie1.Layer = TextBox_Layer_NO_PLOT.Text
                                        BTrecord.AppendEntity(Linie1)
                                        Trans1.AddNewlyCreatedDBObject(Linie1, True)

                                        Dim Col_int As New Point3dCollection

                                        Linie1.IntersectWith(Polylinie_profil, Intersect.ExtendThis, Col_int, IntPtr.Zero, IntPtr.Zero)

                                        If Col_int.Count > 0 Then
                                            y = Col_int(0).Y
                                        End If


                                        Dim Scale1 As Double = 1 '/ Viewport_scale

                                        Dim Block1 As Autodesk.AutoCAD.DatabaseServices.BlockReference
                                        If Replace(Descriptie2, " ", "") = "" Then
                                            Block1 = Insereaza_Block_cu_2atribute(Block_name_string, New Autodesk.AutoCAD.Geometry.Point3d(x, y, 0), Scale1, ComboBox_atrib_chainage.Text, Chainage_string, _
                                                                                  ComboBox_atrib_description_1.Text, Descriptie1, ComboBox_LAYER_TEXT_AND_BLOCKS.Text, BTrecord)

                                        Else
                                            Block1 = Insereaza_Block_cu_3atribute(Block_name_string, New Autodesk.AutoCAD.Geometry.Point3d(x, y, 0), Scale1, ComboBox_atrib_chainage.Text, Chainage_string, _
                                                                               ComboBox_atrib_description_1.Text, Descriptie1, ComboBox_atrib_description_2.Text, Descriptie2, ComboBox_LAYER_TEXT_AND_BLOCKS.Text, BTrecord)
                                        End If




                                        Dim MTEXT1 As New MText
                                        MTEXT1.Contents = Chainage_string & "-" & Descriptie1 & " " & Descriptie2
                                        MTEXT1.Location = New Point3d(x, y01 + 1, 0)
                                        MTEXT1.Rotation = PI / 2
                                        MTEXT1.Layer = TextBox_Layer_NO_PLOT.Text
                                        MTEXT1.TextHeight = 1
                                        MTEXT1.ColorIndex = 7
                                        MTEXT1.Attachment = AttachmentPoint.MiddleLeft
                                        BTrecord.AppendEntity(MTEXT1)
                                        Trans1.AddNewlyCreatedDBObject(MTEXT1, True)
                                    End If
                                End If
                            End If


                        End If





                        'asta e de la INSERT BLOCKS
                    Next

                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using

                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                ' asta e de la lock
            End Using



        Catch ex As Exception

            MsgBox(ex.Message)
        End Try
    End Sub
    Public Overridable Function Insereaza_Block_cu_2atribute(ByVal Nume_block As String, ByVal Punct_inserare As Point3d, ByVal ScaleXYZ As Double, ByVal Atribut_field_name1 As String, ByVal Atribut_valoare1 As String, _
                                                             ByVal Atribut_field_name2 As String, ByVal Atribut_valoare2 As String, ByVal Layer1 As String, ByVal Spatiu As BlockTableRecord)
        Dim dlock As DocumentLock = Nothing
        Dim bt As BlockTable
        Dim Block_table_record1 As BlockTableRecord = Nothing
        Dim br As BlockReference
        Dim id As ObjectId
        Dim db As Autodesk.AutoCAD.DatabaseServices.Database = HostApplicationServices.WorkingDatabase
        Using trans As Transaction = db.TransactionManager.StartTransaction
            Dim ed As Autodesk.AutoCAD.EditorInput.Editor = Application.DocumentManager.MdiActiveDocument.Editor

            'insert block and rename it
            Try
                Try
                    dlock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                Catch ex As Exception
                    Dim aex As New System.Exception("Error locking document for InsertBlock: " & Nume_block & ": ", ex)
                    Throw aex
                End Try
                bt = trans.GetObject(db.BlockTableId, OpenMode.ForWrite)
                If bt.Has(Nume_block) Then
                    Block_table_record1 = trans.GetObject(bt.Item(Nume_block), OpenMode.ForRead)

                    Spatiu = trans.GetObject(Spatiu.ObjectId, OpenMode.ForWrite)
                    'Set the Attribute Value
                    Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection
                    Dim ent As Entity
                    Dim Block_table_record1enum As BlockTableRecordEnumerator
                    br = New BlockReference(Punct_inserare, Block_table_record1.ObjectId)
                    br.Layer = Layer1
                    br.ScaleFactors = New Autodesk.AutoCAD.Geometry.Scale3d(ScaleXYZ, ScaleXYZ, ScaleXYZ)

                    Spatiu.AppendEntity(br)
                    trans.AddNewlyCreatedDBObject(br, True)
                    If Not Atribut_field_name1 = "" Or Not Atribut_field_name2 = "" Then
                        attColl = br.AttributeCollection
                        Block_table_record1enum = Block_table_record1.GetEnumerator
                        While Block_table_record1enum.MoveNext
                            ent = Block_table_record1enum.Current.GetObject(OpenMode.ForWrite)
                            If TypeOf ent Is AttributeDefinition Then
                                Dim attdef As AttributeDefinition = ent
                                Dim attref As New AttributeReference
                                attref.SetAttributeFromBlock(attdef, br.BlockTransform)
                                'attref.TextString = attref.Tag

                                If Not Atribut_valoare1 = "" Then
                                    If attref.Tag = Atribut_field_name1 Then
                                        attref.TextString = Atribut_valoare1
                                    End If
                                End If

                                If Not Atribut_valoare2 = "" Then
                                    If attref.Tag = Atribut_field_name2 Then
                                        attref.TextString = Atribut_valoare2
                                    End If
                                End If


                                attColl.AppendAttribute(attref)
                                trans.AddNewlyCreatedDBObject(attref, True)


                            End If
                        End While
                    End If 'daca avem valoare la atribut

                    trans.Commit()
                End If ' ASTA E DE LA  If bt.Has(Nume_block)

            Catch ex As System.Exception
                Dim aex2 As New System.Exception("Error in inserting new block: " & Nume_block & ": ", ex)
                Throw aex2
            Finally
                If Not trans Is Nothing Then trans.Dispose()
                If Not dlock Is Nothing Then dlock.Dispose()
            End Try
        End Using
        Return br
    End Function

    Public Overridable Function Insereaza_Block_cu_3atribute(ByVal Nume_block As String, ByVal Punct_inserare As Point3d, ByVal ScaleXYZ As Double, ByVal Atribut_field_name1 As String, ByVal Atribut_valoare1 As String, _
                                                         ByVal Atribut_field_name2 As String, ByVal Atribut_valoare2 As String, ByVal Atribut_field_name3 As String, ByVal Atribut_valoare3 As String, ByVal Layer1 As String, ByVal Spatiu As BlockTableRecord)
        Dim dlock As DocumentLock = Nothing
        Dim bt As BlockTable
        Dim Block_table_record1 As BlockTableRecord = Nothing
        Dim br As BlockReference
        Dim id As ObjectId
        Dim db As Autodesk.AutoCAD.DatabaseServices.Database = HostApplicationServices.WorkingDatabase
        Using trans As Transaction = db.TransactionManager.StartTransaction
            Dim ed As Autodesk.AutoCAD.EditorInput.Editor = Application.DocumentManager.MdiActiveDocument.Editor

            'insert block and rename it
            Try
                Try
                    dlock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                Catch ex As Exception
                    Dim aex As New System.Exception("Error locking document for InsertBlock: " & Nume_block & ": ", ex)
                    Throw aex
                End Try
                bt = trans.GetObject(db.BlockTableId, OpenMode.ForWrite)
                If bt.Has(Nume_block) Then
                    Block_table_record1 = trans.GetObject(bt.Item(Nume_block), OpenMode.ForRead)

                    Spatiu = trans.GetObject(Spatiu.ObjectId, OpenMode.ForWrite)
                    'Set the Attribute Value
                    Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection
                    Dim ent As Entity
                    Dim Block_table_record1enum As BlockTableRecordEnumerator
                    br = New BlockReference(Punct_inserare, Block_table_record1.ObjectId)
                    br.Layer = Layer1
                    br.ScaleFactors = New Autodesk.AutoCAD.Geometry.Scale3d(ScaleXYZ, ScaleXYZ, ScaleXYZ)

                    Spatiu.AppendEntity(br)
                    trans.AddNewlyCreatedDBObject(br, True)
                    If Not Atribut_field_name1 = "" Or Not Atribut_field_name2 = "" Or Not Atribut_field_name3 = "" Then
                        attColl = br.AttributeCollection
                        Block_table_record1enum = Block_table_record1.GetEnumerator
                        While Block_table_record1enum.MoveNext
                            ent = Block_table_record1enum.Current.GetObject(OpenMode.ForWrite)
                            If TypeOf ent Is AttributeDefinition Then
                                Dim attdef As AttributeDefinition = ent
                                Dim attref As New AttributeReference
                                attref.SetAttributeFromBlock(attdef, br.BlockTransform)
                                'attref.TextString = attref.Tag

                                If Not Atribut_valoare1 = "" Then
                                    If attref.Tag = Atribut_field_name1 Then
                                        attref.TextString = Atribut_valoare1
                                    End If
                                End If

                                If Not Atribut_valoare2 = "" Then
                                    If attref.Tag = Atribut_field_name2 Then
                                        attref.TextString = Atribut_valoare2
                                    End If
                                End If

                                If Not Atribut_valoare3 = "" Then
                                    If attref.Tag = Atribut_field_name3 Then
                                        attref.TextString = Atribut_valoare3
                                    End If
                                End If

                                attColl.AppendAttribute(attref)
                                trans.AddNewlyCreatedDBObject(attref, True)


                            End If
                        End While
                    End If 'daca avem valoare la atribut

                    trans.Commit()
                End If ' ASTA E DE LA  If bt.Has(Nume_block)

            Catch ex As System.Exception
                Dim aex2 As New System.Exception("Error in inserting new block: " & Nume_block & ": ", ex)
                Throw aex2
            Finally
                If Not trans Is Nothing Then trans.Dispose()
                If Not dlock Is Nothing Then dlock.Dispose()
            End Try
        End Using
        Return br
    End Function

    Public Sub CREAZA_LEADER(ByVal Punct3D As Point3d, ByVal Text_content As String, ByVal Mtext_Height As Double, ByVal Landing_gap As Double, ByVal Arrow_size As Double, ByVal Dog_length As Double, ByVal Layer As String)
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
        ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument

                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                    Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                    Dim Mtext1 As New MText
                    Mtext1.Contents = Text_content
                    Mtext1.TextHeight = Mtext_Height
                    Mtext1.ColorIndex = 0


                    Dim Curent_UCS As Matrix3d = Editor1.CurrentUserCoordinateSystem

                    Dim Mleader1 As New MLeader
                    Dim Nr1 As Integer = Mleader1.AddLeader()
                    Dim Nr2 As Integer = Mleader1.AddLeaderLine(Nr1)
                    Mleader1.AddFirstVertex(Nr2, Punct3D.TransformBy(Curent_UCS))
                    Mleader1.AddLastVertex(Nr2, New Point3d(Punct3D.TransformBy(Curent_UCS).X + 1,
                                                           Punct3D.TransformBy(Curent_UCS).Y + 1, 0))

                    Mleader1.ContentType = ContentType.MTextContent
                    Mleader1.MText = Mtext1
                    Mleader1.Layer = Layer
                    Mleader1.LandingGap = Landing_gap
                    Mleader1.ArrowSize = Arrow_size
                    Mleader1.DoglegLength = Dog_length

                    Mleader1.TransformBy(Matrix3d.Displacement(New Point3d(0, 0, 0).GetVectorTo(New Point3d(0, 0, Punct3D.Z))))


                    BTrecord.AppendEntity(Mleader1)
                    Trans1.AddNewlyCreatedDBObject(Mleader1, True)
                    Trans1.Commit()
                End Using
            End Using
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub pn_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_point_name.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Or e.KeyCode = Windows.Forms.Keys.Tab Then
            With TextBox_East
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub
    Private Sub TextBox_point_name_TextChanged(sender As Object, e As EventArgs) Handles TextBox_point_name.TextChanged
        With TextBox_point_name
            .Text = .Text.ToUpper
        End With
    End Sub
    Private Sub chainage_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_CHAINAGE.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Or e.KeyCode = Windows.Forms.Keys.Tab Then
            With TextBox_elevation
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub
    Private Sub TextBox_CHAINAGE_TextChanged(sender As Object, e As EventArgs) Handles TextBox_CHAINAGE.TextChanged
        With TextBox_CHAINAGE
            .Text = .Text.ToUpper
        End With
    End Sub
    Private Sub east_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_East.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Or e.KeyCode = Windows.Forms.Keys.Tab Then
            With TextBox_NORTH
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub
    Private Sub TextBox_East_TextChanged(sender As Object, e As EventArgs) Handles TextBox_East.TextChanged
        With TextBox_East
            .Text = .Text.ToUpper
        End With
    End Sub
    Private Sub north_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_NORTH.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Or e.KeyCode = Windows.Forms.Keys.Tab Then
            With TextBox_elevation
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub
    Private Sub TextBox_NORTH_TextChanged(sender As Object, e As EventArgs) Handles TextBox_NORTH.TextChanged
        With TextBox_NORTH
            .Text = .Text.ToUpper
        End With
    End Sub
    Private Sub elev_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_elevation.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Or e.KeyCode = Windows.Forms.Keys.Tab Then
            With TextBox_Description
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub
    Private Sub TextBox_elevation_TextChanged(sender As Object, e As EventArgs) Handles TextBox_elevation.TextChanged
        With TextBox_elevation
            .Text = .Text.ToUpper
        End With
    End Sub
    Private Sub descr_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_Description.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Or e.KeyCode = Windows.Forms.Keys.Tab Then
            With TextBox_description_extra
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub
    Private Sub TextBox_Description_TextChanged(sender As Object, e As EventArgs) Handles TextBox_Description.TextChanged
        With TextBox_Description
            .Text = .Text.ToUpper
        End With
    End Sub
    Private Sub TextBox_description_extra_kd(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_description_extra.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Or e.KeyCode = Windows.Forms.Keys.Tab Then
            With TextBox_row_start
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub
    Private Sub TextBox_description_extra_TextChanged(sender As Object, e As EventArgs) Handles TextBox_description_extra.TextChanged
        With TextBox_description_extra
            .Text = .Text.ToUpper
        End With
    End Sub

    Private Sub chainage2_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_chainage2.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Or e.KeyCode = Windows.Forms.Keys.Tab Then
            With TextBox_elevation2
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub
    Private Sub TextBox_chainage2_TextChanged(sender As Object, e As EventArgs) Handles TextBox_chainage2.TextChanged
        With TextBox_chainage2
            .Text = .Text.ToUpper
        End With
    End Sub
    Private Sub east2_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_East2.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Or e.KeyCode = Windows.Forms.Keys.Tab Then
            With TextBox_NORTH2
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub
    Private Sub TextBox_East2_TextChanged(sender As Object, e As EventArgs) Handles TextBox_East2.TextChanged
        With TextBox_East2
            .Text = .Text.ToUpper
        End With
    End Sub
    Private Sub north2_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_NORTH2.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Or e.KeyCode = Windows.Forms.Keys.Tab Then
            With TextBox_elevation2
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub
    Private Sub TextBox_NORTH2_TextChanged(sender As Object, e As EventArgs) Handles TextBox_NORTH2.TextChanged
        With TextBox_NORTH2
            .Text = .Text.ToUpper
        End With
    End Sub
    Private Sub elev2_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_elevation2.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Or e.KeyCode = Windows.Forms.Keys.Tab Then
            With TextBox_row_start
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub
    Private Sub TextBox_elevation2_TextChanged(sender As Object, e As EventArgs) Handles TextBox_elevation2.TextChanged
        With TextBox_elevation2
            .Text = .Text.ToUpper
        End With
    End Sub
    Private Sub strt_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_row_start.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Or e.KeyCode = Windows.Forms.Keys.Tab Then
            With TextBox_row_end
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub
    Private Sub prtrt_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_printing_scale.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Or e.KeyCode = Windows.Forms.Keys.Tab Then
            With TextBox_viewport_scale
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub
    Private Sub TextBox_viewport_scale_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_viewport_scale.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Or e.KeyCode = Windows.Forms.Keys.Tab Then
            With TextBox_Hscale
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub
    Private Sub TextBox_Hscale_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_Hscale.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Or e.KeyCode = Windows.Forms.Keys.Tab Then
            With TextBox_Vscale
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub
    Private Sub TextBox_Vscale_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_Vscale.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Or e.KeyCode = Windows.Forms.Keys.Tab Then
            With TextBox_Hincr
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub
    Private Sub TextBox_Hincr_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_Hincr.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Or e.KeyCode = Windows.Forms.Keys.Tab Then
            With TextBox_Vincr
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub
    Private Sub TextBox_vincr_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_Vincr.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Or e.KeyCode = Windows.Forms.Keys.Tab Then
            With TextBox_L_elevation
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub
    Private Sub TextBox_L_elevation_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_L_elevation.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Or e.KeyCode = Windows.Forms.Keys.Tab Then
            With TextBox_H_Elevation
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub
    Private Sub TextBox_h_elevation_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_H_Elevation.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Or e.KeyCode = Windows.Forms.Keys.Tab Then
            With TextBox_Minimum_chainage
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub
    Private Sub TextBox_Minimum_chainage_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_Minimum_chainage.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Or e.KeyCode = Windows.Forms.Keys.Tab Then
            With TextBox_Maximum_chainage
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub
    Private Sub TextBox_Maximum_chainage_key_d(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_Maximum_chainage.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Or e.KeyCode = Windows.Forms.Keys.Tab Then
            With TextBox_row_0
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub



    Private Sub TextBox_color_index_grid_lines_key_d(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_color_index_grid_lines.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Or e.KeyCode = Windows.Forms.Keys.Tab Then
            With TextBox_text_height
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub





    Private Sub CheckBox_Hydrostatic_style_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_Hydrostatic_style.CheckedChanged
        Try
            If CheckBox_Hydrostatic_style.Checked = True Then
                ComboBox_text_styles.Text = "Standard"
                If ComboBox_layer_grid_lines.Items.Contains("Grid") = True Then
                    ComboBox_layer_grid_lines.Text = "Grid"
                End If
                If ComboBox_LINETYPE.Items.Contains("TCDASH2") = True Then
                    ComboBox_layer_grid_lines.Text = "TCDASH2"
                End If
                If ComboBox_LAYER_TEXT_AND_BLOCKS.Items.Contains("TEXT") = True Then
                    ComboBox_LAYER_TEXT_AND_BLOCKS.Text = "TEXT"
                End If
                If ComboBox_LAYER_PROFILE_POLYLINE.Items.Contains("PGRADE") = True Then
                    ComboBox_LAYER_PROFILE_POLYLINE.Text = "PGRADE"
                End If
                TextBox_Vscale.Text = 200
            Else
                TextBox_Vscale.Text = 1000
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub



    Private Sub CheckBox_add_description_on_graph_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_add_description_on_graph.CheckedChanged
        If CheckBox_add_description_on_graph.Checked = True Then
            Panel_BLOCKS.Visible = True
        Else
            Panel_BLOCKS.Visible = False
        End If
    End Sub

    Private Sub Panel_BLOCKS_Click(sender As Object, e As EventArgs) Handles Panel_BLOCKS.Click
        Incarca_existing_Blocks_with_attributes_to_combobox(ComboBox_blocks)
        ComboBox_atrib_description_1.Items.Clear()
        ComboBox_atrib_chainage.Items.Clear()
        ComboBox_atrib_description_2.Items.Clear()
    End Sub

    Private Sub Panel_Columns_Click(sender As Object, e As EventArgs) Handles Panel_Columns.Click
        Try
            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
            W1 = Get_active_worksheet_from_Excel()
            Dim Celula_a1 As Microsoft.Office.Interop.Excel.Range
            Celula_a1 = W1.Range("A1")

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub








End Class