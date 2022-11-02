Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class Alignment_engineering_Form

    Dim Extra_index_dupa_removal As Integer = 0
    Dim Data_table_Crossing As System.Data.DataTable
    Dim Nr_pagina As Integer
    Dim Freeze_operations As Boolean = False

    Private Sub Alignment_engineering_Form_Load(sender As Object, e As EventArgs) Handles Me.Load
        Data_table_Crossing = New System.Data.DataTable
        Data_table_Crossing.Columns.Add("DESCRIPTION1", GetType(String))
        Data_table_Crossing.Columns.Add("DESCRIPTION2", GetType(String))
        Data_table_Crossing.Columns.Add("ID_NO", GetType(String))
        Data_table_Crossing.Columns.Add("STA", GetType(Double))
        Data_table_Crossing.Columns.Add("BEGINSTA", GetType(Double))
        Data_table_Crossing.Columns.Add("ENDSTA", GetType(Double))
        Data_table_Crossing.Columns.Add("LENGTH", GetType(Double))
        Data_table_Crossing.Columns.Add("MATERIAL", GetType(String))
        Data_table_Crossing.Columns.Add("PREVIOUS_MATERIAL", GetType(String))
        Data_table_Crossing.Columns.Add("COVER", GetType(String))
        Data_table_Crossing.Columns.Add("CROSSING_TYPE", GetType(String))
        Data_table_Crossing.Columns.Add("BLOCKNAME", GetType(String))
        Data_table_Crossing.Columns.Add("SHEET", GetType(Double))
        Data_table_Crossing.Columns.Add("CP", GetType(String))
        Data_table_Crossing.Columns.Add("EXTRA_LENGTH", GetType(Double))
        Data_table_Crossing.Columns.Add("WARNING_SIGN", GetType(Boolean))
        Data_table_Crossing.Columns.Add("BUOYANCY", GetType(Double))
    End Sub

    Private Sub Button_Read_excel_Click(sender As Object, e As EventArgs) Handles Button_Read_excel.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                If IsNumeric(TextBox_PAGE_START_NO_excel.Text) = False Then
                    With TextBox_PAGE_START_NO_excel
                        .Text = ""
                        .Focus()
                    End With
                    MsgBox("Please specify the start page number:")
                    Freeze_operations = False
                    Exit Sub
                End If

                If TextBox_description1_COL_XL.Text = "" Then
                    MsgBox("Please specify the description 1 EXCEL COLUMN!")
                    Freeze_operations = False
                    Exit Sub
                End If

                If TextBox_description2_COL_XL.Text = "" Then
                    MsgBox("Please specify the description 2 EXCEL COLUMN!")
                    Freeze_operations = False
                    Exit Sub
                End If

                If TextBox_ID_NO_COL_XL.Text = "" Then
                    MsgBox("Please specify the ID_NO EXCEL COLUMN!")
                    Freeze_operations = False
                    Exit Sub
                End If

                If TextBox_chainage_col_xl.Text = "" Then
                    MsgBox("Please specify the station EXCEL COLUMN!")
                    Freeze_operations = False
                    Exit Sub
                End If

                If TextBox_LENGTH_col_xl.Text = "" Then
                    MsgBox("Please specify the length EXCEL COLUMN!")
                    Freeze_operations = False
                    Exit Sub
                End If

                If TextBox_MATERIAL_col_xl.Text = "" Then
                    MsgBox("Please specify the material EXCEL COLUMN!")
                    Freeze_operations = False
                    Exit Sub
                End If

                If TextBox_Cover_col_xl.Text = "" Then
                    Freeze_operations = False
                    Exit Sub
                End If

                If TextBox_Crossing_type_col_XL.Text = "" Then
                    MsgBox("Please specify the crossing type EXCEL COLUMN!")
                    Freeze_operations = False
                    Exit Sub
                End If

                If TextBox_Buoyancy_dist.Text = "" Then
                    MsgBox("Please specify the Buoyancy EXCEL COLUMN!")
                    Freeze_operations = False
                    Exit Sub
                End If

                If TextBox_ROW_START.Text = "" Then
                    MsgBox("Please specify the EXCEL START ROW!")
                    Freeze_operations = False
                    Exit Sub
                End If

                If IsNumeric(TextBox_ROW_START.Text) = False Then
                    With TextBox_ROW_START
                        .Text = ""
                        .Focus()
                    End With
                    MsgBox("Please specify start row")
                    Freeze_operations = False
                    Exit Sub
                End If
                If TextBox_ROW_END.Text = "" Then
                    MsgBox("Please specify the EXCEL END ROW!")
                    Freeze_operations = False
                    Exit Sub
                End If

                If IsNumeric(TextBox_ROW_END.Text) = False Then
                    With TextBox_ROW_END
                        .Text = ""
                        .Focus()
                    End With
                    MsgBox("Please specify END row")
                    Freeze_operations = False
                    Exit Sub
                End If

                If Val(TextBox_ROW_START.Text) < 1 Then
                    With TextBox_ROW_START
                        .Text = ""
                        .Focus()
                    End With
                    MsgBox("Start row can't be smaller than 1")
                    Freeze_operations = False
                    Exit Sub
                End If

                If Val(TextBox_ROW_END.Text) < 1 Then
                    With TextBox_ROW_END
                        .Text = ""
                        .Focus()
                    End With
                    MsgBox("End row can't be smaller than 1")
                    Freeze_operations = False
                    Exit Sub
                End If

                If Val(TextBox_ROW_END.Text) < Val(TextBox_ROW_START.Text) Then
                    MsgBox("END row smaller than start row")
                    Freeze_operations = False
                    Exit Sub
                End If

                If ComboBox_nps.Text = "" Or IsNumeric(ComboBox_nps.Text) = False Then
                    MsgBox("Specify The Nominal Pipe Size!")
                    Freeze_operations = False
                    Exit Sub
                End If

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
                Dim start1 As Integer = CInt(TextBox_ROW_START.Text)
                Dim end1 As Integer = CInt(TextBox_ROW_END.Text)
                Dim Col1 As String = TextBox_description1_COL_XL.Text.ToUpper
                Dim Col2 As String = TextBox_description2_COL_XL.Text.ToUpper
                Dim Col3 As String = TextBox_ID_NO_COL_XL.Text.ToUpper
                Dim Col4 As String = TextBox_chainage_col_xl.Text.ToUpper
                Dim Col5 As String = TextBox_LENGTH_col_xl.Text.ToUpper
                Dim Col6 As String = TextBox_MATERIAL_col_xl.Text.ToUpper
                Dim Col7 As String = TextBox_Cover_col_xl.Text.ToUpper
                Dim Col8 As String = TextBox_Crossing_type_col_XL.Text.ToUpper
                Panel_items_from_excel.Controls.Clear()
                Dim Y_index_panel As Integer = 0

                Extra_index_dupa_removal = 0
                Data_table_Crossing.Rows.Clear()
                Dim Index_Data_table As Integer = 0

                Nr_pagina = CInt(TextBox_PAGE_START_NO_excel.Text)

                Dim Chainage_previous As Double = 0
                Dim Cumulative_Distance_for_warning_signs As Double = 0
                Dim Maximum_distance_for_warning_signs As Double = 500
                Dim Minimum_elbow_angle_for_warning_signs As Double = 30

                Dim Station_previous_warning_sign As Double = 0


                For i = start1 To end1
                    Dim Description1 As String = W1.Range(Col1 & i).Value
                    Dim Description2 As String = W1.Range(Col2 & i).Value
                    Dim Drawing_Id_No As String = W1.Range(Col3 & i).Value

                    Dim Station As Double = -1
                    If IsNumeric(Replace(W1.Range(Col4 & i).Value, "+", "")) = True Then
                        Station = CDbl(Replace(W1.Range(Col4 & i).Value, "+", ""))


                        If Station < Chainage_previous Then
                            W1.Range(Col4 & i).Select()
                            MsgBox("The previous station is bigger than current station" & Col4 & i)
                            Freeze_operations = False
                            Exit Sub
                        End If

                        Cumulative_Distance_for_warning_signs = Station - Station_previous_warning_sign

                        Dim Length As Double = -1
                        If IsNumeric(W1.Range(Col5 & i).Value) = True Then
                            Length = Round(W1.Range(Col5 & i).Value, 1)
                        End If

                        Dim Cover As Double = -1
                        Dim Cover_text As String

                        If IsNumeric(W1.Range(Col7 & i).Value) = True Then
                            Cover = Round(W1.Range(Col7 & i).Value * 10, System.MidpointRounding.AwayFromZero) / 10
                        Else
                            Cover_text = W1.Range(Col7 & i).Value
                        End If

                        Dim Material As String = ""
                        If Not Replace(W1.Range(Col6 & i).Value, " ", "") = "" Then
                            Material = W1.Range(Col6 & i).Value
                        End If

                        If Material.ToUpper = "M" Then
                            Data_table_Crossing.Rows.Add()
                            Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = "M"
                            Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "MATCHLINE"
                            Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "MATCHLINE"
                            Data_table_Crossing.Rows(Index_Data_table).Item("STA") = Station
                            If Index_Data_table > 0 Then
                                For j = Index_Data_table - 1 To 0 Step -1
                                    If IsDBNull(Data_table_Crossing.Rows(j).Item("MATERIAL")) = False Then
                                        If IsNumeric(Data_table_Crossing.Rows(j).Item("MATERIAL")) = True Then
                                            Data_table_Crossing.Rows(Index_Data_table).Item("PREVIOUS_MATERIAL") = Data_table_Crossing.Rows(j).Item("MATERIAL")
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If

                            Dim TextBox11 As New Windows.Forms.TextBox
                            TextBox11.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
                            TextBox11.Location = New System.Drawing.Point(3, 3 + 27 * Y_index_panel)
                            TextBox11.Size = New System.Drawing.Size(127, 22)
                            TextBox11.Text = Get_chainage_from_double(Station, 1)
                            TextBox11.BackColor = Drawing.Color.White
                            TextBox11.ForeColor = Drawing.Color.Black
                            Panel_items_from_excel.Controls.Add(TextBox11)

                            Dim TextBox12 As New Windows.Forms.TextBox
                            TextBox12.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
                            TextBox12.Location = New System.Drawing.Point(136, 3 + 27 * Y_index_panel)
                            TextBox12.Size = New System.Drawing.Size(419, 22)

                            If Not i = start1 Then
                                TextBox12.Text = "MATCHLINE between page " & Nr_pagina & " and " & (Nr_pagina + 1).ToString
                            Else
                                TextBox12.Text = "MATCHLINE START PAGE " & Nr_pagina
                            End If

                            TextBox12.BackColor = Drawing.Color.White
                            TextBox12.ForeColor = Drawing.Color.Black
                            Panel_items_from_excel.Controls.Add(TextBox12)

                            Y_index_panel = Y_index_panel + 1
                            Data_table_Crossing.Rows(Index_Data_table).Item("SHEET") = Nr_pagina
                            If Not i = start1 Then Nr_pagina = Nr_pagina + 1
                            Index_Data_table = Index_Data_table + 1
                        End If

                        If Material = "T" Then
                            If IsNumeric(Replace(W1.Range(Col7 & i).Value, "TL ", "")) = True Then
                                Cover = Round(IsNumeric(Replace(W1.Range(Col7 & i).Value, "TL ", "")) * 10, System.MidpointRounding.AwayFromZero) / 10
                            End If
                        End If

                        Dim Crossing_type As String = W1.Range(Col8 & i).Value

                        If IsNothing(Description1) = True Then Description1 = ""
                        If IsNothing(Description2) = True Then Description2 = ""
                        If IsNothing(Drawing_Id_No) = True Then Drawing_Id_No = ""
                        If IsNothing(Crossing_type) = True Then Crossing_type = ""

                        If Not Replace(Description1, " ", "") = "" And Not Material.ToUpper = "M" Then
                            Data_table_Crossing.Rows.Add()
                            Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = Description1
                            Dim Buoyancy_dist As Double = 0
                            Dim Buoyancy_text As String = W1.Range(TextBox_Buoyancy_dist.Text & i).Value
                            If IsNumeric(Buoyancy_text) = True Then
                                Buoyancy_dist = CDbl(Buoyancy_text)
                                Data_table_Crossing.Rows(Index_Data_table).Item("BUOYANCY") = Buoyancy_dist
                            End If










                            If Not Replace(Description2, " ", "") = "" Then
                                Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION2") = Description2
                            End If

                            If Crossing_type.ToUpper = "CM" Or Crossing_type.ToUpper = "PL" Or Crossing_type.ToUpper = "RD" Then
                                If Not Cover = -1 Then
                                    Data_table_Crossing.Rows(Index_Data_table).Item("COVER") = Get_String_Rounded(Cover, 1) & "m CVR"
                                Else
                                    If Cover_text = "" Then
                                        If Not Crossing_type.ToUpper = "RD" Then
                                            Data_table_Crossing.Rows(Index_Data_table).Item("COVER") = "UNKN CVR"
                                        End If
                                    Else
                                        Data_table_Crossing.Rows(Index_Data_table).Item("COVER") = Cover_text
                                    End If
                                End If
                            End If

                            If Not Replace(Drawing_Id_No, " ", "") = "" Then
                                Data_table_Crossing.Rows(Index_Data_table).Item("ID_NO") = Drawing_Id_No
                            End If
                            If Not Replace(Material, " ", "") = "" Then
                                Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = Material

                                If Material.ToUpper = "T" Then
                                    Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "TRANSITION"
                                    If Description2.ToUpper = "FAKE" Then
                                        Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION2") = "FAKE"
                                    End If
                                    Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "Pipe_Transition"
                                    If Not Cover = -1 Then
                                        Data_table_Crossing.Rows(Index_Data_table).Item("EXTRA_LENGTH") = Cover
                                    End If
                                    If Index_Data_table > 0 Then
                                        For j = Index_Data_table - 1 To 0 Step -1
                                            If IsDBNull(Data_table_Crossing.Rows(j).Item("MATERIAL")) = False Then
                                                If IsNumeric(Data_table_Crossing.Rows(j).Item("MATERIAL")) = True Then
                                                    Data_table_Crossing.Rows(Index_Data_table).Item("PREVIOUS_MATERIAL") = Data_table_Crossing.Rows(j).Item("MATERIAL")
                                                    Exit For
                                                End If
                                            End If
                                        Next
                                    End If
                                End If

                                If Index_Data_table > 0 Then
                                    For j = Index_Data_table - 1 To 0 Step -1
                                        If IsDBNull(Data_table_Crossing.Rows(j).Item("MATERIAL")) = False Then
                                            If IsNumeric(Data_table_Crossing.Rows(j).Item("MATERIAL")) = True Then
                                                Data_table_Crossing.Rows(Index_Data_table).Item("PREVIOUS_MATERIAL") = Data_table_Crossing.Rows(j).Item("MATERIAL")
                                                Exit For
                                            End If
                                        End If
                                    Next
                                End If



                            End If
                            If Not Replace(Crossing_type, " ", "") = "" Then
                                Data_table_Crossing.Rows(Index_Data_table).Item("CROSSING_TYPE") = Crossing_type
                            End If

                            Data_table_Crossing.Rows(Index_Data_table).Item("STA") = Station


                            Dim IS_elbow As Boolean = False
                            Dim Len_elbow As Double = -1
                            Dim Is_elbow_length_mistake As Boolean = False

                            If Description1.ToUpper.Contains("ELBOW") = True Or Description1.ToUpper.Contains("OVERBEND") = True Or Description1.ToUpper.Contains("SAGBEND") = True Then
                                IS_elbow = True
                                If Not Length = -1 Then
                                    Data_table_Crossing.Rows(Index_Data_table).Item("LENGTH") = Length
                                    Dim degree_elbow As String = extrage_numar_din_text_de_la_inceputul_textului(Description1.ToUpper)


                                    If IsNumeric(degree_elbow) = True Then
                                        If CheckBox_Warning_Signs.Checked = False Then
                                            If CheckBox_extra_warning_signs.Checked = True And CDbl(degree_elbow) >= Minimum_elbow_angle_for_warning_signs Then
                                                Data_table_Crossing.Rows(Index_Data_table).Item("WARNING_SIGN") = True
                                                Index_Data_table = Index_Data_table + Insert_extra_warning_signs(Maximum_distance_for_warning_signs, Cumulative_Distance_for_warning_signs, Station_previous_warning_sign)
                                                Station_previous_warning_sign = Station
                                                Cumulative_Distance_for_warning_signs = 0
                                            End If
                                        End If



                                        Dim Unghi As Double = CDbl(degree_elbow) * PI / 180
                                        If IsNumeric(ComboBox_nps.Text) = True Then
                                            Dim Diam As Double = 2 * Get_from_NPS_radius_for_pipes_from_inches_to_milimeters(CDbl(ComboBox_nps.Text)) / 1000

                                            Len_elbow = Round(2 * (3 * Diam * Tan(Unghi / 2) + 1), 3)

                                            If Abs(Length - 2 * (3 * Diam * Tan(Unghi / 2) + 1)) > 0.1 Then
                                                Is_elbow_length_mistake = True
                                                If CheckBox_elbow_mistake.Checked = True Then
                                                    MsgBox("3*D Elbow length issue at station " & Get_chainage_from_double(Station, 1) & vbCrLf _
                                                           & "a " & degree_elbow & " 3*D elbow length is " & Round(2 * (3 * Diam * Tan(Unghi / 2) + 1), 3) & "m")
                                                End If
                                            End If
                                        Else
                                            MsgBox("Specify the pipe diameter for station " & Get_chainage_from_double(Station, 1))

                                        End If
                                    Else
                                        MsgBox("3*D Elbow degree value missing at station " & Get_chainage_from_double(Station, 1))
                                    End If

                                    Dim StartCha As Double = Station - Length / 2
                                    Dim EndCha As Double = Station + Length / 2
                                    Data_table_Crossing.Rows(Index_Data_table).Item("BEGINSTA") = StartCha
                                    Data_table_Crossing.Rows(Index_Data_table).Item("ENDSTA") = EndCha

                                    If IsNumeric(Material) = False Then
                                        W1.Range(Col6 & i).Select()
                                        MsgBox("An elbow needs the material specified. See cell" & Col6 & i)
                                        Freeze_operations = False
                                        Exit Sub
                                    End If
                                Else
                                    W1.Range(Col5 & i).Select()
                                    MsgBox("Non numerical value at " & Col5 & i)
                                    Freeze_operations = False
                                    Exit Sub
                                End If
                                Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "Elbow_al"
                            End If

                            If Crossing_type.ToUpper.Contains("FACILITY") = True Then
                                Dim StartCha As Double = -1
                                Dim EndCha As Double = -1
                                If i > start1 Then
                                    If IsNumeric(Replace(W1.Range(Col4 & i - 1).Value, "+", "")) = True Then
                                        StartCha = Round(CDbl(Replace(W1.Range(Col4 & i - 1).Value, "+", "")), 1)
                                    End If
                                End If
                                If i < end1 Then
                                    If IsNumeric(Replace(W1.Range(Col4 & i + 1).Value, "+", "")) = True Then
                                        EndCha = Round(CDbl(Replace(W1.Range(Col4 & i + 1).Value, "+", "")), 1)
                                    End If
                                End If

                                If Not StartCha = -1 And Not EndCha = -1 Then
                                    Data_table_Crossing.Rows(Index_Data_table).Item("LENGTH") = Abs(EndCha - StartCha)
                                    Data_table_Crossing.Rows(Index_Data_table).Item("BEGINSTA") = StartCha
                                    Data_table_Crossing.Rows(Index_Data_table).Item("ENDSTA") = EndCha
                                Else
                                    W1.Range(Col4 & i).Select()
                                    MsgBox("Non numerical value at " & Col4 & i)
                                    Freeze_operations = False
                                    Exit Sub
                                End If
                                Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "FACILITY"
                            End If

                            If Crossing_type.ToUpper = "CM" Then
                                Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "Cable_alignment_crossing"
                                If CheckBox_Warning_Signs.Checked = False Then
                                    Data_table_Crossing.Rows(Index_Data_table).Item("WARNING_SIGN") = True


                                    Station_previous_warning_sign = Station
                                    Index_Data_table = Index_Data_table + Insert_extra_warning_signs(Maximum_distance_for_warning_signs, Cumulative_Distance_for_warning_signs, Station_previous_warning_sign)
                                    Station_previous_warning_sign = Station
                                    Cumulative_Distance_for_warning_signs = 0
                                End If
                            End If

                            If Crossing_type.ToUpper = "RD" Or Crossing_type.ToUpper = "WC" Or Crossing_type.ToUpper = "OT" Or Crossing_type.ToUpper = "RR" Or Crossing_type.ToUpper = "WC*" Then
                                Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "General_alignment_crossing"
                                If Crossing_type.ToUpper = "RD" Or Crossing_type.ToUpper = "WC" Or Crossing_type.ToUpper = "RR" Then
                                    If CheckBox_Warning_Signs.Checked = False Then
                                        Data_table_Crossing.Rows(Index_Data_table).Item("WARNING_SIGN") = True
                                        Index_Data_table = Index_Data_table + Insert_extra_warning_signs(Maximum_distance_for_warning_signs, Cumulative_Distance_for_warning_signs, Station_previous_warning_sign)
                                        Station_previous_warning_sign = Station
                                        Cumulative_Distance_for_warning_signs = 0
                                    End If
                                End If
                            End If

                            If Crossing_type.ToUpper = "WS" Then
                                Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "WARNING_SIGN"
                            End If

                            If Crossing_type.ToUpper = "PL" Then
                                Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "Pipe_alignment_crossing"
                                If CheckBox_Warning_Signs.Checked = False Then
                                    Data_table_Crossing.Rows(Index_Data_table).Item("WARNING_SIGN") = True
                                    Index_Data_table = Index_Data_table + Insert_extra_warning_signs(Maximum_distance_for_warning_signs, Cumulative_Distance_for_warning_signs, Station_previous_warning_sign)
                                    Station_previous_warning_sign = Station
                                    Cumulative_Distance_for_warning_signs = 0
                                End If
                            End If

                            If Replace(Description1.ToUpper, " ", "").Contains("TESTSECTION") = True Or Crossing_type.ToUpper = "TST" Then
                                Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "Test_section_alignment_crossing"
                            End If


                            If Material.ToUpper = "SA1" Then
                                Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = "SA1"
                                Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "SCREW ANCHOR START"
                                Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "SCREWANCHOR_START"
                            End If

                            If Material.ToUpper = "SA2" Then
                                Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = "SA2"
                                Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "SCREW ANCHOR END"
                                Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "SCREWANCHOR_END"
                            End If

                            If Material.ToUpper = "SBW1" Then
                                Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = "SBW1"
                                Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "SAND BAG START"
                                Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "SANDBAG_START"
                            End If

                            If Material.ToUpper = "SBW2" Then
                                Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = "SBW2"
                                Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "SAND BAG END"
                                Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "SANDBAG_END"
                            End If


                            If Material.ToUpper = "CCC1" Then
                                Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = "CCC1"
                                Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "CONCRETE START"
                                Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "CONCRETE_START"
                            End If

                            If Material.ToUpper = "CCC2" Then
                                Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = "CCC2"
                                Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "CONCRETE END"
                                Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "CONCRETE_END"
                            End If

                            If Material.ToUpper = "RW1" Then
                                Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = "RW1"
                                Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "RIVER WEIGHT START"
                                Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "RIVER_WEIGHT_START"
                            End If

                            If Material.ToUpper = "RW2" Then
                                Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = "RW2"
                                Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "RIVER WEIGHT END"
                                Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "RIVER_WEIGHT_END"
                            End If

                            If Material.ToUpper = "CP" Then
                                Dim Ch_2 As Double = 0
                                If i < end1 Then
                                    If IsNumeric(Replace(W1.Range(Col4 & i + 1).Value, "+", "")) = True Then
                                        Ch_2 = Round(CDbl(Replace(W1.Range(Col4 & i + 1).Value, "+", "")), 1)
                                    End If
                                End If

                                Dim Index_forCp As Integer = 0
                                For ab = 1 To 10
                                    If Index_Data_table - ab > 0 Then
                                        If IsDBNull(Data_table_Crossing.Rows(Index_Data_table - ab).Item("STA")) = False Then
                                            Index_forCp = Index_Data_table - ab
                                            Exit For
                                        End If
                                    End If

                                Next

                                If Index_forCp > 0 Then
                                    If Data_table_Crossing.Rows(Index_forCp).Item("STA") = Station Then
                                        If Data_table_Crossing.Rows(Index_forCp).Item("BLOCKNAME") = "Cable_alignment_crossing" _
                                            Or Data_table_Crossing.Rows(Index_forCp).Item("BLOCKNAME") = "General_alignment_crossing" _
                                            Or Data_table_Crossing.Rows(Index_forCp).Item("BLOCKNAME") = "Pipe_alignment_crossing" _
                                            Or Data_table_Crossing.Rows(Index_forCp).Item("BLOCKNAME") = "Test_section_alignment_crossing" Then
                                            Data_table_Crossing.Rows(Index_forCp).Item("CP") = "CP"
                                        End If

                                    ElseIf Ch_2 = Station Then
                                        Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = "CP"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "CATHODIC_PROTECTION"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("STA") = Station
                                    Else
                                        Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = "CP"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "CATHODIC_PROTECTION"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("CP") = "CP"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("STA") = Station
                                    End If
                                End If
                            End If

                            Dim TextBox11 As New Windows.Forms.TextBox
                            TextBox11.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
                            TextBox11.Location = New System.Drawing.Point(3, 3 + 27 * Y_index_panel)
                            TextBox11.Size = New System.Drawing.Size(127, 22)
                            TextBox11.Text = Get_chainage_from_double(Station, 1)
                            TextBox11.BackColor = Drawing.Color.White
                            TextBox11.ForeColor = Drawing.Color.Black

                            Panel_items_from_excel.Controls.Add(TextBox11)

                            Dim txt_listbox = Description1
                            If Not Description2 = "" Then txt_listbox = txt_listbox & " " & Description2
                            If IS_elbow = True Then txt_listbox = txt_listbox & " [" & Length & "m]"

                            Dim TextBox12 As New Windows.Forms.TextBox
                            TextBox12.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
                            TextBox12.Location = New System.Drawing.Point(136, 3 + 27 * Y_index_panel)
                            TextBox12.Size = New System.Drawing.Size(419, 22)
                            TextBox12.Text = txt_listbox
                            TextBox12.BackColor = Drawing.Color.White
                            TextBox12.ForeColor = Drawing.Color.Black

                            Panel_items_from_excel.Controls.Add(TextBox12)

                            If IS_elbow = True Then
                                Dim TextBox13 As New Windows.Forms.TextBox
                                TextBox13.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
                                TextBox13.Location = New System.Drawing.Point(561, 3 + 27 * Y_index_panel)
                                TextBox13.Size = New System.Drawing.Size(60, 22)
                                TextBox13.ForeColor = Drawing.Color.Black
                                TextBox13.BackColor = Drawing.Color.White
                                TextBox13.Text = Len_elbow & "m"

                                If Is_elbow_length_mistake = True Then
                                    TextBox12.BackColor = Drawing.Color.Red
                                End If

                                Panel_items_from_excel.Controls.Add(TextBox13)
                            End If


                            Y_index_panel = Y_index_panel + 1

                            Index_Data_table = Index_Data_table + 1
                        Else '  If Not Replace(Description1, " ", "") = ""

                            If Not Material.ToUpper = "M" Then
                                If Crossing_type.ToUpper.Contains("FACILITY") = True Then
                                    Dim StartCha As Double = -1
                                    Dim EndCha As Double = -1
                                    If i > start1 Then
                                        If IsNumeric(Replace(W1.Range(Col4 & i - 1).Value, "+", "")) = True Then
                                            StartCha = Round(CDbl(Replace(W1.Range(Col4 & i - 1).Value, "+", "")), 1)
                                        End If
                                    End If
                                    If i < end1 Then
                                        If IsNumeric(Replace(W1.Range(Col4 & i + 1).Value, "+", "")) = True Then
                                            EndCha = Round(CDbl(Replace(W1.Range(Col4 & i + 1).Value, "+", "")), 1)
                                        End If
                                    End If
                                    If Not StartCha = -1 And Not EndCha = -1 Then
                                        Data_table_Crossing.Rows.Add()
                                        If Not Replace(Material, " ", "") = "" Then Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = Material
                                        Data_table_Crossing.Rows(Index_Data_table).Item("LENGTH") = Abs(EndCha - StartCha)
                                        Data_table_Crossing.Rows(Index_Data_table).Item("BEGINSTA") = StartCha
                                        Data_table_Crossing.Rows(Index_Data_table).Item("ENDSTA") = EndCha
                                        Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "FACILITY"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("SHEET") = Nr_pagina

                                        Dim TextBox11 As New Windows.Forms.TextBox
                                        TextBox11.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
                                        TextBox11.Location = New System.Drawing.Point(3, 3 + 27 * Y_index_panel)
                                        TextBox11.Size = New System.Drawing.Size(127, 22)
                                        TextBox11.Text = Get_chainage_from_double(Station, 1)
                                        TextBox11.BackColor = Drawing.Color.White
                                        TextBox11.ForeColor = Drawing.Color.Black
                                        Panel_items_from_excel.Controls.Add(TextBox11)

                                        Dim TextBox12 As New Windows.Forms.TextBox
                                        TextBox12.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
                                        TextBox12.Location = New System.Drawing.Point(136, 3 + 27 * Y_index_panel)
                                        TextBox12.Size = New System.Drawing.Size(419, 22)
                                        TextBox12.Text = "FACILITY"
                                        TextBox12.BackColor = Drawing.Color.White
                                        TextBox12.ForeColor = Drawing.Color.Black
                                        Panel_items_from_excel.Controls.Add(TextBox12)

                                        Y_index_panel = Y_index_panel + 1

                                        Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "FACILITY"
                                        Index_Data_table = Index_Data_table + 1
                                    Else
                                        W1.Range(Col4 & i).Select()
                                        MsgBox("Non numerical value at " & Col4 & i)
                                        Freeze_operations = False
                                        Exit Sub
                                    End If
                                End If




                                Select Case Material.ToUpper
                                    Case "T"

                                        Data_table_Crossing.Rows.Add()
                                        Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = Material.ToUpper
                                        Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "TRANSITION"
                                        If Not Replace(Station, " ", "") = "" Then
                                            Data_table_Crossing.Rows(Index_Data_table).Item("STA") = Station
                                        Else
                                            If Index_Data_table = 0 Then
                                                Data_table_Crossing.Rows(Index_Data_table).Item("STA") = "0"
                                                Station = 0
                                            End If
                                        End If
                                        If Not Cover = -1 Then
                                            Data_table_Crossing.Rows(Index_Data_table).Item("EXTRA_LENGTH") = Cover
                                        End If
                                        If Index_Data_table > 0 Then
                                            For j = Index_Data_table - 1 To 0 Step -1
                                                If IsDBNull(Data_table_Crossing.Rows(j).Item("MATERIAL")) = False Then
                                                    If IsNumeric(Data_table_Crossing.Rows(j).Item("MATERIAL")) = True Then
                                                        Data_table_Crossing.Rows(Index_Data_table).Item("PREVIOUS_MATERIAL") = Data_table_Crossing.Rows(j).Item("MATERIAL")
                                                        Exit For
                                                    End If
                                                End If
                                            Next
                                        End If
                                        Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "Pipe_Transition"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("SHEET") = Nr_pagina
                                        Dim TextBox11 As New Windows.Forms.TextBox
                                        TextBox11.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
                                        TextBox11.Location = New System.Drawing.Point(3, 3 + 27 * Y_index_panel)
                                        TextBox11.Size = New System.Drawing.Size(127, 22)
                                        TextBox11.Text = Get_chainage_from_double(Station, 1)
                                        TextBox11.BackColor = Drawing.Color.White
                                        TextBox11.ForeColor = Drawing.Color.Black
                                        Panel_items_from_excel.Controls.Add(TextBox11)

                                        Dim TextBox12 As New Windows.Forms.TextBox
                                        TextBox12.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
                                        TextBox12.Location = New System.Drawing.Point(136, 3 + 27 * Y_index_panel)
                                        TextBox12.Size = New System.Drawing.Size(419, 22)
                                        TextBox12.Text = "TRANSITION"
                                        TextBox12.BackColor = Drawing.Color.White
                                        TextBox12.ForeColor = Drawing.Color.Black
                                        Panel_items_from_excel.Controls.Add(TextBox12)

                                        Y_index_panel = Y_index_panel + 1
                                        Index_Data_table = Index_Data_table + 1



                                    Case "SA1"

                                        Data_table_Crossing.Rows.Add()
                                        Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = "SA1"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "SCREW ANCHOR START"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("STA") = Station

                                        Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "SCREWANCHOR_START"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("SHEET") = Nr_pagina

                                        Dim TextBox11 As New Windows.Forms.TextBox
                                        TextBox11.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
                                        TextBox11.Location = New System.Drawing.Point(3, 3 + 27 * Y_index_panel)
                                        TextBox11.Size = New System.Drawing.Size(127, 22)
                                        TextBox11.Text = Get_chainage_from_double(Station, 1)
                                        TextBox11.BackColor = Drawing.Color.White
                                        TextBox11.ForeColor = Drawing.Color.Black
                                        Panel_items_from_excel.Controls.Add(TextBox11)

                                        Dim TextBox12 As New Windows.Forms.TextBox
                                        TextBox12.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
                                        TextBox12.Location = New System.Drawing.Point(136, 3 + 27 * Y_index_panel)
                                        TextBox12.Size = New System.Drawing.Size(419, 22)
                                        TextBox12.Text = "SCREW ANCHOR START"
                                        TextBox12.BackColor = Drawing.Color.White
                                        TextBox12.ForeColor = Drawing.Color.Black
                                        Panel_items_from_excel.Controls.Add(TextBox12)

                                        Y_index_panel = Y_index_panel + 1

                                        Index_Data_table = Index_Data_table + 1

                                    Case "SA2"

                                        Data_table_Crossing.Rows.Add()
                                        Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = "SA2"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "SCREW ANCHOR END"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("STA") = Station

                                        Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "SCREWANCHOR_END"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("SHEET") = Nr_pagina

                                        Dim TextBox11 As New Windows.Forms.TextBox
                                        TextBox11.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
                                        TextBox11.Location = New System.Drawing.Point(3, 3 + 27 * Y_index_panel)
                                        TextBox11.Size = New System.Drawing.Size(127, 22)
                                        TextBox11.Text = Get_chainage_from_double(Station, 1)
                                        TextBox11.BackColor = Drawing.Color.White
                                        TextBox11.ForeColor = Drawing.Color.Black
                                        Panel_items_from_excel.Controls.Add(TextBox11)

                                        Dim TextBox12 As New Windows.Forms.TextBox
                                        TextBox12.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
                                        TextBox12.Location = New System.Drawing.Point(136, 3 + 27 * Y_index_panel)
                                        TextBox12.Size = New System.Drawing.Size(419, 22)
                                        TextBox12.Text = "SCREW ANCHOR END"
                                        TextBox12.BackColor = Drawing.Color.White
                                        TextBox12.ForeColor = Drawing.Color.Black
                                        Panel_items_from_excel.Controls.Add(TextBox12)

                                        Y_index_panel = Y_index_panel + 1


                                        Index_Data_table = Index_Data_table + 1

                                    Case "SBW1"

                                        Data_table_Crossing.Rows.Add()
                                        Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = "SBW1"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "SAND BAG START"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("STA") = Station
                                        Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "SANDBAG_START"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("SHEET") = Nr_pagina

                                        Dim TextBox11 As New Windows.Forms.TextBox
                                        TextBox11.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
                                        TextBox11.Location = New System.Drawing.Point(3, 3 + 27 * Y_index_panel)
                                        TextBox11.Size = New System.Drawing.Size(127, 22)
                                        TextBox11.Text = Get_chainage_from_double(Station, 1)
                                        TextBox11.BackColor = Drawing.Color.White
                                        TextBox11.ForeColor = Drawing.Color.Black
                                        Panel_items_from_excel.Controls.Add(TextBox11)

                                        Dim TextBox12 As New Windows.Forms.TextBox
                                        TextBox12.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
                                        TextBox12.Location = New System.Drawing.Point(136, 3 + 27 * Y_index_panel)
                                        TextBox12.Size = New System.Drawing.Size(419, 22)
                                        TextBox12.Text = "SAND BAG START"
                                        TextBox12.BackColor = Drawing.Color.White
                                        TextBox12.ForeColor = Drawing.Color.Black
                                        Panel_items_from_excel.Controls.Add(TextBox12)

                                        Y_index_panel = Y_index_panel + 1


                                        Index_Data_table = Index_Data_table + 1

                                    Case "SBW2"

                                        Data_table_Crossing.Rows.Add()
                                        Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = "SBW2"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "SAND BAG END"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("STA") = Station

                                        Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "SANDBAG_END"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("SHEET") = Nr_pagina

                                        Dim TextBox11 As New Windows.Forms.TextBox
                                        TextBox11.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
                                        TextBox11.Location = New System.Drawing.Point(3, 3 + 27 * Y_index_panel)
                                        TextBox11.Size = New System.Drawing.Size(127, 22)
                                        TextBox11.Text = Get_chainage_from_double(Station, 1)
                                        TextBox11.BackColor = Drawing.Color.White
                                        TextBox11.ForeColor = Drawing.Color.Black
                                        Panel_items_from_excel.Controls.Add(TextBox11)

                                        Dim TextBox12 As New Windows.Forms.TextBox
                                        TextBox12.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
                                        TextBox12.Location = New System.Drawing.Point(136, 3 + 27 * Y_index_panel)
                                        TextBox12.Size = New System.Drawing.Size(419, 22)
                                        TextBox12.Text = "SAND BAG END"
                                        TextBox12.BackColor = Drawing.Color.White
                                        TextBox12.ForeColor = Drawing.Color.Black
                                        Panel_items_from_excel.Controls.Add(TextBox12)

                                        Y_index_panel = Y_index_panel + 1


                                        Index_Data_table = Index_Data_table + 1

                                    Case "CCC1"

                                        Data_table_Crossing.Rows.Add()
                                        Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = "CCC1"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "CONCRETE START"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("STA") = Station


                                        Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "CONCRETE_START"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("SHEET") = Nr_pagina

                                        Dim TextBox11 As New Windows.Forms.TextBox
                                        TextBox11.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
                                        TextBox11.Location = New System.Drawing.Point(3, 3 + 27 * Y_index_panel)
                                        TextBox11.Size = New System.Drawing.Size(127, 22)
                                        TextBox11.Text = Get_chainage_from_double(Station, 1)
                                        TextBox11.BackColor = Drawing.Color.White
                                        TextBox11.ForeColor = Drawing.Color.Black
                                        Panel_items_from_excel.Controls.Add(TextBox11)

                                        Dim TextBox12 As New Windows.Forms.TextBox
                                        TextBox12.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
                                        TextBox12.Location = New System.Drawing.Point(136, 3 + 27 * Y_index_panel)
                                        TextBox12.Size = New System.Drawing.Size(419, 22)
                                        TextBox12.Text = "CONCRETE START"
                                        TextBox12.BackColor = Drawing.Color.White
                                        TextBox12.ForeColor = Drawing.Color.Black
                                        Panel_items_from_excel.Controls.Add(TextBox12)

                                        Y_index_panel = Y_index_panel + 1


                                        Index_Data_table = Index_Data_table + 1

                                    Case "CCC2"

                                        Data_table_Crossing.Rows.Add()
                                        Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = "CCC2"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "CONCRETE END"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("STA") = Station

                                        Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "CONCRETE_END"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("SHEET") = Nr_pagina

                                        Dim TextBox11 As New Windows.Forms.TextBox
                                        TextBox11.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
                                        TextBox11.Location = New System.Drawing.Point(3, 3 + 27 * Y_index_panel)
                                        TextBox11.Size = New System.Drawing.Size(127, 22)
                                        TextBox11.Text = Get_chainage_from_double(Station, 1)
                                        TextBox11.BackColor = Drawing.Color.White
                                        TextBox11.ForeColor = Drawing.Color.Black
                                        Panel_items_from_excel.Controls.Add(TextBox11)

                                        Dim TextBox12 As New Windows.Forms.TextBox
                                        TextBox12.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
                                        TextBox12.Location = New System.Drawing.Point(136, 3 + 27 * Y_index_panel)
                                        TextBox12.Size = New System.Drawing.Size(419, 22)
                                        TextBox12.Text = "CONCRETE END"
                                        TextBox12.BackColor = Drawing.Color.White
                                        TextBox12.ForeColor = Drawing.Color.Black
                                        Panel_items_from_excel.Controls.Add(TextBox12)

                                        Y_index_panel = Y_index_panel + 1


                                        Index_Data_table = Index_Data_table + 1


                                    Case "RW1"

                                        Data_table_Crossing.Rows.Add()
                                        Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = "RW1"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "RIVER WEIGHT START"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("STA") = Station

                                        Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "RIVER_WEIGHT"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("SHEET") = Nr_pagina

                                        Dim TextBox11 As New Windows.Forms.TextBox
                                        TextBox11.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
                                        TextBox11.Location = New System.Drawing.Point(3, 3 + 27 * Y_index_panel)
                                        TextBox11.Size = New System.Drawing.Size(127, 22)
                                        TextBox11.Text = Get_chainage_from_double(Station, 1)
                                        TextBox11.BackColor = Drawing.Color.White
                                        TextBox11.ForeColor = Drawing.Color.Black
                                        Panel_items_from_excel.Controls.Add(TextBox11)

                                        Dim TextBox12 As New Windows.Forms.TextBox
                                        TextBox12.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
                                        TextBox12.Location = New System.Drawing.Point(136, 3 + 27 * Y_index_panel)
                                        TextBox12.Size = New System.Drawing.Size(419, 22)
                                        TextBox12.Text = "RIVER WEIGHT START"
                                        TextBox12.BackColor = Drawing.Color.White
                                        TextBox12.ForeColor = Drawing.Color.Black
                                        Panel_items_from_excel.Controls.Add(TextBox12)

                                        Y_index_panel = Y_index_panel + 1

                                        Index_Data_table = Index_Data_table + 1




                                    Case "RW2"

                                        Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = "RW2"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "RIVER WEIGHT END"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("STA") = Station

                                        Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "RIVER_WEIGHT"
                                        Data_table_Crossing.Rows(Index_Data_table).Item("SHEET") = Nr_pagina
                                        Dim TextBox11 As New Windows.Forms.TextBox
                                        TextBox11.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
                                        TextBox11.Location = New System.Drawing.Point(3, 3 + 27 * Y_index_panel)
                                        TextBox11.Size = New System.Drawing.Size(127, 22)
                                        TextBox11.Text = Get_chainage_from_double(Station, 1)
                                        TextBox11.BackColor = Drawing.Color.White
                                        TextBox11.ForeColor = Drawing.Color.Black
                                        Panel_items_from_excel.Controls.Add(TextBox11)

                                        Dim TextBox12 As New Windows.Forms.TextBox
                                        TextBox12.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
                                        TextBox12.Location = New System.Drawing.Point(136, 3 + 27 * Y_index_panel)
                                        TextBox12.Size = New System.Drawing.Size(419, 22)
                                        TextBox12.Text = "RIVER WEIGHT END"
                                        TextBox12.BackColor = Drawing.Color.White
                                        TextBox12.ForeColor = Drawing.Color.Black
                                        Panel_items_from_excel.Controls.Add(TextBox12)

                                        Y_index_panel = Y_index_panel + 1

                                        Index_Data_table = Index_Data_table + 1




                                    Case "CP"




                                        Dim Ch_2 As Double = 0
                                        If i < end1 Then
                                            If IsNumeric(Replace(W1.Range(Col4 & i + 1).Value, "+", "")) = True Then
                                                Ch_2 = Round(CDbl(Replace(W1.Range(Col4 & i + 1).Value, "+", "")), 1)
                                            End If
                                        End If





                                        Dim Index_forCp As Integer = 0
                                        For ab = 1 To 10
                                            If Index_Data_table - ab > 0 Then
                                                If IsDBNull(Data_table_Crossing.Rows(Index_Data_table - ab).Item("STA")) = False Then
                                                    Index_forCp = Index_Data_table - ab
                                                    Exit For
                                                End If
                                            End If

                                        Next

                                        If Index_forCp > 0 Then
                                            If Data_table_Crossing.Rows(Index_forCp).Item("STA") = Station Then

                                                If Data_table_Crossing.Rows(Index_forCp).Item("BLOCKNAME") = "Cable_alignment_crossing" _
                                                    Or Data_table_Crossing.Rows(Index_forCp).Item("BLOCKNAME") = "General_alignment_crossing" _
                                                    Or Data_table_Crossing.Rows(Index_forCp).Item("BLOCKNAME") = "Pipe_alignment_crossing" _
                                                    Or Data_table_Crossing.Rows(Index_forCp).Item("BLOCKNAME") = "Test_section_alignment_crossing" Then
                                                    Data_table_Crossing.Rows(Index_forCp).Item("CP") = "CP"
                                                End If

                                            ElseIf Ch_2 = Station Then
                                                Data_table_Crossing.Rows.Add()
                                                Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = "CP"
                                                Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "CATHODIC_PROTECTION"
                                                Data_table_Crossing.Rows(Index_Data_table).Item("SHEET") = Nr_pagina
                                                Data_table_Crossing.Rows(Index_Data_table).Item("STA") = Station
                                                Dim TextBox11 As New Windows.Forms.TextBox
                                                TextBox11.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
                                                TextBox11.Location = New System.Drawing.Point(3, 3 + 27 * Y_index_panel)
                                                TextBox11.Size = New System.Drawing.Size(127, 22)
                                                TextBox11.Text = Get_chainage_from_double(Station, 1)
                                                TextBox11.BackColor = Drawing.Color.White
                                                TextBox11.ForeColor = Drawing.Color.Black
                                                Panel_items_from_excel.Controls.Add(TextBox11)

                                                Dim TextBox12 As New Windows.Forms.TextBox
                                                TextBox12.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
                                                TextBox12.Location = New System.Drawing.Point(136, 3 + 27 * Y_index_panel)
                                                TextBox12.Size = New System.Drawing.Size(419, 22)
                                                TextBox12.Text = "CP"
                                                TextBox12.BackColor = Drawing.Color.White
                                                TextBox12.ForeColor = Drawing.Color.Black
                                                Panel_items_from_excel.Controls.Add(TextBox12)

                                                Y_index_panel = Y_index_panel + 1
                                                Index_Data_table = Index_Data_table + 1
                                            Else
                                                Data_table_Crossing.Rows.Add()
                                                Data_table_Crossing.Rows(Index_Data_table).Item("STA") = Station
                                                Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = "CP"
                                                Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "CATHODIC_PROTECTION"
                                                Data_table_Crossing.Rows(Index_Data_table).Item("CP") = "CP"
                                                Data_table_Crossing.Rows(Index_Data_table).Item("SHEET") = Nr_pagina
                                                Dim TextBox11 As New Windows.Forms.TextBox
                                                TextBox11.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
                                                TextBox11.Location = New System.Drawing.Point(3, 3 + 27 * Y_index_panel)
                                                TextBox11.Size = New System.Drawing.Size(127, 22)
                                                TextBox11.Text = Get_chainage_from_double(Station, 1)
                                                TextBox11.BackColor = Drawing.Color.White
                                                TextBox11.ForeColor = Drawing.Color.Black
                                                Panel_items_from_excel.Controls.Add(TextBox11)

                                                Dim TextBox12 As New Windows.Forms.TextBox
                                                TextBox12.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
                                                TextBox12.Location = New System.Drawing.Point(136, 3 + 27 * Y_index_panel)
                                                TextBox12.Size = New System.Drawing.Size(419, 22)
                                                TextBox12.Text = "CP"
                                                TextBox12.BackColor = Drawing.Color.White
                                                TextBox12.ForeColor = Drawing.Color.Black
                                                Panel_items_from_excel.Controls.Add(TextBox12)

                                                Y_index_panel = Y_index_panel + 1
                                                Index_Data_table = Index_Data_table + 1
                                            End If
                                        End If



                                    Case Else

                                        If IsNumeric(Material) = True Then
                                            Data_table_Crossing.Rows.Add()
                                            Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = Material
                                            Data_table_Crossing.Rows(Index_Data_table).Item("SHEET") = Nr_pagina
                                            Index_Data_table = Index_Data_table + 1
                                        End If


                                End Select
                            End If

                        End If ' If Not Replace(Description1, " ", "") = ""


                        Chainage_previous = Station


                    Else
                        Dim Material As String = ""
                        If Not Replace(W1.Range(Col6 & i).Value, " ", "") = "" Then
                            Material = W1.Range(Col6 & i).Value
                        End If

                        If IsNumeric(Material) = True Then
                            Data_table_Crossing.Rows.Add()
                            Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = Material
                            Data_table_Crossing.Rows(Index_Data_table).Item("SHEET") = Nr_pagina
                            Index_Data_table = Index_Data_table + 1
                        Else
                            'W1.Range(Col4 & i).Select()
                            'MsgBox("Non numerical value at " & Col4 & i)
                            'Freeze_operations = False
                            'Exit Sub
                        End If



                    End If


                Next

                Dim sTR1 As String = ""


                If Data_table_Crossing.Rows.Count > 0 Then
                    Transfer_datatable_to_new_excel_spreadsheet(Data_table_Crossing)

                    Label_DATA_LOADED.Text = "DATA LOADED"
                End If

            Catch ex As Exception

                MsgBox(ex.Message)

            End Try
            Freeze_operations = False
        End If
    End Sub

    Private Function Insert_extra_warning_signs(ByVal Maximum_distance_for_warning_signs As Double, ByVal Cumulative_Distance_for_warning_signs As Double, ByVal Station_previous_warning_sign As Double) As Integer

        If Cumulative_Distance_for_warning_signs >= Maximum_distance_for_warning_signs Then

            Dim Dist_amount As Double = 2 * Maximum_distance_for_warning_signs
            Dim Index_new_entries As Integer = 2

            Do Until Dist_amount <= Maximum_distance_for_warning_signs
                Dist_amount = Cumulative_Distance_for_warning_signs / Index_new_entries
                If Dist_amount > Maximum_distance_for_warning_signs Then
                    Index_new_entries = Index_new_entries + 1
                End If
            Loop
            For k = 1 To Index_new_entries - 1
                Dim Sta_for_WS As Double = Station_previous_warning_sign + k * Dist_amount

                Dim Index_insert As Integer
                For Index_insert = 0 To Data_table_Crossing.Rows.Count - 1
                    If IsDBNull(Data_table_Crossing.Rows(Index_insert).Item("STA")) = False Then
                        Dim Sta2 As Double = Data_table_Crossing.Rows(Index_insert).Item("STA")
                        If Sta2 > Sta_for_WS Then
                            Exit For
                        End If
                    End If

                Next


                Dim Row1 As System.Data.DataRow
                Row1 = Data_table_Crossing.NewRow()
                Row1("BLOCKNAME") = "WARNING_SIGN"
                Row1("STA") = Sta_for_WS
                Data_table_Crossing.Rows.InsertAt(Row1, Index_insert)
            Next

            Return Index_new_entries - 1

        Else
            Return 0
        End If

    End Function

    Private Sub Button_list_to_DWG_Click(sender As Object, e As EventArgs) Handles Button_list_to_DWG.Click
        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Empty_array() As ObjectId
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Try

                Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult

                Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select the START Point:")
                PP1.AllowNone = True
                Point1 = Editor1.GetPoint(PP1)

                If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                    Editor1.WriteMessage(vbLf & "Command:")
                    Freeze_operations = False
                    Exit Sub
                End If

                Using lock1 As DocumentLock = ThisDrawing.LockDocument
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)

                        If Data_table_Crossing.Rows.Count > 0 Then
                            Dim X, Y, Z As Double
                            X = Point1.Value.X
                            Y = Point1.Value.Y
                            Z = 0

                            Dim ChainageT1 As Double
                            Dim ChainageT2 As Double
                            Dim Extra_length As Double = 0
                            Dim Is_matchline As Boolean = False

                            Dim ChainageSA1_for_Matchline_calcs As Double
                            Dim ChainageSA1 As Double
                            Dim ChainageSA2 As Double
                            Dim Numar_SA_inserted As Integer
                            Dim Extra_length_SA As Double = 0
                            Dim Is_matchline_screw_anchor As Boolean = False
                            Dim Is_ScrewAnchor As Boolean = False
                            Dim Is_primulT_for_SA As Boolean = True

                            Dim ChainageSB1_for_Matchline_calcs As Double
                            Dim ChainageSB1 As Double
                            Dim ChainageSB2 As Double
                            Dim Numar_SB_inserted As Integer
                            Dim Extra_length_SB As Double = 0
                            Dim Is_matchline_sand_bag As Boolean = False
                            Dim Is_SandBag As Boolean = False
                            Dim Is_primulT_for_SB As Boolean = True


                            Dim ChainageRW1_for_Matchline_calcs As Double
                            Dim ChainageRW1 As Double
                            Dim ChainageRW2 As Double
                            Dim Numar_RW_inserted As Integer
                            Dim Extra_length_RW As Double = 0
                            Dim Is_matchline_river_weight As Boolean = False
                            Dim Is_Riverweight As Boolean = False
                            Dim Is_primulT_for_rw As Boolean = True

                            Dim ChainageCCC1 As Double
                            Dim ChainageCCC2 As Double
                            Dim Is_CONCRETE As Boolean = False
                            Dim Is_CP As Boolean = False

                            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                            Dim Index_row_excel As Integer = 2





                            For i = 0 To Data_table_Crossing.Rows.Count - 1
                                Dim Material1 As String = ""
                                If IsDBNull(Data_table_Crossing.Rows(i).Item("SHEET")) = False Then
                                    If IsNumeric(Data_table_Crossing.Rows(i).Item("SHEET")) = True Then
                                        Nr_pagina = Data_table_Crossing.Rows(i).Item("SHEET")
                                    End If
                                End If

                                If IsDBNull(Data_table_Crossing.Rows(i).Item("MATERIAL")) = False Then
                                    If IsNumeric(Data_table_Crossing.Rows(i).Item("MATERIAL")) = True Then
                                        Material1 = Data_table_Crossing.Rows(i).Item("MATERIAL")
                                    End If
                                End If

                                Dim Material_previous As String = ""
                                If IsDBNull(Data_table_Crossing.Rows(i).Item("PREVIOUS_MATERIAL")) = False Then
                                    If IsNumeric(Data_table_Crossing.Rows(i).Item("PREVIOUS_MATERIAL")) = True Then
                                        Material_previous = Data_table_Crossing.Rows(i).Item("PREVIOUS_MATERIAL")
                                    End If
                                End If

                                If IsDBNull(Data_table_Crossing.Rows(i).Item("BLOCKNAME")) = False Then
                                    Dim Block_name As String = Data_table_Crossing.Rows(i).Item("BLOCKNAME")
                                    'Dim Pagina_de_procesat As String = ComboBox_sheet_no.Text

                                    'If Not Replace(Pagina_de_procesat, " ", "") = "" Then
                                    'If Not Data_table_Crossing.Rows(i).Item("SHEET").ToString = Pagina_de_procesat Then
                                    'GoTo END1
                                    'End If
                                    'End If



                                    Select Case Block_name
                                        Case "CONCRETE_START"
                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                                                ChainageCCC1 = Data_table_Crossing.Rows(i).Item("STA")
                                                Is_CONCRETE = True
                                                Extra_length = Extra_length + 14
                                                If RadioButton_right_to_left.Checked = True Then
                                                    X = X - 14
                                                Else
                                                    X = X + 14
                                                End If
                                            End If
                                        Case "CONCRETE_END"
                                            If Is_CONCRETE = True Then
                                                If Not ChainageCCC1 = 0 Then
                                                    If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                                                        ChainageCCC2 = Data_table_Crossing.Rows(i).Item("STA")


                                                        If ChainageCCC2 < ChainageCCC1 Then
                                                            MsgBox("Concrete issue at station " & ChainageCCC2)
                                                            Freeze_operations = False
                                                            Exit Sub
                                                        End If

                                                        Dim nume_block_sa As String = "Concrete_alignment1"


                                                        Dim Colectie_atr_name As New Specialized.StringCollection
                                                        Dim Colectie_atr_value As New Specialized.StringCollection


                                                        Colectie_atr_name.Add("BEGINSTA")
                                                        Colectie_atr_name.Add("ENDSTA")

                                                        If RadioButton_right_to_left.Checked = True Then
                                                            Colectie_atr_value.Add(Get_chainage_from_double(ChainageCCC2, 1))
                                                            Colectie_atr_value.Add(Get_chainage_from_double(ChainageCCC1, 1))

                                                        Else
                                                            Colectie_atr_value.Add(Get_chainage_from_double(ChainageCCC1, 1))
                                                            Colectie_atr_value.Add(Get_chainage_from_double(ChainageCCC2, 1))
                                                        End If

                                                        Dim Val1, Val2 As Double
                                                        Val1 = Round(ChainageCCC1, 1)
                                                        Val2 = Round(ChainageCCC2, 1)
                                                        Colectie_atr_name.Add("LENGTH")
                                                        Colectie_atr_value.Add(Get_String_Rounded(Abs(Val2 - Val1), 1))

                                                        Dim Extra_shift_x As Double = 0
                                                        If RadioButton_left_to_right.Checked = True Then
                                                            Extra_shift_x = 28
                                                        End If

                                                        InsertBlock_with_multiple_atributes(nume_block_sa & ".dwg", nume_block_sa, New Point3d(X - Extra_shift_x, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_name, Colectie_atr_value)

                                                        If RadioButton_right_to_left.Checked = True Then
                                                            X = X - 14
                                                        Else
                                                            X = X + 14
                                                        End If

                                                        Is_CONCRETE = False
                                                        ChainageCCC1 = 0
                                                        Extra_length = Extra_length + 14
                                                    End If
                                                Else
                                                    MsgBox("Concrete issue at station " & ChainageCCC2)
                                                    Freeze_operations = False
                                                    Exit Sub
                                                End If ' asta e de la  If Not ChainageCCC1 = 0
                                            Else
                                                MsgBox("Concrete issue at station " & ChainageCCC2)
                                                Freeze_operations = False
                                                Exit Sub
                                            End If ' asta e de la If Is_CCC = True


                                        Case "SCREWANCHOR_START"
                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                                                ChainageSA1 = Data_table_Crossing.Rows(i).Item("STA")
                                                ChainageSA1_for_Matchline_calcs = ChainageSA1
                                                If RadioButton_right_to_left.Checked = True Then
                                                    X = X - 14
                                                Else
                                                    X = X + 14
                                                End If
                                                Extra_length = Extra_length + 14
                                                Is_primulT_for_SA = True
                                                Is_ScrewAnchor = True
                                                Numar_SA_inserted = 0
                                            Else
                                                ChainageSA1 = 0
                                            End If
                                        Case "SCREWANCHOR_END"
                                            If Is_ScrewAnchor = True Then
                                                If Not ChainageSA1 = 0 Then
                                                    If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                                                        ChainageSA2 = Data_table_Crossing.Rows(i).Item("STA")


                                                        If ChainageSA2 < ChainageSA1 Then
                                                            MsgBox("Screw Anchors issue at station " & ChainageSA2)
                                                            Freeze_operations = False
                                                            Exit Sub
                                                        End If

                                                        Dim nume_block_sa As String = "Screw_anchor_alignment1"

                                                        If RadioButton_right_to_left.Checked = True Then
                                                            Select Case Round(Extra_length_SA, 0)
                                                                Case 0
                                                                    nume_block_sa = "Screw_anchor_alignment1"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_right1"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 14
                                                                    nume_block_sa = "Screw_anchor_alignment1"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_right1"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If

                                                                Case 28
                                                                    nume_block_sa = "Screw_anchor_alignment2"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_right2"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If

                                                                Case 42
                                                                    nume_block_sa = "Screw_anchor_alignment3"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_right3"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If

                                                                Case 56
                                                                    nume_block_sa = "Screw_anchor_alignment4"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_right4"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If

                                                                Case 70
                                                                    nume_block_sa = "Screw_anchor_alignment5"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_right5"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If

                                                                Case 84
                                                                    nume_block_sa = "Screw_anchor_alignment6"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_right6"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If

                                                                Case 98
                                                                    nume_block_sa = "Screw_anchor_alignment7"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_right7"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 112
                                                                    nume_block_sa = "Screw_anchor_alignment8"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_right8"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 126
                                                                    nume_block_sa = "Screw_anchor_alignment9"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_right9"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 140
                                                                    nume_block_sa = "Screw_anchor_alignment10"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_right10"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 154
                                                                    nume_block_sa = "Screw_anchor_alignment11"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_right11"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 168
                                                                    nume_block_sa = "Screw_anchor_alignment12"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_right12"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 182
                                                                    nume_block_sa = "Screw_anchor_alignment13"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_right13"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 196
                                                                    nume_block_sa = "Screw_anchor_alignment14"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_right14"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 210
                                                                    nume_block_sa = "Screw_anchor_alignment15"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_right15"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 224
                                                                    nume_block_sa = "Screw_anchor_alignment16"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_right16"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 238
                                                                    nume_block_sa = "Screw_anchor_alignment17"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_right17"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 252
                                                                    nume_block_sa = "Screw_anchor_alignment18"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_right18"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 266
                                                                    nume_block_sa = "Screw_anchor_alignment19"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_right19"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case Else
                                                                    nume_block_sa = "Screw_anchor_alignment20"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_right20"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If

                                                            End Select
                                                        End If

                                                        If RadioButton_left_to_right.Checked = True Then
                                                            Select Case Round(Extra_length_SA, 0)
                                                                Case 0
                                                                    nume_block_sa = "Screw_anchor_alignment1"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_left1"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 14
                                                                    nume_block_sa = "Screw_anchor_alignment1"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_left1"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If

                                                                Case 28
                                                                    nume_block_sa = "Screw_anchor_alignment2"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_left2"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If

                                                                Case 42
                                                                    nume_block_sa = "Screw_anchor_alignment3"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_left3"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If

                                                                Case 56
                                                                    nume_block_sa = "Screw_anchor_alignment4"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_left4"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If

                                                                Case 70
                                                                    nume_block_sa = "Screw_anchor_alignment5"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_left5"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If

                                                                Case 84
                                                                    nume_block_sa = "Screw_anchor_alignment6"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_left6"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If

                                                                Case 98
                                                                    nume_block_sa = "Screw_anchor_alignment7"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_left7"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 112
                                                                    nume_block_sa = "Screw_anchor_alignment8"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_left8"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 126
                                                                    nume_block_sa = "Screw_anchor_alignment9"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_left9"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 140
                                                                    nume_block_sa = "Screw_anchor_alignment10"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_left10"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 154
                                                                    nume_block_sa = "Screw_anchor_alignment11"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_left11"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 168
                                                                    nume_block_sa = "Screw_anchor_alignment12"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_left12"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 182
                                                                    nume_block_sa = "Screw_anchor_alignment13"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_left13"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 196
                                                                    nume_block_sa = "Screw_anchor_alignment14"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_left14"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 210
                                                                    nume_block_sa = "Screw_anchor_alignment15"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_left15"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 224
                                                                    nume_block_sa = "Screw_anchor_alignment16"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_left16"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 238
                                                                    nume_block_sa = "Screw_anchor_alignment17"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_left17"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 252
                                                                    nume_block_sa = "Screw_anchor_alignment18"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_left18"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 266
                                                                    nume_block_sa = "Screw_anchor_alignment19"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_left19"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case Else
                                                                    nume_block_sa = "Screw_anchor_alignment20"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_sa = "Screw_anchor_alignment_match_left20"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If

                                                            End Select
                                                        End If


                                                        Dim Colectie_atr_name As New Specialized.StringCollection
                                                        Dim Colectie_atr_value As New Specialized.StringCollection


                                                        Colectie_atr_name.Add("BEGINSTA")
                                                        Colectie_atr_name.Add("ENDSTA")

                                                        If RadioButton_right_to_left.Checked = True Then
                                                            Colectie_atr_value.Add(Get_chainage_from_double(ChainageSA2, 1))
                                                            Colectie_atr_value.Add(Get_chainage_from_double(ChainageSA1, 1))

                                                        Else
                                                            Colectie_atr_value.Add(Get_chainage_from_double(ChainageSA1, 1))
                                                            Colectie_atr_value.Add(Get_chainage_from_double(ChainageSA2, 1))
                                                        End If

                                                        Dim Val1, Val2, Val3 As Double
                                                        Val1 = Round(ChainageSA1, 1)
                                                        Val2 = Round(ChainageSA2, 1)
                                                        Val3 = Round(ChainageSA1_for_Matchline_calcs, 1)


                                                        If IsDBNull(Data_table_Crossing.Rows(i).Item("BUOYANCY")) = False Then
                                                            Dim Spacing_SA As Double = Round(Data_table_Crossing.Rows(i).Item("BUOYANCY"), 1)
                                                            If Spacing_SA > 0 Then
                                                                Dim Length_SA As Double = Round(Abs(Val2 - Val1), 1)

                                                                Dim NR_SA As Integer = CInt(1 + Length_SA / Spacing_SA)

                                                                If Numar_SA_inserted > 0 Then
                                                                    Dim Length_SA_for_match As Double = Round(Abs(Val2 - Val3), 1)
                                                                    Dim NR_SA_for_match As Integer = CInt(1 + Length_SA_for_match / Spacing_SA)
                                                                    NR_SA = NR_SA_for_match - Numar_SA_inserted
                                                                    Numar_SA_inserted = 0
                                                                End If



                                                                Colectie_atr_name.Add("NO_TYPE")
                                                                Colectie_atr_name.Add("SPACING")
                                                                Colectie_atr_value.Add(NR_SA & " SA")
                                                                Colectie_atr_value.Add(Get_String_Rounded(Spacing_SA, 1) & " C/C")

                                                            End If


                                                        End If


                                                        If Extra_length_SA = 0 Then
                                                            If RadioButton_right_to_left.Checked = True Then
                                                                X = X - 14
                                                            Else
                                                                X = X + 14
                                                            End If
                                                        End If

                                                        Dim Valoare_x_shift As Double = 0
                                                        If RadioButton_left_to_right.Checked = True Then
                                                            If Extra_length_SA = 0 Then
                                                                Valoare_x_shift = 28
                                                            Else
                                                                Valoare_x_shift = Extra_length_SA + 14
                                                            End If
                                                        End If



                                                        InsertBlock_with_multiple_atributes(nume_block_sa & ".dwg", nume_block_sa, New Point3d(X - Valoare_x_shift, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_name, Colectie_atr_value)

                                                        If RadioButton_right_to_left.Checked = True Then
                                                            X = X - 14
                                                        Else
                                                            X = X + 14
                                                        End If



                                                        If Extra_length_SA = 0 Then
                                                            Extra_length = Extra_length + 28
                                                        Else
                                                            Extra_length = Extra_length + 14
                                                        End If

                                                        Extra_length_SA = 0
                                                        Is_ScrewAnchor = False
                                                        Is_primulT_for_SA = False
                                                        Is_matchline_screw_anchor = False
                                                        ChainageSA1 = 0
                                                    End If
                                                Else
                                                    MsgBox("Screw Anchors issue at station (Screw Anchor Start = 0)" & ChainageSA2)
                                                    Freeze_operations = False
                                                    Exit Sub
                                                End If ' asta e de la  If Not ChainageSA1 = 0


                                            Else
                                                MsgBox("Screw Anchors issue at station (No Screw Anchor Start)" & ChainageSA2)
                                                Freeze_operations = False
                                                Exit Sub
                                            End If ' asta e de la If Is_ScrewAnchor = True

                                        Case "SANDBAG_START"
                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                                                ChainageSB1 = Data_table_Crossing.Rows(i).Item("STA")
                                                ChainageSB1_for_Matchline_calcs = ChainageSB1
                                                If RadioButton_right_to_left.Checked = True Then
                                                    X = X - 14
                                                Else
                                                    X = X + 14
                                                End If
                                                Extra_length = Extra_length + 14
                                                Is_primulT_for_SB = True
                                                Is_SandBag = True
                                                Numar_SB_inserted = 0
                                            End If
                                        Case "SANDBAG_END"
                                            If Is_SandBag = True Then
                                                If Not ChainageSB1 = 0 Then
                                                    If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                                                        ChainageSB2 = Data_table_Crossing.Rows(i).Item("STA")


                                                        If ChainageSB2 < ChainageSB1 Then
                                                            MsgBox("Sand bags issue at station " & ChainageSB2)
                                                            Freeze_operations = False
                                                            Exit Sub
                                                        End If

                                                        Dim nume_block_SB As String = "Sand_Bag_alignment1"

                                                        If RadioButton_right_to_left.Checked = True Then
                                                            Select Case Round(Extra_length_SB, 0)
                                                                Case 0
                                                                    nume_block_SB = "Sand_Bag_alignment1"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_right1"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 14
                                                                    nume_block_SB = "Sand_Bag_alignment1"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_right1"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If

                                                                Case 28
                                                                    nume_block_SB = "Sand_Bag_alignment2"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_right2"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If

                                                                Case 42
                                                                    nume_block_SB = "Sand_Bag_alignment3"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_right3"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If

                                                                Case 56
                                                                    nume_block_SB = "Sand_Bag_alignment4"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_right4"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If

                                                                Case 70
                                                                    nume_block_SB = "Sand_Bag_alignment5"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_right5"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If

                                                                Case 84
                                                                    nume_block_SB = "Sand_Bag_alignment6"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_right6"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If

                                                                Case 98
                                                                    nume_block_SB = "Sand_Bag_alignment7"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_right7"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 112
                                                                    nume_block_SB = "Sand_Bag_alignment8"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_right8"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 126
                                                                    nume_block_SB = "Sand_Bag_alignment9"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_right9"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 140
                                                                    nume_block_SB = "Sand_Bag_alignment10"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_right10"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 154
                                                                    nume_block_SB = "Sand_Bag_alignment11"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_right11"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 168
                                                                    nume_block_SB = "Sand_Bag_alignment12"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_right12"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 182
                                                                    nume_block_SB = "Sand_Bag_alignment13"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_right13"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 196
                                                                    nume_block_SB = "Sand_Bag_alignment14"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_right14"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 210
                                                                    nume_block_SB = "Sand_Bag_alignment15"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_right15"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 224
                                                                    nume_block_SB = "Sand_Bag_alignment16"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_right16"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 238
                                                                    nume_block_SB = "Sand_Bag_alignment17"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_right17"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 252
                                                                    nume_block_SB = "Sand_Bag_alignment18"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_right18"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 266
                                                                    nume_block_SB = "Sand_Bag_alignment19"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_right19"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case Else
                                                                    nume_block_SB = "Sand_Bag_alignment20"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_right20"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If

                                                            End Select

                                                        End If

                                                        If RadioButton_left_to_right.Checked = True Then
                                                            Select Case Round(Extra_length_SB, 0)
                                                                Case 0
                                                                    nume_block_SB = "Sand_Bag_alignment1"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_left1"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 14
                                                                    nume_block_SB = "Sand_Bag_alignment1"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_left1"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If

                                                                Case 28
                                                                    nume_block_SB = "Sand_Bag_alignment2"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_left2"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If

                                                                Case 42
                                                                    nume_block_SB = "Sand_Bag_alignment3"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_left3"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If

                                                                Case 56
                                                                    nume_block_SB = "Sand_Bag_alignment4"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_left4"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If

                                                                Case 70
                                                                    nume_block_SB = "Sand_Bag_alignment5"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_left5"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If

                                                                Case 84
                                                                    nume_block_SB = "Sand_Bag_alignment6"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_left6"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If

                                                                Case 98
                                                                    nume_block_SB = "Sand_Bag_alignment7"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_left7"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 112
                                                                    nume_block_SB = "Sand_Bag_alignment8"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_left8"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 126
                                                                    nume_block_SB = "Sand_Bag_alignment9"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_left9"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 140
                                                                    nume_block_SB = "Sand_Bag_alignment10"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_left10"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 154
                                                                    nume_block_SB = "Sand_Bag_alignment11"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_left11"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 168
                                                                    nume_block_SB = "Sand_Bag_alignment12"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_left12"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 182
                                                                    nume_block_SB = "Sand_Bag_alignment13"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_left13"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 196
                                                                    nume_block_SB = "Sand_Bag_alignment14"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_left14"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 210
                                                                    nume_block_SB = "Sand_Bag_alignment15"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_left15"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 224
                                                                    nume_block_SB = "Sand_Bag_alignment16"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_left16"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 238
                                                                    nume_block_SB = "Sand_Bag_alignment17"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_left17"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 252
                                                                    nume_block_SB = "Sand_Bag_alignment18"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_left18"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case 266
                                                                    nume_block_SB = "Sand_Bag_alignment19"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_left19"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If
                                                                Case Else
                                                                    nume_block_SB = "Sand_Bag_alignment20"

                                                                    If Is_matchline_screw_anchor = True Then
                                                                        nume_block_SB = "Sand_Bag_alignment_match_left20"
                                                                        Is_matchline_screw_anchor = False
                                                                    End If

                                                            End Select

                                                        End If


                                                        Dim Colectie_atr_name As New Specialized.StringCollection
                                                        Dim Colectie_atr_value As New Specialized.StringCollection


                                                        Colectie_atr_name.Add("BEGINSTA")
                                                        Colectie_atr_name.Add("ENDSTA")

                                                        If RadioButton_right_to_left.Checked = True Then
                                                            Colectie_atr_value.Add(Get_chainage_from_double(ChainageSB2, 1))
                                                            Colectie_atr_value.Add(Get_chainage_from_double(ChainageSB1, 1))

                                                        Else
                                                            Colectie_atr_value.Add(Get_chainage_from_double(ChainageSB1, 1))
                                                            Colectie_atr_value.Add(Get_chainage_from_double(ChainageSB2, 1))
                                                        End If

                                                        Dim Val1, Val2, Val3 As Double
                                                        Val1 = Round(ChainageSB1, 1)
                                                        Val2 = Round(ChainageSB2, 1)
                                                        Val3 = Round(ChainageSB1_for_Matchline_calcs, 1)

                                                        If IsDBNull(Data_table_Crossing.Rows(i).Item("BUOYANCY")) = False Then
                                                            Dim Spacing_SB As Double = Round(Data_table_Crossing.Rows(i).Item("BUOYANCY"), 1)
                                                            If Spacing_SB > 0 Then
                                                                Dim Length_SB As Double = Round(Abs(Val2 - Val1), 1)

                                                                Dim NR_SB As Integer = CInt(1 + Length_SB / Spacing_SB)

                                                                If Numar_SB_inserted > 0 Then
                                                                    Dim Length_SB_for_match As Double = Round(Abs(Val2 - Val3), 1)
                                                                    Dim NR_SB_for_match As Integer = CInt(1 + Length_SB_for_match / Spacing_SB)
                                                                    NR_SB = NR_SB_for_match - Numar_SB_inserted
                                                                    Numar_SB_inserted = 0
                                                                End If



                                                                Colectie_atr_name.Add("NO_TYPE")
                                                                Colectie_atr_name.Add("SPACING")
                                                                Colectie_atr_value.Add(NR_SB & " SBW")
                                                                Colectie_atr_value.Add(Get_String_Rounded(Spacing_SB, 1) & " C/C")

                                                            End If


                                                        End If


                                                        If Extra_length_SB = 0 Then
                                                            If RadioButton_right_to_left.Checked = True Then
                                                                X = X - 14
                                                            Else
                                                                X = X + 14
                                                            End If
                                                        End If

                                                        Dim Valoare_x_shift As Double = 0
                                                        If RadioButton_left_to_right.Checked = True Then
                                                            If Extra_length_SB = 0 Then
                                                                Valoare_x_shift = 28
                                                            Else
                                                                Valoare_x_shift = Extra_length_SB + 14
                                                            End If
                                                        End If

                                                        InsertBlock_with_multiple_atributes(nume_block_SB & ".dwg", nume_block_SB, New Point3d(X - Valoare_x_shift, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_name, Colectie_atr_value)

                                                        If RadioButton_right_to_left.Checked = True Then
                                                            X = X - 14
                                                        Else
                                                            X = X + 14
                                                        End If



                                                        If Extra_length_SB = 0 Then
                                                            Extra_length = Extra_length + 28
                                                        Else
                                                            Extra_length = Extra_length + 14
                                                        End If

                                                        Extra_length_SB = 0
                                                        Is_SandBag = False
                                                        Is_primulT_for_SB = False
                                                        Is_matchline_sand_bag = False
                                                        ChainageSB1 = 0
                                                    End If
                                                Else
                                                    MsgBox("Sand bags issue at station " & ChainageSB2)
                                                    Freeze_operations = False
                                                    Exit Sub
                                                End If ' asta e de la  If Not ChainageSA1 = 0
                                            Else
                                                MsgBox("Sand bags issue at station " & ChainageSB2)
                                                Freeze_operations = False
                                                Exit Sub
                                            End If ' asta e de la If Is_ScrewAnchor = True

                                        Case "CATHODIC_PROTECTION"
                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                                                Dim Colectie_atr_name As New Specialized.StringCollection
                                                Dim Colectie_atr_value As New Specialized.StringCollection

                                                Dim Chainage_cp As Double = Data_table_Crossing.Rows(i).Item("STA")

                                                If IsDBNull(Data_table_Crossing.Rows(i).Item("CP")) = False Then
                                                    Colectie_atr_value.Add(Get_chainage_from_double(Chainage_cp, 1))
                                                    Colectie_atr_name.Add("STA")
                                                    InsertBlock_with_multiple_atributes("CATHODIC_PROTECTION.dwg", "CATHODIC_PROTECTION", New Point3d(X, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_name, Colectie_atr_value)

                                                    If RadioButton_right_to_left.Checked = True Then
                                                        X = X - 14
                                                    Else
                                                        X = X + 14
                                                    End If
                                                    Extra_length = Extra_length + 14
                                                    If Is_ScrewAnchor = True Then
                                                        Extra_length_SA = Extra_length_SA + 14
                                                    End If
                                                    If Is_SandBag = True Then
                                                        Extra_length_SB = Extra_length_SB + 14
                                                    End If

                                                Else

                                                    Is_CP = True
                                                End If

                                            End If

                                        Case "WARNING_SIGN"
                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                                                Dim Colectie_atr_name As New Specialized.StringCollection
                                                Dim Colectie_atr_value As New Specialized.StringCollection

                                                Dim Chainage_WS As Double = Data_table_Crossing.Rows(i).Item("STA")


                                                Colectie_atr_value.Add(Get_chainage_from_double(Chainage_WS, 1))
                                                Colectie_atr_name.Add("STA")

                                                If CheckBox_extra_warning_signs.Checked = False Then
                                                    InsertBlock_with_multiple_atributes("Warning_sign_with_invisible_STA.dwg", "Warning_sign_with_invisible_STA", New Point3d(X, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_name, Colectie_atr_value)
                                                Else
                                                    InsertBlock_with_multiple_atributes("Warning_sign_alignment.dwg", "Warning_sign_alignment", New Point3d(X, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_name, Colectie_atr_value)
                                                End If


                                                If RadioButton_right_to_left.Checked = True Then
                                                    X = X - 14
                                                Else
                                                    X = X + 14
                                                End If
                                                Extra_length = Extra_length + 14
                                                If Is_ScrewAnchor = True Then
                                                    Extra_length_SA = Extra_length_SA + 14
                                                End If
                                                If Is_SandBag = True Then
                                                    Extra_length_SB = Extra_length_SB + 14
                                                End If



                                            End If

                                        Case "Pipe_Transition"
                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                                                Dim Colectie_atr_namet As New Specialized.StringCollection
                                                Dim Colectie_atr_valuet As New Specialized.StringCollection


                                                ChainageT1 = ChainageT2
                                                ChainageT2 = Data_table_Crossing.Rows(i).Item("STA")

                                                Dim Nume_pipeW As String = "heavy_wall_1"
                                                If Not i = 0 Then
                                                    If Extra_length = 0 Then
                                                        If RadioButton_right_to_left.Checked = True Then
                                                            X = X - 14
                                                        Else
                                                            X = X + 14
                                                        End If
                                                    End If
                                                End If

                                                Dim NumarT1 As Double = Round(ChainageT1, 1)
                                                Dim NumarT2 As Double = Round(ChainageT2, 1)

                                                If NumarT1 = NumarT2 Then
                                                    If RadioButton_right_to_left.Checked = True Then
                                                        X = X + 28
                                                    Else
                                                        X = X - 28
                                                    End If
                                                End If
                                                Dim Description2 As String = ""
                                                If IsDBNull(Data_table_Crossing.Rows(i).Item("DESCRIPTION2")) = False Then
                                                    Description2 = Data_table_Crossing.Rows(i).Item("DESCRIPTION2")
                                                End If

                                                If Not Description2 = "FAKE" Then
                                                    InsertBlock_with_multiple_atributes("Pipe_Transition.dwg", "Pipe_Transition", New Point3d(X, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_namet, Colectie_atr_valuet)
                                                End If



                                                If NumarT1 < NumarT2 Then

                                                    If RadioButton_right_to_left.Checked = True Then
                                                        Select Case Round(Extra_length, 0)
                                                            Case 0
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_1"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_1"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_1"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_1"
                                                                End If

                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_right_1"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_right_1"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_right_1"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_1"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If

                                                            Case 14
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_1"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_1"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_1"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_1"
                                                                End If

                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_right_1"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_right_1"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_right_1"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_1"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If

                                                            Case 28
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_2"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_2"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_2"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_2"
                                                                End If


                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_right_2"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_right_2"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_right_2"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_2"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If

                                                            Case 42
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_3"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_3"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_3"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_3"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_right_3"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_right_3"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_right_3"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_3"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If


                                                            Case 56
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_4"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_4"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_4"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_4"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_right_4"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_right_4"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_right_4"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_4"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If

                                                            Case 70
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_5"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_5"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_5"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_5"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_right_5"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_right_5"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_right_5"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_5"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If

                                                            Case 84
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_6"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_6"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_6"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_6"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_right_6"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_right_6"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_right_6"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_6"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If

                                                            Case 98
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_7"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_7"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_7"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_7"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_right_7"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_right_7"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_right_7"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_7"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If

                                                            Case 112
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_8"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_8"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_8"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_8"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_right_8"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_right_8"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_right_8"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_8"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If
                                                            Case 126
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_9"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_9"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_9"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_9"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_right_9"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_right_9"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_right_9"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_9"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If
                                                            Case 140
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_10"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_10"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_10"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_10"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_right_10"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_right_10"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_right_10"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_10"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If
                                                            Case 154
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_11"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_11"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_11"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_11"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_right_11"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_right_11"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_right_11"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_11"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If
                                                            Case 168
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_12"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_12"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_12"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_12"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_right_12"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_right_12"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_right_12"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_12"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If
                                                            Case 182
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_13"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_13"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_13"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_13"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_right_13"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_right_13"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_right_13"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_13"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If
                                                            Case 196
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_14"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_14"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_14"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_14"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_right_14"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_right_14"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_right_14"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_14"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If
                                                            Case 210
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_15"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_15"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_15"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_15"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_right_15"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_right_15"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_right_15"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_15"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If
                                                            Case 224
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_16"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "Line_pipe_16"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_16"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_16"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_16"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_right_16"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_right_16"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_right_16"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_16"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If
                                                            Case 238
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_17"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_17"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_17"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_17"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_right_17"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_right_17"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_right_17"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_17"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If
                                                            Case 252
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_18"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_18"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_18"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_18"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_right_18"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_right_18"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_right_18"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_18"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If
                                                            Case 266
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_19"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_19"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_19"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_19"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_right_19"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_right_19"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_right_19"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_19"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If
                                                            Case Else
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_20"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_20"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_20"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_20"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_right_20"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_right_20"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_right_20"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_20"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If
                                                        End Select
                                                    End If

                                                    If RadioButton_left_to_right.Checked = True Then
                                                        Select Case Round(Extra_length, 0)
                                                            Case 0
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_1"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_1"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_1"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_1"
                                                                End If

                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_left_1"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_left_1"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_left_1"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_1"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If

                                                            Case 14
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_1"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_1"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_1"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_1"
                                                                End If

                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_left_1"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_left_1"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_left_1"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_1"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If

                                                            Case 28
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_2"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_2"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_2"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_2"
                                                                End If


                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_left_2"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_left_2"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_left_2"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_2"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If

                                                            Case 42
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_3"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_3"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_3"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_3"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_left_3"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_left_3"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_left_3"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_3"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If


                                                            Case 56
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_4"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_4"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_4"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_4"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_left_4"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_left_4"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_left_4"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_4"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If

                                                            Case 70
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_5"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_5"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_5"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_5"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_left_5"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_left_5"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_left_5"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_5"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If

                                                            Case 84
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_6"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_6"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_6"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_6"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_left_6"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_left_6"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_left_6"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_6"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If

                                                            Case 98
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_7"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_7"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_7"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_7"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_left_7"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_left_7"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_left_7"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_7"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If

                                                            Case 112
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_8"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_8"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_8"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_8"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_left_8"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_left_8"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_left_8"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_8"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If
                                                            Case 126
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_9"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_9"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_9"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_9"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_left_9"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_left_9"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_left_9"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_9"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If
                                                            Case 140
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_10"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_10"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_10"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_10"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_left_10"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_left_10"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_left_10"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_10"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If
                                                            Case 154
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_11"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_11"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_11"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_11"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_left_11"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_left_11"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_left_11"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_11"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If
                                                            Case 168
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_12"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_12"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_12"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_12"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_left_12"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_left_12"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_left_12"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_12"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If
                                                            Case 182
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_13"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_13"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_13"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_13"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_left_13"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_left_13"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_left_13"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_13"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If
                                                            Case 196
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_14"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_14"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_14"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_14"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_left_14"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_left_14"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_left_14"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_14"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If
                                                            Case 210
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_15"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_15"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_15"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_15"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_left_15"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_left_15"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_left_15"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_15"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If
                                                            Case 224
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_16"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "Line_pipe_16"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_16"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_16"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_16"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_left_16"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_left_16"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_left_16"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_16"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If
                                                            Case 238
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_17"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_17"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_17"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_17"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_left_17"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_left_17"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_left_17"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_17"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If
                                                            Case 252
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_18"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_18"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_18"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_18"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_left_18"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_left_18"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_left_18"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_18"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If
                                                            Case 266
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_19"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_19"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_19"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_19"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_left_19"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_left_19"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_left_19"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_19"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If
                                                            Case Else
                                                                If Material_previous = "1" Then
                                                                    Nume_pipeW = "mat_20"
                                                                ElseIf Material_previous = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "Line_pipe_20"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_20"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_20"
                                                                End If
                                                                If Is_matchline = True Then
                                                                    If Material_previous = "1" Then
                                                                        Nume_pipeW = "mat1_match_left_20"
                                                                    ElseIf Material_previous = "2" Then
                                                                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                            Nume_pipeW = "line_pipe_match_left_20"
                                                                        Else
                                                                            Nume_pipeW = "heavy_wall_match_left_20"
                                                                        End If
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_20"
                                                                    End If
                                                                    Is_matchline = False
                                                                End If
                                                        End Select
                                                    End If



                                                    Dim Colectie_atr_name As New Specialized.StringCollection
                                                    Dim Colectie_atr_value As New Specialized.StringCollection


                                                    Colectie_atr_name.Add("BEGINSTA")
                                                    Colectie_atr_name.Add("ENDSTA")

                                                    If RadioButton_right_to_left.Checked = True Then
                                                        Colectie_atr_value.Add(Get_chainage_from_double(ChainageT2, 1))
                                                        Colectie_atr_value.Add(Get_chainage_from_double(ChainageT1, 1))
                                                    Else
                                                        Colectie_atr_value.Add(Get_chainage_from_double(ChainageT1, 1))
                                                        Colectie_atr_value.Add(Get_chainage_from_double(ChainageT2, 1))
                                                    End If

                                                    Dim Val1, Val2 As Double
                                                    Val1 = Round(ChainageT1, 1)
                                                    Val2 = Round(ChainageT2, 1)

                                                    Colectie_atr_name.Add("LENGTH")
                                                    Dim String_len As String = Get_String_Rounded(Abs(Val2 - Val1), 1)
                                                    If IsDBNull(Data_table_Crossing.Rows(i).Item("EXTRA_LENGTH")) = False Then
                                                        String_len = String_len & " (" & Get_String_Rounded(Data_table_Crossing.Rows(i).Item("EXTRA_LENGTH"), 1) & " TRUE LENGTH)"
                                                    End If

                                                    Colectie_atr_value.Add(String_len)
                                                    Colectie_atr_name.Add("MAT")
                                                    Colectie_atr_value.Add(Material_previous)

                                                    If Not i = 0 Then
                                                        If RadioButton_right_to_left.Checked = True Then
                                                            InsertBlock_with_multiple_atributes(Nume_pipeW & ".dwg", Nume_pipeW, New Point3d(X, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_name, Colectie_atr_value)
                                                        Else
                                                            Dim Valoare_x_shift As Double
                                                            If Extra_length = 0 Then
                                                                Valoare_x_shift = 28
                                                            Else
                                                                Valoare_x_shift = Extra_length + 14
                                                            End If

                                                            InsertBlock_with_multiple_atributes(Nume_pipeW & ".dwg", Nume_pipeW, New Point3d(X - Valoare_x_shift, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_name, Colectie_atr_value)
                                                        End If

                                                    End If

                                                End If

                                                If Is_ScrewAnchor = True Then
                                                    If Is_primulT_for_SA = True Then
                                                        Extra_length_SA = Extra_length_SA + 14
                                                        Is_primulT_for_SA = False
                                                    Else
                                                        Extra_length_SA = Extra_length_SA + 28
                                                    End If
                                                End If

                                                If Is_SandBag = True Then
                                                    If Is_primulT_for_SB = True Then
                                                        Extra_length_SB = Extra_length_SB + 14
                                                        Is_primulT_for_SB = False
                                                    Else
                                                        Extra_length_SB = Extra_length_SB + 28
                                                    End If
                                                End If

                                                Extra_length = 0



                                                If RadioButton_right_to_left.Checked = True Then
                                                    X = X - 14
                                                Else
                                                    X = X + 14
                                                End If

                                            End If


                                        Case "Elbow_al"


                                            Dim Colectie_atr_name_el As New Specialized.StringCollection
                                            Dim Colectie_atr_value_el As New Specialized.StringCollection

                                            Dim StartC As String = ""
                                            Dim EndC As String = ""
                                            Dim StaC As String = ""

                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("ENDSTA")) = False Then
                                                EndC = Data_table_Crossing.Rows(i).Item("ENDSTA")
                                            End If

                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("BEGINSTA")) = False Then
                                                StartC = Data_table_Crossing.Rows(i).Item("BEGINSTA")
                                            End If

                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                                                StaC = Data_table_Crossing.Rows(i).Item("STA")
                                            End If
                                            If IsNumeric(StaC) = True Then
                                                StaC = Get_chainage_from_double(CDbl(StaC), 1)
                                            End If

                                            If IsNumeric(StartC) = True Then
                                                StartC = Get_chainage_from_double(CDbl(StartC), 1)
                                            End If

                                            If IsNumeric(EndC) = True Then
                                                EndC = Get_chainage_from_double(CDbl(EndC), 1)
                                            End If


                                            ChainageT1 = ChainageT2

                                            ChainageT2 = CDbl(Replace(StartC, "+", ""))



                                            ChainageT1 = Round(ChainageT1, 1)
                                            ChainageT2 = Round(ChainageT2, 1)

                                            If ChainageT2 = ChainageT1 Then
                                                If RadioButton_right_to_left.Checked = True Then
                                                    X = X + 14
                                                Else
                                                    X = X - 14
                                                End If
                                            End If

                                            If ChainageT1 < ChainageT2 Then
                                                Dim Nume_pipeW As String = "heavy_wall_1"
                                                If Not i = 0 Then
                                                    If Extra_length = 0 Then
                                                        If RadioButton_right_to_left.Checked = True Then
                                                            X = X - 14
                                                        Else
                                                            X = X + 14
                                                        End If
                                                    End If
                                                End If


                                                If RadioButton_right_to_left.Checked = True Then
                                                    Select Case Round(Extra_length, 0)
                                                        Case 0
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_1"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_1"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_1"
                                                                End If

                                                            Else
                                                                Nume_pipeW = "heavy_wall_1"
                                                            End If

                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_right_1"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_right_1"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_1"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_1"
                                                                End If
                                                                Is_matchline = False
                                                            End If

                                                        Case 14
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_1"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_1"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_1"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_1"
                                                            End If

                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_right_1"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_right_1"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_1"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_1"
                                                                End If
                                                                Is_matchline = False
                                                            End If

                                                        Case 28
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_2"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_2"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_2"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_2"
                                                            End If


                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_right_2"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_right_2"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_2"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_2"
                                                                End If
                                                                Is_matchline = False
                                                            End If

                                                        Case 42
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_3"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_3"

                                                                Else
                                                                    Nume_pipeW = "heavy_wall_3"
                                                                End If

                                                            Else
                                                                Nume_pipeW = "heavy_wall_3"
                                                            End If


                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_right_3"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then

                                                                        Nume_pipeW = "line_pipe_match_right_3"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_3"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_3"
                                                                End If
                                                                Is_matchline = False
                                                            End If


                                                        Case 56
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_4"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then

                                                                    Nume_pipeW = "Line_pipe_4"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_4"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_4"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_right_4"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_right_4"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_4"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_4"
                                                                End If
                                                                Is_matchline = False
                                                            End If

                                                        Case 70
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_5"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_5"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_5"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_5"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_right_5"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_right_5"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_5"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_5"
                                                                End If
                                                                Is_matchline = False
                                                            End If

                                                        Case 84
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_6"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_6"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_6"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_6"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_right_6"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_right_6"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_6"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_6"
                                                                End If
                                                                Is_matchline = False
                                                            End If

                                                        Case 98
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_7"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_7"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_7"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_7"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_right_7"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_right_7"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_7"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_7"
                                                                End If
                                                                Is_matchline = False
                                                            End If

                                                        Case 112
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_8"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_8"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_8"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_8"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_right_8"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_right_8"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_8"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_8"
                                                                End If
                                                                Is_matchline = False
                                                            End If
                                                        Case 126
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_9"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_9"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_9"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_9"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_right_9"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_right_9"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_9"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_9"
                                                                End If
                                                                Is_matchline = False
                                                            End If
                                                        Case 140
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_10"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_10"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_10"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_10"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_right_10"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_right_10"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_10"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_10"
                                                                End If
                                                                Is_matchline = False
                                                            End If
                                                        Case 154
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_11"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_11"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_11"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_11"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_right_11"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_right_11"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_11"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_11"
                                                                End If
                                                                Is_matchline = False
                                                            End If
                                                        Case 168
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_12"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_12"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_12"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_12"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_right_12"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_right_12"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_12"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_12"
                                                                End If
                                                                Is_matchline = False
                                                            End If
                                                        Case 182
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_13"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_13"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_13"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_13"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_right_13"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_right_13"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_13"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_13"
                                                                End If
                                                                Is_matchline = False
                                                            End If
                                                        Case 196
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_14"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_14"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_14"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_14"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_right_14"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_right_14"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_14"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_14"
                                                                End If
                                                                Is_matchline = False
                                                            End If
                                                        Case 210
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_15"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_15"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_15"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_15"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_right_15"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_right_15"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_15"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_15"
                                                                End If
                                                                Is_matchline = False
                                                            End If
                                                        Case 224
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_16"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_16"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_16"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_16"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_right_16"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_right_16"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_16"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_16"
                                                                End If
                                                                Is_matchline = False
                                                            End If
                                                        Case 238
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_17"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_17"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_17"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_17"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_right_17"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_right_17"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_17"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_17"
                                                                End If
                                                                Is_matchline = False
                                                            End If
                                                        Case 252
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_18"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_18"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_18"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_18"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_right_18"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_right_18"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_18"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_18"
                                                                End If
                                                                Is_matchline = False
                                                            End If
                                                        Case 266
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_19"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_19"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_19"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_19"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_right_19"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_right_19"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_19"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_19"
                                                                End If
                                                                Is_matchline = False
                                                            End If
                                                        Case Else
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_20"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_20"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_20"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_20"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_right_20"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_right_20"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_right_20"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_20"
                                                                End If
                                                                Is_matchline = False
                                                            End If
                                                    End Select
                                                End If

                                                If RadioButton_left_to_right.Checked = True Then
                                                    Select Case Round(Extra_length, 0)
                                                        Case 0
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_1"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_1"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_1"
                                                                End If

                                                            Else
                                                                Nume_pipeW = "heavy_wall_1"
                                                            End If

                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_left_1"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_left_1"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_1"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_1"
                                                                End If
                                                                Is_matchline = False
                                                            End If

                                                        Case 14
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_1"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_1"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_1"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_1"
                                                            End If

                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_left_1"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_left_1"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_1"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_1"
                                                                End If
                                                                Is_matchline = False
                                                            End If

                                                        Case 28
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_2"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_2"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_2"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_2"
                                                            End If


                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_left_2"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_left_2"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_2"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_2"
                                                                End If
                                                                Is_matchline = False
                                                            End If

                                                        Case 42
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_3"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_3"

                                                                Else
                                                                    Nume_pipeW = "heavy_wall_3"
                                                                End If

                                                            Else
                                                                Nume_pipeW = "heavy_wall_3"
                                                            End If


                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_left_3"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then

                                                                        Nume_pipeW = "line_pipe_match_left_3"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_3"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_3"
                                                                End If
                                                                Is_matchline = False
                                                            End If


                                                        Case 56
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_4"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then

                                                                    Nume_pipeW = "Line_pipe_4"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_4"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_4"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_left_4"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_left_4"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_4"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_4"
                                                                End If
                                                                Is_matchline = False
                                                            End If

                                                        Case 70
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_5"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_5"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_5"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_5"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_left_5"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_left_5"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_5"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_5"
                                                                End If
                                                                Is_matchline = False
                                                            End If

                                                        Case 84
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_6"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_6"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_6"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_6"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_left_6"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_left_6"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_6"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_6"
                                                                End If
                                                                Is_matchline = False
                                                            End If

                                                        Case 98
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_7"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_7"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_7"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_7"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_left_7"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_left_7"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_7"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_7"
                                                                End If
                                                                Is_matchline = False
                                                            End If

                                                        Case 112
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_8"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_8"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_8"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_8"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_left_8"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_left_8"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_8"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_8"
                                                                End If
                                                                Is_matchline = False
                                                            End If
                                                        Case 126
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_9"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_9"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_9"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_9"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_left_9"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_left_9"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_9"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_9"
                                                                End If
                                                                Is_matchline = False
                                                            End If
                                                        Case 140
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_10"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_10"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_10"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_10"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_left_10"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_left_10"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_10"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_10"
                                                                End If
                                                                Is_matchline = False
                                                            End If
                                                        Case 154
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_11"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_11"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_11"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_11"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_left_11"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_left_11"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_11"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_11"
                                                                End If
                                                                Is_matchline = False
                                                            End If
                                                        Case 168
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_12"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_12"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_12"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_12"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_left_12"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_left_12"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_12"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_12"
                                                                End If
                                                                Is_matchline = False
                                                            End If
                                                        Case 182
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_13"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_13"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_13"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_13"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_left_13"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_left_13"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_13"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_13"
                                                                End If
                                                                Is_matchline = False
                                                            End If
                                                        Case 196
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_14"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_14"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_14"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_14"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_left_14"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_left_14"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_14"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_14"
                                                                End If
                                                                Is_matchline = False
                                                            End If
                                                        Case 210
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_15"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_15"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_15"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_15"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_left_15"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_left_15"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_15"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_15"
                                                                End If
                                                                Is_matchline = False
                                                            End If
                                                        Case 224
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_16"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_16"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_16"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_16"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_left_16"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_left_16"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_16"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_16"
                                                                End If
                                                                Is_matchline = False
                                                            End If
                                                        Case 238
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_17"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_17"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_17"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_17"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_left_17"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_left_17"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_17"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_17"
                                                                End If
                                                                Is_matchline = False
                                                            End If
                                                        Case 252
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_18"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_18"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_18"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_18"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_left_18"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_left_18"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_18"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_18"
                                                                End If
                                                                Is_matchline = False
                                                            End If
                                                        Case 266
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_19"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_19"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_19"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_19"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_left_19"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_left_19"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_19"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_19"
                                                                End If
                                                                Is_matchline = False
                                                            End If
                                                        Case Else
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat_20"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "Line_pipe_20"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_20"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_20"
                                                            End If
                                                            If Is_matchline = True Then
                                                                If Material1 = "1" Then
                                                                    Nume_pipeW = "mat1_match_left_20"
                                                                ElseIf Material1 = "2" Then
                                                                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                        Nume_pipeW = "line_pipe_match_left_20"
                                                                    Else
                                                                        Nume_pipeW = "heavy_wall_match_left_20"
                                                                    End If
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_20"
                                                                End If
                                                                Is_matchline = False
                                                            End If
                                                    End Select

                                                End If




                                                Dim Colectie_atr_name As New Specialized.StringCollection
                                                Dim Colectie_atr_value As New Specialized.StringCollection


                                                Colectie_atr_name.Add("BEGINSTA")
                                                Colectie_atr_name.Add("ENDSTA")
                                                Dim Pct_ins As New Point3d
                                                If RadioButton_right_to_left.Checked = True Then
                                                    Colectie_atr_value.Add(Get_chainage_from_double(ChainageT2, 1))
                                                    Colectie_atr_value.Add(Get_chainage_from_double(ChainageT1, 1))
                                                    Pct_ins = New Point3d(X, Y, Z)
                                                Else
                                                    Colectie_atr_value.Add(Get_chainage_from_double(ChainageT1, 1))
                                                    Colectie_atr_value.Add(Get_chainage_from_double(ChainageT2, 1))

                                                    Dim Valoare_x_shift As Double
                                                    If Extra_length = 0 Then
                                                        Valoare_x_shift = 28
                                                    Else
                                                        Valoare_x_shift = Extra_length + 14
                                                    End If
                                                    Pct_ins = New Point3d(X - Valoare_x_shift, Y, Z)
                                                End If

                                                Dim Val1, Val2 As Double
                                                Val1 = Round(ChainageT1, 1)
                                                Val2 = Round(ChainageT2, 1)

                                                Colectie_atr_name.Add("LENGTH")
                                                Colectie_atr_value.Add(Get_String_Rounded(Abs(Val2 - Val1), 1))
                                                Colectie_atr_name.Add("MAT")
                                                Colectie_atr_value.Add(Material1)

                                                If Not i = 0 Then
                                                    InsertBlock_with_multiple_atributes(Nume_pipeW & ".dwg", Nume_pipeW, Pct_ins, 1, BTrecord, "TEXT", Colectie_atr_name, Colectie_atr_value)
                                                End If


                                            End If ' If ChainageT1 < ChainageT2


                                            If Is_ScrewAnchor = True Then
                                                Extra_length_SA = Extra_length_SA + 14
                                            End If

                                            If Is_SandBag = True Then
                                                Extra_length_SB = Extra_length_SB + 14
                                            End If

                                            Extra_length = 0


                                            ChainageT2 = Replace(EndC, "+", "")

                                            If RadioButton_right_to_left.Checked = True Then
                                                X = X - 14
                                                Dim Temp As String = StartC
                                                StartC = EndC
                                                EndC = Temp

                                            Else
                                                X = X + 14

                                            End If


                                            Colectie_atr_name_el.Add("STA")
                                            Colectie_atr_value_el.Add(StaC)
                                            Colectie_atr_name_el.Add("BEGINSTA")
                                            Colectie_atr_value_el.Add(StartC)
                                            Colectie_atr_name_el.Add("ENDSTA")
                                            Colectie_atr_value_el.Add(EndC)
                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("LENGTH")) = False Then
                                                Colectie_atr_name_el.Add("LENGTH")
                                                Colectie_atr_value_el.Add(Get_String_Rounded(Data_table_Crossing.Rows(i).Item("LENGTH"), 1))
                                            End If
                                            Dim Desc As String = " "
                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("DESCRIPTION1")) = False Then
                                                Desc = Data_table_Crossing.Rows(i).Item("DESCRIPTION1")
                                            End If
                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("DESCRIPTION2")) = False Then
                                                Desc = Desc & vbCrLf & Data_table_Crossing.Rows(i).Item("DESCRIPTION2")
                                            End If
                                            Colectie_atr_name_el.Add("DESC")
                                            Colectie_atr_value_el.Add(Desc)

                                            If IsNothing(W1) = False Then
                                                Dim degree_elbow As String = extrage_numar_din_text_de_la_inceputul_textului(Desc.ToUpper)
                                                If IsNumeric(degree_elbow) = True Then
                                                    Dim Unghi As Double = CDbl(degree_elbow) * PI / 180
                                                    If IsNumeric(ComboBox_nps.Text) = True Then
                                                        Dim Diam As Double = 2 * Get_from_NPS_radius_for_pipes_from_inches_to_milimeters(CDbl(ComboBox_nps.Text)) / 1000

                                                        W1.Range("A" & Index_row_excel).Value = "ELBOW 3*D " & degree_elbow
                                                        W1.Range("B" & Index_row_excel).Value = StaC
                                                        W1.Range("C" & Index_row_excel).Value = Get_String_Rounded(Data_table_Crossing.Rows(i).Item("LENGTH"), 1)
                                                        W1.Range("D" & Index_row_excel).Value = Get_String_Rounded(2 * (3 * Diam * Tan(Unghi / 2) + 1), 1)

                                                        If Not Round(Data_table_Crossing.Rows(i).Item("LENGTH"), 1) = Round(2 * (3 * Diam * Tan(Unghi / 2) + 1), 1) Then
                                                            W1.Range("F" & Index_row_excel).Value = "Calculated 3xDiameter value not equal with the published value"
                                                        End If
                                                    Else
                                                        W1.Range("A" & Index_row_excel).Value = "ELBOW 3*D " & degree_elbow
                                                        W1.Range("B" & Index_row_excel).Value = StaC
                                                        W1.Range("C" & Index_row_excel).Value = Get_String_Rounded(Data_table_Crossing.Rows(i).Item("LENGTH"), 1)
                                                        W1.Range("F" & Index_row_excel).Value = "NPS not specified"
                                                    End If
                                                Else
                                                    W1.Range("B" & Index_row_excel).Value = StaC
                                                    W1.Range("C" & Index_row_excel).Value = Get_String_Rounded(Data_table_Crossing.Rows(i).Item("LENGTH"), 1)
                                                    W1.Range("F" & Index_row_excel).Value = "Elbow angle not specified"
                                                End If
                                                Dim Elbow_station As Double = CDbl(Replace(StaC, "+", ""))

                                                Index_row_excel = Index_row_excel + 1
                                            End If


                                            Dim iD_no As String = " "
                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("ID_NO")) = False Then
                                                iD_no = Data_table_Crossing.Rows(i).Item("ID_NO")
                                            End If
                                            Colectie_atr_name_el.Add("ID_NO")
                                            Colectie_atr_value_el.Add(iD_no)
                                            Dim Nume_elbow_block As String = "Elbow_alignment"

                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("MATERIAL")) = False Then
                                                If Data_table_Crossing.Rows(i).Item("MATERIAL") = "1" Then
                                                    Nume_elbow_block = "Elbow_alignment_mat1"
                                                End If
                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                    If Data_table_Crossing.Rows(i).Item("MATERIAL") = "2" Then
                                                        Nume_elbow_block = "Elbow_alignment_mat1"
                                                    End If
                                                End If
                                            End If

                                            Dim Pct_ins_elb As New Point3d
                                            If RadioButton_right_to_left.Checked = True Then
                                                Pct_ins_elb = New Point3d(X, Y, Z)
                                            Else
                                                Pct_ins_elb = New Point3d(X - 14, Y, Z)
                                            End If

                                            InsertBlock_with_multiple_atributes(Nume_elbow_block & ".dwg", Nume_elbow_block, Pct_ins_elb, 1, BTrecord, "TEXT", Colectie_atr_name_el, Colectie_atr_value_el)

                                            If CheckBox_extra_warning_signs.Checked = True Then
                                                If IsDBNull(Data_table_Crossing.Rows(i).Item("WARNING_SIGN")) = False Then
                                                    If Data_table_Crossing.Rows(i).Item("WARNING_SIGN") = True Then
                                                        InsertBlock_with_multiple_atributes("Warning_sign_alignment.dwg", "Warning_sign_alignment", New Point3d(Pct_ins_elb.X + 0, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_name_el, Colectie_atr_value_el)
                                                        InsertBlock_with_multiple_atributes("Warning_sign_alignment.dwg", "Warning_sign_alignment", New Point3d(Pct_ins_elb.X + 7, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_name_el, Colectie_atr_value_el)
                                                        InsertBlock_with_multiple_atributes("Warning_sign_alignment.dwg", "Warning_sign_alignment", New Point3d(Pct_ins_elb.X + 14, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_name_el, Colectie_atr_value_el)
                                                    End If
                                                End If

                                            End If

                                            If RadioButton_right_to_left.Checked = True Then
                                                X = X - 14
                                            Else
                                                X = X + 14
                                            End If


                                        Case "FACILITY"

                                            Dim Colectie_atr_name_el As New Specialized.StringCollection
                                            Dim Colectie_atr_value_el As New Specialized.StringCollection

                                            Dim StartC As String = ""
                                            Dim EndC As String = ""
                                            Dim StaC As String = ""

                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("ENDSTA")) = False Then
                                                EndC = Data_table_Crossing.Rows(i).Item("ENDSTA")
                                            End If

                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("BEGINSTA")) = False Then
                                                StartC = Data_table_Crossing.Rows(i).Item("BEGINSTA")
                                            End If

                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                                                StaC = Data_table_Crossing.Rows(i).Item("STA")
                                            End If
                                            If IsNumeric(StaC) = True Then
                                                StaC = Get_chainage_from_double(CDbl(StaC), 1)
                                            End If

                                            If IsNumeric(StartC) = True Then
                                                StartC = Get_chainage_from_double(CDbl(StartC), 1)
                                            End If

                                            If IsNumeric(EndC) = True Then
                                                EndC = Get_chainage_from_double(CDbl(EndC), 1)
                                            End If


                                            ChainageT1 = ChainageT2
                                            ChainageT2 = CDbl(Replace(StartC, "+", ""))


                                            Dim Nume_pipeW As String = "heavy_wall_1"
                                            If Not i = 0 Then
                                                If Extra_length = 0 Then
                                                    If RadioButton_right_to_left.Checked = True Then
                                                        X = X - 14
                                                    Else
                                                        X = X + 14
                                                    End If
                                                End If
                                            End If

                                            If RadioButton_right_to_left.Checked = True Then
                                                Select Case Round(Extra_length, 0)
                                                    Case 0
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_1"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_1"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_1"
                                                            End If

                                                        Else
                                                            Nume_pipeW = "heavy_wall_1"
                                                        End If

                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_right_1"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_1"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_1"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_1"
                                                            End If
                                                            Is_matchline = False
                                                        End If

                                                    Case 14
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_1"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_1"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_1"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_1"
                                                        End If

                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_right_1"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_1"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_1"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_1"
                                                            End If
                                                            Is_matchline = False
                                                        End If

                                                    Case 28
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_2"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_2"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_2"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_2"
                                                        End If


                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_right_2"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_2"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_2"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_2"
                                                            End If
                                                            Is_matchline = False
                                                        End If

                                                    Case 42
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_3"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_3"

                                                            Else
                                                                Nume_pipeW = "heavy_wall_3"
                                                            End If

                                                        Else
                                                            Nume_pipeW = "heavy_wall_3"
                                                        End If


                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_right_3"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then

                                                                    Nume_pipeW = "line_pipe_match_right_3"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_3"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_3"
                                                            End If
                                                            Is_matchline = False
                                                        End If


                                                    Case 56
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_4"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then

                                                                Nume_pipeW = "Line_pipe_4"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_4"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_4"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_right_4"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_4"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_4"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_4"
                                                            End If
                                                            Is_matchline = False
                                                        End If

                                                    Case 70
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_5"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_5"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_5"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_5"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_right_5"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_5"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_5"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_5"
                                                            End If
                                                            Is_matchline = False
                                                        End If

                                                    Case 84
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_6"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_6"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_6"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_6"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_right_6"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_6"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_6"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_6"
                                                            End If
                                                            Is_matchline = False
                                                        End If

                                                    Case 98
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_7"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_7"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_7"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_7"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_right_7"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_7"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_7"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_7"
                                                            End If
                                                            Is_matchline = False
                                                        End If

                                                    Case 112
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_8"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_8"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_8"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_8"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_right_8"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_8"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_8"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_8"
                                                            End If
                                                            Is_matchline = False
                                                        End If
                                                    Case 126
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_9"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_9"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_9"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_9"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_right_9"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_9"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_9"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_9"
                                                            End If
                                                            Is_matchline = False
                                                        End If
                                                    Case 140
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_10"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_10"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_10"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_10"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_right_10"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_10"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_10"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_10"
                                                            End If
                                                            Is_matchline = False
                                                        End If
                                                    Case 154
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_11"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_11"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_11"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_11"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_right_11"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_11"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_11"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_11"
                                                            End If
                                                            Is_matchline = False
                                                        End If
                                                    Case 168
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_12"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_12"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_12"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_12"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_right_12"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_12"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_12"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_12"
                                                            End If
                                                            Is_matchline = False
                                                        End If
                                                    Case 182
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_13"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_13"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_13"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_13"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_right_13"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_13"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_13"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_13"
                                                            End If
                                                            Is_matchline = False
                                                        End If
                                                    Case 196
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_14"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_14"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_14"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_14"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_right_14"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_14"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_14"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_14"
                                                            End If
                                                            Is_matchline = False
                                                        End If
                                                    Case 210
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_15"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_15"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_15"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_15"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_right_15"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_15"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_15"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_15"
                                                            End If
                                                            Is_matchline = False
                                                        End If
                                                    Case 224
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_16"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_16"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_16"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_16"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_right_16"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_16"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_16"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_16"
                                                            End If
                                                            Is_matchline = False
                                                        End If
                                                    Case 238
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_17"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_17"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_17"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_17"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_right_17"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_17"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_17"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_17"
                                                            End If
                                                            Is_matchline = False
                                                        End If
                                                    Case 252
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_18"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_18"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_18"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_18"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_right_18"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_18"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_18"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_18"
                                                            End If
                                                            Is_matchline = False
                                                        End If
                                                    Case 266
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_19"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_19"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_19"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_19"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_right_19"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_19"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_19"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_19"
                                                            End If
                                                            Is_matchline = False
                                                        End If
                                                    Case Else
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_20"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_20"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_20"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_20"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_right_20"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_20"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_20"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_20"
                                                            End If
                                                            Is_matchline = False
                                                        End If
                                                End Select

                                            End If

                                            If RadioButton_left_to_right.Checked = True Then
                                                Select Case Round(Extra_length, 0)
                                                    Case 0
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_1"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_1"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_1"
                                                            End If

                                                        Else
                                                            Nume_pipeW = "heavy_wall_1"
                                                        End If

                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_left_1"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_1"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_1"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_1"
                                                            End If
                                                            Is_matchline = False
                                                        End If

                                                    Case 14
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_1"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_1"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_1"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_1"
                                                        End If

                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_left_1"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_1"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_1"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_1"
                                                            End If
                                                            Is_matchline = False
                                                        End If

                                                    Case 28
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_2"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_2"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_2"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_2"
                                                        End If


                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_left_2"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_2"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_2"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_2"
                                                            End If
                                                            Is_matchline = False
                                                        End If

                                                    Case 42
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_3"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_3"

                                                            Else
                                                                Nume_pipeW = "heavy_wall_3"
                                                            End If

                                                        Else
                                                            Nume_pipeW = "heavy_wall_3"
                                                        End If


                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_left_3"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then

                                                                    Nume_pipeW = "line_pipe_match_left_3"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_3"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_3"
                                                            End If
                                                            Is_matchline = False
                                                        End If


                                                    Case 56
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_4"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then

                                                                Nume_pipeW = "Line_pipe_4"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_4"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_4"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_left_4"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_4"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_4"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_4"
                                                            End If
                                                            Is_matchline = False
                                                        End If

                                                    Case 70
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_5"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_5"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_5"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_5"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_left_5"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_5"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_5"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_5"
                                                            End If
                                                            Is_matchline = False
                                                        End If

                                                    Case 84
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_6"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_6"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_6"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_6"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_left_6"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_6"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_6"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_6"
                                                            End If
                                                            Is_matchline = False
                                                        End If

                                                    Case 98
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_7"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_7"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_7"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_7"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_left_7"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_7"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_7"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_7"
                                                            End If
                                                            Is_matchline = False
                                                        End If

                                                    Case 112
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_8"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_8"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_8"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_8"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_left_8"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_8"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_8"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_8"
                                                            End If
                                                            Is_matchline = False
                                                        End If
                                                    Case 126
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_9"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_9"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_9"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_9"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_left_9"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_9"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_9"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_9"
                                                            End If
                                                            Is_matchline = False
                                                        End If
                                                    Case 140
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_10"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_10"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_10"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_10"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_left_10"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_10"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_10"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_10"
                                                            End If
                                                            Is_matchline = False
                                                        End If
                                                    Case 154
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_11"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_11"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_11"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_11"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_left_11"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_11"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_11"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_11"
                                                            End If
                                                            Is_matchline = False
                                                        End If
                                                    Case 168
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_12"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_12"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_12"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_12"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_left_12"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_12"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_12"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_12"
                                                            End If
                                                            Is_matchline = False
                                                        End If
                                                    Case 182
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_13"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_13"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_13"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_13"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_left_13"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_13"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_13"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_13"
                                                            End If
                                                            Is_matchline = False
                                                        End If
                                                    Case 196
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_14"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_14"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_14"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_14"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_left_14"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_14"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_14"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_14"
                                                            End If
                                                            Is_matchline = False
                                                        End If
                                                    Case 210
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_15"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_15"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_15"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_15"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_left_15"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_15"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_15"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_15"
                                                            End If
                                                            Is_matchline = False
                                                        End If
                                                    Case 224
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_16"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_16"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_16"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_16"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_left_16"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_16"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_16"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_16"
                                                            End If
                                                            Is_matchline = False
                                                        End If
                                                    Case 238
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_17"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_17"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_17"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_17"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_left_17"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_17"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_17"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_17"
                                                            End If
                                                            Is_matchline = False
                                                        End If
                                                    Case 252
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_18"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_18"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_18"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_18"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_left_18"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_18"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_18"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_18"
                                                            End If
                                                            Is_matchline = False
                                                        End If
                                                    Case 266
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_19"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_19"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_19"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_19"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_left_19"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_19"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_19"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_19"
                                                            End If
                                                            Is_matchline = False
                                                        End If
                                                    Case Else
                                                        If Material1 = "1" Then
                                                            Nume_pipeW = "mat_20"
                                                        ElseIf Material1 = "2" Then
                                                            If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                Nume_pipeW = "Line_pipe_20"
                                                            Else
                                                                Nume_pipeW = "heavy_wall_20"
                                                            End If
                                                        Else
                                                            Nume_pipeW = "heavy_wall_20"
                                                        End If
                                                        If Is_matchline = True Then
                                                            If Material1 = "1" Then
                                                                Nume_pipeW = "mat1_match_left_20"
                                                            ElseIf Material1 = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_20"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_20"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_20"
                                                            End If
                                                            Is_matchline = False
                                                        End If
                                                End Select

                                            End If



                                            Dim Colectie_atr_name As New Specialized.StringCollection
                                            Dim Colectie_atr_value As New Specialized.StringCollection

                                            Colectie_atr_name.Add("BEGINSTA")
                                            Colectie_atr_name.Add("ENDSTA")
                                            Dim Pct_ins As New Point3d
                                            If RadioButton_right_to_left.Checked = True Then
                                                Colectie_atr_value.Add(Get_chainage_from_double(ChainageT2, 1))
                                                Colectie_atr_value.Add(Get_chainage_from_double(ChainageT1, 1))
                                                Pct_ins = New Point3d(X, Y, Z)
                                            Else
                                                Colectie_atr_value.Add(Get_chainage_from_double(ChainageT1, 1))
                                                Colectie_atr_value.Add(Get_chainage_from_double(ChainageT2, 1))

                                                Pct_ins = New Point3d(X - (Extra_length + 14), Y, Z)
                                            End If

                                            Dim Val1, Val2 As Double
                                            Val1 = Round(ChainageT1, 1)
                                            Val2 = Round(ChainageT2, 1)

                                            Colectie_atr_name.Add("LENGTH")
                                            Colectie_atr_value.Add(Get_String_Rounded(Abs(Val2 - Val1), 1))
                                            Colectie_atr_name.Add("MAT")
                                            Colectie_atr_value.Add(Material1)

                                            If Not i = 0 Then
                                                InsertBlock_with_multiple_atributes(Nume_pipeW & ".dwg", Nume_pipeW, Pct_ins, 1, BTrecord, "TEXT", Colectie_atr_name, Colectie_atr_value)
                                            End If



                                            If Is_ScrewAnchor = True Then
                                                Extra_length_SA = Extra_length_SA + 28
                                            End If
                                            If Is_SandBag = True Then
                                                Extra_length_SB = Extra_length_SB + 28
                                            End If

                                            Extra_length = 0


                                            ChainageT2 = Replace(EndC, "+", "")

                                            If RadioButton_right_to_left.Checked = True Then
                                                X = X - 28
                                                Dim Temp As String = StartC
                                                StartC = EndC
                                                EndC = Temp

                                            Else
                                                X = X + 28

                                            End If




                                            Colectie_atr_name_el.Add("STA")
                                            Colectie_atr_value_el.Add(StaC)
                                            Colectie_atr_name_el.Add("BEGINSTA")
                                            Colectie_atr_value_el.Add(StartC)
                                            Colectie_atr_name_el.Add("ENDSTA")
                                            Colectie_atr_value_el.Add(EndC)
                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("LENGTH")) = False Then
                                                Colectie_atr_name_el.Add("LENGTH")
                                                Colectie_atr_value_el.Add(Get_String_Rounded(Data_table_Crossing.Rows(i).Item("LENGTH"), 1))
                                            End If

                                            Colectie_atr_name_el.Add("MAT")
                                            Colectie_atr_value_el.Add(Material1)


                                            Dim Nume_facility_block As String = "facility_temp_block"

                                            Dim Valoare_x_shift As Double = 0
                                            If RadioButton_left_to_right.Checked = True Then
                                                Valoare_x_shift = 28
                                            End If

                                            InsertBlock_with_multiple_atributes(Nume_facility_block & ".dwg", Nume_facility_block, New Point3d(X - Valoare_x_shift, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_name_el, Colectie_atr_value_el)


                                            If RadioButton_right_to_left.Checked = True Then
                                                X = X - 14
                                            Else
                                                X = X + 14
                                            End If


                                        Case "Pipe_alignment_crossing"
                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                                                Dim Colectie_atr_name_PIPE As New Specialized.StringCollection
                                                Dim Colectie_atr_value_PIPE As New Specialized.StringCollection

                                                'Nr_crossings = Nr_crossings + 1
                                                Colectie_atr_name_PIPE.Add("STA")
                                                Colectie_atr_value_PIPE.Add(Get_chainage_from_double(Data_table_Crossing.Rows(i).Item("STA"), 1))

                                                Dim Desc As String = " "
                                                If IsDBNull(Data_table_Crossing.Rows(i).Item("DESCRIPTION1")) = False Then
                                                    Desc = Data_table_Crossing.Rows(i).Item("DESCRIPTION1")
                                                End If

                                                If IsDBNull(Data_table_Crossing.Rows(i).Item("DESCRIPTION2")) = False Then
                                                    Desc = Desc & vbCrLf & Data_table_Crossing.Rows(i).Item("DESCRIPTION2")
                                                End If
                                                Colectie_atr_name_PIPE.Add("DESC")
                                                Colectie_atr_value_PIPE.Add(Desc)
                                                Dim iD_no As String = " "
                                                If IsDBNull(Data_table_Crossing.Rows(i).Item("ID_NO")) = False Then
                                                    iD_no = Data_table_Crossing.Rows(i).Item("ID_NO")
                                                End If
                                                Colectie_atr_name_PIPE.Add("ID_NO")
                                                Colectie_atr_value_PIPE.Add(iD_no)
                                                Dim Cover As String = " "
                                                If IsDBNull(Data_table_Crossing.Rows(i).Item("COVER")) = False Then
                                                    Cover = Data_table_Crossing.Rows(i).Item("COVER")
                                                End If
                                                Colectie_atr_name_PIPE.Add("COV")
                                                Colectie_atr_value_PIPE.Add(Cover)

                                                If Is_CP = False And IsDBNull(Data_table_Crossing.Rows(i).Item("CP")) = True Then
                                                    InsertBlock_with_multiple_atributes("Pipe_alignment_crossing.dwg", "Pipe_alignment_crossing", New Point3d(X, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_name_PIPE, Colectie_atr_value_PIPE)
                                                Else
                                                    InsertBlock_with_multiple_atributes("Pipe_alignment_crossing_CP.dwg", "Pipe_alignment_crossing_CP", New Point3d(X, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_name_PIPE, Colectie_atr_value_PIPE)
                                                    Is_CP = False
                                                End If



                                                If IsDBNull(Data_table_Crossing.Rows(i).Item("WARNING_SIGN")) = False Then
                                                    If CheckBox_Warning_Signs.Checked = False Then
                                                        If Data_table_Crossing.Rows(i).Item("WARNING_SIGN") = True Then
                                                            InsertBlock_with_multiple_atributes("Warning_sign_alignment.dwg", "Warning_sign_alignment", New Point3d(X - 4, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_name_PIPE, Colectie_atr_value_PIPE)
                                                            InsertBlock_with_multiple_atributes("Warning_sign_alignment.dwg", "Warning_sign_alignment", New Point3d(X + 4, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_name_PIPE, Colectie_atr_value_PIPE)
                                                        End If
                                                    End If

                                                End If


                                                If RadioButton_right_to_left.Checked = True Then
                                                    X = X - 14
                                                Else
                                                    X = X + 14
                                                End If

                                                Extra_length = Extra_length + 14
                                                If Is_ScrewAnchor = True Then
                                                    Extra_length_SA = Extra_length_SA + 14
                                                End If
                                                If Is_SandBag = True Then
                                                    Extra_length_SB = Extra_length_SB + 14
                                                End If

                                            End If

                                        Case "Cable_alignment_crossing"
                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                                                Dim Colectie_atr_name_PIPE As New Specialized.StringCollection
                                                Dim Colectie_atr_value_PIPE As New Specialized.StringCollection
                                                If IsNumeric(Data_table_Crossing.Rows(i).Item("STA")) = True Then
                                                    ' Nr_crossings = Nr_crossings + 1
                                                    Colectie_atr_name_PIPE.Add("STA")
                                                    Colectie_atr_value_PIPE.Add(Get_chainage_from_double(Data_table_Crossing.Rows(i).Item("STA"), 1))
                                                    Dim Desc As String = " "
                                                    If IsDBNull(Data_table_Crossing.Rows(i).Item("DESCRIPTION1")) = False Then
                                                        Desc = Data_table_Crossing.Rows(i).Item("DESCRIPTION1")
                                                    End If

                                                    If IsDBNull(Data_table_Crossing.Rows(i).Item("DESCRIPTION2")) = False Then
                                                        Desc = Desc & vbCrLf & Data_table_Crossing.Rows(i).Item("DESCRIPTION2")
                                                    End If
                                                    Colectie_atr_name_PIPE.Add("DESC")
                                                    Colectie_atr_value_PIPE.Add(Desc)
                                                    Dim iD_no As String = " "
                                                    If IsDBNull(Data_table_Crossing.Rows(i).Item("ID_NO")) = False Then
                                                        iD_no = Data_table_Crossing.Rows(i).Item("ID_NO")
                                                    End If
                                                    Colectie_atr_name_PIPE.Add("ID_NO")
                                                    Colectie_atr_value_PIPE.Add(iD_no)
                                                    Dim Cover As String = " "
                                                    If IsDBNull(Data_table_Crossing.Rows(i).Item("COVER")) = False Then
                                                        Cover = Data_table_Crossing.Rows(i).Item("COVER")
                                                    End If
                                                    Colectie_atr_name_PIPE.Add("COV")
                                                    Colectie_atr_value_PIPE.Add(Cover)
                                                    If Is_CP = False And IsDBNull(Data_table_Crossing.Rows(i).Item("CP")) = True Then
                                                        InsertBlock_with_multiple_atributes("Cable_alignment_crossing.dwg", "Cable_alignment_crossing", New Point3d(X, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_name_PIPE, Colectie_atr_value_PIPE)
                                                    Else
                                                        InsertBlock_with_multiple_atributes("Cable_alignment_crossing_CP.dwg", "Cable_alignment_crossing_CP", New Point3d(X, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_name_PIPE, Colectie_atr_value_PIPE)
                                                        Is_CP = False
                                                    End If



                                                    If IsDBNull(Data_table_Crossing.Rows(i).Item("WARNING_SIGN")) = False Then
                                                        If CheckBox_Warning_Signs.Checked = False Then
                                                            If Data_table_Crossing.Rows(i).Item("WARNING_SIGN") = True Then
                                                                InsertBlock_with_multiple_atributes("Warning_sign_alignment.dwg", "Warning_sign_alignment", New Point3d(X - 4, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_name_PIPE, Colectie_atr_value_PIPE)
                                                                InsertBlock_with_multiple_atributes("Warning_sign_alignment.dwg", "Warning_sign_alignment", New Point3d(X + 4, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_name_PIPE, Colectie_atr_value_PIPE)
                                                            End If
                                                        End If

                                                    End If


                                                    If RadioButton_right_to_left.Checked = True Then
                                                        X = X - 14
                                                    Else
                                                        X = X + 14
                                                    End If
                                                    Extra_length = Extra_length + 14
                                                    If Is_ScrewAnchor = True Then
                                                        Extra_length_SA = Extra_length_SA + 14
                                                    End If
                                                    If Is_SandBag = True Then
                                                        Extra_length_SB = Extra_length_SB + 14
                                                    End If
                                                End If
                                            End If

                                        Case "General_alignment_crossing"
                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                                                Dim Colectie_atr_name_PIPE As New Specialized.StringCollection
                                                Dim Colectie_atr_value_PIPE As New Specialized.StringCollection
                                                If IsNumeric(Data_table_Crossing.Rows(i).Item("STA")) = True Then
                                                    'Nr_crossings = Nr_crossings + 1
                                                    Colectie_atr_name_PIPE.Add("STA")
                                                    Colectie_atr_value_PIPE.Add(Get_chainage_from_double(Data_table_Crossing.Rows(i).Item("STA"), 1))
                                                    Dim Desc As String = " "
                                                    If IsDBNull(Data_table_Crossing.Rows(i).Item("DESCRIPTION1")) = False Then
                                                        Desc = Data_table_Crossing.Rows(i).Item("DESCRIPTION1")
                                                    End If

                                                    If IsDBNull(Data_table_Crossing.Rows(i).Item("DESCRIPTION2")) = False Then
                                                        Desc = Desc & vbCrLf & Data_table_Crossing.Rows(i).Item("DESCRIPTION2")
                                                    End If
                                                    Colectie_atr_name_PIPE.Add("DESC")
                                                    Colectie_atr_value_PIPE.Add(Desc)
                                                    Dim iD_no As String = " "
                                                    If IsDBNull(Data_table_Crossing.Rows(i).Item("ID_NO")) = False Then
                                                        iD_no = Data_table_Crossing.Rows(i).Item("ID_NO")
                                                    End If

                                                    Colectie_atr_name_PIPE.Add("ID_NO")
                                                    Colectie_atr_value_PIPE.Add(iD_no)

                                                    Dim Cover As String = " "
                                                    If IsDBNull(Data_table_Crossing.Rows(i).Item("COVER")) = False Then
                                                        Cover = Data_table_Crossing.Rows(i).Item("COVER")
                                                        Colectie_atr_name_PIPE.Add("COV")
                                                        Colectie_atr_value_PIPE.Add(Cover)
                                                    End If

                                                    If Is_CP = False And IsDBNull(Data_table_Crossing.Rows(i).Item("CP")) = True Then
                                                        InsertBlock_with_multiple_atributes("General_alignment_crossing.dwg", "General_alignment_crossing", New Point3d(X, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_name_PIPE, Colectie_atr_value_PIPE)
                                                    Else
                                                        InsertBlock_with_multiple_atributes("General_alignment_crossing_CP.dwg", "General_alignment_crossing_CP", New Point3d(X, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_name_PIPE, Colectie_atr_value_PIPE)
                                                        Is_CP = False
                                                    End If



                                                    If IsDBNull(Data_table_Crossing.Rows(i).Item("WARNING_SIGN")) = False Then
                                                        If CheckBox_Warning_Signs.Checked = False Then
                                                            If Data_table_Crossing.Rows(i).Item("WARNING_SIGN") = True Then
                                                                InsertBlock_with_multiple_atributes("Warning_sign_alignment.dwg", "Warning_sign_alignment", New Point3d(X - 4, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_name_PIPE, Colectie_atr_value_PIPE)
                                                                InsertBlock_with_multiple_atributes("Warning_sign_alignment.dwg", "Warning_sign_alignment", New Point3d(X + 4, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_name_PIPE, Colectie_atr_value_PIPE)
                                                            End If
                                                        End If
                                                    End If



                                                    If RadioButton_right_to_left.Checked = True Then
                                                        X = X - 14
                                                    Else
                                                        X = X + 14
                                                    End If
                                                    Extra_length = Extra_length + 14
                                                    If Is_ScrewAnchor = True Then
                                                        Extra_length_SA = Extra_length_SA + 14
                                                    End If
                                                    If Is_SandBag = True Then
                                                        Extra_length_SB = Extra_length_SB + 14
                                                    End If
                                                End If
                                            End If
                                        Case "Test_section_alignment_crossing"
                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                                                Dim Colectie_atr_name_PIPE As New Specialized.StringCollection
                                                Dim Colectie_atr_value_PIPE As New Specialized.StringCollection
                                                If IsNumeric(Data_table_Crossing.Rows(i).Item("STA")) = True Then
                                                    'Nr_crossings = Nr_crossings + 1
                                                    Colectie_atr_name_PIPE.Add("STA")
                                                    Colectie_atr_value_PIPE.Add(Get_chainage_from_double(Data_table_Crossing.Rows(i).Item("STA"), 1))
                                                    Dim Desc As String = " "
                                                    If IsDBNull(Data_table_Crossing.Rows(i).Item("DESCRIPTION1")) = False Then
                                                        Desc = Data_table_Crossing.Rows(i).Item("DESCRIPTION1")
                                                    End If
                                                    Colectie_atr_name_PIPE.Add("DESC")
                                                    Colectie_atr_value_PIPE.Add(Desc)


                                                    If IsDBNull(Data_table_Crossing.Rows(i).Item("DESCRIPTION2")) = False Then
                                                        Dim iD_no As String = " "
                                                        iD_no = Data_table_Crossing.Rows(i).Item("DESCRIPTION2")
                                                        Colectie_atr_name_PIPE.Add("ID_NO")
                                                        Colectie_atr_value_PIPE.Add(iD_no)
                                                    End If



                                                    If Is_CP = False And IsDBNull(Data_table_Crossing.Rows(i).Item("CP")) = True Then
                                                        InsertBlock_with_multiple_atributes("Test_section_alignment_crossing.dwg", "Test_section_alignment_crossing", New Point3d(X, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_name_PIPE, Colectie_atr_value_PIPE)
                                                    Else
                                                        InsertBlock_with_multiple_atributes("Test_section_alignment_crossing_CP.dwg", "Test_section_alignment_crossing_CP", New Point3d(X, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_name_PIPE, Colectie_atr_value_PIPE)
                                                        Is_CP = False
                                                    End If

                                                    If RadioButton_right_to_left.Checked = True Then
                                                        X = X - 14
                                                    Else
                                                        X = X + 14
                                                    End If
                                                    Extra_length = Extra_length + 14
                                                    If Is_ScrewAnchor = True Then
                                                        Extra_length_SA = Extra_length_SA + 14
                                                    End If
                                                    If Is_SandBag = True Then
                                                        Extra_length_SB = Extra_length_SB + 14
                                                    End If
                                                End If
                                            End If



                                        Case "RIVER_WEIGHT_START"
                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                                                ChainageRW1 = Data_table_Crossing.Rows(i).Item("STA")
                                                ChainageRW1_for_Matchline_calcs = ChainageRW1
                                                If RadioButton_right_to_left.Checked = True Then
                                                    X = X - 14
                                                Else
                                                    X = X + 14
                                                End If
                                                Extra_length = Extra_length + 14
                                                Is_primulT_for_rw = True
                                                Is_Riverweight = True
                                                Numar_RW_inserted = 0
                                            Else
                                                ChainageRW1 = 0
                                            End If

                                        Case "RIVER_WEIGHT_END"
                                            If Is_Riverweight = True Then
                                                If Not ChainageRW1 = 0 Then
                                                    If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                                                        ChainageRW2 = Data_table_Crossing.Rows(i).Item("STA")


                                                        If ChainageRW2 < ChainageRW1 Then
                                                            MsgBox("Screw Anchors issue at station " & ChainageRW2)
                                                            Freeze_operations = False
                                                            Exit Sub
                                                        End If

                                                        Dim nume_block_RW As String = "River_weight_alignment1"

                                                        If RadioButton_right_to_left.Checked = True Then
                                                            Select Case Round(Extra_length_RW, 0)
                                                                Case 0
                                                                    nume_block_RW = "River_weight_alignment1"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_right1"
                                                                        Is_matchline_river_weight = False
                                                                    End If
                                                                Case 14
                                                                    nume_block_RW = "River_weight_alignment1"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_right1"
                                                                        Is_matchline_river_weight = False
                                                                    End If

                                                                Case 28
                                                                    nume_block_RW = "River_weight_alignment2"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_right2"
                                                                        Is_matchline_river_weight = False
                                                                    End If

                                                                Case 42
                                                                    nume_block_RW = "River_weight_alignment3"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_right3"
                                                                        Is_matchline_river_weight = False
                                                                    End If

                                                                Case 56
                                                                    nume_block_RW = "River_weight_alignment4"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_right4"
                                                                        Is_matchline_river_weight = False
                                                                    End If

                                                                Case 70
                                                                    nume_block_RW = "River_weight_alignment5"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_right5"
                                                                        Is_matchline_river_weight = False
                                                                    End If

                                                                Case 84
                                                                    nume_block_RW = "River_weight_alignment6"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_right6"
                                                                        Is_matchline_river_weight = False
                                                                    End If

                                                                Case 98
                                                                    nume_block_RW = "River_weight_alignment7"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_right7"
                                                                        Is_matchline_river_weight = False
                                                                    End If
                                                                Case 112
                                                                    nume_block_RW = "River_weight_alignment8"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_right8"
                                                                        Is_matchline_river_weight = False
                                                                    End If
                                                                Case 126
                                                                    nume_block_RW = "River_weight_alignment9"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_right9"
                                                                        Is_matchline_river_weight = False
                                                                    End If
                                                                Case 140
                                                                    nume_block_RW = "River_weight_alignment10"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_right10"
                                                                        Is_matchline_river_weight = False
                                                                    End If
                                                                Case 154
                                                                    nume_block_RW = "River_weight_alignment11"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_right11"
                                                                        Is_matchline_river_weight = False
                                                                    End If
                                                                Case 168
                                                                    nume_block_RW = "River_weight_alignment12"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_right12"
                                                                        Is_matchline_river_weight = False
                                                                    End If
                                                                Case 182
                                                                    nume_block_RW = "River_weight_alignment13"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_right13"
                                                                        Is_matchline_river_weight = False
                                                                    End If
                                                                Case 196
                                                                    nume_block_RW = "River_weight_alignment14"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_right14"
                                                                        Is_matchline_river_weight = False
                                                                    End If
                                                                Case 210
                                                                    nume_block_RW = "River_weight_alignment15"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_right15"
                                                                        Is_matchline_river_weight = False
                                                                    End If
                                                                Case 224
                                                                    nume_block_RW = "River_weight_alignment16"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_right16"
                                                                        Is_matchline_river_weight = False
                                                                    End If
                                                                Case 238
                                                                    nume_block_RW = "River_weight_alignment17"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_right17"
                                                                        Is_matchline_river_weight = False
                                                                    End If
                                                                Case 252
                                                                    nume_block_RW = "River_weight_alignment18"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_right18"
                                                                        Is_matchline_river_weight = False
                                                                    End If
                                                                Case 266
                                                                    nume_block_RW = "River_weight_alignment19"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_right19"
                                                                        Is_matchline_river_weight = False
                                                                    End If
                                                                Case Else
                                                                    nume_block_RW = "River_weight_alignment20"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_right20"
                                                                        Is_matchline_river_weight = False
                                                                    End If

                                                            End Select
                                                        End If

                                                        If RadioButton_left_to_right.Checked = True Then
                                                            Select Case Round(Extra_length_RW, 0)
                                                                Case 0
                                                                    nume_block_RW = "River_weight_alignment1"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_left1"
                                                                        Is_matchline_river_weight = False
                                                                    End If
                                                                Case 14
                                                                    nume_block_RW = "River_weight_alignment1"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_left1"
                                                                        Is_matchline_river_weight = False
                                                                    End If

                                                                Case 28
                                                                    nume_block_RW = "River_weight_alignment2"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_left2"
                                                                        Is_matchline_river_weight = False
                                                                    End If

                                                                Case 42
                                                                    nume_block_RW = "River_weight_alignment3"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_left3"
                                                                        Is_matchline_river_weight = False
                                                                    End If

                                                                Case 56
                                                                    nume_block_RW = "River_weight_alignment4"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_left4"
                                                                        Is_matchline_river_weight = False
                                                                    End If

                                                                Case 70
                                                                    nume_block_RW = "River_weight_alignment5"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_left5"
                                                                        Is_matchline_river_weight = False
                                                                    End If

                                                                Case 84
                                                                    nume_block_RW = "River_weight_alignment6"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_left6"
                                                                        Is_matchline_river_weight = False
                                                                    End If

                                                                Case 98
                                                                    nume_block_RW = "River_weight_alignment7"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_left7"
                                                                        Is_matchline_river_weight = False
                                                                    End If
                                                                Case 112
                                                                    nume_block_RW = "River_weight_alignment8"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_left8"
                                                                        Is_matchline_river_weight = False
                                                                    End If
                                                                Case 126
                                                                    nume_block_RW = "River_weight_alignment9"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_left9"
                                                                        Is_matchline_river_weight = False
                                                                    End If
                                                                Case 140
                                                                    nume_block_RW = "River_weight_alignment10"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_left10"
                                                                        Is_matchline_river_weight = False
                                                                    End If
                                                                Case 154
                                                                    nume_block_RW = "River_weight_alignment11"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_left11"
                                                                        Is_matchline_river_weight = False
                                                                    End If
                                                                Case 168
                                                                    nume_block_RW = "River_weight_alignment12"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_left12"
                                                                        Is_matchline_river_weight = False
                                                                    End If
                                                                Case 182
                                                                    nume_block_RW = "River_weight_alignment13"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_left13"
                                                                        Is_matchline_river_weight = False
                                                                    End If
                                                                Case 196
                                                                    nume_block_RW = "River_weight_alignment14"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_left14"
                                                                        Is_matchline_river_weight = False
                                                                    End If
                                                                Case 210
                                                                    nume_block_RW = "River_weight_alignment15"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_left15"
                                                                        Is_matchline_river_weight = False
                                                                    End If
                                                                Case 224
                                                                    nume_block_RW = "River_weight_alignment16"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_left16"
                                                                        Is_matchline_river_weight = False
                                                                    End If
                                                                Case 238
                                                                    nume_block_RW = "River_weight_alignment17"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_left17"
                                                                        Is_matchline_river_weight = False
                                                                    End If
                                                                Case 252
                                                                    nume_block_RW = "River_weight_alignment18"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_left18"
                                                                        Is_matchline_river_weight = False
                                                                    End If
                                                                Case 266
                                                                    nume_block_RW = "River_weight_alignment19"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_left19"
                                                                        Is_matchline_river_weight = False
                                                                    End If
                                                                Case Else
                                                                    nume_block_RW = "River_weight_alignment20"

                                                                    If Is_matchline_river_weight = True Then
                                                                        nume_block_RW = "River_weight_alignment_match_left20"
                                                                        Is_matchline_river_weight = False
                                                                    End If

                                                            End Select
                                                        End If


                                                        Dim Colectie_atr_name As New Specialized.StringCollection
                                                        Dim Colectie_atr_value As New Specialized.StringCollection


                                                        Colectie_atr_name.Add("BEGINSTA")
                                                        Colectie_atr_name.Add("ENDSTA")

                                                        If RadioButton_right_to_left.Checked = True Then
                                                            Colectie_atr_value.Add(Get_chainage_from_double(ChainageRW2, 1))
                                                            Colectie_atr_value.Add(Get_chainage_from_double(ChainageRW1, 1))

                                                        Else
                                                            Colectie_atr_value.Add(Get_chainage_from_double(ChainageRW1, 1))
                                                            Colectie_atr_value.Add(Get_chainage_from_double(ChainageRW2, 1))
                                                        End If

                                                        Dim Val1, Val2, Val3 As Double
                                                        Val1 = Round(ChainageRW1, 1)
                                                        Val2 = Round(ChainageRW2, 1)
                                                        Val3 = Round(ChainageRW1_for_Matchline_calcs, 1)

                                                        If IsDBNull(Data_table_Crossing.Rows(i).Item("BUOYANCY")) = False Then
                                                            Dim Spacing_rw As Double = Round(Data_table_Crossing.Rows(i).Item("BUOYANCY"), 1)
                                                            If Spacing_rw > 0 Then
                                                                Dim Length_rw As Double = Round(Abs(Val2 - Val1), 1)

                                                                Dim NR_rw As Integer = CInt(1 + Length_rw / Spacing_rw)

                                                                If Numar_RW_inserted > 0 Then
                                                                    Dim Length_rw_for_match As Double = Round(Abs(Val2 - Val3), 1)
                                                                    Dim NR_rw_for_match As Integer = CInt(1 + Length_rw_for_match / Spacing_rw)
                                                                    NR_rw = NR_rw_for_match - Numar_RW_inserted
                                                                    Numar_RW_inserted = 0
                                                                End If



                                                                Colectie_atr_name.Add("NO_TYPE")
                                                                Colectie_atr_name.Add("SPACING")
                                                                Colectie_atr_value.Add(NR_rw & " RW")
                                                                Colectie_atr_value.Add(Get_String_Rounded(Spacing_rw, 1) & " C/C")

                                                            End If


                                                        End If


                                                        If Extra_length_RW = 0 Then
                                                            If RadioButton_right_to_left.Checked = True Then
                                                                X = X - 14
                                                            Else
                                                                X = X + 14
                                                            End If
                                                        End If

                                                        Dim Valoare_x_shift As Double = 0
                                                        If RadioButton_left_to_right.Checked = True Then
                                                            If Extra_length_RW = 0 Then
                                                                Valoare_x_shift = 28
                                                            Else
                                                                Valoare_x_shift = Extra_length_RW + 14
                                                            End If
                                                        End If



                                                        InsertBlock_with_multiple_atributes(nume_block_RW & ".dwg", nume_block_RW, New Point3d(X - Valoare_x_shift, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_name, Colectie_atr_value)

                                                        If RadioButton_right_to_left.Checked = True Then
                                                            X = X - 14
                                                        Else
                                                            X = X + 14
                                                        End If



                                                        If Extra_length_RW = 0 Then
                                                            Extra_length = Extra_length + 28
                                                        Else
                                                            Extra_length = Extra_length + 14
                                                        End If

                                                        Extra_length_RW = 0
                                                        Is_Riverweight = False
                                                        Is_primulT_for_rw = False
                                                        Is_matchline_river_weight = False
                                                        ChainageRW1 = 0
                                                    End If
                                                Else
                                                    MsgBox("River Weight issue at station (River Weight Start = 0)" & ChainageRW2)
                                                    Freeze_operations = False
                                                    Exit Sub
                                                End If ' asta e de la  If Not Chainagerw1 = 0


                                            Else
                                                MsgBox("River Weight issue at station (No River Weight Start)" & ChainageRW2)
                                                Freeze_operations = False
                                                Exit Sub
                                            End If ' asta e de la If is_riverweight = True









                                        Case "MATCHLINE"
                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                                                ChainageT1 = ChainageT2
                                                ChainageT2 = Data_table_Crossing.Rows(i).Item("STA")

                                                Dim Nume_pipeW As String = "heavy_wall_1"
                                                If Not i = 0 Then
                                                    If Extra_length = 0 Then
                                                        If RadioButton_right_to_left.Checked = True Then
                                                            X = X - 14
                                                        Else
                                                            X = X + 14
                                                        End If
                                                    End If
                                                End If


                                                If RadioButton_left_to_right.Checked = True Then
                                                    Select Case Round(Extra_length, 0)

                                                        Case 0
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_right_1"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_1"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_1"
                                                                End If

                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_1"
                                                            End If
                                                        Case 14
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_right_1"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_1"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_1"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_1"
                                                            End If
                                                        Case 28
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_right_2"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_2"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_2"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_2"
                                                            End If
                                                        Case 42
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_right_3"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_3"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_3"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_3"
                                                            End If
                                                        Case 56
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_right_4"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_4"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_4"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_4"
                                                            End If
                                                        Case 70
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_right_5"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_5"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_5"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_5"
                                                            End If
                                                        Case 84
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_right_6"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_6"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_6"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_6"
                                                            End If
                                                        Case 98
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_right_7"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_7"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_7"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_7"
                                                            End If
                                                        Case 112
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_right_8"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_8"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_8"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_8"
                                                            End If
                                                        Case 126
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_right_9"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_9"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_9"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_9"
                                                            End If
                                                        Case 140
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_right_10"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_10"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_10"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_10"
                                                            End If
                                                        Case 154
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_right_11"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_11"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_11"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_11"
                                                            End If
                                                        Case 168
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_right_12"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_12"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_12"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_12"
                                                            End If
                                                        Case 182
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_right_13"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_13"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_13"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_13"
                                                            End If
                                                        Case 196
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_right_14"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_14"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_14"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_14"
                                                            End If
                                                        Case 210
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_right_15"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_15"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_15"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_15"
                                                            End If
                                                        Case 224
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_right_16"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_16"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_16"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_16"
                                                            End If
                                                        Case 238
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_right_17"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_17"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_17"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_17"
                                                            End If
                                                        Case 252
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_right_18"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_18"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_18"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_18"
                                                            End If
                                                        Case 266
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_right_19"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_19"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_19"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_19"
                                                            End If
                                                        Case Else
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_right_20"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_right_20"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_right_20"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_right_20"
                                                            End If
                                                    End Select
                                                End If

                                                If RadioButton_right_to_left.Checked = True Then
                                                    Select Case Round(Extra_length, 0)

                                                        Case 0
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_left_1"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_1"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_1"
                                                                End If

                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_1"
                                                            End If
                                                        Case 14
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_left_1"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_1"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_1"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_1"
                                                            End If
                                                        Case 28
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_left_2"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_2"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_2"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_2"
                                                            End If
                                                        Case 42
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_left_3"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_3"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_3"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_3"
                                                            End If
                                                        Case 56
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_left_4"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_4"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_4"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_4"
                                                            End If
                                                        Case 70
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_left_5"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_5"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_5"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_5"
                                                            End If
                                                        Case 84
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_left_6"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_6"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_6"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_6"
                                                            End If
                                                        Case 98
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_left_7"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_7"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_7"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_7"
                                                            End If
                                                        Case 112
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_left_8"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_8"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_8"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_8"
                                                            End If
                                                        Case 126
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_left_9"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_9"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_9"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_9"
                                                            End If
                                                        Case 140
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_left_10"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_10"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_10"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_10"
                                                            End If
                                                        Case 154
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_left_11"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_11"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_11"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_11"
                                                            End If
                                                        Case 168
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_left_12"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_12"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_12"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_12"
                                                            End If
                                                        Case 182
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_left_13"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_13"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_13"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_13"
                                                            End If
                                                        Case 196
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_left_14"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_14"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_14"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_14"
                                                            End If
                                                        Case 210
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_left_15"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_15"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_15"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_15"
                                                            End If
                                                        Case 224
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_left_16"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_16"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_16"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_16"
                                                            End If
                                                        Case 238
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_left_17"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_17"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_17"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_17"
                                                            End If
                                                        Case 252
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_left_18"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_18"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_18"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_18"
                                                            End If
                                                        Case 266
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_left_19"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_19"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_19"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_19"
                                                            End If
                                                        Case Else
                                                            If Material_previous = "1" Then
                                                                Nume_pipeW = "mat1_match_left_20"
                                                            ElseIf Material_previous = "2" Then
                                                                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                                                                    Nume_pipeW = "line_pipe_match_left_20"
                                                                Else
                                                                    Nume_pipeW = "heavy_wall_match_left_20"
                                                                End If
                                                            Else
                                                                Nume_pipeW = "heavy_wall_match_left_20"
                                                            End If
                                                    End Select

                                                End If


                                                Dim Colectie_atr_name As New Specialized.StringCollection
                                                Dim Colectie_atr_value As New Specialized.StringCollection


                                                Colectie_atr_name.Add("BEGINSTA")
                                                Colectie_atr_name.Add("ENDSTA")
                                                Dim Pct_ins As New Point3d
                                                If RadioButton_right_to_left.Checked = True Then
                                                    Colectie_atr_value.Add(Get_chainage_from_double(ChainageT2, 1))
                                                    Colectie_atr_value.Add(Get_chainage_from_double(ChainageT1, 1))
                                                    Pct_ins = New Point3d(X, Y, Z)
                                                Else
                                                    Colectie_atr_value.Add(Get_chainage_from_double(ChainageT1, 1))
                                                    Colectie_atr_value.Add(Get_chainage_from_double(ChainageT2, 1))

                                                    Dim Valoare_x_shift As Double
                                                    If Extra_length = 0 Then
                                                        Valoare_x_shift = 28
                                                    Else
                                                        Valoare_x_shift = Extra_length + 14
                                                    End If
                                                    Pct_ins = New Point3d(X - Valoare_x_shift, Y, Z)




                                                End If

                                                Dim Val1, Val2 As Double
                                                Val1 = Round(ChainageT1, 1)
                                                Val2 = Round(ChainageT2, 1)

                                                Colectie_atr_name.Add("LENGTH")
                                                Colectie_atr_value.Add(Get_String_Rounded(Abs(Val2 - Val1), 1))
                                                Colectie_atr_name.Add("MAT")
                                                Colectie_atr_value.Add(Material_previous)

                                                If Not i = 0 Then
                                                    InsertBlock_with_multiple_atributes(Nume_pipeW & ".dwg", Nume_pipeW, Pct_ins, 1, BTrecord, "TEXT", Colectie_atr_name, Colectie_atr_value)
                                                End If






                                                If Is_ScrewAnchor = True Then
                                                    Dim nume_block_sa As String = "Screw_anchor_alignment1"

                                                    If RadioButton_right_to_left.Checked = True Then
                                                        Select Case Round(Extra_length_SA, 0)
                                                            Case 0
                                                                nume_block_sa = "Screw_anchor_alignment_match_left1"
                                                            Case 14
                                                                nume_block_sa = "Screw_anchor_alignment_match_left1"
                                                            Case 28
                                                                nume_block_sa = "Screw_anchor_alignment_match_left2"
                                                            Case 42
                                                                nume_block_sa = "Screw_anchor_alignment_match_left3"
                                                            Case 56
                                                                nume_block_sa = "Screw_anchor_alignment_match_left4"
                                                            Case 70
                                                                nume_block_sa = "Screw_anchor_alignment_match_left5"
                                                            Case 84
                                                                nume_block_sa = "Screw_anchor_alignment_match_left6"
                                                            Case 98
                                                                nume_block_sa = "Screw_anchor_alignment_match_left7"
                                                            Case 112
                                                                nume_block_sa = "Screw_anchor_alignment_match_left8"
                                                            Case 126
                                                                nume_block_sa = "Screw_anchor_alignment_match_left9"
                                                            Case 140
                                                                nume_block_sa = "Screw_anchor_alignment_match_left10"
                                                            Case 154
                                                                nume_block_sa = "Screw_anchor_alignment_match_left11"
                                                            Case 168
                                                                nume_block_sa = "Screw_anchor_alignment_match_left12"
                                                            Case 182
                                                                nume_block_sa = "Screw_anchor_alignment_match_left13"
                                                            Case 196
                                                                nume_block_sa = "Screw_anchor_alignment_match_left14"
                                                            Case 210
                                                                nume_block_sa = "Screw_anchor_alignment_match_left15"
                                                            Case 224
                                                                nume_block_sa = "Screw_anchor_alignment_match_left16"
                                                            Case 238
                                                                nume_block_sa = "Screw_anchor_alignment_match_left17"
                                                            Case 252
                                                                nume_block_sa = "Screw_anchor_alignment_match_left18"
                                                            Case 266
                                                                nume_block_sa = "Screw_anchor_alignment_match_left19"
                                                            Case Else
                                                                nume_block_sa = "Screw_anchor_alignment_match_left20"
                                                        End Select
                                                    End If

                                                    If RadioButton_left_to_right.Checked = True Then
                                                        Select Case Round(Extra_length_SA, 0)
                                                            Case 0
                                                                nume_block_sa = "Screw_anchor_alignment_match_right1"
                                                            Case 14
                                                                nume_block_sa = "Screw_anchor_alignment_match_right1"
                                                            Case 28
                                                                nume_block_sa = "Screw_anchor_alignment_match_right2"
                                                            Case 42
                                                                nume_block_sa = "Screw_anchor_alignment_match_right3"
                                                            Case 56
                                                                nume_block_sa = "Screw_anchor_alignment_match_right4"
                                                            Case 70
                                                                nume_block_sa = "Screw_anchor_alignment_match_right5"
                                                            Case 84
                                                                nume_block_sa = "Screw_anchor_alignment_match_right6"
                                                            Case 98
                                                                nume_block_sa = "Screw_anchor_alignment_match_right7"
                                                            Case 112
                                                                nume_block_sa = "Screw_anchor_alignment_match_right8"
                                                            Case 126
                                                                nume_block_sa = "Screw_anchor_alignment_match_right9"
                                                            Case 140
                                                                nume_block_sa = "Screw_anchor_alignment_match_right10"
                                                            Case 154
                                                                nume_block_sa = "Screw_anchor_alignment_match_right11"
                                                            Case 168
                                                                nume_block_sa = "Screw_anchor_alignment_match_right12"
                                                            Case 182
                                                                nume_block_sa = "Screw_anchor_alignment_match_right13"
                                                            Case 196
                                                                nume_block_sa = "Screw_anchor_alignment_match_right14"
                                                            Case 210
                                                                nume_block_sa = "Screw_anchor_alignment_match_right15"
                                                            Case 224
                                                                nume_block_sa = "Screw_anchor_alignment_match_right16"
                                                            Case 238
                                                                nume_block_sa = "Screw_anchor_alignment_match_right17"
                                                            Case 252
                                                                nume_block_sa = "Screw_anchor_alignment_match_right18"
                                                            Case 266
                                                                nume_block_sa = "Screw_anchor_alignment_match_right19"
                                                            Case Else
                                                                nume_block_sa = "Screw_anchor_alignment_match_right20"
                                                        End Select
                                                    End If


                                                    Dim Colectie_atr_nameSA As New Specialized.StringCollection
                                                    Dim Colectie_atr_valueSA As New Specialized.StringCollection


                                                    Colectie_atr_nameSA.Add("BEGINSTA")
                                                    Colectie_atr_nameSA.Add("ENDSTA")

                                                    If RadioButton_right_to_left.Checked = True Then
                                                        Colectie_atr_valueSA.Add(Get_chainage_from_double(ChainageT2, 1))
                                                        Colectie_atr_valueSA.Add(Get_chainage_from_double(ChainageSA1, 1))

                                                    Else
                                                        Colectie_atr_valueSA.Add(Get_chainage_from_double(ChainageSA1, 1))
                                                        Colectie_atr_valueSA.Add(Get_chainage_from_double(ChainageT2, 1))
                                                    End If

                                                    Val1 = Round(ChainageSA1, 1)
                                                    Val2 = Round(ChainageT2, 1)


                                                    If IsDBNull(Data_table_Crossing.Rows(i).Item("BUOYANCY")) = False Then
                                                        Dim Spacing_SA As Double = Round(Data_table_Crossing.Rows(i).Item("BUOYANCY"), 1)
                                                        If Spacing_SA > 0 Then
                                                            Dim Length_SA As Double = Round(Abs(Val2 - Val1), 1)
                                                            Dim NR_SA As Integer = CInt(1 + Length_SA / Spacing_SA)
                                                            Colectie_atr_nameSA.Add("NO_TYPE")
                                                            Colectie_atr_nameSA.Add("SPACING")
                                                            Colectie_atr_valueSA.Add(NR_SA & " SA")
                                                            Colectie_atr_valueSA.Add(Get_String_Rounded(Spacing_SA, 1) & " C/C")

                                                            Numar_SA_inserted = NR_SA
                                                        End If


                                                    End If

                                                    If Extra_length_SA = 0 Then
                                                        If RadioButton_right_to_left.Checked = True Then
                                                            X = X - 14
                                                        Else
                                                            X = X + 14
                                                        End If
                                                    End If


                                                    InsertBlock_with_multiple_atributes(nume_block_sa & ".dwg", nume_block_sa, New Point3d(X, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_nameSA, Colectie_atr_valueSA)

                                                    Is_matchline_screw_anchor = True
                                                    Extra_length_SA = 0
                                                    ChainageSA1 = ChainageT2
                                                End If 'If Is_ScrewAnchor = True


                                                If Is_SandBag = True Then
                                                    Dim nume_block_SB As String = "Sand_Bag_alignment1"

                                                    If RadioButton_right_to_left.Checked = True Then
                                                        Select Case Round(Extra_length_SB, 0)
                                                            Case 0
                                                                nume_block_SB = "Sand_Bag_alignment_match_left1"
                                                            Case 14
                                                                nume_block_SB = "Sand_Bag_alignment_match_left1"
                                                            Case 28
                                                                nume_block_SB = "Sand_Bag_alignment_match_left2"
                                                            Case 42
                                                                nume_block_SB = "Sand_Bag_alignment_match_left3"
                                                            Case 56
                                                                nume_block_SB = "Sand_Bag_alignment_match_left4"
                                                            Case 70
                                                                nume_block_SB = "Sand_Bag_alignment_match_left5"
                                                            Case 84
                                                                nume_block_SB = "Sand_Bag_alignment_match_left6"
                                                            Case 98
                                                                nume_block_SB = "Sand_Bag_alignment_match_left7"
                                                            Case 112
                                                                nume_block_SB = "Sand_Bag_alignment_match_left8"
                                                            Case 126
                                                                nume_block_SB = "Sand_Bag_alignment_match_left9"
                                                            Case 140
                                                                nume_block_SB = "Sand_Bag_alignment_match_left10"
                                                            Case 154
                                                                nume_block_SB = "Sand_Bag_alignment_match_left11"
                                                            Case 168
                                                                nume_block_SB = "Sand_Bag_alignment_match_left12"
                                                            Case 182
                                                                nume_block_SB = "Sand_Bag_alignment_match_left13"
                                                            Case 196
                                                                nume_block_SB = "Sand_Bag_alignment_match_left14"
                                                            Case 210
                                                                nume_block_SB = "Sand_Bag_alignment_match_left15"
                                                            Case 224
                                                                nume_block_SB = "Sand_Bag_alignment_match_left16"
                                                            Case 238
                                                                nume_block_SB = "Sand_Bag_alignment_match_left17"
                                                            Case 252
                                                                nume_block_SB = "Sand_Bag_alignment_match_left18"
                                                            Case 266
                                                                nume_block_SB = "Sand_Bag_alignment_match_left19"
                                                            Case Else
                                                                nume_block_SB = "Sand_Bag_alignment_match_left20"
                                                        End Select
                                                    End If

                                                    If RadioButton_left_to_right.Checked = True Then
                                                        Select Case Round(Extra_length_SB, 0)
                                                            Case 0
                                                                nume_block_SB = "Sand_Bag_alignment_match_right1"
                                                            Case 14
                                                                nume_block_SB = "Sand_Bag_alignment_match_right1"
                                                            Case 28
                                                                nume_block_SB = "Sand_Bag_alignment_match_right2"
                                                            Case 42
                                                                nume_block_SB = "Sand_Bag_alignment_match_right3"
                                                            Case 56
                                                                nume_block_SB = "Sand_Bag_alignment_match_right4"
                                                            Case 70
                                                                nume_block_SB = "Sand_Bag_alignment_match_right5"
                                                            Case 84
                                                                nume_block_SB = "Sand_Bag_alignment_match_right6"
                                                            Case 98
                                                                nume_block_SB = "Sand_Bag_alignment_match_right7"
                                                            Case 112
                                                                nume_block_SB = "Sand_Bag_alignment_match_right8"
                                                            Case 126
                                                                nume_block_SB = "Sand_Bag_alignment_match_right9"
                                                            Case 140
                                                                nume_block_SB = "Sand_Bag_alignment_match_right10"
                                                            Case 154
                                                                nume_block_SB = "Sand_Bag_alignment_match_right11"
                                                            Case 168
                                                                nume_block_SB = "Sand_Bag_alignment_match_right12"
                                                            Case 182
                                                                nume_block_SB = "Sand_Bag_alignment_match_right13"
                                                            Case 196
                                                                nume_block_SB = "Sand_Bag_alignment_match_right14"
                                                            Case 210
                                                                nume_block_SB = "Sand_Bag_alignment_match_right15"
                                                            Case 224
                                                                nume_block_SB = "Sand_Bag_alignment_match_right16"
                                                            Case 238
                                                                nume_block_SB = "Sand_Bag_alignment_match_right17"
                                                            Case 252
                                                                nume_block_SB = "Sand_Bag_alignment_match_right18"
                                                            Case 266
                                                                nume_block_SB = "Sand_Bag_alignment_match_right19"
                                                            Case Else
                                                                nume_block_SB = "Sand_Bag_alignment_match_right20"
                                                        End Select
                                                    End If

                                                    Dim Colectie_atr_nameSB As New Specialized.StringCollection
                                                    Dim Colectie_atr_valueSB As New Specialized.StringCollection

                                                    Colectie_atr_nameSB.Add("BEGINSTA")
                                                    Colectie_atr_nameSB.Add("ENDSTA")

                                                    If RadioButton_right_to_left.Checked = True Then
                                                        Colectie_atr_valueSB.Add(Get_chainage_from_double(ChainageT2, 1))
                                                        Colectie_atr_valueSB.Add(Get_chainage_from_double(ChainageSB1, 1))

                                                    Else
                                                        Colectie_atr_valueSB.Add(Get_chainage_from_double(ChainageSB1, 1))
                                                        Colectie_atr_valueSB.Add(Get_chainage_from_double(ChainageT2, 1))
                                                    End If

                                                    Val1 = Round(ChainageSB1, 1)
                                                    Val2 = Round(ChainageT2, 1)


                                                    If IsDBNull(Data_table_Crossing.Rows(i).Item("BUOYANCY")) = False Then
                                                        Dim Spacing_SB As Double = Round(Data_table_Crossing.Rows(i).Item("BUOYANCY"), 1)
                                                        If Spacing_SB > 0 Then
                                                            Dim Length_SB As Double = Round(Abs(Val2 - Val1), 1)
                                                            Dim NR_SB As Integer = CInt(1 + Length_SB / Spacing_SB)
                                                            Colectie_atr_nameSB.Add("NO_TYPE")
                                                            Colectie_atr_nameSB.Add("SPACING")
                                                            Colectie_atr_valueSB.Add(NR_SB & " SBW")
                                                            Colectie_atr_valueSB.Add(Get_String_Rounded(Spacing_SB, 1) & " C/C")

                                                            Numar_SB_inserted = NR_SB
                                                        End If


                                                    End If

                                                    If Extra_length_SB = 0 Then
                                                        If RadioButton_right_to_left.Checked = True Then
                                                            X = X - 14
                                                        Else
                                                            X = X + 14
                                                        End If
                                                    End If

                                                    InsertBlock_with_multiple_atributes(nume_block_SB & ".dwg", nume_block_SB, New Point3d(X, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_nameSB, Colectie_atr_valueSB)

                                                    Is_matchline_sand_bag = True
                                                    Extra_length_SB = 0
                                                    ChainageSB1 = ChainageT2
                                                End If 'If Is_SAND_BAG = True


                                                If Is_Riverweight = True Then
                                                    Dim nume_block_RW As String = "River_weight_alignment1"

                                                    If RadioButton_right_to_left.Checked = True Then
                                                        Select Case Round(Extra_length_RW, 0)
                                                            Case 0
                                                                nume_block_RW = "River_weight_alignment_match_left1"
                                                            Case 14
                                                                nume_block_RW = "River_weight_alignment_match_left1"
                                                            Case 28
                                                                nume_block_RW = "River_weight_alignment_match_left2"
                                                            Case 42
                                                                nume_block_RW = "River_weight_alignment_match_left3"
                                                            Case 56
                                                                nume_block_RW = "River_weight_alignment_match_left4"
                                                            Case 70
                                                                nume_block_RW = "River_weight_alignment_match_left5"
                                                            Case 84
                                                                nume_block_RW = "River_weight_alignment_match_left6"
                                                            Case 98
                                                                nume_block_RW = "River_weight_alignment_match_left7"
                                                            Case 112
                                                                nume_block_RW = "River_weight_alignment_match_left8"
                                                            Case 126
                                                                nume_block_RW = "River_weight_alignment_match_left9"
                                                            Case 140
                                                                nume_block_RW = "River_weight_alignment_match_left10"
                                                            Case 154
                                                                nume_block_RW = "River_weight_alignment_match_left11"
                                                            Case 168
                                                                nume_block_RW = "River_weight_alignment_match_left12"
                                                            Case 182
                                                                nume_block_RW = "River_weight_alignment_match_left13"
                                                            Case 196
                                                                nume_block_RW = "River_weight_alignment_match_left14"
                                                            Case 210
                                                                nume_block_RW = "River_weight_alignment_match_left15"
                                                            Case 224
                                                                nume_block_RW = "River_weight_alignment_match_left16"
                                                            Case 238
                                                                nume_block_RW = "River_weight_alignment_match_left17"
                                                            Case 252
                                                                nume_block_RW = "River_weight_alignment_match_left18"
                                                            Case 266
                                                                nume_block_RW = "River_weight_alignment_match_left19"
                                                            Case Else
                                                                nume_block_RW = "River_weight_alignment_match_left20"
                                                        End Select
                                                    End If

                                                    If RadioButton_left_to_right.Checked = True Then
                                                        Select Case Round(Extra_length_RW, 0)
                                                            Case 0
                                                                nume_block_RW = "River_weight_alignment_match_right1"
                                                            Case 14
                                                                nume_block_RW = "River_weight_alignment_match_right1"
                                                            Case 28
                                                                nume_block_RW = "River_weight_alignment_match_right2"
                                                            Case 42
                                                                nume_block_RW = "River_weight_alignment_match_right3"
                                                            Case 56
                                                                nume_block_RW = "River_weight_alignment_match_right4"
                                                            Case 70
                                                                nume_block_RW = "River_weight_alignment_match_right5"
                                                            Case 84
                                                                nume_block_RW = "River_weight_alignment_match_right6"
                                                            Case 98
                                                                nume_block_RW = "River_weight_alignment_match_right7"
                                                            Case 112
                                                                nume_block_RW = "River_weight_alignment_match_right8"
                                                            Case 126
                                                                nume_block_RW = "River_weight_alignment_match_right9"
                                                            Case 140
                                                                nume_block_RW = "River_weight_alignment_match_right10"
                                                            Case 154
                                                                nume_block_RW = "River_weight_alignment_match_right11"
                                                            Case 168
                                                                nume_block_RW = "River_weight_alignment_match_right12"
                                                            Case 182
                                                                nume_block_RW = "River_weight_alignment_match_right13"
                                                            Case 196
                                                                nume_block_RW = "River_weight_alignment_match_right14"
                                                            Case 210
                                                                nume_block_RW = "River_weight_alignment_match_right15"
                                                            Case 224
                                                                nume_block_RW = "River_weight_alignment_match_right16"
                                                            Case 238
                                                                nume_block_RW = "River_weight_alignment_match_right17"
                                                            Case 252
                                                                nume_block_RW = "River_weight_alignment_match_right18"
                                                            Case 266
                                                                nume_block_RW = "River_weight_alignment_match_right19"
                                                            Case Else
                                                                nume_block_RW = "River_weight_alignment_match_right20"
                                                        End Select
                                                    End If


                                                    Dim Colectie_atr_nameRW As New Specialized.StringCollection
                                                    Dim Colectie_atr_valueRW As New Specialized.StringCollection


                                                    Colectie_atr_nameRW.Add("BEGINSTA")
                                                    Colectie_atr_nameRW.Add("ENDSTA")

                                                    If RadioButton_right_to_left.Checked = True Then
                                                        Colectie_atr_valueRW.Add(Get_chainage_from_double(ChainageT2, 1))
                                                        Colectie_atr_valueRW.Add(Get_chainage_from_double(ChainageRW1, 1))

                                                    Else
                                                        Colectie_atr_valueRW.Add(Get_chainage_from_double(ChainageRW1, 1))
                                                        Colectie_atr_valueRW.Add(Get_chainage_from_double(ChainageT2, 1))
                                                    End If

                                                    Val1 = Round(ChainageSA1, 1)
                                                    Val2 = Round(ChainageT2, 1)


                                                    If IsDBNull(Data_table_Crossing.Rows(i).Item("BUOYANCY")) = False Then
                                                        Dim Spacing_RW As Double = Round(Data_table_Crossing.Rows(i).Item("BUOYANCY"), 1)
                                                        If Spacing_RW > 0 Then
                                                            Dim Length_RW As Double = Round(Abs(Val2 - Val1), 1)
                                                            Dim NR_RW As Integer = CInt(1 + Length_RW / Spacing_RW)
                                                            Colectie_atr_nameRW.Add("NO_TYPE")
                                                            Colectie_atr_nameRW.Add("SPACING")
                                                            Colectie_atr_valueRW.Add(NR_RW & " RW")
                                                            Colectie_atr_valueRW.Add(Get_String_Rounded(Spacing_RW, 1) & " C/C")

                                                            Numar_RW_inserted = NR_RW
                                                        End If


                                                    End If

                                                    If Extra_length_RW = 0 Then
                                                        If RadioButton_right_to_left.Checked = True Then
                                                            X = X - 14
                                                        Else
                                                            X = X + 14
                                                        End If
                                                    End If


                                                    InsertBlock_with_multiple_atributes(nume_block_RW & ".dwg", nume_block_RW, New Point3d(X, Y, Z), 1, BTrecord, "TEXT", Colectie_atr_nameRW, Colectie_atr_valueRW)

                                                    Is_matchline_river_weight = True
                                                    Extra_length_RW = 0
                                                    ChainageRW1 = ChainageT2
                                                End If 'If Is_ScrewAnchor = True




                                            End If '  If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False

                                            Is_matchline = True
                                            Extra_length = 0

                                            Dim New_text As New DBText
                                            New_text.Layer = "0"
                                            If RadioButton_right_to_left.Checked = True Then

                                                New_text.Justify = AttachmentPoint.MiddleRight
                                                New_text.AlignmentPoint = New Point3d(X - 20, Y, 0)
                                            Else

                                                New_text.Justify = AttachmentPoint.MiddleLeft
                                                New_text.AlignmentPoint = New Point3d(X + 20, Y, 0)
                                            End If

                                            New_text.TextString = CStr(Nr_pagina)
                                            New_text.Height = 25

                                            BTrecord.AppendEntity(New_text)
                                            Trans1.AddNewlyCreatedDBObject(New_text, True)

                                            X = Point1.Value.X
                                            If Not i = 0 Then Y = Y - 130
                                    End Select


                                    If i = Data_table_Crossing.Rows.Count - 1 Then
                                        If Not Block_name = "MATCHLINE" Then
                                            Dim New_text As New DBText
                                            New_text.Layer = "0"
                                            If RadioButton_right_to_left.Checked = True Then
                                                New_text.Justify = AttachmentPoint.MiddleRight
                                                New_text.AlignmentPoint = New Point3d(X - 20, Y, 0)

                                            Else
                                                New_text.Justify = AttachmentPoint.MiddleLeft
                                                New_text.AlignmentPoint = New Point3d(X + 20, Y, 0)

                                            End If

                                            New_text.TextString = CStr(Nr_pagina)
                                            New_text.Height = 25

                                            BTrecord.AppendEntity(New_text)
                                            Trans1.AddNewlyCreatedDBObject(New_text, True)
                                        End If
                                    End If

                                End If ' If IsDBNull(Data_table_Crossing.Rows(i).Item("BLOCKNAME")) = False

end1:
                            Next '  For i = 0 To Data_table_Crossing.Rows.Count - 1




                        End If '   If Data_table_Crossing.Rows.Count > 0

                        Trans1.Commit()
                    End Using
                End Using

                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
            Catch ex As Exception
                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False

        End If


    End Sub

    Private Sub Button_create_layout_Click(sender As Object, e As EventArgs) Handles Button_create_layout.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            If IsNumeric(TextBox_Layouts_number.Text) = False Then
                With TextBox_Layouts_number
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify the layouts number:")
                Freeze_operations = False
                Exit Sub
            End If

            If IsNumeric(TextBox_VIEWPORT_GAP.Text) = False Then
                With TextBox_VIEWPORT_GAP
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify the gap distance:")
                Freeze_operations = False
                Exit Sub
            End If
            If CDbl(TextBox_VIEWPORT_GAP.Text) < 0 Then
                With TextBox_VIEWPORT_GAP
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify the gap distance:")
                Freeze_operations = False
                Exit Sub
            End If
            If IsNumeric(TextBox_viewport_height.Text) = False Then
                With TextBox_viewport_height
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify the viewport heigth:")
                Freeze_operations = False
                Exit Sub
            End If
            If CDbl(TextBox_viewport_height.Text) < 0 Then
                With TextBox_viewport_height
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify the viewport heigth:")
                Freeze_operations = False
                Exit Sub
            End If
            If IsNumeric(TextBox_viewport_length.Text) = False Then
                With TextBox_viewport_length
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify the viewport width:")
                Freeze_operations = False
                Exit Sub
            End If
            If CDbl(TextBox_viewport_length.Text) < 0 Then
                With TextBox_viewport_length
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify the viewport width:")
                Freeze_operations = False
                Exit Sub
            End If
            If IsNumeric(TextBox_viewport_scale.Text) = False Then
                With TextBox_viewport_scale
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify the viewport scale:")
                Freeze_operations = False
                Exit Sub
            End If
            If CDbl(TextBox_viewport_scale.Text) < 0 Then
                With TextBox_viewport_scale
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify the viewport scale:")
                Freeze_operations = False
                Exit Sub
            End If



            If IsNumeric(TextBox_Layouts_number.Text) = False Then
                With TextBox_Layouts_number
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify layouts number:")
                Freeze_operations = False
                Exit Sub
            End If
            If CInt(TextBox_Layouts_number.Text) < 1 Then
                With TextBox_Layouts_number
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify the end number:")
                Freeze_operations = False
                Exit Sub
            End If

            Dim End1 As Integer = CInt(TextBox_Layouts_number.Text)




            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
            ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()


            Try

                Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                        Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        Dim BTrecordMS As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BTrecordMS = Trans1.GetObject(BlockTable_data1(BlockTableRecord.ModelSpace), OpenMode.ForRead)
                        Dim BTrecordPS As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord


                        Dim Rezultat_point As Autodesk.AutoCAD.EditorInput.PromptPointResult
                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                        Dim curent_ucs_matrix As Matrix3d = ThisDrawing.Editor.CurrentUserCoordinateSystem

                        Rezultat_point = ThisDrawing.Editor.GetPoint(vbLf & "Specify the start point of first line")


                        If Rezultat_point.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                            Freeze_operations = False
                            Exit Sub
                        End If

                        Creaza_layer("VPORT", 7, "Viewports", False)



                        Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current

                        Dim exista_layoutul As Boolean = False
                        Dim Layoutdict As DBDictionary

                        Layoutdict = Trans1.GetObject(ThisDrawing.Database.LayoutDictionaryId, OpenMode.ForRead)

                        Dim nr_layouts As Integer = Layoutdict.Count


                        For Each entry As DBDictionaryEntry In Layoutdict
                            If entry.Key = "TEMPLATE" Then
                                exista_layoutul = True
                                Exit For
                            End If
                        Next
                        If exista_layoutul = False Then
                            MsgBox("TEMPLATE layout missing")
                            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                            Freeze_operations = False
                            Exit Sub
                        End If
                        Dim Index_Layout As Integer = nr_layouts
                        Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId("TEMPLATE"), OpenMode.ForRead)
                        Dim Nume_nou As String = TextBox_PAGE_START_NO_Layouts.Text
                        Dim Nume_nou_numar As Integer
                        If IsNumeric(Nume_nou) = False Then
                            Nume_nou_numar = 1
                        Else
                            Nume_nou_numar = CInt(TextBox_PAGE_START_NO_Layouts.Text)
                        End If
                        If exista_layoutul = True Then

                            For i = 1 To End1

                                Nume_nou = (Nume_nou_numar + i - 1).ToString

                                Dim exista_layoutul_nou As Boolean = True
                                Dim Increment As Integer = 1

                                Do Until exista_layoutul_nou = False
                                    For Each entry As DBDictionaryEntry In Layoutdict
                                        If entry.Key = Nume_nou Then
                                            exista_layoutul_nou = True
                                            Exit For
                                        Else
                                            exista_layoutul_nou = False
                                        End If
                                    Next
                                    If exista_layoutul_nou = True Then
                                        Nume_nou = Nr_pagina.ToString & "_" & Increment.ToString
                                        Increment = Increment + 1
                                    End If
                                Loop

                                LayoutManager1.CloneLayout(Layout1.LayoutName, Nume_nou, Index_Layout)
                                Dim Tilemode1 As Integer = Application.GetSystemVariable("TILEMODE")
                                Dim CVport1 As Integer = Application.GetSystemVariable("CVPORT")

                                If Tilemode1 = 0 Then
                                    If CVport1 = 2 Then
                                        Editor1.SwitchToPaperSpace()
                                    End If
                                Else
                                    Application.SetSystemVariable("TILEMODE", 0)
                                End If

                                Dim Layout2 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(Nume_nou), OpenMode.ForRead)



                                LayoutManager1.CurrentLayout = Nume_nou


                                BTrecordPS = Trans1.GetObject(BlockTable_data1(BlockTableRecord.PaperSpace), OpenMode.ForWrite)

                                Dim Dist_between_bands As Double = 130
                                If IsNumeric(TextBox_Distance_between_bands.Text) = True Then
                                    Dist_between_bands = CDbl(TextBox_Distance_between_bands.Text)
                                End If

                                Dim Point_target_MS1 As Point3d
                                If RadioButton_right_to_left.Checked = True Then
                                    Point_target_MS1 = New Point3d(Rezultat_point.Value.X - 270, (Rezultat_point.Value.Y + 5) - Dist_between_bands * (i - 1), 0)
                                Else
                                    Point_target_MS1 = New Point3d(Rezultat_point.Value.X + 270, (Rezultat_point.Value.Y + 5) - Dist_between_bands * (i - 1), 0)
                                End If

                                Dim Viewport1 As New Viewport
                                Viewport1.SetDatabaseDefaults()
                                Viewport1.CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(CDbl(TextBox_VIEWPORT_X.Text) + CDbl(TextBox_viewport_length.Text) / 2, _
                                                                                              CDbl(TextBox_VIEWPORT_Y.Text) + 2.5 * CDbl(TextBox_viewport_height.Text) + 2 * CDbl(TextBox_VIEWPORT_GAP.Text), _
                                                                                              0) ' asta e pozitia viewport in paper space
                                Viewport1.Height = CDbl(TextBox_viewport_height.Text)
                                Viewport1.Width = CDbl(TextBox_viewport_length.Text)
                                Viewport1.Layer = "VPORT"

                                Viewport1.ViewDirection = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis
                                Viewport1.ViewTarget = Point_target_MS1 ' asta e pozitia viewport in MODEL space
                                Viewport1.ViewCenter = Autodesk.AutoCAD.Geometry.Point2d.Origin
                                Viewport1.TwistAngle = 0 ' asta e PT TWIST

                                BTrecordPS.AppendEntity(Viewport1)
                                Trans1.AddNewlyCreatedDBObject(Viewport1, True)

                                Viewport1.On = True
                                Viewport1.CustomScale = CDbl(TextBox_viewport_scale.Text)
                                Viewport1.Locked = True

                                Dim Point_target_MS2 As Point3d
                                If RadioButton_right_to_left.Checked = True Then
                                    Point_target_MS2 = New Point3d(Rezultat_point.Value.X - 660, (Rezultat_point.Value.Y + 5) - Dist_between_bands * (i - 1), 0)
                                Else
                                    Point_target_MS2 = New Point3d(Rezultat_point.Value.X + 660, (Rezultat_point.Value.Y + 5) - Dist_between_bands * (i - 1), 0)
                                End If

                                Dim Viewport2 As New Viewport
                                Viewport2.SetDatabaseDefaults()
                                Viewport2.CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(CDbl(TextBox_VIEWPORT_X.Text) + CDbl(TextBox_viewport_length.Text) / 2, _
                                                                                              CDbl(TextBox_VIEWPORT_Y.Text) + 1.5 * CDbl(TextBox_viewport_height.Text) + CDbl(TextBox_VIEWPORT_GAP.Text), _
                                                                                              0) ' asta e pozitia viewport in paper space
                                Viewport2.Height = CDbl(TextBox_viewport_height.Text)
                                Viewport2.Width = CDbl(TextBox_viewport_length.Text)
                                Viewport2.Layer = "VPORT"

                                Viewport2.ViewDirection = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis
                                Viewport2.ViewTarget = Point_target_MS2 ' asta e pozitia viewport in MODEL space
                                Viewport2.ViewCenter = Autodesk.AutoCAD.Geometry.Point2d.Origin
                                Viewport2.TwistAngle = 0 ' asta e PT TWIST

                                BTrecordPS.AppendEntity(Viewport2)
                                Trans1.AddNewlyCreatedDBObject(Viewport2, True)

                                Viewport2.On = True
                                Viewport2.CustomScale = CDbl(TextBox_viewport_scale.Text)
                                Viewport2.Locked = True

                                Dim Point_target_MS3 As Point3d
                                If RadioButton_right_to_left.Checked = True Then
                                    Point_target_MS3 = New Point3d(Rezultat_point.Value.X - 1050, (Rezultat_point.Value.Y + 5) - Dist_between_bands * (i - 1), 0)
                                Else
                                    Point_target_MS3 = New Point3d(Rezultat_point.Value.X + 1050, (Rezultat_point.Value.Y + 5) - Dist_between_bands * (i - 1), 0)
                                End If

                                Dim Viewport3 As New Viewport
                                Viewport3.SetDatabaseDefaults()
                                Viewport3.CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(CDbl(TextBox_VIEWPORT_X.Text) + CDbl(TextBox_viewport_length.Text) / 2, _
                                                                                              CDbl(TextBox_VIEWPORT_Y.Text) + 0.5 * CDbl(TextBox_viewport_height.Text), _
                                                                                              0) ' asta e pozitia viewport in paper space
                                Viewport3.Height = CDbl(TextBox_viewport_height.Text)
                                Viewport3.Width = CDbl(TextBox_viewport_length.Text)
                                Viewport3.Layer = "VPORT"

                                Viewport3.ViewDirection = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis
                                Viewport3.ViewTarget = Point_target_MS3 ' asta e pozitia viewport in MODEL space
                                Viewport3.ViewCenter = Autodesk.AutoCAD.Geometry.Point2d.Origin
                                Viewport3.TwistAngle = 0 ' asta e PT TWIST

                                BTrecordPS.AppendEntity(Viewport3)
                                Trans1.AddNewlyCreatedDBObject(Viewport3, True)

                                Viewport3.On = True
                                Viewport3.CustomScale = CDbl(TextBox_viewport_scale.Text)
                                Viewport3.Locked = True
                                Index_Layout = Index_Layout + 1
                            Next
                        End If

                        Trans1.Commit()

                    End Using
                End Using

            Catch ex As Exception
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
            Freeze_operations = False
        End If

    End Sub

    Private Sub Button_CHECK_AND_HIGHLIGHT_ELBOWS_Click(sender As Object, e As EventArgs) Handles Button_CHECK_AND_HIGHLIGHT_ELBOWS.Click
        If Freeze_operations = False Then
            Freeze_operations = True

            If IsNumeric(ComboBox_nps.Text) = False Then
                MsgBox("Specify The Nominal Pipe Size!")
                Freeze_operations = False
                Exit Sub
            End If

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Try
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()




                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select engineering bands:"
                Object_Prompt.SingleOnly = False
                Rezultat1 = Editor1.GetSelection(Object_Prompt)

                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                    If IsNothing(Rezultat1) = False Then

                        Using Lock_dwg As DocumentLock = ThisDrawing.LockDocument


                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                                BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                                For i = 0 To Rezultat1.Value.Count - 1
                                    Dim Ent1 As Entity = Rezultat1.Value(i).ObjectId.GetObject(OpenMode.ForRead)
                                    If TypeOf Ent1 Is BlockReference Then
                                        Dim Block1 As BlockReference = Ent1
                                        If Block1.Name.ToUpper.Contains("ELBOW") = True Then
                                            Dim BEGINSTA As Double = -1
                                            Dim ENDSTA As Double = -1
                                            Dim Length_of_elbow As Double = -1
                                            Dim Elbow_angle As Double = -1

                                            Dim BeginSTA_string As String = ""
                                            Dim ENDSTA_string As String = ""
                                            Dim DESC_string As String = ""
                                            Dim id_no_string As String = ""
                                            Dim STA_string As String = ""


                                            If Block1.AttributeCollection.Count > 0 Then
                                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                                For Each id In attColl
                                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                                    If attref.Tag.ToUpper = "BEGINSTA" Then
                                                        Dim Continut As String = attref.TextString
                                                        If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                            BEGINSTA = CDbl(Replace(Continut, "+", ""))
                                                        End If
                                                        BeginSTA_string = Continut
                                                    End If
                                                    If attref.Tag.ToUpper = "ENDSTA" Then
                                                        Dim Continut As String = attref.TextString
                                                        If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                            ENDSTA = CDbl(Replace(Continut, "+", ""))
                                                        End If
                                                        ENDSTA_string = Continut
                                                    End If
                                                    If attref.Tag.ToUpper = "LENGTH" Then
                                                        Dim Continut As String = attref.TextString
                                                        If IsNumeric(Continut) = True Then
                                                            Length_of_elbow = CDbl(Continut)
                                                        End If
                                                    End If
                                                    If attref.Tag.ToUpper = "DESC" Then
                                                        Dim Continut As String = attref.TextString
                                                        Elbow_angle = extrage_numar_din_text_de_la_inceputul_textului(Continut)
                                                        DESC_string = Continut
                                                    End If
                                                    If attref.Tag.ToUpper = "ID_NO" Then
                                                        Dim Continut As String = attref.TextString
                                                        id_no_string = Continut
                                                    End If
                                                    If attref.Tag.ToUpper = "STA" Then
                                                        Dim Continut As String = attref.TextString
                                                        STA_string = Continut
                                                    End If
                                                Next

                                                If (Not Elbow_angle = -1 And Not Length_of_elbow = -1) Or Not Abs(BEGINSTA - ENDSTA) = Length_of_elbow Then


                                                    Dim Diam As Double = 2 * Get_from_NPS_radius_for_pipes_from_inches_to_milimeters(CDbl(ComboBox_nps.Text)) / 1000
                                                    Dim True_length As Double
                                                    True_length = Round(2 * (3 * Diam * Tan((Elbow_angle * PI / 180) / 2) + 1), 3)

                                                    If Abs(Length_of_elbow - True_length) > 0.1 Then
                                                        Dim Colectie_atr_name As New Specialized.StringCollection
                                                        Dim Colectie_atr_value As New Specialized.StringCollection

                                                        Colectie_atr_name.Add("TRUE_LENGTH")
                                                        Colectie_atr_value.Add(Round(True_length, 3))



                                                        Colectie_atr_name.Add("BEGINSTA")
                                                        Colectie_atr_value.Add(BeginSTA_string)
                                                        Colectie_atr_name.Add("ENDSTA")
                                                        Colectie_atr_value.Add(ENDSTA_string)

                                                        Colectie_atr_name.Add("STA")
                                                        Colectie_atr_value.Add(STA_string)
                                                        Colectie_atr_name.Add("LENGTH")
                                                        Colectie_atr_value.Add(Length_of_elbow)


                                                        Colectie_atr_name.Add("DESC")
                                                        Colectie_atr_value.Add(DESC_string)
                                                        Colectie_atr_name.Add("ID_NO")
                                                        Colectie_atr_value.Add(id_no_string)

                                                        InsertBlock_with_multiple_atributes(Block1.Name & "_RED.dwg", Block1.Name & "_RED", Block1.Position, 1, BTrecord, Block1.Layer, Colectie_atr_name, Colectie_atr_value)
                                                        Block1.UpgradeOpen()
                                                        Block1.Erase()

                                                    End If



                                                End If


                                            End If
                                        End If
                                    End If
                                Next






                                Trans1.Commit()





                            End Using
                        End Using

                    End If
                End If


                Dim Empty_array() As ObjectId
                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
            Catch ex As Exception
                Editor1.WriteMessage(vbLf & "Command:")
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If

    End Sub

    Private Sub Button_Display_matchlines_Click(sender As Object, e As EventArgs) Handles Button_Display_matchlines.Click

        If Freeze_operations = False Then
            Freeze_operations = True

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Try
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()


                Dim Empty_array() As ObjectId
                Editor1.SetImpliedSelection(Empty_array)

                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select engineering blocks:"
                Object_Prompt.SingleOnly = False
                Rezultat1 = Editor1.GetSelection(Object_Prompt)

                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                    If IsNothing(Rezultat1) = False Then

                        Using Lock_dwg As DocumentLock = ThisDrawing.LockDocument


                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                                BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                                Dim Max_sta As Double = -1
                                Dim Min_sta As Double = -1

                                For i = 0 To Rezultat1.Value.Count - 1
                                    Dim Ent1 As Entity = Rezultat1.Value(i).ObjectId.GetObject(OpenMode.ForRead)
                                    If TypeOf Ent1 Is BlockReference Then
                                        Dim Block1 As BlockReference = Ent1

                                        If Block1.AttributeCollection.Count > 0 Then
                                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                            Dim BEGINSTA As Double = -1
                                            Dim ENDSTA As Double = -1
                                            For Each id In attColl
                                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                                If attref.Tag.ToUpper = "BEGINSTA" Then

                                                    Dim Continut As String = attref.TextString
                                                    If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                        BEGINSTA = CDbl(Replace(Continut, "+", ""))


                                                    End If

                                                End If
                                                If attref.Tag.ToUpper = "ENDSTA" Then

                                                    Dim Continut As String = attref.TextString
                                                    If IsNumeric(Replace(Continut, "+", "")) = True Then


                                                        ENDSTA = CDbl(Replace(Continut, "+", ""))
                                                    End If

                                                End If
                                            Next

                                            If Not BEGINSTA = -1 And Not ENDSTA = -1 Then
                                                If BEGINSTA > ENDSTA Then
                                                    Dim T As Double = BEGINSTA
                                                    BEGINSTA = ENDSTA
                                                    ENDSTA = T
                                                End If

                                            End If

                                            If Not ENDSTA = -1 Then
                                                If Max_sta = -1 Then
                                                    Max_sta = ENDSTA
                                                End If
                                            End If

                                            If Not BEGINSTA = -1 Then
                                                If Min_sta = -1 Then
                                                    Min_sta = BEGINSTA
                                                End If
                                            End If

                                            If Not Min_sta = -1 And Not BEGINSTA = -1 Then
                                                If BEGINSTA < Min_sta Then
                                                    Min_sta = BEGINSTA
                                                End If
                                            End If

                                            If Not Max_sta = -1 And Not ENDSTA = -1 Then
                                                If ENDSTA > Max_sta Then
                                                    Max_sta = ENDSTA
                                                End If
                                            End If




                                        End If

                                    End If
                                Next

                                If Not Max_sta = -1 And Not Min_sta = -1 Then
                                    Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult

                                    Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify position:")
                                    PP1.AllowNone = True
                                    Point1 = Editor1.GetPoint(PP1)

                                    If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                        Editor1.WriteMessage(vbLf & "Command:")
                                        Freeze_operations = False
                                        Exit Sub
                                    End If

                                    Dim Mtext1 As New MText
                                    Mtext1.Layer = 0
                                    Mtext1.TextHeight = 5
                                    Mtext1.Attachment = AttachmentPoint.MiddleCenter
                                    Mtext1.Contents = Get_chainage_from_double(Min_sta, 1) & vbCrLf & Get_chainage_from_double(Max_sta, 1) & vbCrLf _
                                        & "{\H0.1x;________________________________________________________}" & vbCrLf & _
                                        Get_String_Rounded(Round(Max_sta, 1) - Round(Min_sta, 1), 1)

                                    Mtext1.Location = Point1.Value
                                    BTrecord.AppendEntity(Mtext1)
                                    Trans1.AddNewlyCreatedDBObject(Mtext1, True)

                                    Trans1.Commit()
                                End If

                            End Using
                        End Using

                    End If
                End If


                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
            Catch ex As Exception
                Editor1.WriteMessage(vbLf & "Command:")
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If

    End Sub

    Private Sub Button_warning_signs_to_excel_Click(sender As Object, e As EventArgs) Handles Button_warning_signs_to_excel.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Dim Empty_array() As ObjectId
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Editor1.SetImpliedSelection(Empty_array)
            Try
                Try
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt.MessageForAdding = vbLf & "Select engineering bands elements:"
                    Object_Prompt.SingleOnly = False
                    Rezultat1 = Editor1.GetSelection(Object_Prompt)
                    If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        If IsNothing(Rezultat1) = False Then

                            Using Lock_dwg As DocumentLock = ThisDrawing.LockDocument


                                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                                    Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                    BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                                    Dim DataTable_WS As New System.Data.DataTable
                                    DataTable_WS.Columns.Add("X", GetType(Double))
                                    DataTable_WS.Columns.Add("Y", GetType(Double))
                                    DataTable_WS.Columns.Add("TYPE", GetType(String))

                                    Dim DataTable_Xing As New System.Data.DataTable
                                    DataTable_Xing.Columns.Add("DESCRIPTION", GetType(String))
                                    DataTable_Xing.Columns.Add("STA", GetType(Double))
                                    DataTable_Xing.Columns.Add("X", GetType(Double))
                                    DataTable_Xing.Columns.Add("Y", GetType(Double))


                                    Dim Idx1 As Integer = 0
                                    Dim Idx2 As Integer = 0

                                    For i = 0 To Rezultat1.Value.Count - 1
                                        Dim Ent1 As Entity = Trans1.GetObject(Rezultat1.Value(i).ObjectId, OpenMode.ForRead)
                                        If TypeOf Ent1 Is BlockReference Then
                                            Dim Block1 As BlockReference = Ent1
                                            If Block1.Name.ToUpper.Contains("WARNING") = True Then
                                                DataTable_WS.Rows.Add()
                                                DataTable_WS.Rows(Idx1).Item("X") = Block1.Position.X
                                                DataTable_WS.Rows(Idx1).Item("Y") = Block1.Position.Y
                                                DataTable_WS.Rows(Idx1).Item("TYPE") = "WS"
                                                Idx1 = Idx1 + 1
                                            End If

                                            If Block1.Name.ToUpper.Contains("TRIVIEW") = True Then
                                                DataTable_WS.Rows.Add()
                                                DataTable_WS.Rows(Idx1).Item("X") = Block1.Position.X
                                                DataTable_WS.Rows(Idx1).Item("Y") = Block1.Position.Y
                                                DataTable_WS.Rows(Idx1).Item("TYPE") = "TM"
                                                Idx1 = Idx1 + 1
                                            End If

                                            If Block1.Name.ToUpper.Contains("WARNING") = False And Block1.Name.ToUpper.Contains("TRIVIEW") = False Then
                                                If Block1.AttributeCollection.Count > 0 Then
                                                    Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                                    Dim Sta As Double = -1
                                                    Dim Descr As String = ""
                                                    For Each id In attColl
                                                        Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                                        If attref.Tag.ToUpper = "STA" Then
                                                            Dim Continut As String = attref.TextString
                                                            If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                                Sta = CDbl(Replace(Continut, "+", ""))
                                                            End If
                                                        End If
                                                        If attref.Tag.ToUpper = "DESC" Then
                                                            Dim Continut As String = attref.TextString
                                                            If Not Continut = "" Then
                                                                Descr = Continut
                                                            End If
                                                        End If
                                                    Next
                                                    If Not Sta = -1 And Not Descr = "" Then
                                                        DataTable_Xing.Rows.Add()
                                                        DataTable_Xing.Rows(Idx2).Item("X") = Block1.Position.X
                                                        DataTable_Xing.Rows(Idx2).Item("Y") = Block1.Position.Y
                                                        DataTable_Xing.Rows(Idx2).Item("DESCRIPTION") = Descr
                                                        DataTable_Xing.Rows(Idx2).Item("STA") = Sta
                                                        Idx2 = Idx2 + 1
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next

                                    Dim row1 As Integer = 1
                                    If IsNumeric(TextBox_excel_row_warning_signs.Text) = True Then
                                        row1 = CInt(TextBox_excel_row_warning_signs.Text)
                                    End If

                                    If DataTable_WS.Rows.Count > 0 And DataTable_Xing.Rows.Count > 0 Then
                                        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
                                        For i = 0 To DataTable_WS.Rows.Count - 1
                                            Dim Station1 As Double = -1
                                            Dim Descr1 As String = ""
                                            Dim Dist_min As Double = 1000
                                            Dim Type1 As String = DataTable_WS.Rows(i).Item("TYPE")

                                            Dim X_ws As Double = DataTable_WS.Rows(i).Item("X")
                                            For j = 0 To DataTable_Xing.Rows.Count - 1
                                                Dim X_xing As Double = DataTable_Xing.Rows(j).Item("X")
                                                If Abs(X_ws - X_xing) < Dist_min Then
                                                    Station1 = DataTable_Xing.Rows(j).Item("STA")
                                                    Descr1 = DataTable_Xing.Rows(j).Item("DESCRIPTION")
                                                    Dist_min = Abs(X_ws - X_xing)
                                                End If
                                            Next

                                            If Not Station1 = -1 And Not Descr1 = "" And Not Dist_min = 1000 Then
                                                W1.Range("A" & row1).Value = Station1
                                                W1.Range("B" & row1).Value = Descr1
                                                W1.Range("C" & row1).Value = Dist_min
                                                W1.Range("D" & row1).Value = Type1
                                                row1 = row1 + 1
                                            End If

                                            TextBox_excel_row_warning_signs.Text = row1
                                        Next
                                    End If

                                    Trans1.Commit()

                                    MsgBox("DONE")

                                End Using
                            End Using

                        End If
                    End If





                Catch ex As System.SystemException
                    MsgBox(ex.Message)
                End Try
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        End If
    End Sub


    Private Sub Button_water_to_excel_Click(sender As Object, e As EventArgs) Handles Button_water_to_excel.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Dim Empty_array() As ObjectId
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Editor1.SetImpliedSelection(Empty_array)
            Try
                Try
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt.MessageForAdding = vbLf & "Select station bands elements:"
                    Object_Prompt.SingleOnly = False
                    Rezultat1 = Editor1.GetSelection(Object_Prompt)
                    If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        If IsNothing(Rezultat1) = False Then

                            Using Lock_dwg As DocumentLock = ThisDrawing.LockDocument


                                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                                    Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                    BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                                    Dim DataTable_Xing As New System.Data.DataTable
                                    DataTable_Xing.Columns.Add("DESCRIPTION", GetType(String))
                                    DataTable_Xing.Columns.Add("STA", GetType(Double))



                                    Dim Idx1 As Integer = 0
                                    Dim row1 As Integer = 1
                                    If IsNumeric(TextBox_excel_row_water.Text) = True Then
                                        row1 = CInt(TextBox_excel_row_water.Text)
                                    End If

                                    Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()

                                    For i = 0 To Rezultat1.Value.Count - 1
                                        Dim Ent1 As Entity = Trans1.GetObject(Rezultat1.Value(i).ObjectId, OpenMode.ForRead)
                                        If TypeOf Ent1 Is DBText Then
                                            Dim Text1 As DBText = Ent1


                                            If Text1.TextString.ToUpper.Contains("WATERBODY") = True Or Text1.TextString.ToUpper.Contains("STREAM") = True Or Text1.TextString.ToUpper.Contains("CENTERLINE") = True Or Text1.TextString.ToUpper.Contains("WETLAND") = True Then

                                                Dim sTA1 As String = extrage_STATION_din_text_de_la_inceputul_textului(Text1.TextString)
                                                If IsNumeric(Replace(sTA1, "+", "")) = True Then

                                                    W1.Range("A" & row1).Value = Replace(sTA1, "+", "")
                                                    W1.Range("B" & row1).Value = Text1.TextString.Replace(sTA1 & " ", "")
                                                    W1.Range("C" & row1).Value = ThisDrawing.Database.OriginalFileName

                                                    row1 = row1 + 1


                                                End If


                                            End If
                                        End If

                                        If TypeOf Ent1 Is MText Then
                                            Dim Text1 As MText = Ent1


                                            If Text1.Text.ToUpper.Contains("WATERBODY") = True Or Text1.Text.ToUpper.Contains("STREAM") = True Or Text1.Text.ToUpper.Contains("WETLAND") = True Or Text1.Text.ToUpper.Contains("CENTERLINE") = True Then

                                                Dim sTA1 As String = extrage_STATION_din_text_de_la_inceputul_textului(Text1.Text)
                                                If IsNumeric(Replace(sTA1, "+", "")) = True Then

                                                    W1.Range("A" & row1).Value = Replace(sTA1, "+", "")
                                                    W1.Range("B" & row1).Value = Text1.Text.Replace(sTA1 & " ", "")
                                                    W1.Range("C" & row1).Value = ThisDrawing.Database.OriginalFileName

                                                    row1 = row1 + 1


                                                End If


                                            End If
                                        End If


                                    Next






                                    
                                    Trans1.Commit()
                                    TextBox_excel_row_water.Text = row1
                                    MsgBox("DONE")

                                End Using
                            End Using

                        End If
                    End If





                Catch ex As System.SystemException
                    MsgBox(ex.Message)
                End Try
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        End If
    End Sub

End Class
