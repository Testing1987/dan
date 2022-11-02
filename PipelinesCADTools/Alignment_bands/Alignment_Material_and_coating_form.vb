Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class ALIGNMENT_MATERIAL_AND_COATING_FORM
    Dim Colectie1 As New Specialized.StringCollection
    Dim Extra_index_dupa_removal As Integer = 0
    Dim Data_table_Crossing As System.Data.DataTable
    Dim Data_table_Stations As System.Data.DataTable
    Dim Data_table_Centerline As System.Data.DataTable
    Dim Data_table_Match_rotatations As System.Data.DataTable
    Dim Data_table_station_equation As System.Data.DataTable
    Dim Poly_centerline As Polyline
    Dim Nr_pagina As Integer
    Dim Freeze_operations As Boolean = False
    Dim Empty_array() As ObjectId
    Private Sub ALIGNMENT_MATERIAL_AND_COATING_FORM_Load(sender As Object, e As EventArgs) Handles Me.Load
        Data_table_Crossing = New System.Data.DataTable
        Data_table_Crossing.Columns.Add("DESCRIPTION1", GetType(String))
        Data_table_Crossing.Columns.Add("DESCRIPTION2", GetType(String))
        Data_table_Crossing.Columns.Add("STA", GetType(Double))
        Data_table_Crossing.Columns.Add("BEGINSTA", GetType(Double))
        Data_table_Crossing.Columns.Add("ENDSTA", GetType(Double))
        Data_table_Crossing.Columns.Add("LENGTH", GetType(Double))
        Data_table_Crossing.Columns.Add("MATERIAL", GetType(String))
        Data_table_Crossing.Columns.Add("PREVIOUS_MATERIAL", GetType(String))
        Data_table_Crossing.Columns.Add("BLOCKNAME", GetType(String))
        Data_table_Crossing.Columns.Add("SHEET", GetType(Double))
        Data_table_Stations = New System.Data.DataTable
        Data_table_Stations.Columns.Add("STATION", GetType(Double))
    End Sub

    Private Sub Button_Read_excel_Click(sender As Object, e As EventArgs) Handles Button_Read_excel.Click
        If Freeze_operations = False Then
            Freeze_operations = True

            Try

                If TextBox_description1_COL_XL.Text = "" Then
                    MsgBox("Please specify the description EXCEL COLUMN!")
                    Freeze_operations = False
                    Exit Sub
                End If


                If TextBox_chainage_col_xl.Text = "" Then
                    MsgBox("Please specify the station EXCEL COLUMN!")
                    Freeze_operations = False
                    Exit Sub
                End If


                If TextBox_MATERIAL_col_xl.Text = "" Then
                    MsgBox("Please specify the material EXCEL COLUMN!")
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




                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
                Dim start1 As Integer = CInt(TextBox_ROW_START.Text)
                Dim end1 As Integer = CInt(TextBox_ROW_END.Text)
                Dim Col_descr As String = TextBox_description1_COL_XL.Text.ToUpper
                Dim Col_stations As String = TextBox_chainage_col_xl.Text.ToUpper
                Dim Col_mat As String = TextBox_MATERIAL_col_xl.Text.ToUpper

                Extra_index_dupa_removal = 0
                Data_table_Crossing.Rows.Clear()
                Dim Index_Data_table As Integer = 0

                Dim Index_Data_table_station As Integer = 0
                Data_table_Stations = New System.Data.DataTable
                Data_table_Stations.Columns.Add("STATION", GetType(Double))

                Nr_pagina = 1

                Dim Station_previous As Double = 0
                For i = start1 To end1
                    Dim Description1 As String = W1.Range(Col_descr & i).Value

                    Dim Station1 As Double = -1
                    If IsNumeric(Replace(W1.Range(Col_stations & i).Value, "+", "")) = True Then
                        Station1 = CDbl(Replace(W1.Range(Col_stations & i).Value, "+", ""))


                        Data_table_Stations.Rows.Add()
                        Data_table_Stations.Rows(Index_Data_table_station).Item("STATION") = Station1
                        Index_Data_table_station = Index_Data_table_station + 1

                    End If

                    Dim Material As String = ""
                    If Not Replace(W1.Range(Col_mat & i).Value, " ", "") = "" Then
                        Material = W1.Range(Col_mat & i).Value
                    End If

                    If Material.ToUpper = "MATCHLINE" And Not Station1 = -1 Then
                        Data_table_Crossing.Rows.Add()
                        Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = "M"
                        Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "MATCHLINE"
                        Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "MATCHLINE"
                        Data_table_Crossing.Rows(Index_Data_table).Item("STA") = Station1



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

                        Data_table_Crossing.Rows(Index_Data_table).Item("SHEET") = Nr_pagina
                        If Not i = start1 Then Nr_pagina = Nr_pagina + 1
                        Index_Data_table = Index_Data_table + 1
                    End If


                    If IsNothing(Description1) = True Then Description1 = ""

                    If Not Replace(Description1, " ", "") = "" And Not Material.ToUpper = "MATCHLINE" Then
                        Data_table_Crossing.Rows.Add()
                        Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = Description1

                        If Station1 = -1 Then
                            W1.Range(Col_stations & i).Select()
                            MsgBox("Non numerical value at " & Col_stations & i)
                            Freeze_operations = False
                            Exit Sub
                        Else
                            If Station1 < Station_previous Then
                                W1.Range(Col_stations & i).Select()
                                MsgBox("The previous station is bigger than current station" & Col_stations & i)
                                Freeze_operations = False
                                Exit Sub
                            End If
                            Station_previous = Station1
                        End If



                        If Not Replace(Material, " ", "") = "" Then
                            Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = Material
                            If Material.ToUpper = "T" Then
                                Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "TRANSITION"
                                Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "Pipe_Transition"
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

                            If Material.ToUpper = "MC" Then
                                Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "TRANSITION"

                                Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION2") = "FAKE"

                                Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "Pipe_Transition"
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

                            If Not Station1 = -1 Then
                                Data_table_Crossing.Rows(Index_Data_table).Item("STA") = Station1
                            End If

                            If Description1.ToUpper.Contains("ELBOW") = True Then
                                Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "Elbow_al"
                            End If

                        End If












                        Index_Data_table = Index_Data_table + 1
                    Else
                        If Not Replace(Material, " ", "") = "" And Not Station1 = -1 And IsNumeric(Material) = True Then
                            Data_table_Crossing.Rows.Add()
                            Data_table_Crossing.Rows(Index_Data_table).Item("STA") = Station1
                            Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = Material
                            Data_table_Crossing.Rows(Index_Data_table).Item("SHEET") = Nr_pagina
                            Index_Data_table = Index_Data_table + 1
                        End If


                    End If ' If Not Replace(Description1, " ", "") = ""







                Next

                Add_to_clipboard_Data_table(Data_table_Crossing)

            Catch ex As Exception
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub

    Private Function Line_pipe_heavy_wall_pipe_transition(ByVal Mat1 As Integer, ByVal Is_matchline As Boolean) As String

        Select Case Mat1
            Case 1


                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat1_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat1_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If

                End If

                If CheckBox_Show_mat1_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 2


                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If
            Case 3


                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat3_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat3_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat3_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If
            Case 4


                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat4_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat4_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat4_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If
            Case 5


                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat5_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat5_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If


                If CheckBox_Show_mat5_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If
            Case 6


                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat6_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat6_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat6_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 7


                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat7_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat7_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat7_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 8


                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat8_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat8_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat8_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If
            Case 9


                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat9_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat9_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat9_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 10


                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat10_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat10_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat10_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 11


                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat11_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat11_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat11_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 12


                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat12_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat12_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat12_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If


            Case 13
                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat13_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat13_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat13_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 14
                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat14_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat14_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat14_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 15
                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat15_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat15_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat15_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 16
                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat16_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat16_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat16_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 17
                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat17_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat17_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat17_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 18
                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat18_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat18_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat18_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 19
                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat19_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat19_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat19_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 20
                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat20_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat20_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat20_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 21
                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat21_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat21_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat21_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 22
                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat22_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat22_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat22_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 23
                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat23_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat23_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat23_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 24
                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat24_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat24_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat24_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 25
                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat25_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat25_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat25_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 26
                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat26_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat26_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat26_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 27
                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat27_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat27_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat27_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 28
                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat28_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat28_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat28_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 29
                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat29_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat29_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat29_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 30
                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat30_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat30_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat30_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 31
                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat31_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat31_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat31_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 32
                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat32_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat32_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat32_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 33
                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat33_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat33_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat33_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 34
                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat34_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat34_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat34_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 35
                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat35_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat35_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat35_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 36
                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat36_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat36_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat36_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 37
                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat37_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat37_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat37_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 38
                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat38_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat38_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat38_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 39
                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat39_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat39_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat39_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If

            Case 40
                If Is_matchline = True Then
                    If RadioButton_right_to_left.Checked = True Then
                        If CheckBox_Show_mat40_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_right1"
                        Else
                            Return "Heavy_Wall_x_match_right1"
                        End If
                    End If
                    If RadioButton_left_to_right.Checked = True Then
                        If CheckBox_Show_mat40_as_linepipe.Checked = True Then
                            Return "Line_Pipe_x_match_left1"
                        Else
                            Return "Heavy_Wall_x_match_left1"
                        End If
                    End If
                End If

                If CheckBox_Show_mat40_as_linepipe.Checked = True Then
                    Return "Line_Pipe_x1"
                Else
                    Return "Heavy_Wall_x1"
                End If


            Case Else
                Return "Line_Pipe_x1"

        End Select



    End Function

    Private Function Line_pipe_heavy_wall_pipe_matchline(ByVal Mat1 As Integer) As String

        Select Case Mat1
            Case 1
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat1_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat1_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If
            Case 2
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat2_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If

            Case 3
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat3_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat3_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If
            Case 4
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat4_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat4_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If
            Case 5
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat5_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat5_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If
            Case 6
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat6_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat6_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If
            Case 7
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat7_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat7_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If
            Case 8
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat8_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat8_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If
            Case 9
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat9_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat9_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If
            Case 10
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat10_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat10_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If
            Case 11
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat11_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat11_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If
            Case 12
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat12_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat12_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If

            Case 13
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat13_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat13_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If

            Case 14
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat14_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat14_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If

            Case 15
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat15_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat15_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If

            Case 16
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat16_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat16_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If

            Case 17
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat17_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat17_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If

            Case 18
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat18_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat18_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If

            Case 19
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat19_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat19_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If

            Case 20
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat20_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat20_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If

            Case 21
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat21_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat21_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If

            Case 22
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat22_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat22_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If

            Case 23
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat23_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat23_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If

            Case 24
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat24_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat24_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If

            Case 25
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat25_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat25_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If

            Case 26
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat26_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat26_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If

            Case 27
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat27_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat27_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If

            Case 28
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat28_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat28_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If

            Case 29
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat29_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat29_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If

            Case 30
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat30_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat30_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If

            Case 31
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat31_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat31_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If

            Case 32
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat32_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat32_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If

            Case 33
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat33_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat33_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If

            Case 34
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat34_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat34_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If

            Case 35
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat35_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat35_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If

            Case 36
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat36_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat36_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If

            Case 37
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat37_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat37_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If

            Case 38
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat38_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat38_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If

            Case 39
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat39_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat39_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If

            Case 40
                If RadioButton_left_to_right.Checked = True Then
                    If CheckBox_Show_mat40_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_right1"
                    Else
                        Return "Heavy_Wall_x_match_right1"
                    End If
                End If
                If RadioButton_right_to_left.Checked = True Then
                    If CheckBox_Show_mat40_as_linepipe.Checked = True Then
                        Return "Line_Pipe_x_match_left1"
                    Else
                        Return "Heavy_Wall_x_match_left1"
                    End If
                End If


            Case Else
                Return "Line_Pipe_x_match_left1"

        End Select



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
                Dim Colectie_nume_block As New Specialized.StringCollection
                Dim Nr_de_match As Integer = 0

                Using lock1 As DocumentLock = ThisDrawing.LockDocument
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)




                        If Data_table_Crossing.Rows.Count > 0 Then
                            Dim X, Y, Z As Double
                            X = Point1.Value.X
                            Y = Point1.Value.Y
                            Z = 0

                            Dim Path_folder As String = TextBox_folder_for_blocks.Text
                            If Not Strings.Right(Path_folder, 1) = "\" Then
                                Path_folder = Path_folder & "\"
                            End If

                            Dim LayerCurrent As String = TryCast(Trans1.GetObject(ThisDrawing.Database.Clayer, OpenMode.ForRead), LayerTableRecord).Name
                            If LayerCurrent = "" Then LayerCurrent = "0"

                            Dim Station1 As Double
                            Dim Station2 As Double
                            Dim No_of_Xing As Integer = 0
                            Dim Is_matchline As Boolean = False


                            Dim Index_row_excel As Integer = 2

                            Dim Data_table_elbows As New System.Data.DataTable
                            Data_table_elbows.Columns.Add("STA", GetType(Double))
                            Data_table_elbows.Columns.Add("MATERIAL", GetType(Integer))
                            Dim Index_elbow As Integer = 0


                            Dim Strech_distance_1to1 As Double = 100
                            Dim Strech_distance_viewport As Double = -1
                            Dim Strech_distance As Double = 100


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


                                    Select Case Block_name


                                        Case "Pipe_Transition"
                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then

                                                Dim Colectie_atr_namet As New Specialized.StringCollection
                                                Dim Colectie_atr_valuet As New Specialized.StringCollection


                                                Dim Description2 As String = ""
                                                If IsDBNull(Data_table_Crossing.Rows(i).Item("DESCRIPTION2")) = False Then
                                                    Description2 = Data_table_Crossing.Rows(i).Item("DESCRIPTION2")
                                                End If

                                                If Not Description2 = "FAKE" Then
                                                    'InsertBlock_with_multiple_atributes(Path_folder & "Pipe_Transition_x.dwg", "Pipe_Transition_x", New Point3d(X, Y, Z), 1, BTrecord, LayerCurrent, Colectie_atr_namet, Colectie_atr_valuet)
                                                End If


                                                Station1 = Station2
                                                Station2 = Data_table_Crossing.Rows(i).Item("STA")

                                                Dim Nume_block As String = "Heavy_Wall_x1"



                                                If Not i = 0 Then





                                                    Dim Station_rounded1 As Double = Round(Station1, 0)
                                                    Dim Station_rounded2 As Double = Round(Station2, 0)



                                                    If Station_rounded1 < Station_rounded2 Then


                                                        Nume_block = Line_pipe_heavy_wall_pipe_transition(Material_previous, Is_matchline)

                                                        If Is_matchline = True Then
                                                            Is_matchline = False
                                                        End If


                                                        Dim Colectie_atr_name As New Specialized.StringCollection
                                                        Dim Colectie_atr_value As New Specialized.StringCollection


                                                        Colectie_atr_name.Add("BEGINSTA")
                                                        Colectie_atr_name.Add("ENDSTA")

                                                        Dim Station_ref As Double




                                                        If RadioButton_right_to_left.Checked = True Then
                                                            Colectie_atr_value.Add(Get_chainage_feet_from_double(Station2, 0))
                                                            Colectie_atr_value.Add(Get_chainage_feet_from_double(Station1, 0))
                                                            Station_ref = Round(Station2, 0)
                                                        Else
                                                            Colectie_atr_value.Add(Get_chainage_feet_from_double(Station1, 0))
                                                            Colectie_atr_value.Add(Get_chainage_feet_from_double(Station2, 0))
                                                            Station_ref = Round(Station1, 0)

                                                        End If




                                                        Colectie_atr_name.Add("LENGTH")
                                                        Dim String_len As String = Get_String_Rounded(Abs(Station_rounded1 - Station_rounded2), 0) & "'"

                                                        Colectie_atr_value.Add(String_len)
                                                        Colectie_atr_name.Add("MAT")
                                                        Colectie_atr_value.Add(Material_previous)

                                                        Strech_distance_1to1 = Round(Abs(Station_rounded1 - Station_rounded2), 0)
                                                        Dim Twist1 As Double = 0
                                                        Dim CentruMS As New Point3d(0, 0, 0)
                                                        If IsNothing(Data_table_Match_rotatations) = False Then
                                                            If Data_table_Match_rotatations.Rows.Count > 0 Then

                                                                Dim Match1 As Double = -1
                                                                Dim Match2 As Double = -1

                                                                Dim s1 As Double = Station1
                                                                Dim s2 As Double = Station2
                                                                If CheckBox_use_equation.Checked = True Then
                                                                    s1 = Get_length_from_equation_value(Station1)
                                                                    s2 = Get_length_from_equation_value(Station2)
                                                                End If

                                                                For k = 0 To Data_table_Match_rotatations.Rows.Count - 1
                                                                    If IsDBNull(Data_table_Match_rotatations.Rows(k).Item("MATCHLINE1")) = False And _
                                                                        IsDBNull(Data_table_Match_rotatations.Rows(k).Item("MATCHLINE2")) = False And _
                                                                        IsDBNull(Data_table_Match_rotatations.Rows(k).Item("TWIST")) = False And _
                                                                        IsDBNull(Data_table_Match_rotatations.Rows(k).Item("CENTER")) = False Then
                                                                        Dim M1 As Double = Data_table_Match_rotatations.Rows(k).Item("MATCHLINE1")
                                                                        Dim M2 As Double = Data_table_Match_rotatations.Rows(k).Item("MATCHLINE2")
                                                                        Dim T1 As Double = Data_table_Match_rotatations.Rows(k).Item("TWIST")




                                                                        If s1 >= M1 And s1 <= M2 And s2 >= M1 And s2 <= M2 Then
                                                                            Twist1 = T1
                                                                            Match1 = M1
                                                                            Match2 = M2
                                                                            CentruMS = Data_table_Match_rotatations.Rows(k).Item("CENTER")

                                                                            Nr_de_match = Nr_de_match + 1
                                                                            Exit For
                                                                        End If

                                                                    End If
                                                                Next

                                                                If Not Match1 = -1 And Not Match2 = -1 Then
                                                                    Dim Pt1 As New Point3d
                                                                    Dim Pt2 As New Point3d
                                                                    Pt1 = Poly_centerline.GetPointAtDist(S1)
                                                                    Pt2 = Poly_centerline.GetPointAtDist(S2)
                                                                    Dim Pt0 As New Point3d(0, 0, 0)

                                                                    Dim Point1_PS As New Point3d(Pt0.X - (CentruMS.X - Pt1.X) * 1, Pt0.Y - (CentruMS.Y - Pt1.Y) * 1, 0)
                                                                    Point1_PS = Point1_PS.TransformBy(Matrix3d.Rotation(-Twist1, Vector3d.ZAxis, Pt0))
                                                                    Dim Point2_PS As New Point3d(Pt0.X - (CentruMS.X - Pt2.X) * 1, Pt0.Y - (CentruMS.Y - Pt2.Y) * 1, 0)
                                                                    Point2_PS = Point2_PS.TransformBy(Matrix3d.Rotation(-Twist1, Vector3d.ZAxis, Pt0))
                                                                    Strech_distance_viewport = Abs(Point1_PS.X - Point2_PS.X)
                                                                End If

                                                            End If


                                                        End If

                                                        If Not Strech_distance_viewport = -1 Then
                                                            Strech_distance = Strech_distance_viewport
                                                        Else
                                                            Strech_distance = Strech_distance_1to1
                                                        End If

                                                        If RadioButton_left_to_right.Checked = True Then
                                                            X = X + Strech_distance
                                                        Else
                                                            X = X - Strech_distance
                                                        End If


                                                        Dim Block_INS_PT As New Point3d

                                                        If RadioButton_left_to_right.Checked = True Then

                                                            Block_INS_PT = New Point3d(X - Strech_distance, Y, Z)
                                                        Else
                                                            Block_INS_PT = New Point3d(X, Y, Z)
                                                        End If




                                                        Dim Block1 As BlockReference = InsertBlock_with_multiple_atributes(Path_folder & Nume_block & ".dwg", Nume_block, Block_INS_PT, 1, BTrecord, LayerCurrent, Colectie_atr_name, Colectie_atr_value)
                                                        Stretch_block(Block1, "Distance1", Strech_distance)


                                                        If No_of_Xing > 0 Then
                                                            If IsNothing(Data_table_elbows) = False Then
                                                                If Data_table_elbows.Rows.Count > 0 Then

                                                                    Dim ELBOW_INS_PT As New Point3d

                                                                    For k = 0 To No_of_Xing - 1
                                                                        Dim Colectie_atr_name_Elbow As New Specialized.StringCollection
                                                                        Dim Colectie_atr_value_Elbow As New Specialized.StringCollection
                                                                        Colectie_atr_name_Elbow.Add("STA")
                                                                        Colectie_atr_value_Elbow.Add(Get_chainage_feet_from_double(Data_table_elbows.Rows(k).Item("STA"), 0))
                                                                        Colectie_atr_name_Elbow.Add("MAT")
                                                                        Colectie_atr_value_Elbow.Add(Data_table_elbows.Rows(k).Item("MATERIAL"))

                                                                        Dim Station_elbow As Double = Round(Data_table_elbows.Rows(k).Item("STA"), 0)
                                                                        Dim Pt0 As New Point3d(0, 0, 0)
                                                                        Dim Point_ref_PS As New Point3d
                                                                        Dim Point_elb_PS As New Point3d
                                                                        Dim Strech_distance_elbow As Double

                                                                        If IsNothing(Poly_centerline) = False Then
                                                                            Dim Point_ref As New Point3d
                                                                            Point_ref = Poly_centerline.GetPointAtDist(Station_ref)
                                                                            Point_ref_PS = New Point3d(Pt0.X - (CentruMS.X - Point_ref.X) * 1, Pt0.Y - (CentruMS.Y - Point_ref.Y) * 1, 0)
                                                                            Point_ref_PS = Point_ref_PS.TransformBy(Matrix3d.Rotation(-Twist1, Vector3d.ZAxis, Pt0))


                                                                            Dim Point_elbow As New Point3d
                                                                            Point_elbow = Poly_centerline.GetPointAtDist(Station_elbow)
                                                                            Point_elb_PS = New Point3d(Pt0.X - (CentruMS.X - Point_elbow.X) * 1, Pt0.Y - (CentruMS.Y - Point_elbow.Y) * 1, 0)
                                                                            Point_elb_PS = Point_elbow.TransformBy(Matrix3d.Rotation(-Twist1, Vector3d.ZAxis, Pt0))


                                                                            Strech_distance_elbow = Abs(Point_ref_PS.X - Point_elb_PS.X)
                                                                        Else

                                                                            If RadioButton_left_to_right.Checked = True Then
                                                                                Strech_distance_elbow = Station_elbow - Station_ref
                                                                            Else
                                                                                Strech_distance_elbow = Station_ref - Station_elbow
                                                                            End If

                                                                        End If



                                                                        ELBOW_INS_PT = New Point3d(Block_INS_PT.X + Strech_distance_elbow, Y, Z)



                                                                        InsertBlock_with_multiple_atributes(Path_folder & "Elbow_Alignment_x.dwg", "Elbow_Alignment_x", ELBOW_INS_PT, 1, BTrecord, LayerCurrent, _
                                                                                                            Colectie_atr_name_Elbow, Colectie_atr_value_Elbow)
                                                                    Next

                                                                End If
                                                            End If
                                                        End If



                                                        Data_table_elbows = New System.Data.DataTable
                                                        Data_table_elbows.Columns.Add("STA", GetType(Double))
                                                        Data_table_elbows.Columns.Add("MATERIAL", GetType(Integer))

                                                        Index_elbow = 0
                                                        No_of_Xing = 0

                                                    End If








                                                End If



                                            End If




                                        Case "Elbow_al"
                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then

                                                If IsNumeric(Data_table_Crossing.Rows(i).Item("STA")) = True Then
                                                    Data_table_elbows.Rows.Add()
                                                    Data_table_elbows.Rows(Index_elbow).Item("STA") = Data_table_Crossing.Rows(i).Item("STA")

                                                    If IsDBNull(Data_table_Crossing.Rows(i).Item("MATERIAL")) = False Then
                                                        If IsNumeric(Data_table_Crossing.Rows(i).Item("MATERIAL")) = True Then
                                                            Data_table_elbows.Rows(Index_elbow).Item("MATERIAL") = CInt(Data_table_Crossing.Rows(i).Item("MATERIAL"))
                                                        End If
                                                    End If


                                                    Index_elbow = Index_elbow + 1


                                                    No_of_Xing = No_of_Xing + 1
                                                    'If No_of_Xing > 8 Then

                                                    'MsgBox("Please create the number " & No_of_Xing & " block!")
                                                    'No_of_Xing = 8
                                                    'End If
                                                End If
                                            End If

                                        Case "MATCHLINE"
                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                                                Station1 = Station2
                                                Station2 = Data_table_Crossing.Rows(i).Item("STA")

                                                Dim Nume_block_from_match As String = "Line_Pipe_x_match_right1"
                                                If Not i = 0 Then








                                                    Dim Station_rounded1 As Double = Round(Station1, 0)
                                                    Dim Station_rounded2 As Double = Round(Station2, 0)

                                                    Nume_block_from_match = Line_pipe_heavy_wall_pipe_matchline(Material_previous)



                                                    Dim Colectie_atr_name As New Specialized.StringCollection
                                                    Dim Colectie_atr_value As New Specialized.StringCollection


                                                    Colectie_atr_name.Add("BEGINSTA")
                                                    Colectie_atr_name.Add("ENDSTA")


                                                    Dim Station_ref As Double
                                                    If RadioButton_right_to_left.Checked = True Then
                                                        Colectie_atr_value.Add(Get_chainage_feet_from_double(Station2, 0))
                                                        Colectie_atr_value.Add(Get_chainage_feet_from_double(Station1, 0))
                                                        Station_ref = Round(Station2, 0)
                                                    Else
                                                        Colectie_atr_value.Add(Get_chainage_feet_from_double(Station1, 0))
                                                        Colectie_atr_value.Add(Get_chainage_feet_from_double(Station2, 0))
                                                        Station_ref = Round(Station1, 0)
                                                    End If

                                                    If CheckBox_use_equation.Checked = True Then
                                                        Station_ref = Get_length_from_equation_value(Station_ref)
                                                    End If

                                                    Colectie_atr_name.Add("LENGTH")
                                                    Dim String_len As String = Get_String_Rounded(Abs(Station_rounded1 - Station_rounded2), 0) & "'"


                                                    Colectie_atr_value.Add(String_len)
                                                    Colectie_atr_name.Add("MAT")
                                                    Colectie_atr_value.Add(Material_previous)


                                                    Strech_distance = Round(Abs(Station_rounded1 - Station_rounded2), 0)

                                                    Strech_distance_1to1 = Round(Abs(Station_rounded1 - Station_rounded2), 0)
                                                    Dim Twist1 As Double = 0
                                                    Dim CentruMS As New Point3d(0, 0, 0)

                                                    Dim s1 As Double = Station1
                                                    Dim s2 As Double = Station2
                                                    If CheckBox_use_equation.Checked = True Then
                                                        s1 = Get_length_from_equation_value(Station1)
                                                        s2 = Get_length_from_equation_value(Station2)
                                                    End If


                                                    If IsNothing(Data_table_Match_rotatations) = False Then
                                                        If Data_table_Match_rotatations.Rows.Count > 0 Then

                                                            Dim Match1 As Double = -1
                                                            Dim Match2 As Double = -1
                                                            For k = 0 To Data_table_Match_rotatations.Rows.Count - 1
                                                                If IsDBNull(Data_table_Match_rotatations.Rows(k).Item("MATCHLINE1")) = False And _
                                                                    IsDBNull(Data_table_Match_rotatations.Rows(k).Item("MATCHLINE2")) = False And _
                                                                    IsDBNull(Data_table_Match_rotatations.Rows(k).Item("TWIST")) = False Then
                                                                    Dim M1 As Double = Data_table_Match_rotatations.Rows(k).Item("MATCHLINE1")
                                                                    Dim M2 As Double = Data_table_Match_rotatations.Rows(k).Item("MATCHLINE2")
                                                                    Dim T1 As Double = Data_table_Match_rotatations.Rows(k).Item("TWIST")
                                                                    If s1 >= M1 And s1 <= M2 And s2 >= M1 And s2 <= M2 Then
                                                                        Twist1 = T1
                                                                        Match1 = M1
                                                                        Match2 = M2
                                                                        Exit For
                                                                    End If

                                                                End If
                                                            Next

                                                            If Not Match1 = -1 And Not Match2 = -1 Then
                                                                Dim Pt1 As New Point3d
                                                                Dim Pt2 As New Point3d
                                                                Pt1 = Poly_centerline.GetPointAtDist(s1)
                                                                Pt2 = Poly_centerline.GetPointAtDist(s2)
                                                                Dim Pt0 As New Point3d(0, 0, 0)


                                                                Dim Point1_PS As New Point3d(Pt0.X - (CentruMS.X - Pt1.X) * 1, Pt0.Y - (CentruMS.Y - Pt1.Y) * 1, 0)
                                                                Point1_PS = Point1_PS.TransformBy(Matrix3d.Rotation(-Twist1, Vector3d.ZAxis, Pt0))
                                                                Dim Point2_PS As New Point3d(Pt0.X - (CentruMS.X - Pt2.X) * 1, Pt0.Y - (CentruMS.Y - Pt2.Y) * 1, 0)
                                                                Point2_PS = Point2_PS.TransformBy(Matrix3d.Rotation(-Twist1, Vector3d.ZAxis, Pt0))

                                                                Strech_distance_viewport = Abs(Point1_PS.X - Point2_PS.X)
                                                            End If

                                                        End If


                                                    End If

                                                    If Not Strech_distance_viewport = -1 Then
                                                        Strech_distance = Strech_distance_viewport
                                                    Else
                                                        Strech_distance = Strech_distance_1to1
                                                    End If


                                                    If RadioButton_left_to_right.Checked = True Then
                                                        X = X + Strech_distance
                                                    Else
                                                        X = X - Strech_distance
                                                    End If


                                                    Dim Block_INS_PT As New Point3d

                                                    If RadioButton_left_to_right.Checked = True Then

                                                        Block_INS_PT = New Point3d(X - Strech_distance, Y, Z)
                                                    Else
                                                        Block_INS_PT = New Point3d(X, Y, Z)
                                                    End If


                                                    If Not i = 0 Then




                                                        Dim Block1 As BlockReference = InsertBlock_with_multiple_atributes(Path_folder & Nume_block_from_match & ".dwg", Nume_block_from_match, Block_INS_PT, 1, BTrecord, LayerCurrent, Colectie_atr_name, Colectie_atr_value)
                                                        Stretch_block(Block1, "Distance1", Strech_distance)




                                                    End If

                                                    If No_of_Xing > 0 Then
                                                        If IsNothing(Data_table_elbows) = False Then
                                                            If Data_table_elbows.Rows.Count > 0 Then


                                                                Dim ELBOW_INS_PT As New Point3d
                                                                For k = 0 To No_of_Xing - 1
                                                                    Dim Colectie_atr_name_Elbow As New Specialized.StringCollection
                                                                    Dim Colectie_atr_value_Elbow As New Specialized.StringCollection
                                                                    Colectie_atr_name_Elbow.Add("STA")
                                                                    Colectie_atr_value_Elbow.Add(Get_chainage_feet_from_double(Data_table_elbows.Rows(k).Item("STA"), 0))
                                                                    Colectie_atr_name_Elbow.Add("MAT")
                                                                    Colectie_atr_value_Elbow.Add(Data_table_elbows.Rows(k).Item("MATERIAL"))


                                                                    Dim Station_elbow As Double = Round(Data_table_elbows.Rows(k).Item("STA"), 0)

                                                                    If CheckBox_use_equation.Checked = True Then
                                                                        Station_elbow = Get_length_from_equation_value(Station_elbow)
                                                                    End If



                                                                    Dim Pt0 As New Point3d(0, 0, 0)
                                                                    Dim Point_ref_PS As New Point3d
                                                                    Dim Point_elb_PS As New Point3d
                                                                    Dim Strech_distance_elbow As Double = 0

                                                                    If IsNothing(Poly_centerline) = False Then
                                                                        Dim Point_ref As New Point3d
                                                                        Point_ref = Poly_centerline.GetPointAtDist(Station_ref)

                                                                        Point_ref_PS = Point_ref.TransformBy(Matrix3d.Rotation(-Twist1, Vector3d.ZAxis, Pt0))


                                                                        Dim Point_elbow As New Point3d
                                                                        Point_elbow = Poly_centerline.GetPointAtDist(Station_elbow)

                                                                        Point_elb_PS = Point_elbow.TransformBy(Matrix3d.Rotation(-Twist1, Vector3d.ZAxis, Pt0))

                                                                        Strech_distance_elbow = Abs(Point_ref_PS.X - Point_elb_PS.X)
                                                                    Else

                                                                        If RadioButton_left_to_right.Checked = True Then
                                                                            Strech_distance_elbow = Station_elbow - Station_ref
                                                                        Else
                                                                            Strech_distance_elbow = Station_ref - Station_elbow
                                                                        End If

                                                                    End If

                                                                    ELBOW_INS_PT = New Point3d(Block_INS_PT.X + Strech_distance_elbow, Y, Z)

                                                                    InsertBlock_with_multiple_atributes(Path_folder & "Elbow_Alignment_x.dwg", "Elbow_Alignment_x", ELBOW_INS_PT, 1, BTrecord, LayerCurrent, _
                                                                                                        Colectie_atr_name_Elbow, Colectie_atr_value_Elbow)
                                                                Next




                                                                Data_table_elbows = New System.Data.DataTable
                                                                Data_table_elbows.Columns.Add("STA", GetType(Double))
                                                                Data_table_elbows.Columns.Add("MATERIAL", GetType(Integer))

                                                                Index_elbow = 0
                                                                No_of_Xing = 0

                                                            End If
                                                        End If
                                                    End If

                                                End If

                                            End If '  If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False

                                            Is_matchline = True


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
                                            If Not i = 0 Then Y = Y - 150

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

                                Strech_distance = 100
                                Strech_distance_1to1 = 100
                                Strech_distance_viewport = -1
end1:
                            Next '  For i = 0 To Data_table_Crossing.Rows.Count - 1




                        End If '   If Data_table_Crossing.Rows.Count > 0

                        Trans1.Commit()
                    End Using




                End Using

                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
            Catch ex As Exception
                Freeze_operations = False
                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
            End Try


            Freeze_operations = False
        End If


    End Sub


    Private Sub Button_draw_positions_Click(sender As Object, e As EventArgs) Handles Button_draw_positions.Click
        Try
            If IsNothing(Data_table_Stations) = False Then
                If Data_table_Stations.Rows.Count > 0 And Data_table_Centerline.Rows.Count > 0 Then
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim Len1 As Double = 0
                    Dim Height1 As Double = 0
                    Dim Scale1 As Double = 0
                    Dim Angle1 As Double = 0

                    Dim Centru_MS As New Point3d(0, 0, 0)
                    Dim Centru_PS As New Point3d(0, 0, 0)

                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor

                    Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                        Editor1.SetImpliedSelection(Empty_array)
                        ' Dim k As Double = 1
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            Dim BTrecordPS As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                            Dim BTrecordMS As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                            BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            BTrecordPS = Trans1.GetObject(BlockTable_data1(BlockTableRecord.PaperSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                            BTrecordMS = Trans1.GetObject(BlockTable_data1(BlockTableRecord.ModelSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                            Dim Text_style_table As Autodesk.AutoCAD.DatabaseServices.TextStyleTable = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                            Dim Layer_table As Autodesk.AutoCAD.DatabaseServices.LayerTable = Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)


                            Dim Rezultat_viewport As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                            Dim Prompt_viewport As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                            Prompt_viewport.MessageForAdding = vbLf & "Select the viewport"

                            Prompt_viewport.SingleOnly = False
                            Rezultat_viewport = Editor1.GetSelection(Prompt_viewport)

                            If Rezultat_viewport.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Editor1.SetImpliedSelection(Empty_array)
                                Editor1.WriteMessage(vbLf & "Command:")
                                Exit Sub
                            End If

                            Dim Exista_viewport As Boolean = False

                            For i = 0 To Rezultat_viewport.Value.Count - 1
                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Rezultat_viewport.Value.Item(i)
                                Dim Ent1 As Entity
                                Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Viewport Then
                                    Dim Viewport1 As Viewport = Ent1

                                    Editor1.SwitchToModelSpace()
                                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("CVPORT", Viewport1.Number)
                                    'Editor1.CurrentUserCoordinateSystem = WCS_align()
                                    'Dim GraphicsManager As Autodesk.AutoCAD.GraphicsSystem.Manager = ThisDrawing.GraphicsManager
                                    'Dim View0 As Autodesk.AutoCAD.GraphicsSystem.View = GraphicsManager.GetGsView(CShort(Application.GetSystemVariable("CVPORT")), True) ' acad 2013

                                    'Dim View0 As Autodesk.AutoCAD.GraphicsSystem.View = GraphicsManager.GetCurrentAcGsView(CShort(Application.GetSystemVariable("CVPORT"))) ' acad 2015


                                    Angle1 = Viewport1.TwistAngle
                                    'Len1 = Viewport1.Width
                                    'Height1 = Viewport1.Height
                                    Scale1 = Viewport1.CustomScale
                                    Centru_MS = Application.GetSystemVariable("VIEWCTR")
                                    Centru_PS = Viewport1.CenterPoint
                                    Exista_viewport = True
                                    Editor1.SwitchToPaperSpace()
                                End If
                                Exit For
                            Next

                            Dim Point_MS As New Point3d
                            Dim Point_PS As New Point3d

                            If Exista_viewport = True Then


                                Dim Match1 As Double = 0
                                If IsNumeric(Replace(TextBox_matchline_start.Text, "+", "")) = True Then Match1 = CDbl(Replace(TextBox_matchline_start.Text, "+", ""))
                                Dim Match2 As Double = 0
                                If IsNumeric(Replace(TextBox_matchline_end.Text, "+", "")) = True Then Match2 = CDbl(Replace(TextBox_matchline_end.Text, "+", ""))

                                If Match1 > Match2 Then
                                    Dim Temp As Double = Match1
                                    Match1 = Match2
                                    Match2 = Temp
                                End If

                                If Match1 = 0 And Match2 = 0 Then
                                    Editor1.SetImpliedSelection(Empty_array)
                                    Editor1.WriteMessage(vbLf & "Command:")
                                    MsgBox("Please specify the matchlines")
                                    Exit Sub
                                End If

                                Dim Y1 As Double = 0
                                Dim Y2 As Double = 0
                                If IsNumeric(TextBox_Y1.Text) = True Then
                                    Y1 = CDbl(TextBox_Y1.Text)
                                End If

                                If IsNumeric(TextBox_Y2.Text) = True Then
                                    Y2 = CDbl(TextBox_Y2.Text)
                                End If



                                If Y1 = Y2 Then
                                    MsgBox("y1=y2")
                                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

                                End If

                                If Y1 > Y2 Then
                                    Dim Temp As Double = Y1
                                    Y1 = Y2
                                    Y2 = Temp
                                End If




                                Dim Text_offset_vert_line As Double = 5
                                If IsNumeric(TextBox_textOffset.Text) = True Then
                                    Text_offset_vert_line = CDbl(TextBox_textOffset.Text)
                                End If

                                Dim TextHeight As Double = 16
                                If IsNumeric(TextBox_text_height.Text) = True Then
                                    TextHeight = CDbl(TextBox_text_height.Text)
                                End If

                                Dim Mtext_height As Double = TextHeight
                                Dim Mtext_rotation As Double = PI / 2

                                Dim Block_scale As Double = 1

                                Dim Poly_2D As New Polyline
                                For i = 0 To Data_table_Centerline.Rows.Count - 1
                                    Poly_2D.AddVertexAt(i, New Point2d(Data_table_Centerline.Rows(i).Item("X"), Data_table_Centerline.Rows(i).Item("Y")), 0, 0, 0)
                                Next

                                Dim Point_M1_MS As New Point3d(0, 0, 0)
                                Dim Point_M2_MS As New Point3d(0, 0, 0)

                                If Poly_2D.Length >= Match2 Then
                                    Point_M1_MS = Poly_2D.GetPointAtDist(Match1)
                                    Point_M2_MS = Poly_2D.GetPointAtDist(Match2)
                                End If

                                Dim Point_next_MS As New Point3d(0, 0, 0)
                                Dim Point_next_PS As New Point3d(0, 0, 0)

                                Dim Point_M1_PS As New Point3d(Centru_PS.X - (Centru_MS.X - Point_M1_MS.X) * Scale1, Centru_PS.Y - (Centru_MS.Y - Point_M1_MS.Y) * Scale1, 0)
                                Point_M1_PS = Point_M1_PS.TransformBy(Matrix3d.Rotation(Angle1, Vector3d.ZAxis, Centru_PS))

                                Creaza_layer("NO PLOT", 40, "NO PLOT", False)

                                For i = 0 To Data_table_Stations.Rows.Count - 1
                                    If IsDBNull(Data_table_Stations.Rows(i).Item("STATION")) = False Then
                                        Dim Station As Double = Data_table_Stations.Rows(i).Item("STATION")



                                        If Station >= Match1 And Station <= Match2 Then
                                            Point_next_MS = Poly_2D.GetPointAtDist(Station)
                                            Point_next_PS = New Point3d(Centru_PS.X - (Centru_MS.X - Point_next_MS.X) * Scale1, Centru_PS.Y - (Centru_MS.Y - Point_next_MS.Y) * Scale1, 0)
                                            Point_next_PS = Point_next_PS.TransformBy(Matrix3d.Rotation(Angle1, Vector3d.ZAxis, Centru_PS))
                                            Point_M1_PS = New Point3d(Point_next_PS.X, Point_M1_PS.Y, 0)
                                            Exit For
                                        End If
                                    End If
                                Next






                                Dim Linie2 As New Line(New Point3d(Point_M1_PS.X, Y1, 0), New Point3d(Point_M1_PS.X, Y2, 0))

                                Linie2.Layer = "NO PLOT"

                                BTrecordPS.AppendEntity(Linie2)
                                Trans1.AddNewlyCreatedDBObject(Linie2, True)




                                Dim Mtext2 As New MText
                                Mtext2.Location = New Point3d(Point_M1_PS.X - Text_offset_vert_line, Y1 + Text_offset_vert_line, 0)
                                Mtext2.TextHeight = Mtext_height
                                Mtext2.Rotation = Mtext_rotation
                                Mtext2.Layer = "NO PLOT"
                                Mtext2.Contents = Get_chainage_feet_from_double(Match1, 0)

                                Mtext2.Attachment = AttachmentPoint.BottomLeft
                                BTrecordPS.AppendEntity(Mtext2)
                                Trans1.AddNewlyCreatedDBObject(Mtext2, True)



                                For i = 0 To Data_table_Stations.Rows.Count - 1
                                    If IsDBNull(Data_table_Stations.Rows(i).Item("STATION")) = False Then
                                        Dim Station As Double = Data_table_Stations.Rows(i).Item("STATION")

                                        If Station >= Match1 And Station <= Match2 Then

                                            Point_MS = Poly_2D.GetPointAtDist(Station)
                                            Point_PS = New Point3d(Centru_PS.X - (Centru_MS.X - Point_MS.X) * Scale1, Centru_PS.Y - (Centru_MS.Y - Point_MS.Y) * Scale1, 0)
                                            Point_PS = Point_PS.TransformBy(Matrix3d.Rotation(Angle1, Vector3d.ZAxis, Centru_PS))
                                            Dim Linie1 As New Line(New Point3d(Point_PS.X, Y1, 0), New Point3d(Point_PS.X, Y2, 0))
                                            Linie1.Layer = "NO PLOT"

                                            BTrecordPS.AppendEntity(Linie1)

                                            Trans1.AddNewlyCreatedDBObject(Linie1, True)
                                            Dim Mtext As New MText
                                            Mtext.Location = New Point3d(Point_PS.X - Text_offset_vert_line, Y1 + Text_offset_vert_line, 0)

                                            Mtext.TextHeight = Mtext_height
                                            Mtext.Rotation = Mtext_rotation
                                            Mtext.Layer = "NO PLOT"
                                            Mtext.Contents = Get_chainage_feet_from_double(Station, 0)


                                            Mtext.Attachment = AttachmentPoint.BottomLeft

                                            BTrecordPS.AppendEntity(Mtext)
                                            Trans1.AddNewlyCreatedDBObject(Mtext, True)







                                        End If
                                    End If
                                Next





                                Dim Point_M2_PS As New Point3d(Centru_PS.X - (Centru_MS.X - Point_M2_MS.X) * Scale1, Centru_PS.Y - (Centru_MS.Y - Point_M2_MS.Y) * Scale1, 0)
                                Point_M2_PS = Point_M2_PS.TransformBy(Matrix3d.Rotation(Angle1, Vector3d.ZAxis, Centru_PS))







                                Dim Linie3 As New Line(New Point3d(Point_M2_PS.X, Y1, 0), New Point3d(Point_M2_PS.X, Y2, 0))
                                Linie3.Layer = "NO PLOT"

                                BTrecordPS.AppendEntity(Linie3)
                                Trans1.AddNewlyCreatedDBObject(Linie3, True)

                                Dim mtext3 As New MText
                                mtext3.Location = New Point3d(Point_M2_PS.X - Text_offset_vert_line, Y1 + Text_offset_vert_line, 0)
                                mtext3.TextHeight = Mtext_height
                                mtext3.Rotation = Mtext_rotation
                                mtext3.Layer = "NO PLOT"

                                mtext3.Contents = Get_chainage_feet_from_double(Match2, 0)

                                mtext3.Attachment = AttachmentPoint.BottomLeft




                                BTrecordPS.AppendEntity(mtext3)
                                Trans1.AddNewlyCreatedDBObject(mtext3, True)




                            End If ' asta e de la exista viewport

                            Trans1.Commit()
                            ' asta e de la tranzactie
                        End Using





                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                        ' asta e de la lock
                    End Using


                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                End If
            End If


        Catch ex As Exception
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_load_CL_Click(sender As Object, e As EventArgs) Handles Button_load_CL.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor


        Editor1.SetImpliedSelection(Empty_array)
        Try
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Colectie1 = New Specialized.StringCollection


            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptEntityResult

            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select centerline:")

            Object_Prompt.SetRejectMessage(vbLf & "Please select a lightweight polyline")
            Object_Prompt.AddAllowedClass(GetType(Polyline), True)

            Rezultat1 = Editor1.GetEntity(Object_Prompt)


            If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")

                Editor1.SetImpliedSelection(Empty_array)
                Exit Sub
            End If

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then

                    Using Lock_dwg As DocumentLock = ThisDrawing.LockDocument


                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim Ent1 As Entity = Trans1.GetObject(Rezultat1.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                            Dim Poly1 As Polyline = Ent1
                            Data_table_Centerline = New System.Data.DataTable
                            Data_table_Centerline.Columns.Add("X", GetType(Double))
                            Data_table_Centerline.Columns.Add("Y", GetType(Double))

                            For i = 0 To Poly1.NumberOfVertices - 1
                                Data_table_Centerline.Rows.Add()
                                Data_table_Centerline.Rows(i).Item("X") = Poly1.GetPoint3dAt(i).X
                                Data_table_Centerline.Rows(i).Item("Y") = Poly1.GetPoint3dAt(i).Y

                            Next
                        End Using
                    End Using

                End If
            End If



            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")

        Catch ex As Exception
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub Button_pick_match_from_text_Click(sender As Object, e As EventArgs) Handles Button_pick_match_from_text.Click

        Try
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
            Editor1 = ThisDrawing.Editor

            Dim Station1 As Double = 0
            Dim Station2 As Double = 0

            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

            Using lock As DocumentLock = ThisDrawing.LockDocument
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select text/mtext objects:"

                Object_Prompt.SingleOnly = False
                Rezultat1 = Editor1.GetSelection(Object_Prompt)

                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    If IsNothing(Rezultat1) = False Then
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                            BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)

                            For i = 1 To Rezultat1.Value.Count
                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Rezultat1.Value.Item(i - 1)
                                Dim Ent1 As Entity
                                Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                                If TypeOf Ent1 Is DBText Then
                                    Dim Text1 As DBText = Ent1
                                    If Text1.TextString.Contains("+") Then
                                        Dim Text_chainage As String = extrage_station_din_text_de_la_sfarsitul_textului(Text1.TextString)
                                        If IsNumeric(Text_chainage) = True Then
                                            If Station1 = 0 Then
                                                Station1 = CDbl(Text_chainage)
                                            Else
                                                Station2 = CDbl(Text_chainage)
                                            End If
                                        End If
                                    End If
                                End If

                                If TypeOf Ent1 Is MText Then
                                    Dim mText1 As MText = Ent1
                                    If mText1.Contents.Contains("+") Then
                                        Dim Text_chainage As String = extrage_station_din_text_de_la_sfarsitul_textului(mText1.Text)
                                        If IsNumeric(Text_chainage) = True Then
                                            If Station1 = 0 Then
                                                Station1 = CDbl(Text_chainage)
                                            Else
                                                Station2 = CDbl(Text_chainage)
                                            End If

                                        End If
                                    End If
                                End If
                            Next

                            Editor1.Regen()
                            Trans1.Commit()
                        End Using

                    End If
                End If

                If Station1 < Station2 Then
                    TextBox_matchline_start.Text = Station1
                    TextBox_matchline_end.Text = Station2
                Else
                    TextBox_matchline_start.Text = Station2
                    TextBox_matchline_end.Text = Station1
                End If



                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            End Using




        Catch ex As Exception

            MsgBox(ex.Message)
        End Try

    End Sub
    Public Function extrage_station_din_text_de_la_sfarsitul_textului(ByVal string1 As String) As String
        Try
            Dim Numar As String = ""

            For i = string1.Length To 1 Step -1
                Dim Litera As String = Mid(string1, i, 1)

                Select Case Litera
                    Case "."
                        Numar = Litera & Numar
                    Case "0"
                        Numar = Litera & Numar
                    Case "1"
                        Numar = Litera & Numar
                    Case "2"
                        Numar = Litera & Numar
                    Case "3"
                        Numar = Litera & Numar
                    Case "4"
                        Numar = Litera & Numar
                    Case "5"
                        Numar = Litera & Numar
                    Case "6"
                        Numar = Litera & Numar
                    Case "7"
                        Numar = Litera & Numar
                    Case "8"
                        Numar = Litera & Numar
                    Case "9"
                        Numar = Litera & Numar
                    Case "-"
                        If i = 1 Then Numar = Litera & Numar
                    Case "+"

                    Case Else
                        Exit For
                End Select
            Next



            Return Numar

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function



    Private Sub Button_pick_POSITION_ZERO_Click(sender As Object, e As EventArgs) Handles Button_pick_POSITION_ZERO.Click
        Try

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
            ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()


            Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Pt_rezult As Autodesk.AutoCAD.EditorInput.PromptPointResult
                    Dim Prompt_pt As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify 0+00 position:")

                    Prompt_pt.AllowNone = True
                    Pt_rezult = Editor1.GetPoint(Prompt_pt)

                    If Pt_rezult.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        TextBox_X_MS.Text = Get_String_Rounded(Pt_rezult.Value.X, 4)
                        TextBox_Y_MS.Text = Get_String_Rounded(Pt_rezult.Value.Y, 4)
                    End If





                End Using
            End Using

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
    End Sub

    Private Sub Button_pick_viewport_corner_Click(sender As Object, e As EventArgs) Handles Button_pick_viewport_corner.Click
        Try

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
            ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()


            Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)


                    Dim Pt_rezult As Autodesk.AutoCAD.EditorInput.PromptPointResult
                    Dim Prompt_pt As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify viewport lower left corner:")

                    Prompt_pt.AllowNone = True
                    Pt_rezult = Editor1.GetPoint(Prompt_pt)

                    If Pt_rezult.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        TextBox_X_PS.Text = Get_String_Rounded(Pt_rezult.Value.X, 4)
                        TextBox_Y_PS.Text = Get_String_Rounded(Pt_rezult.Value.Y, 4)
                    End If




                End Using
            End Using

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
    End Sub

    Private Sub Button_draw_viewport_Click(sender As Object, e As EventArgs) Handles Button_draw_Viewport.Click


        Try
            If IsNumeric(TextBox_page.Text) = False Then
                MsgBox("Please specify the page!")
                Exit Sub
            End If
            If IsNumeric(TextBox_SCALE.Text) = False Then
                MsgBox("Please specify the viewport scale!")
                Exit Sub
            End If
            If IsNumeric(TextBox_Height.Text) = False Then
                MsgBox("Please specify the viewport height!")
                Exit Sub
            End If
            If IsNumeric(TextBox_Width.Text) = False Then
                MsgBox("Please specify the viewport width!")
                Exit Sub
            End If
            If IsNumeric(TextBox_BAND_SPACING.Text) = False Then
                MsgBox("Please specify the distance between bands!")
                Exit Sub
            End If
            If IsNumeric(TextBox_X_MS.Text) = False Then
                MsgBox("Please specify the X of the 0+00 station!")
                Exit Sub
            End If
            If IsNumeric(TextBox_Y_MS.Text) = False Then
                MsgBox("Please specify the Y of the 0+00 station!")
                Exit Sub
            End If

            If IsNumeric(TextBox_X_PS.Text) = False Then
                MsgBox("Please specify the X of the viewport corner!")
                Exit Sub
            End If
            If IsNumeric(TextBox_Y_PS.Text) = False Then
                MsgBox("Please specify the Y of the viewport corner!")
                Exit Sub
            End If

            Dim Page1 As Integer = CInt(TextBox_page.Text)

            Dim Scale1 As Double = CDbl(TextBox_SCALE.Text)

            Dim Spacing1 As Double = CDbl(TextBox_BAND_SPACING.Text)

            Dim H1 As Double = CDbl(TextBox_Height.Text)
            Dim W1 As Double = CDbl(TextBox_Width.Text)

            Dim x_MS As Double = CDbl(TextBox_X_MS.Text)
            Dim y_MS As Double = CDbl(TextBox_Y_MS.Text)

            Dim x_pS As Double = CDbl(TextBox_X_PS.Text)
            Dim y_PS As Double = CDbl(TextBox_Y_PS.Text)

            Dim DeltaY As Double
            If IsNumeric(TextBox_shift_viewport.Text) = True Then
                DeltaY = CDbl(TextBox_shift_viewport.Text)
            End If

            If Scale1 <= 0 Or Page1 <= 0 Or Spacing1 <= 0 Or H1 <= 0 Or W1 <= 0 Then
                MsgBox("Negative values not allowed")
                Exit Sub
            End If


            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
            ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()


            Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Creaza_layer("VP", 4, "VIEWPORT", False)
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    Dim BTrecordMS As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BTrecordMS = Trans1.GetObject(BlockTable_data1(BlockTableRecord.ModelSpace), OpenMode.ForRead)
                    Dim BTrecordPS As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BTrecordPS = Trans1.GetObject(BlockTable_data1(BlockTableRecord.PaperSpace), OpenMode.ForWrite)

                    Dim Point_target As Point3d

                    If RadioButton_left_right_viewport.Checked = True Then
                        Point_target = New Point3d(x_MS + (W1 / 2) / Scale1, y_MS - Spacing1 * (Page1 - 1) + DeltaY / Scale1, 0)
                    Else
                        Point_target = New Point3d(x_MS - (W1 / 2) / Scale1, y_MS - Spacing1 * (Page1 - 1) + DeltaY / Scale1, 0)
                    End If

                    Dim Viewport1 As New Viewport
                    Viewport1.SetDatabaseDefaults()
                    Viewport1.CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(x_pS + W1 / 2, y_PS + H1 / 2, 0) ' asta e pozitia viewport in paper space
                    Viewport1.Height = H1
                    Viewport1.Width = W1
                    Viewport1.Layer = "VP"

                    Viewport1.ViewDirection = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis
                    Viewport1.ViewTarget = Point_target ' asta e pozitia viewport in MODEL space
                    Viewport1.ViewCenter = Autodesk.AutoCAD.Geometry.Point2d.Origin
                    Viewport1.TwistAngle = 0 ' asta e PT TWIST

                    BTrecordPS.AppendEntity(Viewport1)
                    Trans1.AddNewlyCreatedDBObject(Viewport1, True)

                    Viewport1.On = True
                    Viewport1.CustomScale = Scale1
                    Viewport1.Locked = True



                    Trans1.Commit()

                End Using
            End Using

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")


    End Sub


    Private Sub Button_read_centerline_Click(sender As Object, e As EventArgs) Handles Button_read_centerline.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor


            Editor1.SetImpliedSelection(Empty_array)
            Try
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                Colectie1 = New Specialized.StringCollection


                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptEntityResult

                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select centerline:")

                Object_Prompt.SetRejectMessage(vbLf & "Please select a lightweight polyline")
                Object_Prompt.AddAllowedClass(GetType(Polyline), True)

                Rezultat1 = Editor1.GetEntity(Object_Prompt)


                If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.WriteMessage(vbLf & "Command:")
                    Freeze_operations = False
                    Editor1.SetImpliedSelection(Empty_array)
                    Exit Sub
                End If

                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                    If IsNothing(Rezultat1) = False Then

                        Using Lock_dwg As DocumentLock = ThisDrawing.LockDocument


                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                                Dim Ent1 As Entity = Trans1.GetObject(Rezultat1.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                Dim Poly1 As Polyline = Ent1
                                Poly_centerline = Poly1

                                Data_table_Centerline = New System.Data.DataTable
                                Data_table_Centerline.Columns.Add("X", GetType(Double))
                                Data_table_Centerline.Columns.Add("Y", GetType(Double))
                                Data_table_Match_rotatations = New System.Data.DataTable
                                Data_table_Match_rotatations.Columns.Add("MATCHLINE1", GetType(Double))
                                Data_table_Match_rotatations.Columns.Add("MATCHLINE2", GetType(Double))
                                Data_table_Match_rotatations.Columns.Add("TWIST", GetType(Double))
                                Data_table_Match_rotatations.Columns.Add("CENTER", GetType(Point3d))

                                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                                Dim LayerTable1 As Autodesk.AutoCAD.DatabaseServices.LayerTable
                                LayerTable1 = Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                Dim Index_dataT As Double = 0


                                For Each objID As ObjectId In BTrecord
                                    Dim Rectangle_poly As Entity = Trans1.GetObject(objID, OpenMode.ForRead)
                                    If TypeOf Rectangle_poly Is Polyline Then
                                        Dim Viewport_poly As Polyline = Rectangle_poly
                                        Dim LayerTableRecord1 As Autodesk.AutoCAD.DatabaseServices.LayerTableRecord
                                        LayerTableRecord1 = LayerTable1(Viewport_poly.Layer).GetObject(OpenMode.ForRead)


                                        If Viewport_poly.Closed = True And Viewport_poly.NumberOfVertices = 4 And LayerTableRecord1.IsOff = False And LayerTableRecord1.IsFrozen = False Then
                                            Dim Col_int As New Point3dCollection
                                            Col_int = Intersect_on_both_operands(Poly1, Viewport_poly)
                                            If Col_int.Count > 0 Then
                                                Dim Station1 As Double = 0
                                                Dim Station2 As Double = 0
                                                Dim Twist1 As Double = 0
                                                Dim Nr_values As Integer = Col_int.Count
                                                If Nr_values > 2 Then Nr_values = 2

                                                If Nr_values = 1 Then
                                                    Dim Point_on_poly = Col_int(0)
                                                    Station2 = Poly1.GetDistAtPoint(Point_on_poly)
                                                    Twist1 = GET_Bearing_rad(Viewport_poly.GetPoint2dAt(3).X, Viewport_poly.GetPoint2dAt(3).Y, Viewport_poly.GetPoint2dAt(2).X, Viewport_poly.GetPoint2dAt(2).Y)
                                                End If

                                                If Nr_values = 2 Then
                                                    Dim Point_on_poly = Col_int(0)
                                                    Station1 = Poly1.GetDistAtPoint(Point_on_poly)
                                                    Twist1 = GET_Bearing_rad(Viewport_poly.GetPoint2dAt(3).X, Viewport_poly.GetPoint2dAt(3).Y, Viewport_poly.GetPoint2dAt(2).X, Viewport_poly.GetPoint2dAt(2).Y)
                                                    Point_on_poly = Col_int(1)
                                                    Station2 = Poly1.GetDistAtPoint(Point_on_poly)
                                                End If

                                                If Station1 > Station2 Then
                                                    Dim T As Double = Station1
                                                    Station1 = Station2
                                                    Station2 = T
                                                End If


                                                Data_table_Match_rotatations.Rows.Add()
                                                Data_table_Match_rotatations.Rows(Index_dataT).Item("MATCHLINE1") = Round(Station1, 2)
                                                Data_table_Match_rotatations.Rows(Index_dataT).Item("MATCHLINE2") = Round(Station2, 2)
                                                Data_table_Match_rotatations.Rows(Index_dataT).Item("TWIST") = Twist1
                                                Data_table_Match_rotatations.Rows(Index_dataT).Item("CENTER") = _
                                                    New Point3d((Viewport_poly.GetPoint3dAt(0).X + Viewport_poly.GetPoint3dAt(2).X) / 2, _
                                                                (Viewport_poly.GetPoint3dAt(0).Y + Viewport_poly.GetPoint3dAt(2).Y) / 2, 0)

                                                Index_dataT = Index_dataT + 1


                                            End If
                                        End If

                                    End If


                                Next

                                If Data_table_Match_rotatations.Rows.Count > 0 Then
                                    For i = 0 To Data_table_Match_rotatations.Rows.Count - 1

                                        Dim Station1 As Double = 0
                                        Dim Station2 As Double = 0
                                        Dim Twist1 As Double = 0

                                        If IsDBNull(Data_table_Match_rotatations.Rows(i).Item("MATCHLINE1")) = False Then
                                            Station1 = Data_table_Match_rotatations.Rows(i).Item("MATCHLINE1")
                                        End If
                                        If IsDBNull(Data_table_Match_rotatations.Rows(i).Item("MATCHLINE2")) = False Then
                                            Station2 = Data_table_Match_rotatations.Rows(i).Item("MATCHLINE2")
                                        End If
                                        If IsDBNull(Data_table_Match_rotatations.Rows(i).Item("TWIST")) = False Then
                                            Twist1 = Data_table_Match_rotatations.Rows(i).Item("TWIST") * 180 / PI
                                        End If

                                        'MsgBox("From " & Station1 & vbCrLf & "To " & Station2 & vbCrLf & "TW = " & Twist1)

                                    Next
                                End If


                                For i = 0 To Poly1.NumberOfVertices - 1
                                    Data_table_Centerline.Rows.Add()
                                    Data_table_Centerline.Rows(i).Item("X") = Poly1.GetPoint3dAt(i).X
                                    Data_table_Centerline.Rows(i).Item("Y") = Poly1.GetPoint3dAt(i).Y

                                Next

                                Trans1.Commit()
                            End Using
                        End Using

                    End If
                End If



                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")

            Catch ex As Exception
                Freeze_operations = False
                Editor1.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If

    End Sub

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
                    Dim Station_back As String = W1.Range(Column_Sta_Back & i).Value2
                    Dim Station_ahead As String = W1.Range(Column_sta_ahead & i).Value2
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

    Public Function Get_length_from_equation_value(ByVal Station_with_eq As Double) As Double
        Dim Valoare As Double = 0
        If IsNothing(Data_table_station_equation) = False Then
            If Data_table_station_equation.Rows.Count > 0 Then
                For i = 0 To Data_table_station_equation.Rows.Count - 1
                    If IsDBNull(Data_table_station_equation.Rows(i).Item("STATION_BACK")) = False And IsDBNull(Data_table_station_equation.Rows(i).Item("STATION_AHEAD")) = False Then
                        Dim Station_back As Double = Data_table_station_equation.Rows(i).Item("STATION_BACK")
                        Dim Station_ahead As Double = Data_table_station_equation.Rows(i).Item("STATION_AHEAD")

                        If Station_with_eq < Station_ahead Then
                            Exit For
                        End If

                        Valoare = Valoare + Station_ahead - Station_back

                    End If
                Next
            End If


        End If


        Return Station_with_eq - Valoare
    End Function
End Class