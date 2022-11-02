Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Public Class Alignment_material_worksheet_form
    Dim Colectie1 As New Specialized.StringCollection
    Dim Extra_index_dupa_removal As Integer = 0
    Dim Data_table_Crossing As System.Data.DataTable
    Dim Data_table_Matchline As System.Data.DataTable
    Dim Data_table_Elbows As System.Data.DataTable
    Dim Data_table_Class_location As System.Data.DataTable
    Dim Data_table_Hydrotest As System.Data.DataTable
    Dim Data_table_Stress As System.Data.DataTable
    Dim Data_table_Buoyancy As System.Data.DataTable

    Dim Nr_pagina As Integer = 1
    Private Sub Alignment_material_worksheet_form_Load(sender As Object, e As EventArgs) Handles Me.Load
        Data_table_Crossing = New System.Data.DataTable
        Data_table_Crossing.Columns.Add("DESCRIPTION1", GetType(String))
        Data_table_Crossing.Columns.Add("DESCRIPTION2", GetType(String))
        Data_table_Crossing.Columns.Add("ID_NO", GetType(String))
        Data_table_Crossing.Columns.Add("STA", GetType(Double))
        Data_table_Crossing.Columns.Add("BEGINSTA", GetType(Double))
        Data_table_Crossing.Columns.Add("ENDSTA", GetType(Double))
        Data_table_Crossing.Columns.Add("LENGTH", GetType(Double))
        Data_table_Crossing.Columns.Add("MATERIAL", GetType(Integer))
        Data_table_Crossing.Columns.Add("PREVIOUS_MATERIAL", GetType(String))
        Data_table_Crossing.Columns.Add("COVER", GetType(String))
        Data_table_Crossing.Columns.Add("CROSSING_TYPE", GetType(String))
        Data_table_Crossing.Columns.Add("BLOCKNAME", GetType(String))
        Data_table_Crossing.Columns.Add("SHEET", GetType(Integer))
        Data_table_Crossing.Columns.Add("WT", GetType(String))
        Data_table_Crossing.Columns.Add("CLASS", GetType(Integer))
        Data_table_Crossing.Columns.Add("BUOYANCYTYPE", GetType(String))

        Data_table_Matchline = New System.Data.DataTable
        Data_table_Matchline.Columns.Add("STATION", GetType(Double))
        Data_table_Matchline.Columns.Add("PAGINA", GetType(Integer))

        Data_table_Elbows = New System.Data.DataTable
        Data_table_Elbows.Columns.Add("STATION", GetType(Double))
        Data_table_Elbows.Columns.Add("ANGLE", GetType(String))
        Data_table_Elbows.Columns.Add("MATCHED", GetType(Boolean))
        Data_table_Elbows.Columns.Add("LENGTH", GetType(Double))
        Data_table_Elbows.Columns.Add("VALUE", GetType(Double))

        Data_table_Class_location = New System.Data.DataTable
        Data_table_Class_location.Columns.Add("CLASS", GetType(Integer))
        Data_table_Class_location.Columns.Add("BEGINSTA", GetType(Double))
        Data_table_Class_location.Columns.Add("ENDSTA", GetType(Double))
        Data_table_Class_location.Columns.Add("WT", GetType(String))
        Data_table_Class_location.Columns.Add("MATERIAL", GetType(Integer))

        Data_table_Hydrotest = New System.Data.DataTable
        Data_table_Hydrotest.Columns.Add("MATERIAL", GetType(Integer))
        Data_table_Hydrotest.Columns.Add("BEGINSTA", GetType(Double))
        Data_table_Hydrotest.Columns.Add("ENDSTA", GetType(Double))

        Data_table_Stress = New System.Data.DataTable
        Data_table_Stress.Columns.Add("WT", GetType(String))
        Data_table_Stress.Columns.Add("BEGINSTA", GetType(Double))
        Data_table_Stress.Columns.Add("ENDSTA", GetType(Double))

        Data_table_Buoyancy = New System.Data.DataTable
        Data_table_Buoyancy.Columns.Add("BEGINSTA", GetType(Double))
        Data_table_Buoyancy.Columns.Add("ENDSTA", GetType(Double))
        Data_table_Buoyancy.Columns.Add("TYPE", GetType(String))
        Data_table_Buoyancy.Columns.Add("SPACING", GetType(Double))

        ComboBox_load_type.SelectedIndex = 0
        Label_col_1.Visible = True
        Label_col_2.Visible = False
        Label_col_3.Visible = False
        Label_col_4.Visible = False
        Label_col_5.Visible = False

        TextBox_col_1.Visible = True
        TextBox_col_2.Visible = False
        TextBox_col_3.Visible = False
        TextBox_col_4.Visible = False
        TextBox_col_5.Visible = False


    End Sub

    Private Sub ComboBox_load_type_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_load_type.SelectedIndexChanged

        TextBox_col_1.BackColor = Drawing.Color.White
        TextBox_col_2.BackColor = Drawing.Color.White
        TextBox_col_3.BackColor = Drawing.Color.White
        TextBox_col_4.BackColor = Drawing.Color.White
        TextBox_col_5.BackColor = Drawing.Color.White

        Select Case ComboBox_load_type.SelectedIndex
            Case 0
                Label_col_1.Visible = True
                Label_col_1.Text = "Match Lines" & vbCrLf & "Column"
                Label_col_2.Visible = False
                Label_col_3.Visible = False
                Label_col_4.Visible = False
                Label_col_5.Visible = False

                TextBox_col_1.Visible = True
                TextBox_col_1.Text = "A"
                TextBox_col_2.Visible = False
                TextBox_col_3.Visible = False
                TextBox_col_4.Visible = False
                TextBox_col_5.Visible = False

            Case 1
                Label_col_1.Visible = True
                Label_col_1.Text = "Class"
                Label_col_2.Visible = True
                Label_col_2.Text = "Crossing" & vbCrLf & "Type"
                Label_col_3.Visible = True
                Label_col_3.Text = "Material"
                Label_col_4.Visible = True
                Label_col_4.Text = "Thickness"
                Label_col_5.Visible = False

                TextBox_col_1.Visible = True
                TextBox_col_1.Text = "A"
                TextBox_col_2.Visible = True
                TextBox_col_2.Text = "B"
                TextBox_col_3.Visible = True
                TextBox_col_3.Text = "C"
                TextBox_col_4.Visible = True
                TextBox_col_4.Text = "D"
                TextBox_col_5.Visible = False

            Case 2
                Label_col_1.Visible = True
                Label_col_1.Text = "Column" & vbCrLf & "Station"
                Label_col_2.Visible = True
                Label_col_2.Text = "Column" & vbCrLf & "Value"
                Label_col_3.Visible = False
                Label_col_4.Visible = False
                Label_col_5.Visible = False

                TextBox_col_1.Visible = True
                TextBox_col_1.Text = "A"
                TextBox_col_2.Visible = True
                TextBox_col_2.Text = "B"
                TextBox_col_3.Visible = False
                TextBox_col_4.Visible = False
                TextBox_col_5.Visible = False

            Case 3
                Label_col_1.Visible = True
                Label_col_1.Text = "Class"
                Label_col_2.Visible = True
                Label_col_2.Text = "Begin" & vbCrLf & "Station"
                Label_col_3.Visible = True
                Label_col_3.Text = "End" & vbCrLf & "Station"
                Label_col_4.Visible = True
                Label_col_4.Text = "Thickness"
                Label_col_5.Visible = True
                Label_col_5.Text = "Material"

                TextBox_col_1.Visible = True
                TextBox_col_1.Text = "A"
                TextBox_col_2.Visible = True
                TextBox_col_2.Text = "B"
                TextBox_col_3.Visible = True
                TextBox_col_3.Text = "C"
                TextBox_col_4.Visible = True
                TextBox_col_4.Text = "E"
                TextBox_col_5.Visible = True
                TextBox_col_5.Text = "F"

            Case 4
                Label_col_1.Visible = True
                Label_col_1.Text = "Material"
                Label_col_2.Visible = True
                Label_col_2.Text = "Begin" & vbCrLf & "Station"
                Label_col_3.Visible = True
                Label_col_3.Text = "End" & vbCrLf & "Station"
                Label_col_4.Visible = False
                Label_col_5.Visible = False

                TextBox_col_1.Visible = True
                TextBox_col_1.Text = "B"
                TextBox_col_2.Visible = True
                TextBox_col_2.Text = "D"
                TextBox_col_3.Visible = True
                TextBox_col_3.Text = "F"
                TextBox_col_4.Visible = False
                TextBox_col_5.Visible = False

            Case 5
                Label_col_1.Visible = True
                Label_col_1.Text = "Begin" & vbCrLf & "Station"
                Label_col_2.Visible = True
                Label_col_2.Text = "End" & vbCrLf & "Station"
                Label_col_3.Visible = True
                Label_col_3.Text = "Thickness"
                Label_col_4.Visible = False
                Label_col_5.Visible = False

                TextBox_col_1.Visible = True
                TextBox_col_1.Text = "P"
                TextBox_col_2.Visible = True
                TextBox_col_2.Text = "Q"
                TextBox_col_3.Visible = True
                TextBox_col_3.Text = "S"
                TextBox_col_4.Visible = False
                TextBox_col_5.Visible = False

            Case 6
                Label_col_1.Visible = True
                Label_col_1.Text = "Item" & vbCrLf & "Type"
                Label_col_2.Visible = True
                Label_col_2.Text = "Spacing"
                Label_col_3.Visible = True
                Label_col_3.Text = "Begin" & vbCrLf & "Station"
                Label_col_4.Visible = True
                Label_col_4.Text = "End" & vbCrLf & "Station"
                Label_col_5.Visible = False

                TextBox_col_1.Visible = True
                TextBox_col_1.Text = "A"
                TextBox_col_2.Visible = True
                TextBox_col_2.Text = "D"
                TextBox_col_3.Visible = True
                TextBox_col_3.Text = "H"
                TextBox_col_4.Visible = True
                TextBox_col_4.Text = "I"
                TextBox_col_5.Visible = False

        End Select
    End Sub

    Private Sub Button_Read_excel_Click(sender As Object, e As EventArgs) Handles Button_Read_excel.Click
        Try

            If TextBox_description1_COL_XL.Text = "" Then
                MsgBox("Please specify the description 1 EXCEL COLUMN!")
                Exit Sub
            End If

            If TextBox_description2_COL_XL.Text = "" Then
                MsgBox("Please specify the description 2 EXCEL COLUMN!")
                Exit Sub
            End If

            If TextBox_ID_NO_COL_XL.Text = "" Then
                MsgBox("Please specify the ID_NO EXCEL COLUMN!")
                Exit Sub
            End If

            If TextBox_chainage_col_xl.Text = "" Then
                MsgBox("Please specify the station EXCEL COLUMN!")
                Exit Sub
            End If

            If TextBox_LENGTH_col_xl.Text = "" Then
                MsgBox("Please specify the length EXCEL COLUMN!")
                Exit Sub
            End If

            If TextBox_MATERIAL_col_xl.Text = "" Then
                MsgBox("Please specify the material EXCEL COLUMN!")
                Exit Sub
            End If

            If TextBox_Cover_col_xl.Text = "" Then
                MsgBox("Please specify the cover EXCEL COLUMN!")
                Exit Sub
            End If

            If TextBox_Crossing_type_col_XL.Text = "" Then
                MsgBox("Please specify the crossing type EXCEL COLUMN!")
                Exit Sub
            End If

            If TextBox_ROW_START.Text = "" Then
                MsgBox("Please specify the EXCEL START ROW!")
                Exit Sub
            End If

            If IsNumeric(TextBox_ROW_START.Text) = False Then
                With TextBox_ROW_START
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify start row")

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

                Exit Sub
            End If

            If Val(TextBox_ROW_START.Text) < 1 Then
                With TextBox_ROW_START
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Start row can't be smaller than 1")

                Exit Sub
            End If

            If Val(TextBox_ROW_END.Text) < 1 Then
                With TextBox_ROW_END
                    .Text = ""
                    .Focus()
                End With
                MsgBox("End row can't be smaller than 1")

                Exit Sub
            End If

            If Val(TextBox_ROW_END.Text) < Val(TextBox_ROW_START.Text) Then
                MsgBox("END row smaller than start row")

                Exit Sub
            End If
            TextBox_col_1.BackColor = Drawing.Color.White
            TextBox_col_2.BackColor = Drawing.Color.White
            TextBox_col_3.BackColor = Drawing.Color.White
            TextBox_col_4.BackColor = Drawing.Color.White
            TextBox_col_5.BackColor = Drawing.Color.White

            ascunde_butoanele_pentru_forms(Me, Colectie1)

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
            Panel_crossings.Controls.Clear()

            Extra_index_dupa_removal = 0
            Data_table_Crossing.Rows.Clear()
            Dim Index_Data_table As Integer = 0


            Dim Chainage_previous As Double = 0
            Panel_crossings.Controls.Clear()
            For i = start1 To end1
                Dim Description1 As String = W1.Range(Col1 & i).Value
                Dim Description2 As String = W1.Range(Col2 & i).Value
                Dim Id_No As String = W1.Range(Col3 & i).Value


                If Not W1.Range(Col1 & i).Interior.Color = 192 Then
                    Dim Chainage As Double = -1
                    If IsNumeric(Replace(W1.Range(Col4 & i).Value, "+", "")) = True Then
                        Chainage = Round(CDbl(Replace(W1.Range(Col4 & i).Value, "+", "")), 1)
                    End If

                    Dim Length As Double = -1
                    If IsNumeric(W1.Range(Col5 & i).Value) = True Then
                        Length = Round(W1.Range(Col5 & i).Value, 1)
                    End If
                    Dim Material As String = ""
                    If Not Replace(W1.Range(Col6 & i).Value, " ", "") = "" Then
                        Material = W1.Range(Col6 & i).Value
                    End If


                    Dim Crossing_type As String = W1.Range(Col8 & i).Value

                    If IsNothing(Description1) = True Then Description1 = ""
                    If IsNothing(Description2) = True Then Description2 = ""
                    If IsNothing(Id_No) = True Then Id_No = ""
                    If IsNothing(Crossing_type) = True Then Crossing_type = ""




                    If Not Replace(Description1, " ", "") = "" Then
                        Data_table_Crossing.Rows.Add()
                        Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = Description1

                        If Chainage = -1 Then
                            W1.Range(Col4 & i).Select()
                            MsgBox("Non numerical value at " & Col4 & i)
                            afiseaza_butoanele_pentru_forms(Me, Colectie1)
                            Exit Sub
                        Else
                            If Chainage < Chainage_previous Then
                                W1.Range(Col4 & i).Select()
                                MsgBox("The previous station is bigger than current station" & Col4 & i)
                                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                Exit Sub
                            End If
                            Chainage_previous = Chainage
                        End If

                        If Not Replace(Description2, " ", "") = "" Then
                            Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION2") = Description2
                        End If

                        If Not Replace(Id_No, " ", "") = "" Then
                            Data_table_Crossing.Rows(Index_Data_table).Item("ID_NO") = Id_No
                        End If

                        If Not Replace(Material, " ", "") = "" Then
                            Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = Material

                        End If

                        If Not Replace(Crossing_type, " ", "") = "" Then
                            Data_table_Crossing.Rows(Index_Data_table).Item("CROSSING_TYPE") = Crossing_type
                        End If

                        If Not Chainage = -1 Then
                            Data_table_Crossing.Rows(Index_Data_table).Item("STA") = Chainage
                        End If

                        If Crossing_type.ToUpper = "RD" Or Crossing_type.ToUpper = "WC" Or Crossing_type.ToUpper = "RR" Or Crossing_type.ToUpper = "TST" Then
                            Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "Mat_worksheet_block1"
                        End If


                        Index_Data_table = Index_Data_table + 1

                    End If ' If Not Replace(Description1, " ", "") = ""
                End If 'E DE LA If Not W1.Range(Col1 & i).Interior.Color = 192
            Next

            adauga_ELBOWS_la_crossing_list()
            adauga_CLASS_LOCATION_la_crossing_list()
            adauga_Hydrotest_la_crossing_list()
            adauga_Stress_la_crossing_list()
            adauga_Buoyancy_la_crossing_list()
            adauga_matchline_la_crossing_list()

            afiseaza_butoanele_pentru_forms(Me, Colectie1)
        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_load_items_from_combobox_Click(sender As Object, e As EventArgs) Handles Button_load_items_from_combobox.Click
        Try
            Select Case ComboBox_load_type.SelectedIndex
                Case 0 ' MATCHLINES
                    If TextBox_col_1.Text = "" Then
                        MsgBox("Please specify the MATCHLINE EXCEL COLUMN!")
                        Exit Sub
                    End If

                    If TextBox_items_row_start.Text = "" Then
                        MsgBox("Please specify the EXCEL START ROW!")
                        Exit Sub
                    End If

                    If IsNumeric(TextBox_items_row_start.Text) = False Then
                        With TextBox_items_row_start
                            .Text = ""
                            .Focus()
                        End With
                        MsgBox("Please specify start row")
                        Exit Sub
                    End If

                    If TextBox_items_row_end.Text = "" Then
                        MsgBox("Please specify the EXCEL END ROW!")
                        Exit Sub
                    End If

                    If IsNumeric(TextBox_items_row_end.Text) = False Then
                        With TextBox_items_row_end
                            .Text = ""
                            .Focus()
                        End With
                        MsgBox("Please specify END row")
                        Exit Sub
                    End If

                    If Val(TextBox_items_row_start.Text) < 1 Then
                        With TextBox_items_row_start
                            .Text = ""
                            .Focus()
                        End With
                        MsgBox("Start row can't be smaller than 1")

                        Exit Sub
                    End If

                    If Val(TextBox_items_row_end.Text) < 1 Then
                        With TextBox_items_row_end
                            .Text = ""
                            .Focus()
                        End With
                        MsgBox("End row can't be smaller than 1")

                        Exit Sub
                    End If

                    If Val(TextBox_items_row_end.Text) < Val(TextBox_items_row_start.Text) Then
                        MsgBox("END row smaller than start row")

                        Exit Sub
                    End If

                    ascunde_butoanele_pentru_forms(Me, Colectie1)

                    Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
                    Dim start1 As Integer = CInt(TextBox_items_row_start.Text)
                    Dim end1 As Integer = CInt(TextBox_items_row_end.Text)

                    Dim Col1 As String = TextBox_col_1.Text.ToUpper

                    Data_table_Matchline.Rows.Clear()
                    Dim Index_Data_table As Integer = 0
                    Dim Nr_pag As Integer = 1
                    TextBox_col_1.BackColor = Drawing.Color.White

                    For i = start1 To end1
                        Dim Chainage As Double = Val(W1.Range(Col1 & i).Value)


                        If IsNothing(Chainage) = True Then Chainage = 0

                        If i = start1 Then
                            If Not Chainage = 0 Then
                                Data_table_Matchline.Rows.Add()
                                Data_table_Matchline.Rows(Index_Data_table).Item("STATION") = 0
                                Data_table_Matchline.Rows(Index_Data_table).Item("PAGINA") = Nr_pag
                                'MsgBox(Nr_pag & " - " & Data_table_Matchline.Rows(Index_Data_table).Item("STATION"))
                                Index_Data_table = Index_Data_table + 1
                                Nr_pag = Nr_pag + 1
                            End If
                        End If

                        Data_table_Matchline.Rows.Add()
                        Data_table_Matchline.Rows(Index_Data_table).Item("STATION") = Chainage
                        Data_table_Matchline.Rows(Index_Data_table).Item("PAGINA") = Nr_pag
                        'MsgBox(Nr_pag & " - " & Data_table_Matchline.Rows(Index_Data_table).Item("STATION"))
                        TextBox_col_1.BackColor = Drawing.Color.Yellow
                        Index_Data_table = Index_Data_table + 1
                        Nr_pag = Nr_pag + 1
                    Next

                    adauga_matchline_la_crossing_list()
                    TextBox_col_1.BackColor = Drawing.Color.Yellow

                Case 2 'ELBOWS

                    TextBox_col_1.BackColor = Drawing.Color.White
                    TextBox_col_2.BackColor = Drawing.Color.White

                    If TextBox_col_1.Text = "" Then
                        MsgBox("Please specify the Deflection station Excel column!")
                        Exit Sub
                    End If

                    If TextBox_col_2.Text = "" Then
                        MsgBox("Please specify the Deflection value Excel column!")
                        Exit Sub
                    End If


                    If TextBox_items_row_start.Text = "" Then
                        MsgBox("Please specify the EXCEL START ROW!")
                        Exit Sub
                    End If

                    If IsNumeric(TextBox_items_row_start.Text) = False Then
                        With TextBox_items_row_start
                            .Text = ""
                            .Focus()
                        End With
                        MsgBox("Please specify start row")

                        Exit Sub
                    End If
                    If TextBox_items_row_end.Text = "" Then
                        MsgBox("Please specify the EXCEL END ROW!")
                        Exit Sub
                    End If

                    If IsNumeric(TextBox_items_row_end.Text) = False Then
                        With TextBox_items_row_end
                            .Text = ""
                            .Focus()
                        End With
                        MsgBox("Please specify END row")

                        Exit Sub
                    End If

                    If Val(TextBox_items_row_start.Text) < 1 Then
                        With TextBox_items_row_start
                            .Text = ""
                            .Focus()
                        End With
                        MsgBox("Start row can't be smaller than 1")

                        Exit Sub
                    End If

                    If Val(TextBox_items_row_end.Text) < 1 Then
                        With TextBox_items_row_end
                            .Text = ""
                            .Focus()
                        End With
                        MsgBox("End row can't be smaller than 1")

                        Exit Sub
                    End If

                    If Val(TextBox_items_row_end.Text) < Val(TextBox_items_row_start.Text) Then
                        MsgBox("END row smaller than start row")

                        Exit Sub
                    End If

                    ascunde_butoanele_pentru_forms(Me, Colectie1)

                    Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
                    Dim start1 As Integer = CInt(TextBox_items_row_start.Text)
                    Dim end1 As Integer = CInt(TextBox_items_row_end.Text)

                    Dim Col1 As String = TextBox_col_1.Text.ToUpper
                    Dim Col2 As String = TextBox_col_2.Text.ToUpper

                    Data_table_Elbows.Rows.Clear()
                    Dim Index_Data_table As Integer = 0

                    For i = start1 To end1
                        Dim Chainage1 As String = Replace(W1.Range(Col1 & i).Value, "+", "")
                        Dim Angle As String = W1.Range(Col2 & i).Value
                        Dim CHAINAGE As Double = 0
                        If IsNumeric(Chainage1) = True Then
                            CHAINAGE = CDbl(Chainage1)
                        End If
                        If CHAINAGE > 0 Then
                            Data_table_Elbows.Rows.Add()
                            Data_table_Elbows.Rows(Index_Data_table).Item("STATION") = CHAINAGE
                            Data_table_Elbows.Rows(Index_Data_table).Item("ANGLE") = Angle
                            'Data_table_Elbows.Rows(Index_Data_table).Item("MATCHED") = True
                            If IsNumeric(ComboBox_nps.Text) = True And IsNumeric(Angle) = True Then
                                Dim Diam As Double = 2 * Get_from_NPS_radius_for_pipes_from_inches_to_milimeters(CDbl(ComboBox_nps.Text)) / 1000
                                Data_table_Elbows.Rows(Index_Data_table).Item("LENGTH") = Round(2 * (3 * Diam * Tan((CDbl(Angle) * PI / 180) / 2) + 1), 1)
                                Data_table_Elbows.Rows(Index_Data_table).Item("VALUE") = Angle
                            Else
                                MsgBox("At " & Get_chainage_from_double(CHAINAGE, 1) & "you have an issue - no Pipe NPS or no angle specified for elbow")
                                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                Exit Sub
                            End If
                            Index_Data_table = Index_Data_table + 1
                        End If
                    Next

                    adauga_ELBOWS_la_crossing_list()

                    TextBox_col_1.BackColor = Drawing.Color.Yellow
                    TextBox_col_2.BackColor = Drawing.Color.Yellow

                Case 3 'CLASS LOCATION

                    TextBox_col_1.BackColor = Drawing.Color.White
                    TextBox_col_2.BackColor = Drawing.Color.White
                    TextBox_col_3.BackColor = Drawing.Color.White
                    TextBox_col_4.BackColor = Drawing.Color.White
                    TextBox_col_5.BackColor = Drawing.Color.White

                    If TextBox_col_1.Text = "" Then
                        MsgBox("Please specify the Class Excel column!")
                        Exit Sub
                    End If

                    If TextBox_col_2.Text = "" Then
                        MsgBox("Please specify the Begin Station Excel column!")
                        Exit Sub
                    End If

                    If TextBox_col_3.Text = "" Then
                        MsgBox("Please specify the End Station Excel column!")
                        Exit Sub
                    End If
                    If TextBox_col_4.Text = "" Then
                        MsgBox("Please specify the Thickness Excel column!")
                        Exit Sub
                    End If

                    If TextBox_col_5.Text = "" Then
                        MsgBox("Please specify the Material Excel column!")
                        Exit Sub
                    End If

                    If TextBox_items_row_start.Text = "" Then
                        MsgBox("Please specify the EXCEL START ROW!")
                        Exit Sub
                    End If

                    If IsNumeric(TextBox_items_row_start.Text) = False Then
                        With TextBox_items_row_start
                            .Text = ""
                            .Focus()
                        End With
                        MsgBox("Please specify start row")

                        Exit Sub
                    End If
                    If TextBox_items_row_end.Text = "" Then
                        MsgBox("Please specify the EXCEL END ROW!")
                        Exit Sub
                    End If

                    If IsNumeric(TextBox_items_row_end.Text) = False Then
                        With TextBox_items_row_end
                            .Text = ""
                            .Focus()
                        End With
                        MsgBox("Please specify END row")

                        Exit Sub
                    End If

                    If Val(TextBox_items_row_start.Text) < 1 Then
                        With TextBox_items_row_start
                            .Text = ""
                            .Focus()
                        End With
                        MsgBox("Start row can't be smaller than 1")

                        Exit Sub
                    End If

                    If Val(TextBox_items_row_end.Text) < 1 Then
                        With TextBox_items_row_end
                            .Text = ""
                            .Focus()
                        End With
                        MsgBox("End row can't be smaller than 1")

                        Exit Sub
                    End If

                    If Val(TextBox_items_row_end.Text) < Val(TextBox_items_row_start.Text) Then
                        MsgBox("END row smaller than start row")

                        Exit Sub
                    End If

                    ascunde_butoanele_pentru_forms(Me, Colectie1)

                    Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
                    Dim start1 As Integer = CInt(TextBox_items_row_start.Text)
                    Dim end1 As Integer = CInt(TextBox_items_row_end.Text)

                    Dim Col1 As String = TextBox_col_1.Text.ToUpper
                    Dim Col2 As String = TextBox_col_2.Text.ToUpper
                    Dim Col3 As String = TextBox_col_3.Text.ToUpper
                    Dim Col4 As String = TextBox_col_4.Text.ToUpper
                    Dim Col5 As String = TextBox_col_5.Text.ToUpper

                    Data_table_Class_location.Rows.Clear()

                    Dim Index_Data_table As Integer = 0

                    For i = start1 To end1
                        Dim Class1 As Integer = W1.Range(Col1 & i).Value
                        Dim BS As String = W1.Range(Col2 & i).Value
                        Dim BeginSTATION As Double
                        If IsNumeric(Replace(BS, "+", "")) = True Then
                            BeginSTATION = CDbl((Replace(BS, "+", "")))
                        End If
                        Dim ES As String = W1.Range(Col3 & i).Value
                        Dim EndSTATION As Double
                        If IsNumeric(Replace(ES, "+", "")) = True Then
                            EndSTATION = CDbl((Replace(ES, "+", "")))
                        End If

                        Dim WallThickness As String = W1.Range(Col4 & i).Value
                        Dim mat As String = W1.Range(Col5 & i).Value
                        Dim Mat1 As Integer
                        If IsNumeric(mat) = True Then
                            Mat1 = CInt(mat)
                        End If

                        If Mat1 > 0 And BeginSTATION >= 0 And EndSTATION > 0 And Class1 > 0 Then
                            Data_table_Class_location.Rows.Add()
                            Data_table_Class_location.Rows(Index_Data_table).Item("CLASS") = Class1
                            Data_table_Class_location.Rows(Index_Data_table).Item("BEGINSTA") = BeginSTATION
                            Data_table_Class_location.Rows(Index_Data_table).Item("ENDSTA") = EndSTATION
                            Data_table_Class_location.Rows(Index_Data_table).Item("WT") = WallThickness
                            Data_table_Class_location.Rows(Index_Data_table).Item("MATERIAL") = Mat1
                            Index_Data_table = Index_Data_table + 1
                        End If
                    Next

                    adauga_CLASS_LOCATION_la_crossing_list()

                    TextBox_col_1.BackColor = Drawing.Color.Yellow
                    TextBox_col_2.BackColor = Drawing.Color.Yellow
                    TextBox_col_3.BackColor = Drawing.Color.Yellow
                    TextBox_col_4.BackColor = Drawing.Color.Yellow
                    TextBox_col_5.BackColor = Drawing.Color.Yellow

                Case 4 'HYDROTEST

                    TextBox_col_1.BackColor = Drawing.Color.White
                    TextBox_col_2.BackColor = Drawing.Color.White
                    TextBox_col_3.BackColor = Drawing.Color.White
                    TextBox_col_4.BackColor = Drawing.Color.White
                    TextBox_col_5.BackColor = Drawing.Color.White

                    If TextBox_col_1.Text = "" Then
                        MsgBox("Please specify the Material Excel column!")
                        Exit Sub
                    End If

                    If TextBox_col_2.Text = "" Then
                        MsgBox("Please specify the Begin Station Excel column!")
                        Exit Sub
                    End If

                    If TextBox_col_3.Text = "" Then
                        MsgBox("Please specify the End Station Excel column!")
                        Exit Sub
                    End If


                    If TextBox_items_row_start.Text = "" Then
                        MsgBox("Please specify the EXCEL START ROW!")
                        Exit Sub
                    End If

                    If IsNumeric(TextBox_items_row_start.Text) = False Then
                        With TextBox_items_row_start
                            .Text = ""
                            .Focus()
                        End With
                        MsgBox("Please specify start row")

                        Exit Sub
                    End If
                    If TextBox_items_row_end.Text = "" Then
                        MsgBox("Please specify the EXCEL END ROW!")
                        Exit Sub
                    End If

                    If IsNumeric(TextBox_items_row_end.Text) = False Then
                        With TextBox_items_row_end
                            .Text = ""
                            .Focus()
                        End With
                        MsgBox("Please specify END row")

                        Exit Sub
                    End If

                    If Val(TextBox_items_row_start.Text) < 1 Then
                        With TextBox_items_row_start
                            .Text = ""
                            .Focus()
                        End With
                        MsgBox("Start row can't be smaller than 1")

                        Exit Sub
                    End If

                    If Val(TextBox_items_row_end.Text) < 1 Then
                        With TextBox_items_row_end
                            .Text = ""
                            .Focus()
                        End With
                        MsgBox("End row can't be smaller than 1")

                        Exit Sub
                    End If

                    If Val(TextBox_items_row_end.Text) < Val(TextBox_items_row_start.Text) Then
                        MsgBox("END row smaller than start row")

                        Exit Sub
                    End If

                    ascunde_butoanele_pentru_forms(Me, Colectie1)

                    Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
                    Dim start1 As Integer = CInt(TextBox_items_row_start.Text)
                    Dim end1 As Integer = CInt(TextBox_items_row_end.Text)

                    Dim Col1 As String = TextBox_col_1.Text.ToUpper
                    Dim Col2 As String = TextBox_col_2.Text.ToUpper
                    Dim Col3 As String = TextBox_col_3.Text.ToUpper


                    Data_table_Hydrotest.Rows.Clear()

                    Dim Index_Data_table As Integer = 0

                    For i = start1 To end1

                        Dim mat As String = W1.Range(Col1 & i).Value
                        Dim Mat1 As Integer
                        If IsNumeric(mat) = True Then
                            Mat1 = CInt(mat)
                        End If

                        Dim BS As String = W1.Range(Col2 & i).Value
                        Dim BeginSTATION As Double
                        If IsNumeric(Replace(BS, "+", "")) = True Then
                            BeginSTATION = CDbl((Replace(BS, "+", "")))
                        End If
                        Dim ES As String = W1.Range(Col3 & i).Value
                        Dim EndSTATION As Double
                        If IsNumeric(Replace(ES, "+", "")) = True Then
                            EndSTATION = CDbl((Replace(ES, "+", "")))
                        End If


                        If Mat1 > 0 And BeginSTATION >= 0 And EndSTATION > 0 Then
                            Data_table_Hydrotest.Rows.Add()

                            Data_table_Hydrotest.Rows(Index_Data_table).Item("BEGINSTA") = BeginSTATION
                            Data_table_Hydrotest.Rows(Index_Data_table).Item("ENDSTA") = EndSTATION

                            Data_table_Hydrotest.Rows(Index_Data_table).Item("MATERIAL") = Mat1
                            Index_Data_table = Index_Data_table + 1
                        End If
                    Next

                    adauga_Hydrotest_la_crossing_list()

                    TextBox_col_1.BackColor = Drawing.Color.Yellow
                    TextBox_col_2.BackColor = Drawing.Color.Yellow
                    TextBox_col_3.BackColor = Drawing.Color.Yellow

                Case 5 'STRESS

                    TextBox_col_1.BackColor = Drawing.Color.White
                    TextBox_col_2.BackColor = Drawing.Color.White
                    TextBox_col_3.BackColor = Drawing.Color.White
                    TextBox_col_4.BackColor = Drawing.Color.White
                    TextBox_col_5.BackColor = Drawing.Color.White

                    If TextBox_col_1.Text = "" Then
                        MsgBox("Please specify the Thickness Excel column!")
                        Exit Sub
                    End If

                    If TextBox_col_2.Text = "" Then
                        MsgBox("Please specify the Begin Station Excel column!")
                        Exit Sub
                    End If

                    If TextBox_col_3.Text = "" Then
                        MsgBox("Please specify the End Station Excel column!")
                        Exit Sub
                    End If


                    If TextBox_items_row_start.Text = "" Then
                        MsgBox("Please specify the EXCEL START ROW!")
                        Exit Sub
                    End If

                    If IsNumeric(TextBox_items_row_start.Text) = False Then
                        With TextBox_items_row_start
                            .Text = ""
                            .Focus()
                        End With
                        MsgBox("Please specify start row")

                        Exit Sub
                    End If
                    If TextBox_items_row_end.Text = "" Then
                        MsgBox("Please specify the EXCEL END ROW!")
                        Exit Sub
                    End If

                    If IsNumeric(TextBox_items_row_end.Text) = False Then
                        With TextBox_items_row_end
                            .Text = ""
                            .Focus()
                        End With
                        MsgBox("Please specify END row")

                        Exit Sub
                    End If

                    If Val(TextBox_items_row_start.Text) < 1 Then
                        With TextBox_items_row_start
                            .Text = ""
                            .Focus()
                        End With
                        MsgBox("Start row can't be smaller than 1")

                        Exit Sub
                    End If

                    If Val(TextBox_items_row_end.Text) < 1 Then
                        With TextBox_items_row_end
                            .Text = ""
                            .Focus()
                        End With
                        MsgBox("End row can't be smaller than 1")

                        Exit Sub
                    End If

                    If Val(TextBox_items_row_end.Text) < Val(TextBox_items_row_start.Text) Then
                        MsgBox("END row smaller than start row")

                        Exit Sub
                    End If

                    ascunde_butoanele_pentru_forms(Me, Colectie1)

                    Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
                    Dim start1 As Integer = CInt(TextBox_items_row_start.Text)
                    Dim end1 As Integer = CInt(TextBox_items_row_end.Text)

                    Dim Col1 As String = TextBox_col_1.Text.ToUpper
                    Dim Col2 As String = TextBox_col_2.Text.ToUpper
                    Dim Col3 As String = TextBox_col_3.Text.ToUpper


                    Data_table_Stress.Rows.Clear()

                    Dim Index_Data_table As Integer = 0

                    For i = start1 To end1

                        Dim Thickness As String
                        Thickness = W1.Range(Col3 & i).Value
                        If Not Replace(Thickness, " ", "") = "" Then
                            Dim NUMAR As String = extrage_numar_din_text_de_la_inceputul_textului(Thickness)
                            If Val(NUMAR) > 0 Then
                                Dim BS As String = W1.Range(Col1 & i).Value
                                Dim BeginSTATION As Double
                                If IsNumeric(Replace(BS, "+", "")) = True Then
                                    BeginSTATION = CDbl((Replace(BS, "+", "")))
                                End If
                                Dim ES As String = W1.Range(Col2 & i).Value
                                Dim EndSTATION As Double
                                If IsNumeric(Replace(ES, "+", "")) = True Then
                                    EndSTATION = CDbl((Replace(ES, "+", "")))
                                End If


                                If BeginSTATION >= 0 And EndSTATION > 0 Then
                                    Data_table_Stress.Rows.Add()

                                    Data_table_Stress.Rows(Index_Data_table).Item("BEGINSTA") = BeginSTATION
                                    Data_table_Stress.Rows(Index_Data_table).Item("ENDSTA") = EndSTATION

                                    Data_table_Stress.Rows(Index_Data_table).Item("WT") = Thickness
                                    Index_Data_table = Index_Data_table + 1
                                End If
                            End If
                        End If

                    Next

                    adauga_Stress_la_crossing_list()

                    TextBox_col_1.BackColor = Drawing.Color.Yellow
                    TextBox_col_2.BackColor = Drawing.Color.Yellow
                    TextBox_col_3.BackColor = Drawing.Color.Yellow

                Case 6 'BUOYANCY

                    TextBox_col_1.BackColor = Drawing.Color.White
                    TextBox_col_2.BackColor = Drawing.Color.White
                    TextBox_col_3.BackColor = Drawing.Color.White
                    TextBox_col_4.BackColor = Drawing.Color.White
                    TextBox_col_5.BackColor = Drawing.Color.White

                    If TextBox_col_1.Text = "" Then
                        MsgBox("Please specify the Buoyancy type Excel column!")
                        Exit Sub
                    End If

                    If TextBox_col_2.Text = "" Then
                        MsgBox("Please specify the Spacing Excel column!")
                        Exit Sub
                    End If

                    If TextBox_col_3.Text = "" Then
                        MsgBox("Please specify the Begin Station Excel column!")
                        Exit Sub
                    End If

                    If TextBox_col_4.Text = "" Then
                        MsgBox("Please specify the End Station Excel column!")
                        Exit Sub
                    End If


                    If TextBox_items_row_start.Text = "" Then
                        MsgBox("Please specify the EXCEL START ROW!")
                        Exit Sub
                    End If

                    If IsNumeric(TextBox_items_row_start.Text) = False Then
                        With TextBox_items_row_start
                            .Text = ""
                            .Focus()
                        End With
                        MsgBox("Please specify start row")

                        Exit Sub
                    End If
                    If TextBox_items_row_end.Text = "" Then
                        MsgBox("Please specify the EXCEL END ROW!")
                        Exit Sub
                    End If

                    If IsNumeric(TextBox_items_row_end.Text) = False Then
                        With TextBox_items_row_end
                            .Text = ""
                            .Focus()
                        End With
                        MsgBox("Please specify END row")

                        Exit Sub
                    End If

                    If Val(TextBox_items_row_start.Text) < 1 Then
                        With TextBox_items_row_start
                            .Text = ""
                            .Focus()
                        End With
                        MsgBox("Start row can't be smaller than 1")

                        Exit Sub
                    End If

                    If Val(TextBox_items_row_end.Text) < 1 Then
                        With TextBox_items_row_end
                            .Text = ""
                            .Focus()
                        End With
                        MsgBox("End row can't be smaller than 1")

                        Exit Sub
                    End If

                    If Val(TextBox_items_row_end.Text) < Val(TextBox_items_row_start.Text) Then
                        MsgBox("END row smaller than start row")

                        Exit Sub
                    End If

                    ascunde_butoanele_pentru_forms(Me, Colectie1)

                    Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
                    Dim start1 As Integer = CInt(TextBox_items_row_start.Text)
                    Dim end1 As Integer = CInt(TextBox_items_row_end.Text)

                    Dim Col1 As String = TextBox_col_1.Text.ToUpper
                    Dim Col2 As String = TextBox_col_2.Text.ToUpper
                    Dim Col3 As String = TextBox_col_3.Text.ToUpper
                    Dim Col4 As String = TextBox_col_4.Text.ToUpper

                    Data_table_Buoyancy.Rows.Clear()

                    Dim Index_Data_table As Integer = 0

                    For i = start1 To end1

                        Dim Type1 As String = W1.Range(Col1 & i).Value

                        Dim Spacing As Double = 0
                        Dim Sp_str As String = W1.Range(Col2 & i).Value

                        If IsNumeric(Sp_str) = True Then
                            Spacing = CDbl(Sp_str)
                        End If

                        Dim BS As String = W1.Range(Col3 & i).Value
                        Dim BeginSTATION As Double
                        If IsNumeric(Replace(BS, "+", "")) = True Then
                            BeginSTATION = CDbl((Replace(BS, "+", "")))
                        End If
                        Dim ES As String = W1.Range(Col4 & i).Value
                        Dim EndSTATION As Double
                        If IsNumeric(Replace(ES, "+", "")) = True Then
                            EndSTATION = CDbl((Replace(ES, "+", "")))
                        End If

                        If Type1.ToUpper.Contains("SCREW") = True Then
                            Type1 = "SA"
                        End If
                        If Type1.ToUpper.Contains("CCC") = True Then
                            Type1 = "CCC"
                        End If
                        If Type1.ToUpper.Contains("WEIGHT") = True Then
                            Type1 = "RW"
                        End If
                        If Type1.ToUpper.Contains("SADDLE") = True Then
                            Type1 = "SBW"
                        End If


                        If BeginSTATION >= 0 And EndSTATION > 0 And Not Type1 = "" Then
                            Data_table_Buoyancy.Rows.Add()

                            Data_table_Buoyancy.Rows(Index_Data_table).Item("BEGINSTA") = BeginSTATION
                            Data_table_Buoyancy.Rows(Index_Data_table).Item("ENDSTA") = EndSTATION
                            Data_table_Buoyancy.Rows(Index_Data_table).Item("TYPE") = Type1
                            Data_table_Buoyancy.Rows(Index_Data_table).Item("SPACING") = Spacing

                            Index_Data_table = Index_Data_table + 1
                        End If

                    Next

                    adauga_Buoyancy_la_crossing_list()

                    TextBox_col_1.BackColor = Drawing.Color.Yellow
                    TextBox_col_2.BackColor = Drawing.Color.Yellow
                    TextBox_col_3.BackColor = Drawing.Color.Yellow
                    TextBox_col_4.BackColor = Drawing.Color.Yellow
            End Select


            afiseaza_butoanele_pentru_forms(Me, Colectie1)
        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub adauga_matchline_la_crossing_list()
        If Data_table_Crossing.Rows.Count > 0 And Data_table_Matchline.Rows.Count > 0 Then



            For i = 0 To Data_table_Crossing.Rows.Count - 1
                If IsDBNull(Data_table_Crossing.Rows(i).Item("BLOCKNAME")) = False Then
                    If Data_table_Crossing.Rows(i).Item("BLOCKNAME") = "MATCHLINE" Then
                        Data_table_Crossing.Rows(i).Item("BLOCKNAME") = DBNull.Value
                        If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                            Data_table_Crossing.Rows(i).Item("STA") = DBNull.Value
                        End If
                    End If
                End If
            Next
            Data_table_Crossing = delete_DBnull_rows_from_data_table(Data_table_Crossing, "STA")



            Dim Index_Data_table As Double = Data_table_Crossing.Rows.Count

            For j = 0 To Data_table_Matchline.Rows.Count - 1
                If IsDBNull(Data_table_Matchline.Rows(j).Item("STATION")) = False Then
                    Data_table_Crossing.Rows.Add()
                    Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "MATCHLINE"
                    Data_table_Crossing.Rows(Index_Data_table).Item("STA") = Data_table_Matchline.Rows(j).Item("STATION")
                    Data_table_Crossing.Rows(Index_Data_table).Item("SHEET") = Data_table_Matchline.Rows(j).Item("PAGINA")

                    Index_Data_table = Index_Data_table + 1
                End If
            Next

            Data_table_Crossing = Sort_data_table(Data_table_Crossing, "STA")



            If Not Data_table_Crossing.Rows(0).Item("BLOCKNAME") = "MATCHLINE" Then
                Dim Done As Boolean = False
                Dim Rowindex As Integer = 1
                Do Until Done = True
                    If Data_table_Crossing.Rows(Rowindex).Item("BLOCKNAME") = "MATCHLINE" Then
                        For Each col As System.Data.DataColumn In Data_table_Crossing.Columns
                            Dim tempVal = Data_table_Crossing.Rows(0).Item(col.ColumnName)
                            Data_table_Crossing.Rows(0).Item(col.ColumnName) = Data_table_Crossing.Rows(Rowindex).Item(col.ColumnName)
                            Data_table_Crossing.Rows(Rowindex).Item(col.ColumnName) = tempVal

                        Next
                        Done = True
                    Else
                        Rowindex = Rowindex + 1
                    End If
                Loop
            End If

            Panel_crossings.Controls.Clear()

            For i = 0 To Data_table_Crossing.Rows.Count - 1
                Dim PAGINA As Integer
                If IsDBNull(Data_table_Crossing.Rows(i).Item("SHEET")) = False Then
                    PAGINA = Data_table_Crossing.Rows(i).Item("SHEET")
                Else
                    Data_table_Crossing.Rows(i).Item("SHEET") = PAGINA
                End If

                ' MsgBox(Data_table_Crossing.Rows(i).Item("STA") & vbCrLf & "PAGE = " & Data_table_Crossing.Rows(i).Item("SHEET"))
            Next





            Dim Ji As Integer = 0 ' asta e pus daca ai dbnull creeaza gap
            For i = 0 To Data_table_Crossing.Rows.Count - 1
                If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                    Dim Continut_textbox As String

                    Dim Backcolor As Drawing.Color = Drawing.Color.Gainsboro

                    If Data_table_Crossing.Rows(i).Item("BLOCKNAME") = "MATCHLINE" Then
                        Continut_textbox = Get_chainage_from_double(Data_table_Crossing.Rows(i).Item("STA"), 1) & " - " & "MATCHLINE"
                    ElseIf Data_table_Crossing.Rows(i).Item("BLOCKNAME") = "CLASS_LOCATION1" Or Data_table_Crossing.Rows(i).Item("BLOCKNAME") = "CLASS_LOCATION2" Then
                        If IsDBNull(Data_table_Crossing.Rows(i).Item("DESCRIPTION1")) = False Then
                            Continut_textbox = Data_table_Crossing.Rows(i).Item("DESCRIPTION1")
                            If IsDBNull(Data_table_Crossing.Rows(i).Item("DESCRIPTION2")) = False Then Continut_textbox = Continut_textbox & " " & Data_table_Crossing.Rows(i).Item("DESCRIPTION2")
                            Continut_textbox = Get_chainage_from_double(Data_table_Crossing.Rows(i).Item("STA"), 1) & "-" & Continut_textbox
                        End If
                    Else
                        If IsDBNull(Data_table_Crossing.Rows(i).Item("DESCRIPTION1")) = False Then
                            Continut_textbox = Data_table_Crossing.Rows(i).Item("DESCRIPTION1")
                            If IsDBNull(Data_table_Crossing.Rows(i).Item("DESCRIPTION2")) = False Then Continut_textbox = Continut_textbox & " " & Data_table_Crossing.Rows(i).Item("DESCRIPTION2")
                            Continut_textbox = Get_chainage_from_double(Data_table_Crossing.Rows(i).Item("STA"), 1) & "-" & Continut_textbox
                        End If
                    End If

                    If Data_table_Crossing.Rows(i).Item("BLOCKNAME") = "MATCHLINE" Then
                        Backcolor = Drawing.Color.SkyBlue
                    ElseIf Data_table_Crossing.Rows(i).Item("BLOCKNAME") = "CLASS_LOCATION1" Or Data_table_Crossing.Rows(i).Item("BLOCKNAME") = "CLASS_LOCATION2" Then
                        Backcolor = Drawing.Color.Cyan
                    Else
                    End If


                    Dim textbox1 As New Windows.Forms.TextBox
                    textbox1 = Adauga_textbox(4, 4 + (21 + 4) * Ji, 427, 21, Continut_textbox & " (" & Data_table_Crossing.Rows(i).Item("SHEET") & ")", True, False, Backcolor)
                    Panel_crossings.Controls.Add(textbox1)
                    Ji = Ji + 1
                End If

            Next

        End If


    End Sub

    Private Sub adauga_ELBOWS_la_crossing_list()

        If Data_table_Crossing.Rows.Count > 0 And Data_table_Elbows.Rows.Count > 0 Then
            For i = 0 To Data_table_Crossing.Rows.Count - 1
                If IsDBNull(Data_table_Crossing.Rows(i).Item("BLOCKNAME")) = False Then
                    If Data_table_Crossing.Rows(i).Item("BLOCKNAME") = "ELBOW" Then
                        Data_table_Crossing.Rows(i).Item("BLOCKNAME") = DBNull.Value
                        If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                            Data_table_Crossing.Rows(i).Item("STA") = DBNull.Value
                        End If
                    End If
                End If
            Next
            Data_table_Crossing = delete_DBnull_rows_from_data_table(Data_table_Crossing, "STA")


            Dim Index_Data_table As Double = Data_table_Crossing.Rows.Count

            For j = 0 To Data_table_Elbows.Rows.Count - 1
                If IsDBNull(Data_table_Elbows.Rows(j).Item("STATION")) = False Then

                    Data_table_Crossing.Rows.Add()
                    Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "ELBOW"
                    Data_table_Crossing.Rows(Index_Data_table).Item("STA") = Data_table_Elbows.Rows(j).Item("STATION")
                    If IsDBNull(Data_table_Elbows.Rows(j).Item("LENGTH")) = False Then
                        Data_table_Crossing.Rows(Index_Data_table).Item("LENGTH") = Data_table_Elbows.Rows(j).Item("LENGTH")
                        If IsDBNull(Data_table_Elbows.Rows(j).Item("VALUE")) = False Then
                            Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "ELBOW " & Data_table_Elbows.Rows(j).Item("VALUE") & Degree_symbol() & " L = " & Data_table_Elbows.Rows(j).Item("LENGTH")
                            Data_table_Crossing.Rows(Index_Data_table).Item("BEGINSTA") = Round(Data_table_Elbows.Rows(j).Item("STATION"), 1) - Ceiling((Data_table_Elbows.Rows(j).Item("LENGTH") * 10) / 2) / 10
                            Data_table_Crossing.Rows(Index_Data_table).Item("ENDSTA") = Round(Data_table_Elbows.Rows(j).Item("STATION"), 1) + Floor((Data_table_Elbows.Rows(j).Item("LENGTH") * 10) / 2) / 10

                        End If
                    End If

                    Index_Data_table = Index_Data_table + 1
                End If
            Next



            Data_table_Crossing = Sort_data_table(Data_table_Crossing, "STA")
            Panel_crossings.Controls.Clear()

            adauga_matchline_la_crossing_list()

        End If


    End Sub

    Private Sub adauga_CLASS_LOCATION_la_crossing_list()

        If Data_table_Crossing.Rows.Count > 0 And Data_table_Class_location.Rows.Count > 0 Then
            For i = 0 To Data_table_Crossing.Rows.Count - 1
                If IsDBNull(Data_table_Crossing.Rows(i).Item("BLOCKNAME")) = False Then
                    If Data_table_Crossing.Rows(i).Item("BLOCKNAME") = "CLASS_LOCATION1" Or Data_table_Crossing.Rows(i).Item("BLOCKNAME") = "CLASS_LOCATION2" Then
                        Data_table_Crossing.Rows(i).Item("BLOCKNAME") = DBNull.Value
                        If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                            Data_table_Crossing.Rows(i).Item("STA") = DBNull.Value
                        End If
                    End If
                End If
            Next
            Data_table_Crossing = delete_DBnull_rows_from_data_table(Data_table_Crossing, "STA")


            Dim Index_Data_table As Double = Data_table_Crossing.Rows.Count

            For j = 0 To Data_table_Class_location.Rows.Count - 1
                If IsDBNull(Data_table_Class_location.Rows(j).Item("BEGINSTA")) = False Then

                    Data_table_Crossing.Rows.Add()

                    Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "CLASS_LOCATION1"
                    Data_table_Crossing.Rows(Index_Data_table).Item("STA") = Data_table_Class_location.Rows(j).Item("BEGINSTA")
                    Data_table_Crossing.Rows(Index_Data_table).Item("BEGINSTA") = Data_table_Class_location.Rows(j).Item("BEGINSTA")
                    Data_table_Crossing.Rows(Index_Data_table).Item("ENDSTA") = Data_table_Class_location.Rows(j).Item("ENDSTA")
                    Data_table_Crossing.Rows(Index_Data_table).Item("WT") = Data_table_Class_location.Rows(j).Item("WT")
                    Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = Data_table_Class_location.Rows(j).Item("MATERIAL")
                    Data_table_Crossing.Rows(Index_Data_table).Item("CLASS") = Data_table_Class_location.Rows(j).Item("CLASS")
                    Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "START OF CLASS " & Data_table_Class_location.Rows(j).Item("CLASS").ToString
                    Index_Data_table = Index_Data_table + 1
                End If
                If IsDBNull(Data_table_Class_location.Rows(j).Item("ENDSTA")) = False Then

                    Data_table_Crossing.Rows.Add()
                    Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "CLASS_LOCATION2"
                    Data_table_Crossing.Rows(Index_Data_table).Item("STA") = Data_table_Class_location.Rows(j).Item("ENDSTA")
                    Data_table_Crossing.Rows(Index_Data_table).Item("BEGINSTA") = Data_table_Class_location.Rows(j).Item("BEGINSTA")
                    Data_table_Crossing.Rows(Index_Data_table).Item("ENDSTA") = Data_table_Class_location.Rows(j).Item("ENDSTA")
                    Data_table_Crossing.Rows(Index_Data_table).Item("WT") = Data_table_Class_location.Rows(j).Item("WT")
                    Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = Data_table_Class_location.Rows(j).Item("MATERIAL")
                    Data_table_Crossing.Rows(Index_Data_table).Item("CLASS") = Data_table_Class_location.Rows(j).Item("CLASS")
                    Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "END OF CLASS " & Data_table_Class_location.Rows(j).Item("CLASS").ToString
                    Index_Data_table = Index_Data_table + 1
                End If
            Next



            Data_table_Crossing = Sort_data_table(Data_table_Crossing, "STA")
            Panel_crossings.Controls.Clear()

            adauga_matchline_la_crossing_list()

        End If


    End Sub

    Private Sub adauga_Hydrotest_la_crossing_list()

        If Data_table_Crossing.Rows.Count > 0 And Data_table_Hydrotest.Rows.Count > 0 Then
            For i = 0 To Data_table_Crossing.Rows.Count - 1
                If IsDBNull(Data_table_Crossing.Rows(i).Item("BLOCKNAME")) = False Then
                    If Data_table_Crossing.Rows(i).Item("BLOCKNAME") = "HYDROTEST1" Or Data_table_Crossing.Rows(i).Item("BLOCKNAME") = "HYDROTEST2" Then
                        Data_table_Crossing.Rows(i).Item("BLOCKNAME") = DBNull.Value
                        If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                            Data_table_Crossing.Rows(i).Item("STA") = DBNull.Value
                        End If
                    End If
                End If
            Next
            Data_table_Crossing = delete_DBnull_rows_from_data_table(Data_table_Crossing, "STA")


            Dim Index_Data_table As Double = Data_table_Crossing.Rows.Count

            For j = 0 To Data_table_Hydrotest.Rows.Count - 1
                If IsDBNull(Data_table_Hydrotest.Rows(j).Item("BEGINSTA")) = False Then

                    Data_table_Crossing.Rows.Add()

                    Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "HYDROTEST1"
                    Data_table_Crossing.Rows(Index_Data_table).Item("STA") = Data_table_Hydrotest.Rows(j).Item("BEGINSTA")
                    Data_table_Crossing.Rows(Index_Data_table).Item("BEGINSTA") = Data_table_Hydrotest.Rows(j).Item("BEGINSTA")
                    Data_table_Crossing.Rows(Index_Data_table).Item("ENDSTA") = Data_table_Hydrotest.Rows(j).Item("ENDSTA")
                    Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = Data_table_Hydrotest.Rows(j).Item("MATERIAL")
                    Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "START OF HYDROTEST MATERIAL " & Data_table_Hydrotest.Rows(j).Item("MATERIAL").ToString
                    Index_Data_table = Index_Data_table + 1
                End If
                If IsDBNull(Data_table_Hydrotest.Rows(j).Item("ENDSTA")) = False Then

                    Data_table_Crossing.Rows.Add()
                    Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "HYDROTEST2"
                    Data_table_Crossing.Rows(Index_Data_table).Item("STA") = Data_table_Hydrotest.Rows(j).Item("ENDSTA")
                    Data_table_Crossing.Rows(Index_Data_table).Item("BEGINSTA") = Data_table_Hydrotest.Rows(j).Item("BEGINSTA")
                    Data_table_Crossing.Rows(Index_Data_table).Item("ENDSTA") = Data_table_Hydrotest.Rows(j).Item("ENDSTA")
                    Data_table_Crossing.Rows(Index_Data_table).Item("MATERIAL") = Data_table_Hydrotest.Rows(j).Item("MATERIAL")

                    Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "END OF HYDROTEST MATERIAL " & Data_table_Hydrotest.Rows(j).Item("MATERIAL").ToString
                    Index_Data_table = Index_Data_table + 1
                End If
            Next



            Data_table_Crossing = Sort_data_table(Data_table_Crossing, "STA")
            Panel_crossings.Controls.Clear()

            adauga_matchline_la_crossing_list()

        End If


    End Sub

    Private Sub adauga_Stress_la_crossing_list()

        If Data_table_Crossing.Rows.Count > 0 And Data_table_Stress.Rows.Count > 0 Then
            For i = 0 To Data_table_Crossing.Rows.Count - 1
                If IsDBNull(Data_table_Crossing.Rows(i).Item("BLOCKNAME")) = False Then
                    If Data_table_Crossing.Rows(i).Item("BLOCKNAME") = "STRESS1" Or Data_table_Crossing.Rows(i).Item("BLOCKNAME") = "STRESS2" Then
                        Data_table_Crossing.Rows(i).Item("BLOCKNAME") = DBNull.Value
                        If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                            Data_table_Crossing.Rows(i).Item("STA") = DBNull.Value
                        End If
                    End If
                End If
            Next
            Data_table_Crossing = delete_DBnull_rows_from_data_table(Data_table_Crossing, "STA")


            Dim Index_Data_table As Double = Data_table_Crossing.Rows.Count

            For j = 0 To Data_table_Stress.Rows.Count - 1
                If IsDBNull(Data_table_Stress.Rows(j).Item("BEGINSTA")) = False Then

                    Data_table_Crossing.Rows.Add()

                    Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "STRESS1"
                    Data_table_Crossing.Rows(Index_Data_table).Item("STA") = Data_table_Stress.Rows(j).Item("BEGINSTA")
                    Data_table_Crossing.Rows(Index_Data_table).Item("BEGINSTA") = Data_table_Stress.Rows(j).Item("BEGINSTA")
                    Data_table_Crossing.Rows(Index_Data_table).Item("ENDSTA") = Data_table_Stress.Rows(j).Item("ENDSTA")
                    Data_table_Crossing.Rows(Index_Data_table).Item("WT") = Data_table_Stress.Rows(j).Item("WT")
                    Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "START OF STRESS THICKNESS " & Data_table_Stress.Rows(j).Item("WT").ToString
                    Index_Data_table = Index_Data_table + 1
                End If
                If IsDBNull(Data_table_Stress.Rows(j).Item("ENDSTA")) = False Then
                    Data_table_Crossing.Rows.Add()
                    Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "STRESS2"
                    Data_table_Crossing.Rows(Index_Data_table).Item("STA") = Data_table_Stress.Rows(j).Item("ENDSTA")
                    Data_table_Crossing.Rows(Index_Data_table).Item("BEGINSTA") = Data_table_Stress.Rows(j).Item("BEGINSTA")
                    Data_table_Crossing.Rows(Index_Data_table).Item("ENDSTA") = Data_table_Stress.Rows(j).Item("ENDSTA")
                    Data_table_Crossing.Rows(Index_Data_table).Item("WT") = Data_table_Stress.Rows(j).Item("WT")

                    Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "END OF STRESS THICKNESS " & Data_table_Stress.Rows(j).Item("WT").ToString
                    Index_Data_table = Index_Data_table + 1
                End If
            Next



            Data_table_Crossing = Sort_data_table(Data_table_Crossing, "STA")
            Panel_crossings.Controls.Clear()

            adauga_matchline_la_crossing_list()

        End If


    End Sub

    Private Sub adauga_Buoyancy_la_crossing_list()

        If Data_table_Crossing.Rows.Count > 0 And Data_table_Buoyancy.Rows.Count > 0 Then
            For i = 0 To Data_table_Crossing.Rows.Count - 1
                If IsDBNull(Data_table_Crossing.Rows(i).Item("BLOCKNAME")) = False Then
                    If Data_table_Crossing.Rows(i).Item("BLOCKNAME") = "SA1" Or Data_table_Crossing.Rows(i).Item("BLOCKNAME") = "SA2" Then

                        Data_table_Crossing.Rows(i).Item("BLOCKNAME") = DBNull.Value
                        If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                            Data_table_Crossing.Rows(i).Item("STA") = DBNull.Value
                        End If
                    End If
                End If
            Next
            Data_table_Crossing = delete_DBnull_rows_from_data_table(Data_table_Crossing, "STA")


            Dim Index_Data_table As Double = Data_table_Crossing.Rows.Count

            For j = 0 To Data_table_Buoyancy.Rows.Count - 1
                If IsDBNull(Data_table_Buoyancy.Rows(j).Item("BEGINSTA")) = False Then
                    Dim Type1 As String = ""
                    Select Case Data_table_Buoyancy.Rows(j).Item("TYPE")
                        Case "SA"
                            Type1 = "SA1"
                        Case "CCC"
                            Type1 = "CCC1"
                        Case "SBW"
                            Type1 = "SBW1"
                        Case "RW"
                            Type1 = "RW1"
                    End Select
                    If Not Type1 = "" Then
                        Data_table_Crossing.Rows.Add()
                        Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "SA1"
                        Data_table_Crossing.Rows(Index_Data_table).Item("BUOYANCYTYPE") = Type1
                        Data_table_Crossing.Rows(Index_Data_table).Item("STA") = Data_table_Buoyancy.Rows(j).Item("BEGINSTA")
                        Data_table_Crossing.Rows(Index_Data_table).Item("BEGINSTA") = Data_table_Buoyancy.Rows(j).Item("BEGINSTA")
                        Data_table_Crossing.Rows(Index_Data_table).Item("ENDSTA") = Data_table_Buoyancy.Rows(j).Item("ENDSTA")
                        Dim Spacing As Double = Data_table_Buoyancy.Rows(j).Item("SPACING")
                        Dim Len1 As Double
                        If Spacing > 0 Then
                            Len1 = Round(Data_table_Buoyancy.Rows(j).Item("ENDSTA") - Data_table_Buoyancy.Rows(j).Item("BEGINSTA"), 1)
                            Dim NR_SA As Integer = CInt(1 + Len1 / Spacing)
                            Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = NR_SA & " " & Data_table_Buoyancy.Rows(j).Item("TYPE") & vbCrLf & Get_String_Rounded(Spacing, 1) & " C/C"
                        Else
                            Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "CONCRETE"
                        End If
                        Index_Data_table = Index_Data_table + 1
                    End If

                End If
                If IsDBNull(Data_table_Buoyancy.Rows(j).Item("ENDSTA")) = False Then
                    Dim Type1 As String = ""
                    Select Case Data_table_Buoyancy.Rows(j).Item("TYPE")
                        Case "SA"
                            Type1 = "SA2"
                        Case "CCC"
                            Type1 = "CCC2"
                        Case "SBW"
                            Type1 = "SBW2"
                        Case "RW"
                            Type1 = "RW2"
                    End Select
                    If Not Type1 = "" Then
                        Data_table_Crossing.Rows.Add()
                        Data_table_Crossing.Rows(Index_Data_table).Item("BLOCKNAME") = "SA2"
                        Data_table_Crossing.Rows(Index_Data_table).Item("BUOYANCYTYPE") = Type1
                        Data_table_Crossing.Rows(Index_Data_table).Item("STA") = Data_table_Buoyancy.Rows(j).Item("ENDSTA")
                        Data_table_Crossing.Rows(Index_Data_table).Item("BEGINSTA") = Data_table_Buoyancy.Rows(j).Item("BEGINSTA")
                        Data_table_Crossing.Rows(Index_Data_table).Item("ENDSTA") = Data_table_Buoyancy.Rows(j).Item("ENDSTA")
                        Dim Spacing As Double = Data_table_Buoyancy.Rows(j).Item("SPACING")
                        Dim Len1 As Double
                        If Spacing > 0 Then
                            Len1 = Round(Data_table_Buoyancy.Rows(j).Item("ENDSTA") - Data_table_Buoyancy.Rows(j).Item("BEGINSTA"), 1)
                            Dim NR_SA As Integer = CInt(1 + Len1 / Spacing)
                            Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = NR_SA & " " & Data_table_Buoyancy.Rows(j).Item("TYPE") & vbCrLf & Get_String_Rounded(Spacing, 1) & " C/C"
                        Else
                            Data_table_Crossing.Rows(Index_Data_table).Item("DESCRIPTION1") = "CONCRETE"
                        End If
                        Index_Data_table = Index_Data_table + 1
                    End If

                End If
            Next



            Data_table_Crossing = Sort_data_table(Data_table_Crossing, "STA")
            Panel_crossings.Controls.Clear()

            adauga_matchline_la_crossing_list()

        End If


    End Sub


    Private Sub Button_list_to_DWG_Click(sender As Object, e As EventArgs) Handles Button_list_to_DWG.Click


        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try



            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    If Data_table_Crossing.Rows.Count > 0 And Data_table_Matchline.Rows.Count > 0 Then

                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord


                        Dim Numele_template_layout As String = "TBLK"

                        Dim X, Z As Double
                        Dim Ycr As Double

                        If RadioButton_left_to_right.Checked = True Then
                            X = 112 + 28
                        Else
                            X = 964 - 28
                        End If

                        Ycr = 599.41

                        Dim Match1 As Double = 0
                        Dim Match2 As Double = 0
                       
                        Dim Class_station1 As Double
                        Dim Class_station2 As Double
                        Dim Is_class_location As Boolean = False
                        Dim text_class_location As String = ""

                        Dim Hydro_station1 As Double
                        Dim Hydro_station2 As Double
                        Dim Is_hydro As Boolean = False
                        Dim text_hydrostation As String = ""

                        Dim Stress_station1 As Double
                        Dim Stress_station2 As Double
                        Dim Is_stress As Boolean = False
                        Dim text_stress As String = ""
                        Dim Len_stress As Double


                        Dim Buoyancy_station1 As Double
                        Dim Buoyancy_station2 As Double
                        Dim Is_Buoyancy As Boolean = False
                        Dim text_Buoyancy As String = ""

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

                                    Case "Mat_worksheet_block1"
                                        If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                                            Dim Colectie_atr_name_PIPE As New Specialized.StringCollection
                                            Dim Colectie_atr_value_PIPE As New Specialized.StringCollection

                                            Colectie_atr_name_PIPE.Add("STA")
                                            Dim Test1 As Double = Data_table_Crossing.Rows(i).Item("STA")

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

                                            Dim Pct_ins As New Point3d
                                            If RadioButton_right_to_left.Checked = True Then
                                                Pct_ins = New Point3d(X, Ycr, Z)
                                            Else
                                                Dim Valoare_x_shift As Double = 28
                                                Pct_ins = New Point3d(X - Valoare_x_shift, Ycr, Z)
                                            End If

                                            Dim diferenta_from_start As Double = Round(Data_table_Crossing.Rows(i).Item("STA"), 1) - Match1
                                            Dim lungime_pag = 825
                                            Dim Diferenta_match As Double = Match2 - Match1
                                            Dim Factor As Double = lungime_pag / Diferenta_match

                                            Dim X0 As Double
                                            If RadioButton_left_to_right.Checked = True Then
                                                X0 = 112
                                                Pct_ins = New Point3d(X0 + Factor * diferenta_from_start, Ycr, Z)
                                            Else
                                                X0 = 964 - 28

                                                Pct_ins = New Point3d(X0 - Factor * diferenta_from_start, Ycr, Z)
                                            End If





                                            InsertBlock_with_multiple_atributes("Mat_worksheet_block1.dwg", "Mat_worksheet_block1", Pct_ins, 1, BTrecord, "TEXT", Colectie_atr_name_PIPE, Colectie_atr_value_PIPE)



                                            If RadioButton_right_to_left.Checked = True Then
                                                X = X - 50
                                            Else
                                                X = X + 50
                                            End If


                                        End If


                                    Case "ELBOW"
                                        If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                                            Dim Colectie_atr_name_PIPE As New Specialized.StringCollection
                                            Dim Colectie_atr_value_PIPE As New Specialized.StringCollection

                                            Colectie_atr_name_PIPE.Add("STA")
                                            Dim Sta1 As Double = Data_table_Crossing.Rows(i).Item("STA")

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

                                            Dim Len1 As Double
                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("LENGTH")) = False Then
                                                Len1 = Data_table_Crossing.Rows(i).Item("LENGTH")
                                            End If
                                            Colectie_atr_name_PIPE.Add("LENGTH")
                                            Colectie_atr_value_PIPE.Add(Get_String_Rounded(Len1, 1))

                                            Dim beginSta As String = " "
                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("BEGINSTA")) = False Then
                                                beginSta = Get_chainage_from_double(Data_table_Crossing.Rows(i).Item("BEGINSTA"), 1)
                                            End If

                                            Dim endSta As String = " "
                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("ENDSTA")) = False Then
                                                endSta = Get_chainage_from_double(Data_table_Crossing.Rows(i).Item("ENDSTA"), 1)
                                            End If


                                            Colectie_atr_name_PIPE.Add("BEGINSTA")
                                            Colectie_atr_name_PIPE.Add("ENDSTA")

                                            If RadioButton_left_to_right.Checked = True Then
                                                Colectie_atr_value_PIPE.Add(beginSta)
                                                Colectie_atr_value_PIPE.Add(endSta)
                                            Else
                                                Colectie_atr_value_PIPE.Add(endSta)
                                                Colectie_atr_value_PIPE.Add(beginSta)
                                            End If



                                            Dim Pct_ins As New Point3d
                                            Dim Yelbow As Double = 512.1
                                            If RadioButton_right_to_left.Checked = True Then
                                                Pct_ins = New Point3d(X, Yelbow, Z)
                                            Else
                                                Dim Valoare_x_shift As Double = 28
                                                Pct_ins = New Point3d(X - Valoare_x_shift, Yelbow, Z)
                                            End If


                                            Dim diferenta_from_start As Double = Round(Data_table_Crossing.Rows(i).Item("STA"), 1) - Match1
                                            Dim lungime_pag = 852
                                            Dim Diferenta_match As Double = Match2 - Match1
                                            Dim Factor As Double = lungime_pag / Diferenta_match

                                            Dim X0 As Double
                                            If RadioButton_left_to_right.Checked = True Then
                                                X0 = 112
                                                Pct_ins = New Point3d(X0 + Factor * diferenta_from_start, Yelbow, Z)
                                            Else
                                                X0 = 964 - 28

                                                Pct_ins = New Point3d(X0 - Factor * diferenta_from_start, Yelbow, Z)
                                            End If

                                            InsertBlock_with_multiple_atributes("Mat_worksheet_block1.dwg", "Mat_worksheet_block1", Pct_ins, 1, BTrecord, "TEXT", Colectie_atr_name_PIPE, Colectie_atr_value_PIPE)

                                            If RadioButton_right_to_left.Checked = True Then
                                                X = X - 50
                                            Else
                                                X = X + 50
                                            End If


                                        End If

                                    Case "CLASS_LOCATION1"
                                        If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then


                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("CLASS")) = False And IsDBNull(Data_table_Crossing.Rows(i).Item("WT")) = False And IsDBNull(Data_table_Crossing.Rows(i).Item("MATERIAL")) = False Then
                                                text_class_location = "CLASS " & Data_table_Crossing.Rows(i).Item("CLASS").ToString & vbCrLf & "(" & Data_table_Crossing.Rows(i).Item("WT").ToString & _
                                                                                           ")" & vbCrLf & Data_table_Crossing.Rows(i).Item("MATERIAL").ToString

                                                Class_station1 = Data_table_Crossing.Rows(i).Item("STA")

                                                Class_station2 = Data_table_Crossing.Rows(i).Item("ENDSTA")

                                                Is_class_location = True

                                            End If


                                        End If


                                    Case "CLASS_LOCATION2"
                                        If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                                            If Is_class_location = True Then

                                                Dim Pct_sus As New Point3d
                                                Dim Pct_jos As New Point3d

                                                Dim Yclass_sus As Double = 512.1
                                                Dim Yclass_jos As Double = 424.8

                                                Dim diferenta_from_start As Double = Round(Data_table_Crossing.Rows(i).Item("STA"), 1) - Match1
                                                Dim lungime_pag As Double = 852
                                                Dim Diferenta_match As Double = Match2 - Match1
                                                Dim Factor As Double = lungime_pag / Diferenta_match

                                                Dim X0 As Double
                                                If RadioButton_left_to_right.Checked = True Then
                                                    X0 = 112
                                                    Pct_jos = New Point3d(X0 + Factor * diferenta_from_start, Yclass_jos, Z)
                                                    Pct_sus = New Point3d(X0 + Factor * diferenta_from_start, Yclass_sus, Z)
                                                Else
                                                    X0 = 964 - 28
                                                    Pct_jos = New Point3d(X0 - Factor * diferenta_from_start, Yclass_jos, Z)
                                                    Pct_sus = New Point3d(X0 - Factor * diferenta_from_start, Yclass_sus, Z)
                                                End If

                                                If Not Round(Data_table_Crossing.Rows(i).Item("STA"), 1) = 0 Then
                                                    Dim Line1 As New Line(Pct_sus, Pct_jos)
                                                    Line1.Layer = "TEXT"
                                                    BTrecord.AppendEntity(Line1)
                                                    Trans1.AddNewlyCreatedDBObject(Line1, True)

                                                    Dim Mtext1 As New MText
                                                    Mtext1.Layer = "TEXT"
                                                    Mtext1.Rotation = PI / 2
                                                    Mtext1.Contents = Get_chainage_from_double(Class_station2, 1)
                                                    Mtext1.TextHeight = 5
                                                    Mtext1.Location = New Point3d(Pct_jos.X + 2, Pct_jos.Y + (Pct_sus.Y - Pct_jos.Y) / 2, 0)
                                                    Mtext1.Attachment = AttachmentPoint.TopCenter
                                                    BTrecord.AppendEntity(Mtext1)
                                                    Trans1.AddNewlyCreatedDBObject(Mtext1, True)

                                                End If










                                                Dim Pct_INS As New Point3d
                                                If Class_station1 > Match1 Then
                                                    diferenta_from_start = Class_station1 + (Class_station2 - Class_station1) * 0.5 - Match1
                                                Else
                                                    diferenta_from_start = (Class_station2 - Match1) * 0.5
                                                End If


                                                If RadioButton_left_to_right.Checked = True Then

                                                    Pct_INS = New Point3d(X0 + Factor * diferenta_from_start, Yclass_jos + (Yclass_sus - Yclass_jos) / 2, Z)

                                                Else

                                                    Pct_INS = New Point3d(X0 - Factor * diferenta_from_start, Yclass_jos + (Yclass_sus - Yclass_jos) / 2, Z)

                                                End If



                                                Dim Mtext2 As New MText
                                                Mtext2.Layer = "TEXT"
                                                Mtext2.Rotation = 0
                                                Mtext2.Contents = text_class_location

                                                Mtext2.TextHeight = 5
                                                Mtext2.Location = Pct_INS
                                                Mtext2.Attachment = AttachmentPoint.MiddleCenter
                                                BTrecord.AppendEntity(Mtext2)
                                                Trans1.AddNewlyCreatedDBObject(Mtext2, True)
                                                text_class_location = ""
                                                Class_station1 = 0
                                                Class_station2 = 0
                                                Is_class_location = False

                                            End If
                                        End If

                                    Case "HYDROTEST1"
                                        If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then



                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("MATERIAL")) = False Then
                                                text_hydrostation = "HYDROTEST " & vbCrLf & Data_table_Crossing.Rows(i).Item("MATERIAL").ToString

                                                Hydro_station2 = Data_table_Crossing.Rows(i).Item("ENDSTA")
                                                Hydro_station1 = Data_table_Crossing.Rows(i).Item("STA")
                                                Is_hydro = True

                                            End If


                                        End If


                                    Case "HYDROTEST2"
                                        If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                                            If Is_hydro = True Then

                                                Dim Pct_sus As New Point3d
                                                Dim Pct_jos As New Point3d

                                                Dim Yhydro_sus As Double = 424.8
                                                Dim Yhydro_jos As Double = 337.5

                                                Dim diferenta_from_start As Double = Round(Data_table_Crossing.Rows(i).Item("STA"), 1) - Match1
                                                Dim lungime_pag As Double = 852
                                                Dim Diferenta_match As Double = Match2 - Match1
                                                Dim Factor As Double = lungime_pag / Diferenta_match

                                                Dim X0 As Double
                                                If RadioButton_left_to_right.Checked = True Then
                                                    X0 = 112
                                                    Pct_jos = New Point3d(X0 + Factor * diferenta_from_start, Yhydro_jos, Z)
                                                    Pct_sus = New Point3d(X0 + Factor * diferenta_from_start, Yhydro_sus, Z)
                                                Else
                                                    X0 = 964 - 28
                                                    Pct_jos = New Point3d(X0 - Factor * diferenta_from_start, Yhydro_jos, Z)
                                                    Pct_sus = New Point3d(X0 - Factor * diferenta_from_start, Yhydro_sus, Z)
                                                End If

                                                If Not Round(Data_table_Crossing.Rows(i).Item("STA"), 1) = 0 Then
                                                    Dim Line1 As New Line(Pct_sus, Pct_jos)
                                                    Line1.Layer = "TEXT"
                                                    BTrecord.AppendEntity(Line1)
                                                    Trans1.AddNewlyCreatedDBObject(Line1, True)

                                                    Dim Mtext1 As New MText
                                                    Mtext1.Layer = "TEXT"
                                                    Mtext1.Rotation = PI / 2
                                                    Mtext1.Contents = Get_chainage_from_double(Hydro_station2, 1)
                                                    Mtext1.TextHeight = 5
                                                    Mtext1.Location = New Point3d(Pct_jos.X + 2, Pct_jos.Y + (Pct_sus.Y - Pct_jos.Y) / 2, 0)
                                                    Mtext1.Attachment = AttachmentPoint.TopCenter
                                                    BTrecord.AppendEntity(Mtext1)
                                                    Trans1.AddNewlyCreatedDBObject(Mtext1, True)

                                                End If










                                                Dim Pct_INS As New Point3d
                                                If Hydro_station1 > Match1 Then
                                                    diferenta_from_start = Hydro_station1 + (Hydro_station2 - Hydro_station1) * 0.5 - Match1
                                                Else
                                                    diferenta_from_start = (Hydro_station2 - Match1) * 0.5
                                                End If


                                                If RadioButton_left_to_right.Checked = True Then

                                                    Pct_INS = New Point3d(X0 + Factor * diferenta_from_start, Yhydro_jos + (Yhydro_sus - Yhydro_jos) / 2, Z)

                                                Else

                                                    Pct_INS = New Point3d(X0 - Factor * diferenta_from_start, Yhydro_jos + (Yhydro_sus - Yhydro_jos) / 2, Z)

                                                End If



                                                Dim Mtext2 As New MText
                                                Mtext2.Layer = "TEXT"
                                                Mtext2.Rotation = 0
                                                Mtext2.Contents = text_hydrostation

                                                Mtext2.TextHeight = 5
                                                Mtext2.Location = Pct_INS
                                                Mtext2.Attachment = AttachmentPoint.MiddleCenter
                                                BTrecord.AppendEntity(Mtext2)
                                                Trans1.AddNewlyCreatedDBObject(Mtext2, True)
                                                text_hydrostation = ""
                                                Hydro_station1 = 0
                                                Hydro_station2 = 0
                                                Is_hydro = False

                                            End If
                                        End If

                                    Case "STRESS1"
                                        If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then



                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("WT")) = False Then
                                                text_stress = "STRESS " & vbCrLf & Data_table_Crossing.Rows(i).Item("WT").ToString

                                                Stress_station1 = Data_table_Crossing.Rows(i).Item("STA")
                                                Stress_station2 = Data_table_Crossing.Rows(i).Item("ENDSTA")
                                                Len_stress = Stress_station2 - Stress_station1

                                                Dim Pct_sus As New Point3d
                                                Dim Pct_jos As New Point3d

                                                Dim YSTRESS_sus As Double = 337.5
                                                Dim YSTRESS_jos As Double = 250.19

                                                Dim diferenta_from_start As Double = Round(Data_table_Crossing.Rows(i).Item("STA"), 1) - Match1
                                                Dim lungime_pag As Double = 852
                                                Dim Diferenta_match As Double = Match2 - Match1
                                                Dim Factor As Double = lungime_pag / Diferenta_match

                                                Dim X0 As Double
                                                If RadioButton_left_to_right.Checked = True Then
                                                    X0 = 112
                                                    Pct_jos = New Point3d(X0 + Factor * diferenta_from_start, YSTRESS_jos, Z)
                                                    Pct_sus = New Point3d(X0 + Factor * diferenta_from_start, YSTRESS_sus, Z)
                                                Else
                                                    X0 = 964 - 28
                                                    Pct_jos = New Point3d(X0 - Factor * diferenta_from_start, YSTRESS_jos, Z)
                                                    Pct_sus = New Point3d(X0 - Factor * diferenta_from_start, YSTRESS_sus, Z)
                                                End If

                                                If Not Round(Data_table_Crossing.Rows(i).Item("STA"), 1) = 0 Then
                                                    Dim Line1 As New Line(Pct_sus, Pct_jos)
                                                    Line1.Layer = "TEXT"
                                                    BTrecord.AppendEntity(Line1)
                                                    Trans1.AddNewlyCreatedDBObject(Line1, True)

                                                    Dim Mtext1 As New MText
                                                    Mtext1.Layer = "TEXT"
                                                    Mtext1.Rotation = PI / 2
                                                    Mtext1.Contents = Get_chainage_from_double(Stress_station1, 1)
                                                    Mtext1.TextHeight = 3
                                                    Mtext1.Location = New Point3d(Pct_jos.X + 2, Pct_jos.Y + (Pct_sus.Y - Pct_jos.Y) / 2, 0)
                                                    Mtext1.Attachment = AttachmentPoint.TopCenter
                                                    BTrecord.AppendEntity(Mtext1)
                                                    Trans1.AddNewlyCreatedDBObject(Mtext1, True)

                                                End If




                                                Is_stress = True

                                            End If


                                        End If


                                    Case "STRESS2"
                                        If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                                            If Is_stress = True Then




                                                Dim Pct_sus As New Point3d
                                                Dim Pct_jos As New Point3d

                                                Dim YSTRESS_sus As Double = 337.5
                                                Dim YSTRESS_jos As Double = 250.19

                                                Dim diferenta_from_start As Double = Round(Data_table_Crossing.Rows(i).Item("STA"), 1) - Match1
                                                Dim lungime_pag As Double = 852
                                                Dim Diferenta_match As Double = Match2 - Match1
                                                Dim Factor As Double = lungime_pag / Diferenta_match

                                                Dim X0 As Double
                                                If RadioButton_left_to_right.Checked = True Then
                                                    X0 = 112
                                                    Pct_jos = New Point3d(X0 + Factor * diferenta_from_start, YSTRESS_jos, Z)
                                                    Pct_sus = New Point3d(X0 + Factor * diferenta_from_start, YSTRESS_sus, Z)
                                                Else
                                                    X0 = 964 - 28
                                                    Pct_jos = New Point3d(X0 - Factor * diferenta_from_start, YSTRESS_jos, Z)
                                                    Pct_sus = New Point3d(X0 - Factor * diferenta_from_start, YSTRESS_sus, Z)
                                                End If

                                                If Not Round(Data_table_Crossing.Rows(i).Item("STA"), 1) = 0 Then
                                                    Dim Line1 As New Line(Pct_sus, Pct_jos)
                                                    Line1.Layer = "TEXT"
                                                    BTrecord.AppendEntity(Line1)
                                                    Trans1.AddNewlyCreatedDBObject(Line1, True)

                                                    Dim Mtext1 As New MText
                                                    Mtext1.Layer = "TEXT"
                                                    Mtext1.Rotation = PI / 2
                                                    Mtext1.Contents = Get_chainage_from_double(Stress_station2, 1)
                                                    Mtext1.TextHeight = 3
                                                    Mtext1.Location = New Point3d(Pct_jos.X + 2, Pct_jos.Y + (Pct_sus.Y - Pct_jos.Y) / 2, 0)
                                                    Mtext1.Attachment = AttachmentPoint.TopCenter
                                                    BTrecord.AppendEntity(Mtext1)
                                                    Trans1.AddNewlyCreatedDBObject(Mtext1, True)

                                                End If










                                                Dim Pct_INS As New Point3d
                                                If Stress_station1 > Match1 Then
                                                    diferenta_from_start = Stress_station1 + (Stress_station2 - Stress_station1) * 0.5 - Match1
                                                Else
                                                    diferenta_from_start = (Stress_station2 - Match1) * 0.5
                                                End If


                                                If RadioButton_left_to_right.Checked = True Then

                                                    Pct_INS = New Point3d(X0 + Factor * diferenta_from_start, 32.5 + YSTRESS_jos + (YSTRESS_sus - YSTRESS_jos) / 2, Z)

                                                Else

                                                    Pct_INS = New Point3d(X0 - Factor * diferenta_from_start, 32.5 + YSTRESS_jos + (YSTRESS_sus - YSTRESS_jos) / 2, Z)

                                                End If



                                                Dim Mtext2 As New MText
                                                Mtext2.Layer = "TEXT"
                                                Mtext2.Rotation = PI / 2
                                                Mtext2.Contents = text_stress
                                                Mtext2.TextHeight = 2.5
                                                Mtext2.Location = Pct_INS
                                                Mtext2.Attachment = AttachmentPoint.MiddleCenter
                                                BTrecord.AppendEntity(Mtext2)
                                                Trans1.AddNewlyCreatedDBObject(Mtext2, True)

                                                Dim Pct_INS1 As New Point3d
                                                If RadioButton_left_to_right.Checked = True Then
                                                    Pct_INS1 = New Point3d(X0 + Factor * diferenta_from_start, -32.5 + YSTRESS_jos + (YSTRESS_sus - YSTRESS_jos) / 2, Z)
                                                Else
                                                    Pct_INS1 = New Point3d(X0 - Factor * diferenta_from_start, -32.5 + YSTRESS_jos + (YSTRESS_sus - YSTRESS_jos) / 2, Z)
                                                End If

                                                Dim Mtext3 As New MText
                                                Mtext3.Layer = "TEXT"
                                                Mtext3.Rotation = PI / 2
                                                Mtext3.Contents = Get_String_Rounded(Len_stress, 1)
                                                Mtext3.TextHeight = 2.5
                                                Mtext3.Location = Pct_INS1
                                                Mtext3.Attachment = AttachmentPoint.MiddleCenter
                                                BTrecord.AppendEntity(Mtext3)
                                                Trans1.AddNewlyCreatedDBObject(Mtext3, True)

                                                Len_stress = 0
                                                Stress_station1 = 0
                                                Stress_station2 = 0
                                                text_stress = ""
                                                Is_stress = False

                                            End If
                                        End If

                                    Case "SA1"
                                        If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then



                                            If IsDBNull(Data_table_Crossing.Rows(i).Item("DESCRIPTION1")) = False Then
                                                text_Buoyancy = Data_table_Crossing.Rows(i).Item("DESCRIPTION1").ToString

                                                Buoyancy_station1 = Data_table_Crossing.Rows(i).Item("STA")
                                                Buoyancy_station2 = Data_table_Crossing.Rows(i).Item("ENDSTA")
                                                Dim Pct_sus As New Point3d
                                                Dim Pct_jos As New Point3d

                                                Dim Ybuoyancy_sus As Double = 250.19
                                                Dim Ybuoyancy_jos As Double = 162.89

                                                Dim diferenta_from_start As Double = Round(Data_table_Crossing.Rows(i).Item("STA"), 1) - Match1
                                                Dim lungime_pag As Double = 852
                                                Dim Diferenta_match As Double = Match2 - Match1
                                                Dim Factor As Double = lungime_pag / Diferenta_match

                                                Dim X0 As Double
                                                If RadioButton_left_to_right.Checked = True Then
                                                    X0 = 112
                                                    Pct_jos = New Point3d(X0 + Factor * diferenta_from_start, Ybuoyancy_jos, Z)
                                                    Pct_sus = New Point3d(X0 + Factor * diferenta_from_start, Ybuoyancy_sus, Z)
                                                Else
                                                    X0 = 964 - 28
                                                    Pct_jos = New Point3d(X0 - Factor * diferenta_from_start, Ybuoyancy_jos, Z)
                                                    Pct_sus = New Point3d(X0 - Factor * diferenta_from_start, Ybuoyancy_sus, Z)
                                                End If

                                                If Not Round(Data_table_Crossing.Rows(i).Item("STA"), 1) = 0 Then
                                                    Dim Line1 As New Line(Pct_sus, Pct_jos)
                                                    Line1.Layer = "TEXT"
                                                    BTrecord.AppendEntity(Line1)
                                                    Trans1.AddNewlyCreatedDBObject(Line1, True)

                                                    Dim Mtext1 As New MText
                                                    Mtext1.Layer = "TEXT"
                                                    Mtext1.Rotation = PI / 2
                                                    Mtext1.Contents = Get_chainage_from_double(Buoyancy_station1, 1)
                                                    Mtext1.TextHeight = 3
                                                    Mtext1.Location = New Point3d(Pct_jos.X + 2, Pct_jos.Y + (Pct_sus.Y - Pct_jos.Y) / 2, 0)
                                                    Mtext1.Attachment = AttachmentPoint.TopCenter
                                                    BTrecord.AppendEntity(Mtext1)
                                                    Trans1.AddNewlyCreatedDBObject(Mtext1, True)

                                                End If




                                                Is_Buoyancy = True

                                            End If


                                        End If


                                    Case "SA2"

                                        If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                                            If Is_Buoyancy = True Then

                                                Dim Pct_sus As New Point3d
                                                Dim Pct_jos As New Point3d

                                                Dim Ybuoyancy_sus As Double = 250.19
                                                Dim Ybuoyancy_jos As Double = 162.89

                                                Dim diferenta_from_start As Double = Round(Data_table_Crossing.Rows(i).Item("STA"), 1) - Match1
                                                Dim lungime_pag As Double = 852
                                                Dim Diferenta_match As Double = Match2 - Match1
                                                Dim Factor As Double = lungime_pag / Diferenta_match

                                                Dim X0 As Double
                                                If RadioButton_left_to_right.Checked = True Then
                                                    X0 = 112
                                                    Pct_jos = New Point3d(X0 + Factor * diferenta_from_start, Ybuoyancy_jos, Z)
                                                    Pct_sus = New Point3d(X0 + Factor * diferenta_from_start, Ybuoyancy_sus, Z)
                                                Else
                                                    X0 = 964 - 28
                                                    Pct_jos = New Point3d(X0 - Factor * diferenta_from_start, Ybuoyancy_jos, Z)
                                                    Pct_sus = New Point3d(X0 - Factor * diferenta_from_start, Ybuoyancy_sus, Z)
                                                End If

                                                If Not Round(Data_table_Crossing.Rows(i).Item("STA"), 1) = 0 Then
                                                    Dim Line1 As New Line(Pct_sus, Pct_jos)
                                                    Line1.Layer = "TEXT"
                                                    BTrecord.AppendEntity(Line1)
                                                    Trans1.AddNewlyCreatedDBObject(Line1, True)

                                                    Dim Mtext1 As New MText
                                                    Mtext1.Layer = "TEXT"
                                                    Mtext1.Rotation = PI / 2
                                                    Mtext1.Contents = Get_chainage_from_double(Buoyancy_station2, 1)
                                                    Mtext1.TextHeight = 3
                                                    Mtext1.Location = New Point3d(Pct_jos.X + 2, Pct_jos.Y + (Pct_sus.Y - Pct_jos.Y) / 2, 0)
                                                    Mtext1.Attachment = AttachmentPoint.TopCenter
                                                    BTrecord.AppendEntity(Mtext1)
                                                    Trans1.AddNewlyCreatedDBObject(Mtext1, True)

                                                End If










                                                Dim Pct_INS As New Point3d
                                                If Buoyancy_station1 > Match1 Then
                                                    diferenta_from_start = Buoyancy_station1 + (Buoyancy_station2 - Buoyancy_station1) * 0.5 - Match1
                                                Else
                                                    diferenta_from_start = (Buoyancy_station2 - Match1) * 0.5
                                                End If


                                                If RadioButton_left_to_right.Checked = True Then

                                                    Pct_INS = New Point3d(X0 + Factor * diferenta_from_start, 32.5 + Ybuoyancy_jos + (Ybuoyancy_sus - Ybuoyancy_jos) / 2, Z)

                                                Else

                                                    Pct_INS = New Point3d(X0 - Factor * diferenta_from_start, 32.5 + Ybuoyancy_jos + (Ybuoyancy_sus - Ybuoyancy_jos) / 2, Z)

                                                End If



                                                Dim Mtext2 As New MText
                                                Mtext2.Layer = "TEXT"
                                                Mtext2.Rotation = PI / 2
                                                Mtext2.Contents = text_Buoyancy

                                                Mtext2.TextHeight = 2.5
                                                Mtext2.Location = Pct_INS
                                                Mtext2.Attachment = AttachmentPoint.MiddleCenter
                                                BTrecord.AppendEntity(Mtext2)
                                                Trans1.AddNewlyCreatedDBObject(Mtext2, True)
                                                text_Buoyancy = ""
                                                Buoyancy_station1 = 0
                                                Buoyancy_station2 = 0
                                                Is_Buoyancy = False

                                            End If
                                        End If


                                    Case "MATCHLINE"
                                        If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                                            Dim Test1 As Double = Data_table_Crossing.Rows(i).Item("STA")

                                            If Is_class_location = True Then
                                                Dim diferenta_from_start As Double
                                                Dim DIF_CLASS As Double
                                                If Class_station1 >= Match1 Then
                                                    DIF_CLASS = Match2 - Class_station1
                                                    diferenta_from_start = (Class_station1 + DIF_CLASS / 2) - Match1
                                                Else
                                                    DIF_CLASS = Match2 - Match1
                                                    diferenta_from_start = DIF_CLASS / 2
                                                End If


                                                Dim lungime_pag = 852
                                                Dim Diferenta_match As Double = Match2 - Match1
                                                Dim Factor As Double = lungime_pag / Diferenta_match
                                                Dim X0 As Double
                                                Dim Pct_INS As New Point3d
                                                Dim Yclass_sus As Double = 512.1
                                                Dim Yclass_jos As Double = 424.8

                                                If RadioButton_left_to_right.Checked = True Then
                                                    X0 = 112
                                                    Pct_INS = New Point3d(X0 + Factor * diferenta_from_start, Yclass_jos + (Yclass_sus - Yclass_jos) / 2, Z)
                                                Else
                                                    X0 = 964 - 28
                                                    Pct_INS = New Point3d(X0 - Factor * diferenta_from_start, Yclass_jos + (Yclass_sus - Yclass_jos) / 2, Z)
                                                End If

                                                Dim Mtext5 As New MText
                                                Mtext5.Layer = "TEXT"
                                                Mtext5.Rotation = 0
                                                Mtext5.Contents = text_class_location
                                                Mtext5.TextHeight = 5
                                                Mtext5.Location = Pct_INS
                                                Mtext5.Attachment = AttachmentPoint.MiddleCenter
                                                BTrecord.AppendEntity(Mtext5)
                                                Trans1.AddNewlyCreatedDBObject(Mtext5, True)



                                            End If

                                            If Is_hydro = True Then
                                                Dim diferenta_from_start As Double
                                                Dim DIF_hydro As Double
                                                If Hydro_station1 >= Match1 Then
                                                    DIF_hydro = Match2 - Hydro_station1
                                                    diferenta_from_start = (Hydro_station1 + DIF_hydro / 2) - Match1
                                                Else
                                                    DIF_hydro = Match2 - Match1
                                                    diferenta_from_start = DIF_hydro / 2
                                                End If


                                                Dim lungime_pag = 852
                                                Dim Diferenta_match As Double = Match2 - Match1
                                                Dim Factor As Double = lungime_pag / Diferenta_match
                                                Dim X0 As Double
                                                Dim Pct_INS As New Point3d
                                                Dim Yhydro_sus As Double = 424.8
                                                Dim Yhydro_jos As Double = 337.5

                                                If RadioButton_left_to_right.Checked = True Then
                                                    X0 = 112
                                                    Pct_INS = New Point3d(X0 + Factor * diferenta_from_start, Yhydro_jos + (Yhydro_sus - Yhydro_jos) / 2, Z)
                                                Else
                                                    X0 = 964 - 28
                                                    Pct_INS = New Point3d(X0 - Factor * diferenta_from_start, Yhydro_jos + (Yhydro_sus - Yhydro_jos) / 2, Z)
                                                End If

                                                Dim Mtext5 As New MText
                                                Mtext5.Layer = "TEXT"
                                                Mtext5.Rotation = 0
                                                Mtext5.Contents = text_hydrostation
                                                Mtext5.TextHeight = 5
                                                Mtext5.Location = Pct_INS
                                                Mtext5.Attachment = AttachmentPoint.MiddleCenter
                                                BTrecord.AppendEntity(Mtext5)
                                                Trans1.AddNewlyCreatedDBObject(Mtext5, True)



                                            End If



                                            If Is_stress = True Then
                                                Dim diferenta_from_start As Double
                                                Dim DIF_STRESS As Double
                                                If Stress_station1 >= Match1 Then
                                                    DIF_STRESS = Match2 - Stress_station1
                                                    diferenta_from_start = (Stress_station1 + DIF_STRESS / 2) - Match1
                                                Else
                                                    DIF_STRESS = Match2 - Match1
                                                    diferenta_from_start = DIF_STRESS / 2
                                                End If


                                                Dim lungime_pag = 852
                                                Dim Diferenta_match As Double = Match2 - Match1
                                                Dim Factor As Double = lungime_pag / Diferenta_match
                                                Dim X0 As Double
                                                Dim Pct_INS As New Point3d
                                                Dim YSTRESS_sus As Double = 337.5
                                                Dim YSTRESS_jos As Double = 250.19

                                                If RadioButton_left_to_right.Checked = True Then
                                                    X0 = 112
                                                    Pct_INS = New Point3d(X0 + Factor * diferenta_from_start, 32.5 + YSTRESS_jos + (YSTRESS_sus - YSTRESS_jos) / 2, Z)
                                                Else
                                                    X0 = 964 - 28
                                                    Pct_INS = New Point3d(X0 - Factor * diferenta_from_start, 32.5 + YSTRESS_jos + (YSTRESS_sus - YSTRESS_jos) / 2, Z)
                                                End If

                                                Dim Mtext5 As New MText
                                                Mtext5.Layer = "TEXT"
                                                Mtext5.Rotation = PI / 2
                                                Mtext5.Contents = text_stress
                                                Mtext5.TextHeight = 2.5
                                                Mtext5.Location = Pct_INS
                                                Mtext5.Attachment = AttachmentPoint.MiddleCenter
                                                BTrecord.AppendEntity(Mtext5)
                                                Trans1.AddNewlyCreatedDBObject(Mtext5, True)

                                                Dim Pct_INS1 As New Point3d
                                                If RadioButton_left_to_right.Checked = True Then
                                                    Pct_INS1 = New Point3d(X0 + Factor * diferenta_from_start, -32.5 + YSTRESS_jos + (YSTRESS_sus - YSTRESS_jos) / 2, Z)
                                                Else
                                                    Pct_INS1 = New Point3d(X0 - Factor * diferenta_from_start, -32.5 + YSTRESS_jos + (YSTRESS_sus - YSTRESS_jos) / 2, Z)
                                                End If

                                                Dim Mtext33 As New MText
                                                Mtext33.Layer = "TEXT"
                                                Mtext33.Rotation = PI / 2
                                                Mtext33.Contents = Get_String_Rounded(Len_stress, 1)
                                                Mtext33.TextHeight = 2.5
                                                Mtext33.Location = Pct_INS1
                                                Mtext33.Attachment = AttachmentPoint.MiddleCenter
                                                BTrecord.AppendEntity(Mtext33)
                                                Trans1.AddNewlyCreatedDBObject(Mtext33, True)


                                            End If

                                            If Is_Buoyancy = True Then
                                                Dim diferenta_from_start As Double
                                                Dim DIF_Buoyancy As Double
                                                If Buoyancy_station1 >= Match1 Then
                                                    DIF_Buoyancy = Match2 - Buoyancy_station1
                                                    diferenta_from_start = (Buoyancy_station1 + DIF_Buoyancy / 2) - Match1
                                                Else
                                                    DIF_Buoyancy = Match2 - Match1
                                                    diferenta_from_start = DIF_Buoyancy / 2
                                                End If


                                                Dim lungime_pag = 852
                                                Dim Diferenta_match As Double = Match2 - Match1
                                                Dim Factor As Double = lungime_pag / Diferenta_match
                                                Dim X0 As Double
                                                Dim Pct_INS As New Point3d
                                                Dim YBuoyancy_sus As Double = 250.19
                                                Dim YBuoyancy_jos As Double = 162.89

                                                If RadioButton_left_to_right.Checked = True Then
                                                    X0 = 112
                                                    Pct_INS = New Point3d(X0 + Factor * diferenta_from_start, 32.5 + YBuoyancy_jos + (YBuoyancy_sus - YBuoyancy_jos) / 2, Z)
                                                Else
                                                    X0 = 964 - 28
                                                    Pct_INS = New Point3d(X0 - Factor * diferenta_from_start, 32.5 + YBuoyancy_jos + (YBuoyancy_sus - YBuoyancy_jos) / 2, Z)
                                                End If

                                                Dim Mtext5 As New MText
                                                Mtext5.Layer = "TEXT"
                                                Mtext5.Rotation = PI / 2
                                                Mtext5.Contents = text_Buoyancy
                                                Mtext5.TextHeight = 2.5
                                                Mtext5.Location = Pct_INS
                                                Mtext5.Attachment = AttachmentPoint.MiddleCenter
                                                BTrecord.AppendEntity(Mtext5)
                                                Trans1.AddNewlyCreatedDBObject(Mtext5, True)
                                            End If



                                            If Nr_pagina > 1 Then
                                                Match1 = Data_table_Matchline(Nr_pagina - 1).Item("STATION")
                                            End If

                                            If Nr_pagina <= Data_table_Matchline.Rows.Count - 1 Then
                                                Match2 = Data_table_Matchline(Nr_pagina).Item("STATION")
                                            End If







                                            Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current

                                            Dim exista_layoutul As Boolean = False
                                            Dim Layoutdict As DBDictionary

                                            Layoutdict = Trans1.GetObject(ThisDrawing.Database.LayoutDictionaryId, OpenMode.ForRead)
                                            Dim nr_layouts As Integer
                                            nr_layouts = Layoutdict.Count

                                            For Each entry As DBDictionaryEntry In Layoutdict
                                                If entry.Key = Numele_template_layout Then
                                                    exista_layoutul = True
                                                    Exit For
                                                End If
                                            Next
                                            Dim Index_Layout As Integer = nr_layouts
                                            Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(Numele_template_layout), OpenMode.ForRead)
                                            Dim Nume_nou As String = (Nr_pagina).ToString

                                            If exista_layoutul = True Then


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

                                            End If

                                            Index_Layout = Index_Layout + 1


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


                                            BTrecord = Trans1.GetObject(BlockTable1(BlockTableRecord.PaperSpace), OpenMode.ForWrite)
                                            If RadioButton_left_to_right.Checked = True Then
                                                X = 112 + 28
                                            Else
                                                X = 964 - 28
                                            End If

                                            Ycr = 599.41




                                            Dim Mtext1 As New MText
                                            Mtext1.Layer = "TEXT"
                                            Mtext1.Rotation = PI / 2
                                            If RadioButton_left_to_right.Checked = True Then
                                                Mtext1.Contents = Get_chainage_from_double(Match1, 1)
                                            Else
                                                Mtext1.Contents = Get_chainage_from_double(Match2, 1)
                                            End If
                                            Mtext1.TextHeight = 5
                                            Mtext1.Location = New Point3d(98.92, 661.14, 0)
                                            Mtext1.Attachment = AttachmentPoint.TopCenter
                                            BTrecord.AppendEntity(Mtext1)
                                            Trans1.AddNewlyCreatedDBObject(Mtext1, True)

                                            Dim Mtext2 As New MText
                                            Mtext2.Layer = "TEXT"
                                            Mtext2.Rotation = PI / 2

                                            If RadioButton_left_to_right.Checked = True Then
                                                Mtext2.Contents = Get_chainage_from_double(Match2, 1)
                                            Else
                                                Mtext2.Contents = Get_chainage_from_double(Match1, 1)
                                            End If

                                            Mtext2.TextHeight = 5
                                            Mtext2.Location = New Point3d(978, 661.14, 0)
                                            Mtext2.Attachment = AttachmentPoint.BottomCenter
                                            BTrecord.AppendEntity(Mtext2)
                                            Trans1.AddNewlyCreatedDBObject(Mtext2, True)

                                            Dim Mtext3 As New MText
                                            Mtext3.Layer = "TEXT"
                                            Mtext3.Rotation = 0
                                            Mtext3.Contents = Nr_pagina.ToString
                                            Mtext3.TextHeight = 6
                                            Mtext3.Location = New Point3d(951, 25.79, 0)
                                            Mtext3.Attachment = AttachmentPoint.MiddleLeft
                                            BTrecord.AppendEntity(Mtext3)
                                            Trans1.AddNewlyCreatedDBObject(Mtext3, True)

                                            Dim Mtext4 As New MText
                                            Mtext4.Layer = "TEXT"
                                            Mtext4.Rotation = 0

                                            If RadioButton_left_to_right.Checked = True Then
                                                Mtext4.Contents = Mtext1.Text & " - " & Mtext2.Text
                                            Else
                                                Mtext4.Contents = Mtext2.Text & " - " & Mtext1.Text
                                            End If
                                            Mtext4.TextHeight = 5
                                            Mtext4.Location = New Point3d(853.5, 42.27, 0)
                                            Mtext4.Attachment = AttachmentPoint.MiddleLeft
                                            BTrecord.AppendEntity(Mtext4)
                                            Trans1.AddNewlyCreatedDBObject(Mtext4, True)



                                            










                                        End If '  If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False



                                End Select

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
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)




    End Sub

    Public Function Adauga_textbox(ByVal x As Double, ByVal y As Double, ByVal Len As Double, ByVal Height As Double, ByVal Continut As String, ByVal Readonly1 As Boolean, ByVal Multiline As Boolean, ByVal BackColor As Drawing.Color) As Windows.Forms.TextBox
        Try
            Dim TextBox1 As New Windows.Forms.TextBox
            TextBox1.Location = New System.Drawing.Point(x, y)
            Dim myfont As New System.Drawing.Font("Arial", 9, Drawing.FontStyle.Bold)
            TextBox1.Font = myfont
            TextBox1.Size = New System.Drawing.Size(Len, Height)
            TextBox1.Text = Continut
            TextBox1.ReadOnly = Readonly1
            TextBox1.Multiline = Multiline
            TextBox1.BackColor = BackColor
            Return TextBox1
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    Private Sub Button_load_deflections()
        Try



            ascunde_butoanele_pentru_forms(Me, Colectie1)

            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
            Dim start1 As Integer
            Dim end1 As Integer

            Dim Col1 As String
            Dim Col2 As String

            Data_table_Elbows.Rows.Clear()
            Dim Index_Data_table As Integer = 0

            For i = start1 To end1
                Dim Chainage1 As String = Replace(W1.Range(Col1 & i).Value, "+", "")
                Dim Angle As String = W1.Range(Col2 & i).Value
                Dim CHAINAGE As Double = 0
                If IsNumeric(Chainage1) = True Then
                    CHAINAGE = CDbl(Chainage1)
                End If
                If CHAINAGE > 0 Then
                    Data_table_Elbows.Rows.Add()
                    Data_table_Elbows.Rows(Index_Data_table).Item("STATION") = CHAINAGE
                    Data_table_Elbows.Rows(Index_Data_table).Item("ANGLE") = Angle
                    'Data_table_Elbows.Rows(Index_Data_table).Item("MATCHED") = True
                    If IsNumeric(ComboBox_nps.Text) = True And IsNumeric(Angle) = True Then
                        Dim Diam As Double = 2 * Get_from_NPS_radius_for_pipes_from_inches_to_milimeters(CDbl(ComboBox_nps.Text)) / 1000
                        Data_table_Elbows.Rows(Index_Data_table).Item("LENGTH") = Round(2 * (3 * Diam * Tan((CDbl(Angle) * PI / 180) / 2) + 1), 1)
                        Data_table_Elbows.Rows(Index_Data_table).Item("VALUE") = Angle
                    Else
                        MsgBox("At " & Get_chainage_from_double(CHAINAGE, 1) & "you have an issue - no Pipe NPS or no angle specified for elbow")
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        Exit Sub
                    End If
                    'TextBox_deflection_station.BackColor = Drawing.Color.Yellow
                    'TextBox_deflection_value.BackColor = Drawing.Color.Yellow
                    Index_Data_table = Index_Data_table + 1
                End If
            Next

            adauga_ELBOWS_la_crossing_list()

            afiseaza_butoanele_pentru_forms(Me, Colectie1)
        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_RESET_Click(sender As Object, e As EventArgs) Handles Button_RESET.Click
        Data_table_Elbows.Rows.Clear()
        Data_table_Matchline.Rows.Clear()
        Data_table_Class_location.Rows.Clear()
        Panel_crossings.Controls.Clear()
        TextBox_col_1.BackColor = Drawing.Color.White
        TextBox_col_2.BackColor = Drawing.Color.White
        TextBox_col_3.BackColor = Drawing.Color.White
        TextBox_col_4.BackColor = Drawing.Color.White
        TextBox_col_5.BackColor = Drawing.Color.White

        If Data_table_Crossing.Rows.Count > 0 Then
            For i = 0 To Data_table_Crossing.Rows.Count - 1
                If IsDBNull(Data_table_Crossing.Rows(i).Item("BLOCKNAME")) = False Then
                    If Data_table_Crossing.Rows(i).Item("BLOCKNAME") = "ELBOW" Then
                        Data_table_Crossing.Rows(i).Item("BLOCKNAME") = DBNull.Value
                        If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                            Data_table_Crossing.Rows(i).Item("STA") = DBNull.Value
                        End If
                        If Data_table_Crossing.Rows(i).Item("BLOCKNAME") = "MATCHLINE" Then
                            Data_table_Crossing.Rows(i).Item("BLOCKNAME") = DBNull.Value
                            If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                                Data_table_Crossing.Rows(i).Item("STA") = DBNull.Value
                            End If
                        End If
                        If Data_table_Crossing.Rows(i).Item("BLOCKNAME") = "CLASS_LOCATION1" Or Data_table_Crossing.Rows(i).Item("BLOCKNAME") = "CLASS_LOCATION2" Then
                            Data_table_Crossing.Rows(i).Item("BLOCKNAME") = DBNull.Value
                            If IsDBNull(Data_table_Crossing.Rows(i).Item("STA")) = False Then
                                Data_table_Crossing.Rows(i).Item("STA") = DBNull.Value
                            End If
                        End If
                    End If
                End If
            Next
            Data_table_Crossing = delete_DBnull_rows_from_data_table(Data_table_Crossing, "STA")
        End If
    End Sub

End Class