Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class Engineering_band_form
    Dim Empty_array() As ObjectId
    Dim Data_table_station_equation As System.Data.DataTable
    Dim Data_table_materials As System.Data.DataTable
    Dim Data_table_elbows As System.Data.DataTable
    Dim Data_table_matchlines As System.Data.DataTable
    Dim Data_table_transitions As System.Data.DataTable
    Dim Data_table_buoyancy_canada_without_matchlines As System.Data.DataTable
    Dim Data_table_buoyancy_canada As System.Data.DataTable
    Dim Data_table_pipes_Canada As System.Data.DataTable
    Dim Data_table_water_canada As System.Data.DataTable
    Dim Data_table_cathodic_canada As System.Data.DataTable

    Dim Data_table_As_built_Pipe_ID As System.Data.DataTable

    Dim Data_table_point_description As System.Data.DataTable

    Dim Data_table_compiled As System.Data.DataTable


    Dim Start_point As New Point3d(50000, -1000, 0)
    Dim Stretch_scale_factor As Double = 1


    Dim Data_table_read_pipe_tally As System.Data.DataTable
    Dim Data_table_read_materials As System.Data.DataTable
    Dim Data_table_read_all_points As System.Data.DataTable
    Dim Data_table_read_fitting As System.Data.DataTable

    Dim Data_table_fitting As System.Data.DataTable
    Dim Data_table_cl_crossing As System.Data.DataTable
    Dim Data_table_cad_weld As System.Data.DataTable
    Dim Data_table_river_weights As System.Data.DataTable

    Dim Data_table_material_count As System.Data.DataTable
    Dim Data_table_material_results As System.Data.DataTable

    Dim PolyCL As Polyline
    Dim PolyCL3D As Polyline3d
    Dim Poly_length As Double

    Dim Data_table_bends As System.Data.DataTable

    Dim Freeze_operations As Boolean = False

    Private Sub Engineering_band_form_Load(sender As Object, e As EventArgs) Handles Me.Load
        TextBox_CSF.Text = "1"
        Panel_Blocks_Click(sender, e)
    End Sub

    Private Sub Panel_Blocks_Click(sender As Object, e As EventArgs) Handles _
        Panel_Blocks_mat.Click, Panel_Blocks_ellbow_as_mat.Click,
        Panel_elbow_as_crossing.Click, Panel_crossing_no_mat.Click,
        Panel_cathodic_protection.Click,
        PaneL_TRANSITION_weld.Click,
        Panel_screw_anchors_multiple.Click, Panel_pipes.Click, Panel43.Click, Panel_mat_count.Click, Panel2.Click, Panel45.Click, Panel44.Click, Panel46.Click

        Incarca_existing_Blocks_with_attributes_to_combobox(ComboBox_blocks_ELBOWS_AS_MAT)
        Incarca_existing_Blocks_with_attributes_to_combobox(ComboBox_blocks_MAT)
        Incarca_existing_Blocks_with_attributes_to_combobox(ComboBox_blocks_fitting)
        Incarca_existing_Blocks_with_attributes_to_combobox(ComboBox_blocks_CL_crossing)
        Incarca_existing_Blocks_with_attributes_to_combobox(ComboBox_blocks_cad_weld)
        Incarca_existing_Blocks_to_combobox(ComboBox_blocks_transition_weld_canada)
        Incarca_existing_Blocks_with_attributes_to_combobox(ComboBox_blocks_screw_anchor_multiple_canada)

        Incarca_existing_Blocks_with_attributes_to_combobox(ComboBox_blocks_pipe_crossings_canada)
        Incarca_existing_Blocks_with_attributes_to_combobox(ComboBox_blocks_RIVER_WEIGHTS)
        Incarca_existing_Blocks_with_attributes_to_combobox(ComboBox_blocks_mat_count)
        Incarca_existing_Blocks_with_attributes_to_combobox(ComboBox_blocks_count1)
        Incarca_existing_Blocks_with_attributes_to_combobox(ComboBox_blocks_count2)
        Incarca_existing_Blocks_with_attributes_to_combobox(ComboBox_blocks_count3)
        Incarca_existing_Blocks_with_attributes_to_combobox(ComboBox_blocks_count4)
    End Sub

    Private Sub ComboBox_blocks_ELBOWS_AS_MAT_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_blocks_ELBOWS_AS_MAT.SelectedIndexChanged
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_ELBOWS_AS_MAT.Text, ComboBox_STA1_att_ELBOW_AS_MAT)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_ELBOWS_AS_MAT.Text, ComboBox_STA2_att_ELBOW_AS_MAT)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_ELBOWS_AS_MAT.Text, ComboBox_EXTRA_att_ELBOW_AS_MAT)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_ELBOWS_AS_MAT.Text, ComboBox_LEN_att_ELBOW_AS_MAT)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_ELBOWS_AS_MAT.Text, ComboBox_MAT_ELBOW_AS_MAT)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_ELBOWS_AS_MAT.Text, ComboBox_dwg_ID_ELBOW_AS_MAT)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_ELBOWS_AS_MAT.Text, ComboBox_descr_ELBOW_AS_MAT)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_ELBOWS_AS_MAT.Text, ComboBox_station_middle_att_ELBOW_AS_MAT)

    End Sub

    Private Sub ComboBox_blocks_mat_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_blocks_MAT.SelectedIndexChanged
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_MAT.Text, ComboBox_mat_STA1)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_MAT.Text, ComboBox_mat_STA2)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_MAT.Text, ComboBox_mat_STA1_duplicate)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_MAT.Text, ComboBox_mat_STA2_duplicate)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_MAT.Text, ComboBox_mat_len)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_MAT.Text, ComboBox_mat_mat)


        With ComboBox_mat_STA1
            If .Items.Contains("STA1") = True Then
                .SelectedIndex = .Items.IndexOf("STA1")
            End If
        End With
        With ComboBox_mat_STA2
            If .Items.Contains("STA2") = True Then
                .SelectedIndex = .Items.IndexOf("STA2")
            End If
        End With

        With ComboBox_mat_STA1_duplicate
            If .Items.Contains("STA11") = True Then
                .SelectedIndex = .Items.IndexOf("STA11")
            End If
        End With
        With ComboBox_mat_STA2_duplicate
            If .Items.Contains("STA22") = True Then
                .SelectedIndex = .Items.IndexOf("STA22")
            End If
        End With

        With ComboBox_mat_len
            If .Items.Contains("LEN") = True Then
                .SelectedIndex = .Items.IndexOf("LEN")
            End If
        End With

        With ComboBox_mat_mat
            If .Items.Contains("MAT") = True Then
                .SelectedIndex = .Items.IndexOf("MAT")
            End If
        End With

    End Sub

    Private Sub ComboBox_blocks_mat_count_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_blocks_mat_count.SelectedIndexChanged

        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_mat_count.Text, ComboBox_mc_STA1)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_mat_count.Text, ComboBox_mc_STA2)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_mat_count.Text, ComboBox_mc_LEN)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_mat_count.Text, ComboBox_mc_MAT)
        With ComboBox_mc_STA1
            If .Items.Contains("STA1") = True Then
                .SelectedIndex = .Items.IndexOf("STA1")
            End If
        End With
        With ComboBox_mc_STA2
            If .Items.Contains("STA2") = True Then
                .SelectedIndex = .Items.IndexOf("STA2")
            End If
        End With

        With ComboBox_mc_LEN
            If .Items.Contains("LEN") = True Then
                .SelectedIndex = .Items.IndexOf("LEN")
            End If
        End With

        With ComboBox_mc_MAT
            If .Items.Contains("MAT") = True Then
                .SelectedIndex = .Items.IndexOf("MAT")
            End If
        End With


    End Sub

    Private Sub ComboBox_blocks_ELBOWS_AS_crossing_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_blocks_fitting.SelectedIndexChanged
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_fitting.Text, ComboBox_STA_att_fitting)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_fitting.Text, ComboBox_DESCRIPTION_att_fitting)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_fitting.Text, ComboBox_MAT_att_fitting)
    End Sub

    Private Sub ComboBox_blocks_crossing_no_mat_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_blocks_CL_crossing.SelectedIndexChanged
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_CL_crossing.Text, ComboBox_STA_att_CL_crossing)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_CL_crossing.Text, ComboBox_DESCRIPTION_att_CL_crossing)
    End Sub

    Private Sub ComboBox_blocks_CATHODIC_PROTECTION_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_blocks_cad_weld.SelectedIndexChanged
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_cad_weld.Text, ComboBox_STA_att_cad_weld)
    End Sub

    Private Sub ComboBox_blocks_sta_count1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_blocks_count1.SelectedIndexChanged
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_count1.Text, ComboBox_atr_count1)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_count1.Text, ComboBox_atr_count1_MAT)
        With ComboBox_atr_count1_MAT
            If .Items.Contains("MAT") = True Then
                .SelectedIndex = .Items.IndexOf("MAT")
            End If
        End With
        With ComboBox_atr_count1
            If .Items.Contains("STA") = True Then
                .SelectedIndex = .Items.IndexOf("STA")
            End If
        End With
    End Sub

    Private Sub ComboBox_blocks_sta_count2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_blocks_count2.SelectedIndexChanged
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_count2.Text, ComboBox_atr_count2)

        With ComboBox_atr_count2
            If .Items.Contains("STA") = True Then
                .SelectedIndex = .Items.IndexOf("STA")
            Else
                If .Items.Contains("STATION") = True Then
                    .SelectedIndex = .Items.IndexOf("STATION")
                End If
            End If
        End With

    End Sub

    Private Sub ComboBox_blocks_sta_count3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_blocks_count3.SelectedIndexChanged
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_count3.Text, ComboBox_atr_count3)
        With ComboBox_atr_count3
            If .Items.Contains("STA") = True Then
                .SelectedIndex = .Items.IndexOf("STA")
            Else
                If .Items.Contains("STATION") = True Then
                    .SelectedIndex = .Items.IndexOf("STATION")
                End If
            End If
        End With
    End Sub

    Private Sub ComboBox_blocks_sta_count4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_blocks_count4.SelectedIndexChanged
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_count4.Text, ComboBox_atr_count4)
        With ComboBox_atr_count4
            If .Items.Contains("STA") = True Then
                .SelectedIndex = .Items.IndexOf("STA")
            Else
                If .Items.Contains("STATION") = True Then
                    .SelectedIndex = .Items.IndexOf("STATION")
                End If
            End If
        End With
    End Sub

    Private Sub ComboBox_blocks_river_weights_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_blocks_RIVER_WEIGHTS.SelectedIndexChanged
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_RIVER_WEIGHTS.Text, ComboBox_STA_att_river_weights)
    End Sub

    Private Sub ComboBox_blocks_screw_mult_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_blocks_screw_anchor_multiple_canada.SelectedIndexChanged
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_screw_anchor_multiple_canada.Text, ComboBox_att_screw_anch_mult_1)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_screw_anchor_multiple_canada.Text, ComboBox_att_screw_anch_mult_2)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_screw_anchor_multiple_canada.Text, ComboBox_att_screw_anch_mult_3)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_screw_anchor_multiple_canada.Text, ComboBox_att_screw_anch_mult_4)

    End Sub

    Private Sub ComboBox_blocks_pipe_crossings_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_blocks_pipe_crossings_canada.SelectedIndexChanged
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_pipe_crossings_canada.Text, ComboBox_att_pipe1)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_pipe_crossings_canada.Text, ComboBox_att_pipe2)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_pipe_crossings_canada.Text, ComboBox_att_pipe4)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_pipe_crossings_canada.Text, ComboBox_att_pipe3)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks_pipe_crossings_canada.Text, ComboBox_att_pipe5)
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

    Private Sub Button_load_materials_Click(sender As Object, e As EventArgs) Handles Button_load_materials.Click

        If Freeze_operations = False Then
            Freeze_operations = True

            Try

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Dim Start1 As Integer = 0
                Dim End1 As Integer = 0
                If IsNumeric(TextBox_ROW_START.Text) = True Then
                    Start1 = CInt(TextBox_ROW_START.Text)
                End If
                If IsNumeric(TextBox_ROW_END.Text) = True Then
                    End1 = CInt(TextBox_ROW_END.Text)
                End If

                If End1 = 0 Or Start1 = 0 Then
                    Freeze_operations = False
                    Exit Sub
                End If
                If End1 < Start1 Then
                    Freeze_operations = False
                    Exit Sub
                End If
                Dim Column_sta1 As String = ""
                Column_sta1 = TextBox_column_mat_Station_start.Text.ToUpper
                Dim Column_sta2 As String = ""
                Column_sta2 = TextBox_column_mat_Station_end.Text.ToUpper
                Dim Column_mat As String = ""
                Column_mat = TextBox_column_material.Text.ToUpper

                If Column_sta1 = "" Or Column_sta2 = "" Or Column_mat = "" Then
                    Freeze_operations = False
                    Exit Sub
                End If

                Data_table_compiled = New System.Data.DataTable


                Data_table_materials = New System.Data.DataTable
                Data_table_materials.Columns.Add("STA1", GetType(Double))
                Data_table_materials.Columns.Add("STA2", GetType(Double))
                Data_table_materials.Columns.Add("MAT", GetType(String))

                Dim Index_data_table As Double



                For i = Start1 To End1
                    Dim Station_string1 As String = W1.Range(Column_sta1 & i).Value2
                    Dim Station_string2 As String = W1.Range(Column_sta2 & i).Value2
                    Dim Material As String = W1.Range(Column_mat & i).Value2
                    If IsNumeric(Station_string1) = True And IsNumeric(Station_string2) = True And Not Material = "" Then
                        Data_table_materials.Rows.Add()
                        Data_table_materials.Rows(Index_data_table).Item("STA1") = CDbl(Station_string1)
                        Data_table_materials.Rows(Index_data_table).Item("STA2") = CDbl(Station_string2)
                        Data_table_materials.Rows(Index_data_table).Item("MAT") = Material
                        Index_data_table = Index_data_table + 1
                    End If

                Next


                Data_table_materials = Sort_data_table(Data_table_materials, "STA1")

                'MsgBox(Data_table_Centerline.Rows.Count)
                Add_to_clipboard_Data_table(Data_table_materials)


                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & Data_table_materials.Rows.Count & " MATERIALS loaded")
            Catch ex As Exception
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try

            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_load_elbows_Click(sender As Object, e As EventArgs) Handles Button_LOAD_Elbows.Click

        If Freeze_operations = False Then
            Freeze_operations = True

            Try

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Dim Start1 As Integer = 0
                Dim End1 As Integer = 0
                If IsNumeric(TextBox_ROW_START.Text) = True Then
                    Start1 = CInt(TextBox_ROW_START.Text)
                End If
                If IsNumeric(TextBox_ROW_END.Text) = True Then
                    End1 = CInt(TextBox_ROW_END.Text)
                End If

                If End1 = 0 Or Start1 = 0 Then
                    Freeze_operations = False
                    Exit Sub
                End If
                If End1 < Start1 Then
                    Freeze_operations = False
                    Exit Sub
                End If
                Dim Column_sta1 As String = ""
                Column_sta1 = TextBox_column_elbow_Station_start.Text.ToUpper
                Dim Column_sta2 As String = ""
                Column_sta2 = TextBox_column_elbow_Station_end.Text.ToUpper
                Dim Column_mat As String = ""
                Column_mat = TextBox_column_elbow_mat.Text.ToUpper

                Dim Column_sta As String = ""
                Column_sta = TextBox_column_elbow_Station_middle.Text.ToUpper

                If Column_sta1 = "" Or Column_sta2 = "" Then
                    Freeze_operations = False
                    Exit Sub
                End If





                Data_table_elbows = New System.Data.DataTable
                Data_table_elbows.Columns.Add("STA1", GetType(Double))
                Data_table_elbows.Columns.Add("STA2", GetType(Double))
                Data_table_elbows.Columns.Add("STA", GetType(Double))
                Data_table_elbows.Columns.Add("MAT", GetType(String))
                Data_table_elbows.Columns.Add("DESCR", GetType(String))
                Data_table_elbows.Columns.Add("ID", GetType(String))

                Dim Index_data_table As Double



                For i = Start1 To End1
                    Dim Station_string As String = W1.Range(Column_sta & i).Value2
                    Dim Station_string1 As String = W1.Range(Column_sta1 & i).Value2
                    Dim Station_string2 As String = W1.Range(Column_sta2 & i).Value2
                    Dim Material As String = W1.Range(Column_mat & i).Value2

                    If IsNumeric(Station_string1) = True And IsNumeric(Station_string2) = True Then
                        Data_table_elbows.Rows.Add()
                        Data_table_elbows.Rows(Index_data_table).Item("STA1") = CDbl(Station_string1)
                        Data_table_elbows.Rows(Index_data_table).Item("STA2") = CDbl(Station_string2)
                        If IsNumeric(Station_string) = True Then
                            Data_table_elbows.Rows(Index_data_table).Item("STA") = CDbl(Station_string)
                        End If
                        If Not Material = "" Then
                            Data_table_elbows.Rows(Index_data_table).Item("MAT") = Material
                        End If
                        Index_data_table = Index_data_table + 1
                    End If

                Next


                Data_table_elbows = Sort_data_table(Data_table_elbows, "STA1")


                Add_to_clipboard_Data_table(Data_table_elbows)


            Catch ex As Exception
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try

            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_read_centerline_and_viewports_Click(sender As Object, e As EventArgs) Handles Button_read_CL_VIEWPORTS_EXCEL.Click, Button_read_matchlines_canada.Click

        If Freeze_operations = False Then
            Freeze_operations = True
            Dim Round1 As Integer = 0
            If IsNumeric(TextBox_dec1.Text) = True Then
                Round1 = CInt(TextBox_dec1.Text)
            End If

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor


            Editor1.SetImpliedSelection(Empty_array)
            Try
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()



                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptEntityResult

                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select centerline:")

                Object_Prompt.SetRejectMessage(vbLf & "Please select a lightweight polyline or a 3d polyline")
                Object_Prompt.AddAllowedClass(GetType(Polyline), True)
                Object_Prompt.AddAllowedClass(GetType(Polyline3d), True)


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


                                Dim PolyCL_for_viewports As Polyline = TryCast(Trans1.GetObject(Rezultat1.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Polyline)
                                If IsNothing(PolyCL_for_viewports) = False Then
                                    Poly_length = PolyCL_for_viewports.Length
                                    PolyCL = PolyCL_for_viewports
                                End If
                                Dim PolyCL3D_for_viewports As Polyline3d
                                PolyCL3D_for_viewports = TryCast(Trans1.GetObject(Rezultat1.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Polyline3d)

                                If IsNothing(PolyCL3D_for_viewports) = False Then
                                    Poly_length = PolyCL3D_for_viewports.Length
                                    PolyCL3D = PolyCL3D_for_viewports
                                    Dim Index_Poly As Integer = 0
                                    PolyCL_for_viewports = New Polyline
                                    For Each vId As Autodesk.AutoCAD.DatabaseServices.ObjectId In PolyCL3D_for_viewports
                                        Dim v3d As Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d = DirectCast(Trans1.GetObject _
                                                (vId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d)

                                        Dim x1 As Double = v3d.Position.X
                                        Dim y1 As Double = v3d.Position.Y
                                        Dim z1 As Double = v3d.Position.Z
                                        PolyCL_for_viewports.AddVertexAt(Index_Poly, New Point2d(x1, y1), 0, 0, 0)
                                        Index_Poly = Index_Poly + 1
                                    Next
                                    PolyCL_for_viewports.Elevation = 0
                                End If

                                If Not PolyCL_for_viewports.Elevation = 0 Then
                                    Freeze_operations = False
                                    MsgBox("CL Polyline is not at elevation 0")
                                    Exit Sub

                                End If

                                Data_table_matchlines = New System.Data.DataTable
                                Data_table_matchlines.Columns.Add("STATION1", GetType(Double))
                                Data_table_matchlines.Columns.Add("STATION2", GetType(Double))
                                Data_table_matchlines.Columns.Add("X1", GetType(Double))
                                Data_table_matchlines.Columns.Add("Y1", GetType(Double))
                                Data_table_matchlines.Columns.Add("X2", GetType(Double))
                                Data_table_matchlines.Columns.Add("Y2", GetType(Double))
                                Data_table_matchlines.Columns.Add("ML_LEN", GetType(Double))


                                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                                Dim LayerTable1 As Autodesk.AutoCAD.DatabaseServices.LayerTable
                                LayerTable1 = Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                Dim Index_dataT As Double = 0


                                For Each objID As ObjectId In BTrecord
                                    Dim Rectangle_poly As Entity = Trans1.GetObject(objID, OpenMode.ForRead)

                                    Dim Executa As Boolean = False
                                    If TypeOf Rectangle_poly Is Polyline Then
                                        If Not Rectangle_poly.ObjectId = PolyCL_for_viewports.ObjectId Then
                                            If IsNothing(PolyCL3D_for_viewports) = False Then
                                                If Not Rectangle_poly.ObjectId = PolyCL3D_for_viewports.ObjectId Then
                                                    Executa = True
                                                End If
                                            Else
                                                Executa = True
                                            End If
                                        End If
                                    End If

                                    If Executa = True Then
                                        Dim Viewport_poly As Polyline = Rectangle_poly
                                        Dim LayerTableRecord1 As Autodesk.AutoCAD.DatabaseServices.LayerTableRecord
                                        LayerTableRecord1 = LayerTable1(Viewport_poly.Layer).GetObject(OpenMode.ForRead)


                                        If Viewport_poly.NumberOfVertices >= 3 And LayerTableRecord1.IsOff = False And LayerTableRecord1.IsFrozen = False Then
                                            Viewport_poly.UpgradeOpen()
                                            Viewport_poly.Elevation = 0
                                            Dim Col_int As New Point3dCollection
                                            Col_int = Intersect_on_both_operands(PolyCL_for_viewports, Viewport_poly)
                                            If Col_int.Count = 2 Then
                                                Dim Station1 As Double = 0
                                                Dim Station2 As Double = 0
                                                Dim Este_zero As Boolean = False
                                                Dim Nr_values As Integer = Col_int.Count
                                                If Nr_values > 2 Then Nr_values = 2


                                                Dim Point_on_poly1 As New Point3d()
                                                Point_on_poly1 = Col_int(0)
                                                Station1 = PolyCL_for_viewports.GetDistAtPoint(Point_on_poly1)

                                                If Round(Station1, 0) = 0 Then
                                                    Este_zero = True
                                                End If

                                                If Not Round(Station1, Round1) = Round(PolyCL_for_viewports.Length, Round1) Then
                                                    If IsNothing(PolyCL3D_for_viewports) = False Then
                                                        Dim Param1 As Double = PolyCL_for_viewports.GetParameterAtPoint(Point_on_poly1)
                                                        Station1 = PolyCL3D_for_viewports.GetDistanceAtParameter(Param1)
                                                    End If
                                                Else
                                                    If IsNothing(PolyCL3D_for_viewports) = False Then
                                                        Station1 = PolyCL3D_for_viewports.Length
                                                    End If
                                                End If

                                                Dim Point_on_poly2 As New Point3d()
                                                Point_on_poly2 = Col_int(1)
                                                Station2 = PolyCL_for_viewports.GetDistAtPoint(Point_on_poly2)
                                                If Not Round(Station2, Round1) = Round(PolyCL_for_viewports.Length, Round1) Then
                                                    If IsNothing(PolyCL3D_for_viewports) = False Then
                                                        Dim Param2 As Double = PolyCL_for_viewports.GetParameterAtPoint(Point_on_poly2)
                                                        Station2 = PolyCL3D_for_viewports.GetDistanceAtParameter(Param2)
                                                    End If
                                                Else
                                                    If IsNothing(PolyCL3D_for_viewports) = False Then
                                                        Station2 = PolyCL3D_for_viewports.Length
                                                    End If
                                                End If



                                                Dim Linie1 As New Line(Point_on_poly1, Point_on_poly2)

                                                Dim vpoly_exploded As New DBObjectCollection
                                                Viewport_poly.Explode(vpoly_exploded)



                                                If Station1 > Station2 Then
                                                    Dim T As Double = Station1
                                                    Station1 = Station2
                                                    Station2 = T
                                                End If

                                                If IsNothing(PolyCL3D_for_viewports) = False Then

                                                Else
                                                    If Station2 > PolyCL_for_viewports.Length Then
                                                        Station2 = PolyCL_for_viewports.Length
                                                    End If
                                                End If



                                                Dim X11, Y11, X21, Y22 As Double
                                                Dim LimitL As Double = 100
                                                For Each Ent4 As Entity In vpoly_exploded
                                                    If TypeOf (Ent4) Is Line Then
                                                        Dim Line2 As Line
                                                        Line2 = Ent4
                                                        If Line2.Length > LimitL Then
                                                            Dim Col_int1 As New Point3dCollection
                                                            Line2.IntersectWith(Linie1, Intersect.OnBothOperands, Col_int1, IntPtr.Zero, IntPtr.Zero)

                                                            If IsNothing(Col_int1) = True Then
                                                                Dim Pt1 As New Point3d
                                                                Dim Pt2 As New Point3d
                                                                Pt1 = Line2.GetClosestPointTo(Point_on_poly1, Vector3d.ZAxis, True)
                                                                Pt2 = Line2.GetClosestPointTo(Point_on_poly2, Vector3d.ZAxis, True)
                                                                If Pt1.GetVectorTo(Pt2).Length > LimitL Then
                                                                    X11 = Pt1.X
                                                                    Y11 = Pt1.Y
                                                                    X21 = Pt2.X
                                                                    Y22 = Pt2.Y
                                                                End If
                                                            Else
                                                                If Col_int1.Count = 0 Then
                                                                    Dim Pt1 As New Point3d
                                                                    Dim Pt2 As New Point3d
                                                                    Pt1 = Line2.GetClosestPointTo(Point_on_poly1, Vector3d.ZAxis, True)
                                                                    Pt2 = Line2.GetClosestPointTo(Point_on_poly2, Vector3d.ZAxis, True)
                                                                    If Pt1.GetVectorTo(Pt2).Length > LimitL Then
                                                                        X11 = Pt1.X
                                                                        Y11 = Pt1.Y
                                                                        X21 = Pt2.X
                                                                        Y22 = Pt2.Y
                                                                    End If

                                                                End If
                                                            End If



                                                        End If


                                                    End If

                                                Next

                                                Dim Lungime_matchline As Double = 0
                                                For i = 0 To Viewport_poly.NumberOfVertices - 2
                                                    Dim Pt1 As Point2d = Viewport_poly.GetPoint2dAt(i)
                                                    Dim Pt2 As Point2d = Viewport_poly.GetPoint2dAt(i + 1)
                                                    Dim Len1 As Double = Pt1.GetVectorTo(Pt2).Length
                                                    If Len1 > Lungime_matchline Then Lungime_matchline = Len1

                                                Next
                                                Dim pt3 As Point2d = Viewport_poly.GetPoint2dAt(Viewport_poly.NumberOfVertices - 1)
                                                Dim Pt0 As Point2d = Viewport_poly.GetPoint2dAt(0)
                                                Dim Len0 As Double = Pt0.GetVectorTo(pt3).Length
                                                If Len0 > Lungime_matchline Then Lungime_matchline = Len0



                                                Data_table_matchlines.Rows.Add()
                                                Data_table_matchlines.Rows(Index_dataT).Item("STATION1") = Round(Station1, Round1)
                                                Data_table_matchlines.Rows(Index_dataT).Item("STATION2") = Round(Station2, Round1)
                                                Data_table_matchlines.Rows(Index_dataT).Item("X1") = X11 'Viewport_poly.GetPointAtParameter(3).X
                                                Data_table_matchlines.Rows(Index_dataT).Item("Y1") = Y11 'Viewport_poly.GetPointAtParameter(3).Y
                                                Data_table_matchlines.Rows(Index_dataT).Item("X2") = X21 'Viewport_poly.GetPointAtParameter(2).X
                                                Data_table_matchlines.Rows(Index_dataT).Item("Y2") = Y22 'Viewport_poly.GetPointAtParameter(2).Y
                                                Data_table_matchlines.Rows(Index_dataT).Item("ML_LEN") = Lungime_matchline
                                                Index_dataT = Index_dataT + 1


                                            End If


                                        End If

                                    End If


                                Next

                                Data_table_matchlines = Sort_data_table(Data_table_matchlines, "STATION1")

                                If Data_table_matchlines.Rows.Count > 0 Then
                                    Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                                    W1 = Get_NEW_worksheet_from_Excel()
                                    Dim Column_match As String = "A"
                                    Dim Idx_col As Integer = W1.Range(Column_match & "1").Column

                                    For i = 0 To Data_table_matchlines.Rows.Count - 1
                                        W1.Range(Column_match & (i + 2).ToString).Value2 = Data_table_matchlines.Rows(i).Item("STATION1") & " - " & Data_table_matchlines.Rows(i).Item("STATION2")

                                        W1.Cells(i + 2, Idx_col + 2).Value2 = Data_table_matchlines.Rows(i).Item("STATION1")
                                        W1.Cells(i + 2, Idx_col + 3).Value2 = Data_table_matchlines.Rows(i).Item("STATION2")
                                        W1.Cells(i + 2, Idx_col + 4).FormulaR1C1 = "=RC[-2]-R[-1]C[-1]"
                                    Next
                                    W1.Cells(Data_table_matchlines.Rows.Count + 2, Idx_col + 4).FormulaR1C1 = "=SUM(R[-" & (Data_table_matchlines.Rows.Count).ToString & "]C:R[-1]C)"


                                End If

                                Trans1.Commit()


                            End Using
                        End Using

                    End If
                End If



                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")

            Catch ex As Exception
                Editor1.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_read_matchlines_Click_OLD(sender As Object, e As EventArgs) 'Handles Button_read_CL_VIEWPORTS_EXCEL.Click, Button_read_matchlines_canada.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Dim Round1 As Integer = 0
            If IsNumeric(TextBox_dec1.Text) = True Then
                Round1 = CInt(TextBox_dec1.Text)
            End If

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor


            Editor1.SetImpliedSelection(Empty_array)
            Try
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()



                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptEntityResult

                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select centerline:")

                Object_Prompt.SetRejectMessage(vbLf & "Please select a lightweight polyline or a 3d polyline")
                Object_Prompt.AddAllowedClass(GetType(Polyline), True)
                Object_Prompt.AddAllowedClass(GetType(Polyline3d), True)


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


                                Dim PolyCL_for_viewports As Polyline = TryCast(Trans1.GetObject(Rezultat1.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Polyline)
                                If IsNothing(PolyCL_for_viewports) = False Then
                                    Poly_length = PolyCL_for_viewports.Length
                                    PolyCL = PolyCL_for_viewports
                                End If
                                Dim PolyCL3D_for_viewports As Polyline3d
                                PolyCL3D_for_viewports = TryCast(Trans1.GetObject(Rezultat1.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Polyline3d)

                                If IsNothing(PolyCL3D_for_viewports) = False Then
                                    Poly_length = PolyCL3D_for_viewports.Length
                                    PolyCL3D = PolyCL3D_for_viewports
                                    Dim Index_Poly As Integer = 0
                                    PolyCL_for_viewports = New Polyline
                                    For Each vId As Autodesk.AutoCAD.DatabaseServices.ObjectId In PolyCL3D_for_viewports
                                        Dim v3d As Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d = DirectCast(Trans1.GetObject _
                                                (vId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d)

                                        Dim x1 As Double = v3d.Position.X
                                        Dim y1 As Double = v3d.Position.Y
                                        Dim z1 As Double = v3d.Position.Z
                                        PolyCL_for_viewports.AddVertexAt(Index_Poly, New Point2d(x1, y1), 0, 0, 0)
                                        Index_Poly = Index_Poly + 1
                                    Next
                                    PolyCL_for_viewports.Elevation = 0
                                End If

                                If Not PolyCL_for_viewports.Elevation = 0 Then
                                    Freeze_operations = False
                                    MsgBox("CL Polyline is not at elevation 0")
                                    Exit Sub

                                End If
                                Data_table_compiled = New System.Data.DataTable
                                Data_table_matchlines = New System.Data.DataTable
                                Data_table_matchlines.Columns.Add("STATION1", GetType(Double))
                                Data_table_matchlines.Columns.Add("STATION2", GetType(Double))
                                Data_table_matchlines.Columns.Add("X1", GetType(Double))
                                Data_table_matchlines.Columns.Add("Y1", GetType(Double))
                                Data_table_matchlines.Columns.Add("X2", GetType(Double))
                                Data_table_matchlines.Columns.Add("Y2", GetType(Double))
                                Data_table_matchlines.Columns.Add("ML_LEN", GetType(Double))




                                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                                Dim LayerTable1 As Autodesk.AutoCAD.DatabaseServices.LayerTable
                                LayerTable1 = Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                Dim Index_dataT As Double = 0


                                For Each objID As ObjectId In BTrecord
                                    Dim Rectangle_poly As Entity = Trans1.GetObject(objID, OpenMode.ForRead)

                                    Dim Executa As Boolean = False
                                    If TypeOf Rectangle_poly Is Polyline Then
                                        If Not Rectangle_poly.ObjectId = PolyCL_for_viewports.ObjectId Then
                                            If IsNothing(PolyCL3D_for_viewports) = False Then
                                                If Not Rectangle_poly.ObjectId = PolyCL3D_for_viewports.ObjectId Then
                                                    Executa = True
                                                End If
                                            Else
                                                Executa = True
                                            End If
                                        End If
                                    End If

                                    If Executa = True Then
                                        Dim Viewport_poly As Polyline = Rectangle_poly
                                        Dim LayerTableRecord1 As Autodesk.AutoCAD.DatabaseServices.LayerTableRecord
                                        LayerTableRecord1 = LayerTable1(Viewport_poly.Layer).GetObject(OpenMode.ForRead)


                                        If Viewport_poly.NumberOfVertices >= 4 And LayerTableRecord1.IsOff = False And LayerTableRecord1.IsFrozen = False Then
                                            Viewport_poly.UpgradeOpen()
                                            Viewport_poly.Elevation = 0




                                            Dim Col_int As New Point3dCollection
                                            Col_int = Intersect_on_both_operands(PolyCL_for_viewports, Viewport_poly)
                                            If Col_int.Count = 2 Then
                                                Dim Station1 As Double = 0
                                                Dim Station2 As Double = 0
                                                Dim Este_zero As Boolean = False
                                                Dim Nr_values As Integer = Col_int.Count
                                                If Nr_values > 2 Then Nr_values = 2


                                                Dim Point_on_poly1 As New Point3d()
                                                Point_on_poly1 = Col_int(0)
                                                Station1 = PolyCL_for_viewports.GetDistAtPoint(Point_on_poly1)

                                                If Round(Station1, 0) = 0 Then
                                                    Este_zero = True
                                                End If

                                                If IsNothing(PolyCL3D_for_viewports) = False Then
                                                    Dim Param1 As Double = PolyCL_for_viewports.GetParameterAtPoint(Point_on_poly1)
                                                    Station1 = PolyCL3D_for_viewports.GetDistanceAtParameter(Param1)
                                                End If

                                                Dim Point_on_poly2 As New Point3d()
                                                Point_on_poly2 = Col_int(1)
                                                Station2 = PolyCL_for_viewports.GetDistAtPoint(Point_on_poly2)

                                                If IsNothing(PolyCL3D_for_viewports) = False Then
                                                    Dim Param2 As Double = PolyCL_for_viewports.GetParameterAtPoint(Point_on_poly2)
                                                    Station2 = PolyCL3D_for_viewports.GetDistanceAtParameter(Param2)
                                                End If


                                                Dim Linie1 As New Line(Point_on_poly1, Point_on_poly2)

                                                Dim vpoly_exploded As New DBObjectCollection
                                                Viewport_poly.Explode(vpoly_exploded)



                                                If Station1 > Station2 Then
                                                    Dim T As Double = Station1
                                                    Station1 = Station2
                                                    Station2 = T
                                                End If

                                                If IsNothing(PolyCL3D_for_viewports) = False Then

                                                Else
                                                    If Station2 > PolyCL_for_viewports.Length Then
                                                        Station2 = PolyCL_for_viewports.Length
                                                    End If
                                                End If



                                                Dim X11, Y11, X21, Y22 As Double
                                                Dim LimitL As Double = 100
                                                For Each Ent4 As Entity In vpoly_exploded
                                                    If TypeOf (Ent4) Is Line Then
                                                        Dim Line2 As Line
                                                        Line2 = Ent4
                                                        If Line2.Length > LimitL Then
                                                            Dim Col_int1 As New Point3dCollection
                                                            Line2.IntersectWith(Linie1, Intersect.OnBothOperands, Col_int1, IntPtr.Zero, IntPtr.Zero)

                                                            If IsNothing(Col_int1) = True Then
                                                                Dim Pt1 As New Point3d
                                                                Dim Pt2 As New Point3d
                                                                Pt1 = Line2.GetClosestPointTo(Point_on_poly1, Vector3d.ZAxis, True)
                                                                Pt2 = Line2.GetClosestPointTo(Point_on_poly2, Vector3d.ZAxis, True)
                                                                If Pt1.GetVectorTo(Pt2).Length > LimitL Then
                                                                    X11 = Pt1.X
                                                                    Y11 = Pt1.Y
                                                                    X21 = Pt2.X
                                                                    Y22 = Pt2.Y
                                                                End If
                                                            Else
                                                                If Col_int1.Count = 0 Then
                                                                    Dim Pt1 As New Point3d
                                                                    Dim Pt2 As New Point3d
                                                                    Pt1 = Line2.GetClosestPointTo(Point_on_poly1, Vector3d.ZAxis, True)
                                                                    Pt2 = Line2.GetClosestPointTo(Point_on_poly2, Vector3d.ZAxis, True)
                                                                    If Pt1.GetVectorTo(Pt2).Length > LimitL Then
                                                                        X11 = Pt1.X
                                                                        Y11 = Pt1.Y
                                                                        X21 = Pt2.X
                                                                        Y22 = Pt2.Y
                                                                    End If

                                                                End If
                                                            End If



                                                        End If


                                                    End If

                                                Next


                                                Dim Lungime_matchline As Double = 0
                                                For i = 0 To Viewport_poly.NumberOfVertices - 2
                                                    Dim Pt1 As Point2d = Viewport_poly.GetPoint2dAt(i)
                                                    Dim Pt2 As Point2d = Viewport_poly.GetPoint2dAt(i + 1)
                                                    Dim Len1 As Double = Pt1.GetVectorTo(Pt2).Length
                                                    If Len1 > Lungime_matchline Then Lungime_matchline = Len1

                                                Next
                                                Dim pt3 As Point2d = Viewport_poly.GetPoint2dAt(Viewport_poly.NumberOfVertices - 1)
                                                Dim Pt0 As Point2d = Viewport_poly.GetPoint2dAt(0)
                                                Dim Len0 As Double = Pt0.GetVectorTo(pt3).Length
                                                If Len0 > Lungime_matchline Then Lungime_matchline = Len0


                                                Data_table_matchlines.Rows.Add()
                                                Data_table_matchlines.Rows(Index_dataT).Item("STATION1") = Round(Station1, 2)
                                                Data_table_matchlines.Rows(Index_dataT).Item("STATION2") = Round(Station2, 2)
                                                Data_table_matchlines.Rows(Index_dataT).Item("X1") = X11 'Viewport_poly.GetPointAtParameter(3).X
                                                Data_table_matchlines.Rows(Index_dataT).Item("Y1") = Y11 'Viewport_poly.GetPointAtParameter(3).Y
                                                Data_table_matchlines.Rows(Index_dataT).Item("X2") = X21 'Viewport_poly.GetPointAtParameter(2).X
                                                Data_table_matchlines.Rows(Index_dataT).Item("Y2") = Y22 'Viewport_poly.GetPointAtParameter(2).Y
                                                Data_table_matchlines.Rows(Index_dataT).Item("ML_LEN") = Lungime_matchline
                                                Index_dataT = Index_dataT + 1


                                            End If


                                        End If

                                    End If


                                Next

                                Add_to_clipboard_Data_table(Data_table_matchlines)

                                Trans1.Commit()


                            End Using
                        End Using

                    End If
                End If



                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")

            Catch ex As Exception
                Editor1.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try
            Freeze_operations = False
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & Data_table_matchlines.Rows.Count & " matchlines loaded")

        End If
    End Sub

    Private Sub Button_create_compiled_Click(sender As Object, e As EventArgs) Handles Button_create_compiled.Click, Button_compile_as_built_canada.Click


        If Freeze_operations = False Then
            Freeze_operations = True
            Dim Round1 As Integer = 0
            If IsNumeric(TextBox_dec1.Text) = True Then
                Round1 = CInt(TextBox_dec1.Text)
            End If
            Dim CSF As Double = 1

            Dim Is_Canada As Boolean = False

            If IsNothing(Data_table_elbows) = False Then
                If Data_table_elbows.Columns.Contains("CANADA") = True Then
                    Is_Canada = True
                End If
            End If
            If IsNothing(Data_table_materials) = False Then
                If Data_table_materials.Columns.Contains("CANADA") = True Then
                    Is_Canada = True
                End If
            End If

            If Is_Canada = True Then
                If IsNumeric(TextBox_CSF.Text) = True Then
                    CSF = CDbl(TextBox_CSF.Text)
                End If
            End If

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Try
                If IsNothing(Data_table_matchlines) = True Then
                    Freeze_operations = False
                    MsgBox("Please load the matchlines")
                    Editor1.WriteMessage(vbLf & "Command:")
                    Exit Sub
                End If
                If Data_table_matchlines.Rows.Count = 0 Then
                    Freeze_operations = False
                    MsgBox("Please load the matchlines")
                    Editor1.WriteMessage(vbLf & "Command:")
                    Exit Sub
                End If

                If IsNothing(Data_table_materials) = True Then
                    Freeze_operations = False
                    MsgBox("Please load the materials")
                    Editor1.WriteMessage(vbLf & "Command:")
                    Exit Sub
                End If
                If Data_table_materials.Rows.Count = 0 Then
                    Freeze_operations = False
                    MsgBox("Please load the materials")
                    Editor1.WriteMessage(vbLf & "Command:")
                    Exit Sub
                End If

                Data_table_compiled = New System.Data.DataTable
                Data_table_compiled.Columns.Add("STA1", GetType(Double))
                Data_table_compiled.Columns.Add("STA2", GetType(Double))
                Data_table_compiled.Columns.Add("MAT", GetType(String))
                Data_table_compiled.Columns.Add("STA", GetType(Double))
                Data_table_compiled.Columns.Add("X1", GetType(Double))
                Data_table_compiled.Columns.Add("Y1", GetType(Double))
                Data_table_compiled.Columns.Add("X2", GetType(Double))
                Data_table_compiled.Columns.Add("Y2", GetType(Double))
                Data_table_compiled.Columns.Add("M1", GetType(Double))
                Data_table_compiled.Columns.Add("M2", GetType(Double))
                Data_table_compiled.Columns.Add("ELLBOW", GetType(Boolean))
                Data_table_compiled.Columns.Add("ELLBOW_DESCR", GetType(String))
                Data_table_compiled.Columns.Add("ELLBOW_ID", GetType(String))
                Data_table_compiled.Columns.Add("DELTAX", GetType(Double))
                Data_table_compiled.Columns.Add("BAND_LENGTH", GetType(Double))
                Data_table_compiled.Columns.Add("PAGE", GetType(Integer))
                Data_table_compiled.Columns.Add("ML_LEN", GetType(Double))


                Dim Poly_center As Curve
                If IsNothing(PolyCL3D) = False Then
                    Poly_center = PolyCL3D
                End If
                If IsNothing(PolyCL) = False Then
                    Poly_center = PolyCL
                End If

                Dim Index_data_table As Integer = 0

                For i = 0 To Data_table_materials.Rows.Count - 1
                    If IsDBNull(Data_table_materials.Rows(i).Item("STA1")) = False And IsDBNull(Data_table_materials.Rows(i).Item("STA2")) = False Then


                        Dim Station1 As Double = Data_table_materials.Rows(i).Item("STA1")
                        Dim Station2 As Double = Data_table_materials.Rows(i).Item("STA2")



                        Dim Material As String = ""


                        If IsDBNull(Data_table_materials.Rows(i).Item("MAT")) = False Then
                            Material = Data_table_materials.Rows(i).Item("MAT")
                        End If

                        Dim I_start As Integer = 0
                        Dim Boolean_go_to_check_s1_s2 As Boolean = False


                        If i = Data_table_materials.Rows.Count - 1 Then
                            Dim debug1 As Double = i
                        End If

123:




                        For j = I_start To Data_table_matchlines.Rows.Count - 1
                            If IsDBNull(Data_table_matchlines.Rows(j).Item("STATION1")) = False And IsDBNull(Data_table_matchlines.Rows(j).Item("STATION2")) = False Then



                                Dim M1 As Double = Data_table_matchlines.Rows(j).Item("STATION1")
                                Dim M2 As Double = Data_table_matchlines.Rows(j).Item("STATION2")
                                Dim Lungime_matchline As Double = Data_table_matchlines.Rows(j).Item("ML_LEN")


                                If IsNothing(PolyCL3D) = False Then
                                    If M2 > PolyCL3D.Length Then
                                        If Round(PolyCL3D.Length, Round1) = Round(M2, Round1) Then
                                            M2 = PolyCL3D.Length
                                        Else
                                            MsgBox("Match2 is bigger than the 3D polyline length" & vbCrLf & M2 & ">" & Round(PolyCL3D.Length, 2))
                                            Freeze_operations = False
                                            Editor1.WriteMessage(vbLf & "Command:")
                                            Exit Sub
                                        End If
                                    End If

                                    If Station2 * CSF > PolyCL3D.Length Then
                                        If Round(PolyCL3D.Length, Round1) = Round(Station2 * CSF, Round1) Then
                                            Station2 = PolyCL3D.Length / CSF
                                        Else
                                            MsgBox("Station2 is bigger than the 3D polyline length" & vbCrLf & Station2 & ">" & Round(PolyCL3D.Length / CSF, 2))
                                            Freeze_operations = False
                                            Editor1.WriteMessage(vbLf & "Command:")
                                            Exit Sub
                                        End If
                                    End If

                                Else
                                    If M2 > PolyCL.Length Then
                                        If Round(PolyCL.Length, Round1) = Round(M2, Round1) Then
                                            M2 = PolyCL.Length
                                        Else
                                            MsgBox("Match2 is bigger than the polyline length" & vbCrLf & M2 & ">" & Round(PolyCL.Length, 2))
                                            Freeze_operations = False
                                            Editor1.WriteMessage(vbLf & "Command:")
                                            Exit Sub
                                        End If
                                    End If


                                    If Station2 > PolyCL.Length Then
                                        If Round(PolyCL.Length, Round1) = Round(Station2 * CSF, Round1) Then
                                            Station2 = PolyCL.Length / CSF
                                        Else
                                            MsgBox("Station2 is bigger than the polyline length" & vbCrLf & Station2 & ">" & Round(PolyCL.Length / CSF, 2))
                                            Freeze_operations = False
                                            Editor1.WriteMessage(vbLf & "Command:")
                                            Exit Sub
                                        End If
                                    End If
                                End If



                                If Boolean_go_to_check_s1_s2 = True Then
                                    If Round(Station1, Round1) = Round(Station2, Round1) Then GoTo LS12end
                                    GoTo LS1S2
                                End If




                                If M1 <= Station1 * CSF And M2 <= Station2 * CSF And M1 <= Station2 * CSF And M2 > Station1 * CSF Then
                                    Dim x1 As Double = Data_table_matchlines.Rows(j).Item("X1")
                                    Dim x2 As Double = Data_table_matchlines.Rows(j).Item("X2")
                                    Dim y1 As Double = Data_table_matchlines.Rows(j).Item("Y1")
                                    Dim y2 As Double = Data_table_matchlines.Rows(j).Item("Y2")

                                    Dim Valoare2 As Double = M2 / CSF
                                    If Not Station1 = Valoare2 Then
                                        Data_table_compiled.Rows.Add()
                                        Data_table_compiled.Rows(Index_data_table).Item("STA1") = Station1
                                        Data_table_compiled.Rows(Index_data_table).Item("STA2") = Valoare2
                                        Data_table_compiled.Rows(Index_data_table).Item("MAT") = Material
                                        Data_table_compiled.Rows(Index_data_table).Item("X2") = x2
                                        Data_table_compiled.Rows(Index_data_table).Item("Y2") = y2
                                        Data_table_compiled.Rows(Index_data_table).Item("X1") = x1
                                        Data_table_compiled.Rows(Index_data_table).Item("Y1") = y1
                                        Data_table_compiled.Rows(Index_data_table).Item("PAGE") = j + 1
                                        Data_table_compiled.Rows(Index_data_table).Item("ELLBOW") = False
                                        Data_table_compiled.Rows(Index_data_table).Item("M1") = M1
                                        Data_table_compiled.Rows(Index_data_table).Item("M2") = M2
                                        Data_table_compiled.Rows(Index_data_table).Item("DELTAX") = DELTA_X_calculate(x1, y1, x2, y2, Station1, M2, CSF, Poly_center)
                                        Data_table_compiled.Rows(Index_data_table).Item("ML_LEN") = Lungime_matchline

                                        Index_data_table = Index_data_table + 1
                                        Station1 = Valoare2
                                        I_start = j + 1
                                        Boolean_go_to_check_s1_s2 = True
                                        GoTo 123
                                    End If

                                End If


                                If Station1 * CSF >= M1 And Station2 * CSF <= M2 And Station1 * CSF < M2 Then
                                    Dim x1 As Double = Data_table_matchlines.Rows(j).Item("X1")
                                    Dim x2 As Double = Data_table_matchlines.Rows(j).Item("X2")
                                    Dim y1 As Double = Data_table_matchlines.Rows(j).Item("Y1")
                                    Dim y2 As Double = Data_table_matchlines.Rows(j).Item("Y2")
                                    Data_table_compiled.Rows.Add()
                                    Data_table_compiled.Rows(Index_data_table).Item("STA1") = Station1
                                    Data_table_compiled.Rows(Index_data_table).Item("STA2") = Station2
                                    Data_table_compiled.Rows(Index_data_table).Item("MAT") = Material
                                    Data_table_compiled.Rows(Index_data_table).Item("X2") = x2
                                    Data_table_compiled.Rows(Index_data_table).Item("Y2") = y2
                                    Data_table_compiled.Rows(Index_data_table).Item("X1") = x1
                                    Data_table_compiled.Rows(Index_data_table).Item("Y1") = y1

                                    Data_table_compiled.Rows(Index_data_table).Item("PAGE") = j + 1
                                    Data_table_compiled.Rows(Index_data_table).Item("ELLBOW") = False
                                    Data_table_compiled.Rows(Index_data_table).Item("M1") = M1
                                    Data_table_compiled.Rows(Index_data_table).Item("M2") = M2
                                    Data_table_compiled.Rows(Index_data_table).Item("DELTAX") = DELTA_X_calculate(x1, y1, x2, y2, Station1, Station2, CSF, Poly_center)
                                    Data_table_compiled.Rows(Index_data_table).Item("ML_LEN") = Lungime_matchline
                                    Index_data_table = Index_data_table + 1
                                    Exit For
                                End If




LS1S2:

                                ' add S1, S2
                                If Station1 * CSF >= M1 And Station2 * CSF <= M2 And Station1 * CSF < M2 Then
                                    Dim x1 As Double = Data_table_matchlines.Rows(j).Item("X1")
                                    Dim x2 As Double = Data_table_matchlines.Rows(j).Item("X2")
                                    Dim y1 As Double = Data_table_matchlines.Rows(j).Item("Y1")
                                    Dim y2 As Double = Data_table_matchlines.Rows(j).Item("Y2")

                                    Data_table_compiled.Rows.Add()
                                    Data_table_compiled.Rows(Index_data_table).Item("STA1") = Station1
                                    Data_table_compiled.Rows(Index_data_table).Item("STA2") = Station2

                                    Data_table_compiled.Rows(Index_data_table).Item("X2") = x2
                                    Data_table_compiled.Rows(Index_data_table).Item("Y2") = y2
                                    Data_table_compiled.Rows(Index_data_table).Item("X1") = x1
                                    Data_table_compiled.Rows(Index_data_table).Item("Y1") = y1
                                    Data_table_compiled.Rows(Index_data_table).Item("PAGE") = j + 1

                                    Data_table_compiled.Rows(Index_data_table).Item("MAT") = Material
                                    Data_table_compiled.Rows(Index_data_table).Item("ELLBOW") = False
                                    Data_table_compiled.Rows(Index_data_table).Item("M1") = M1
                                    Data_table_compiled.Rows(Index_data_table).Item("M2") = M2
                                    Data_table_compiled.Rows(Index_data_table).Item("DELTAX") = DELTA_X_calculate(x1, y1, x2, y2, Station1, Station2, CSF, Poly_center)
                                    Data_table_compiled.Rows(Index_data_table).Item("ML_LEN") = Lungime_matchline

                                    Index_data_table = Index_data_table + 1
                                    Exit For

                                ElseIf Station1 * CSF <= M2 And Station1 * CSF >= M1 Then
                                    Dim x1 As Double = Data_table_matchlines.Rows(j).Item("X1")
                                    Dim x2 As Double = Data_table_matchlines.Rows(j).Item("X2")
                                    Dim y1 As Double = Data_table_matchlines.Rows(j).Item("Y1")
                                    Dim y2 As Double = Data_table_matchlines.Rows(j).Item("Y2")
                                    Dim Valoare2 As Double = M2 / CSF
                                    If Not Station1 = Valoare2 Then
                                        Data_table_compiled.Rows.Add()
                                        Data_table_compiled.Rows(Index_data_table).Item("STA1") = Station1
                                        Data_table_compiled.Rows(Index_data_table).Item("STA2") = Valoare2

                                        Data_table_compiled.Rows(Index_data_table).Item("X2") = x2
                                        Data_table_compiled.Rows(Index_data_table).Item("Y2") = y2
                                        Data_table_compiled.Rows(Index_data_table).Item("X1") = x1
                                        Data_table_compiled.Rows(Index_data_table).Item("Y1") = y1
                                        Data_table_compiled.Rows(Index_data_table).Item("PAGE") = j + 1

                                        Data_table_compiled.Rows(Index_data_table).Item("MAT") = Material
                                        Data_table_compiled.Rows(Index_data_table).Item("ELLBOW") = False
                                        Data_table_compiled.Rows(Index_data_table).Item("M1") = M1
                                        Data_table_compiled.Rows(Index_data_table).Item("M2") = M2
                                        Data_table_compiled.Rows(Index_data_table).Item("DELTAX") = DELTA_X_calculate(x1, y1, x2, y2, Station1, M2, CSF, Poly_center)
                                        Data_table_compiled.Rows(Index_data_table).Item("ML_LEN") = Lungime_matchline

                                        Index_data_table = Index_data_table + 1

                                        Station1 = Valoare2

                                        I_start = j + 1
                                        Boolean_go_to_check_s1_s2 = True
                                        GoTo 123
                                    End If

                                End If
                            End If
                        Next
LS12end:
                    End If
                Next

                If IsNothing(Data_table_elbows) = False Then
                    If Data_table_elbows.Rows.Count > 0 Then
                        For k = 0 To Data_table_elbows.Rows.Count - 1
                            If IsDBNull(Data_table_elbows.Rows(k).Item("STA1")) = False And IsDBNull(Data_table_elbows.Rows(k).Item("STA2")) = False Then
                                Dim Sta_el1 As Double = Data_table_elbows.Rows(k).Item("STA1")
                                Dim Sta_el2 As Double = Data_table_elbows.Rows(k).Item("STA2")


                                Dim Sta_el As String = "XXX"
                                If IsDBNull(Data_table_elbows.Rows(k).Item("STA")) = False Then
                                    Sta_el = Data_table_elbows.Rows(k).Item("STA")
                                End If

                                Dim Descr_el As String = "XXX"
                                If IsDBNull(Data_table_elbows.Rows(k).Item("DESCR")) = False Then
                                    Descr_el = Data_table_elbows.Rows(k).Item("DESCR")
                                End If

                                Dim ID_el As String = "XXX"
                                If IsDBNull(Data_table_elbows.Rows(k).Item("ID")) = False Then
                                    ID_el = Data_table_elbows.Rows(k).Item("ID")
                                End If

                                For i = 0 To Data_table_compiled.Rows.Count - 1
                                    Dim Station1 As Double = Data_table_compiled.Rows(i).Item("STA1")
                                    Dim Station2 As Double = Data_table_compiled.Rows(i).Item("STA2")
                                    If Station1 <= Sta_el1 And Station2 >= Sta_el2 Then
                                        Dim x1 As Double = Data_table_compiled.Rows(i).Item("X1")
                                        Dim y1 As Double = Data_table_compiled.Rows(i).Item("Y1")
                                        Dim x2 As Double = Data_table_compiled.Rows(i).Item("X2")
                                        Dim y2 As Double = Data_table_compiled.Rows(i).Item("Y2")

                                        Data_table_compiled.Rows(i).Item("STA2") = Sta_el1

                                        Data_table_compiled.Rows(i).Item("DELTAX") = DELTA_X_calculate(x1, y1, x2, y2, Station1, Sta_el1, CSF, Poly_center)

                                        Dim Row1 As System.Data.DataRow
                                        Row1 = Data_table_compiled.NewRow()
                                        Row1("STA1") = Sta_el1
                                        Row1("STA2") = Sta_el2
                                        Row1("PAGE") = Data_table_compiled.Rows(i).Item("PAGE")
                                        Row1("X1") = x1
                                        Row1("Y1") = y1
                                        Row1("X2") = x2
                                        Row1("Y2") = y2
                                        Row1("M1") = Data_table_compiled.Rows(i).Item("M1")
                                        Row1("M2") = Data_table_compiled.Rows(i).Item("M2")
                                        Row1("DELTAX") = 0
                                        Row1("ELLBOW") = True
                                        If Not Sta_el = "XXX" Then
                                            Row1("STA") = Sta_el
                                        End If
                                        If Not Descr_el = "XXX" Then
                                            Row1("ELLBOW_DESCR") = Descr_el
                                        End If
                                        If Not ID_el = "XXX" Then
                                            Row1("ELLBOW_ID") = ID_el
                                        End If
                                        If Data_table_elbows.Columns.Contains("MAT") = True Then
                                            Row1("MAT") = Data_table_elbows.Rows(k).Item("MAT")
                                        End If

                                        Data_table_compiled.Rows.InsertAt(Row1, i + 1)

                                        If Not Round(Abs(Station2 - Sta_el2), 1) = 0 Then
                                            Row1 = Data_table_compiled.NewRow()
                                            Row1("STA1") = Sta_el2
                                            Row1("STA2") = Station2
                                            Row1("PAGE") = Data_table_compiled.Rows(i).Item("PAGE")
                                            Row1("X1") = Data_table_compiled.Rows(i).Item("X1")
                                            Row1("Y1") = Data_table_compiled.Rows(i).Item("Y1")
                                            Row1("X2") = Data_table_compiled.Rows(i).Item("X2")
                                            Row1("Y2") = Data_table_compiled.Rows(i).Item("Y2")
                                            Row1("M1") = Data_table_compiled.Rows(i).Item("M1")
                                            Row1("M2") = Data_table_compiled.Rows(i).Item("M2")
                                            Row1("DELTAX") = DELTA_X_calculate(x1, y1, x2, y2, Sta_el2, Station2, CSF, Poly_center)
                                            Row1("ELLBOW") = False
                                            Row1("MAT") = Data_table_compiled.Rows(i).Item("MAT")
                                            Data_table_compiled.Rows.InsertAt(Row1, i + 2)
                                        End If
                                        Exit For
                                    End If

                                    If Station1 * CSF <= Sta_el1 And Station2 * CSF > Sta_el1 And Station2 * CSF < Sta_el2 Then
                                        Dim x1 As Double = Data_table_compiled.Rows(i).Item("X1")
                                        Dim y1 As Double = Data_table_compiled.Rows(i).Item("Y1")
                                        Dim x2 As Double = Data_table_compiled.Rows(i).Item("X2")
                                        Dim y2 As Double = Data_table_compiled.Rows(i).Item("Y2")

                                        Data_table_compiled.Rows(i).Item("STA2") = Sta_el1
                                        Data_table_compiled.Rows(i + 1).Item("STA1") = Sta_el2

                                        Data_table_compiled.Rows(i).Item("DELTAX") = DELTA_X_calculate(x1, y1, x2, y2, Station1, Sta_el1, CSF, Poly_center)
                                        Data_table_compiled.Rows(i + 1).Item("DELTAX") = DELTA_X_calculate(x1, y1, x2, y2, Sta_el2, Station2, CSF, Poly_center)




                                        Dim PAGE As Integer = Data_table_compiled.Rows(i + 1).Item("PAGE")


                                        Dim M1 As Double = Data_table_compiled.Rows(i + 1).Item("M1")
                                        Dim M2 As Double = Data_table_compiled.Rows(i + 1).Item("M2")

                                        Dim Row1 As System.Data.DataRow
                                        Row1 = Data_table_compiled.NewRow()
                                        Row1("STA1") = Sta_el1
                                        Row1("STA2") = Sta_el2
                                        Row1("PAGE") = Data_table_compiled.Rows(i).Item("PAGE")
                                        Row1("X1") = Data_table_compiled.Rows(i).Item("X1")
                                        Row1("Y1") = Data_table_compiled.Rows(i).Item("Y1")
                                        Row1("X2") = Data_table_compiled.Rows(i).Item("X2")
                                        Row1("Y2") = Data_table_compiled.Rows(i).Item("Y2")
                                        Row1("M1") = Data_table_compiled.Rows(i).Item("M1")
                                        Row1("M2") = Data_table_compiled.Rows(i).Item("M2")
                                        Row1("DELTAX") = 0
                                        Row1("ELLBOW") = True
                                        If Not Sta_el = "XXX" Then
                                            Row1("STA") = Sta_el
                                        End If
                                        If Not Descr_el = "XXX" Then
                                            Row1("ELLBOW_DESCR") = Descr_el
                                        End If
                                        If Not ID_el = "XXX" Then
                                            Row1("ELLBOW_ID") = ID_el
                                        End If

                                        If Data_table_elbows.Columns.Contains("MAT") = True Then
                                            Row1("MAT") = Data_table_elbows.Rows(k).Item("MAT")
                                        End If

                                        Data_table_compiled.Rows.InsertAt(Row1, i + 1)

                                        Exit For
                                    End If

                                Next
                            End If


                        Next
                    End If
                End If


                If IsNothing(Data_table_buoyancy_canada_without_matchlines) = False Then
                    If Data_table_buoyancy_canada_without_matchlines.Rows.Count > 0 Then
                        Data_table_buoyancy_canada = New System.Data.DataTable
                        Data_table_buoyancy_canada.Columns.Add("CHAINAGE_START", GetType(Double))
                        Data_table_buoyancy_canada.Columns.Add("CHAINAGE_END", GetType(Double))
                        Data_table_buoyancy_canada.Columns.Add("COUNT", GetType(Integer))
                        Data_table_buoyancy_canada.Columns.Add("SPACING", GetType(Double))
                        Data_table_buoyancy_canada.Columns.Add("DESCRIPTION", GetType(String))
                        Data_table_buoyancy_canada.Columns.Add("X", GetType(Double))
                        Data_table_buoyancy_canada.Columns.Add("Y", GetType(Double))
                        Data_table_buoyancy_canada.Columns.Add("STRETCH", GetType(Double))

                        Dim I_start As Integer = 0
                        Dim boolean_check_S1_S2 As Boolean = False
                        Dim Index_dT1 As Integer = 0

                        For i = 0 To Data_table_buoyancy_canada_without_matchlines.Rows.Count - 1
                            If IsDBNull(Data_table_buoyancy_canada_without_matchlines.Rows(i).Item("CHAINAGE_START")) = False And IsDBNull(Data_table_buoyancy_canada_without_matchlines.Rows(i).Item("CHAINAGE_END")) = False And
                                IsDBNull(Data_table_buoyancy_canada_without_matchlines.Rows(i).Item("COUNT")) = False And IsDBNull(Data_table_buoyancy_canada_without_matchlines.Rows(i).Item("SPACING")) = False And
                                IsDBNull(Data_table_buoyancy_canada_without_matchlines.Rows(i).Item("DESCRIPTION")) = False Then


                                Dim Station1 As Double = Data_table_buoyancy_canada_without_matchlines.Rows(i).Item("CHAINAGE_START")
                                Dim Station2 As Double = Data_table_buoyancy_canada_without_matchlines.Rows(i).Item("CHAINAGE_END")
                                Dim Count1 As Integer = Data_table_buoyancy_canada_without_matchlines.Rows(i).Item("COUNT")
                                Dim Spacing1 As Double = Data_table_buoyancy_canada_without_matchlines.Rows(i).Item("SPACING")
                                Dim Description1 As String = Data_table_buoyancy_canada_without_matchlines.Rows(i).Item("DESCRIPTION")
124:
                                For j = I_start To Data_table_matchlines.Rows.Count - 1
                                    If IsDBNull(Data_table_matchlines.Rows(j).Item("STATION1")) = False And IsDBNull(Data_table_matchlines.Rows(j).Item("STATION2")) = False Then



                                        Dim M1 As Double = Data_table_matchlines.Rows(j).Item("STATION1")
                                        Dim M2 As Double = Data_table_matchlines.Rows(j).Item("STATION2")







                                        If boolean_check_S1_S2 = True Then
                                            GoTo label_add_S1_S21
                                        End If




                                        If M1 <= Station1 * CSF And M2 <= Station2 * CSF And M1 <= Station2 * CSF And M2 >= Station1 * CSF Then

                                            Data_table_buoyancy_canada.Rows.Add()
                                            Data_table_buoyancy_canada.Rows(Index_dT1).Item("CHAINAGE_START") = Station1
                                            Data_table_buoyancy_canada.Rows(Index_dT1).Item("CHAINAGE_END") = M2 / CSF


                                            Dim Nr1 As Double = 1 + (M2 / CSF - Station1) / Spacing1
                                            Data_table_buoyancy_canada.Rows(Index_dT1).Item("COUNT") = Floor(Nr1)
                                            Data_table_buoyancy_canada.Rows(Index_dT1).Item("SPACING") = Spacing1
                                            Data_table_buoyancy_canada.Rows(Index_dT1).Item("DESCRIPTION") = Description1


                                            Index_dT1 = Index_dT1 + 1
                                            Station1 = M2 / CSF
                                            I_start = j + 1
                                            boolean_check_S1_S2 = True
                                            GoTo 124
                                        End If


                                        If Station1 * CSF >= M1 And Station2 * CSF <= M2 Then

                                            Data_table_buoyancy_canada.Rows.Add()
                                            Data_table_buoyancy_canada.Rows(Index_dT1).Item("CHAINAGE_START") = Station1
                                            Data_table_buoyancy_canada.Rows(Index_dT1).Item("CHAINAGE_END") = Station2
                                            Data_table_buoyancy_canada.Rows(Index_dT1).Item("COUNT") = Count1
                                            Data_table_buoyancy_canada.Rows(Index_dT1).Item("SPACING") = Spacing1
                                            Data_table_buoyancy_canada.Rows(Index_dT1).Item("DESCRIPTION") = Description1

                                            Index_dT1 = Index_dT1 + 1
                                            Exit For
                                        End If




label_add_S1_S21:

                                        ' add S1, S2
                                        If Station1 * CSF >= M1 And Station2 * CSF <= M2 Then


                                            Data_table_buoyancy_canada.Rows.Add()
                                            Data_table_buoyancy_canada.Rows(Index_dT1).Item("CHAINAGE_START") = Station1
                                            Data_table_buoyancy_canada.Rows(Index_dT1).Item("CHAINAGE_END") = Station2

                                            Data_table_buoyancy_canada.Rows(Index_dT1).Item("COUNT") = Count1
                                            Data_table_buoyancy_canada.Rows(Index_dT1).Item("SPACING") = Spacing1
                                            Data_table_buoyancy_canada.Rows(Index_dT1).Item("DESCRIPTION") = Description1
                                            Index_dT1 = Index_dT1 + 1
                                            Exit For

                                        ElseIf Station1 * CSF <= M2 And Station1 * CSF >= M1 Then

                                            Data_table_buoyancy_canada.Rows.Add()
                                            Data_table_buoyancy_canada.Rows(Index_dT1).Item("CHAINAGE_START") = Station1
                                            Data_table_buoyancy_canada.Rows(Index_dT1).Item("CHAINAGE_END") = M2 / CSF
                                            Dim Nr1 As Double = 1 + (M2 / CSF - Station1) / Spacing1
                                            Data_table_buoyancy_canada.Rows(Index_dT1).Item("COUNT") = Floor(Nr1)
                                            Data_table_buoyancy_canada.Rows(Index_dT1).Item("SPACING") = Spacing1
                                            Data_table_buoyancy_canada.Rows(Index_dT1).Item("DESCRIPTION") = Description1
                                            Index_dT1 = Index_dT1 + 1

                                            Station1 = M2 / CSF

                                            I_start = j + 1
                                            boolean_check_S1_S2 = True
                                            GoTo 124
                                        End If
                                    End If
                                Next



                            End If
                        Next
                    End If

                End If




                Add_to_clipboard_2_Data_table(Data_table_compiled, Data_table_buoyancy_canada)





            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Editor1.WriteMessage(vbLf & "Data compiled")
            Freeze_operations = False
        End If
    End Sub

    Public Function DELTA_X_calculate(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal Station1 As Double, ByVal Station2 As Double, ByVal CSF As Double, ByVal CurveCL As Curve)
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor

        Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                Dim Poly2D As Polyline
                Dim Poly3D As Polyline3d

                If TypeOf (CurveCL) Is Polyline Then
                    Poly2D = CurveCL
                End If

                If TypeOf (CurveCL) Is Polyline3d Then
                    Poly3D = CurveCL
                    Poly2D = New Polyline
                    Dim Index_poly As Integer = 0
                    For Each vId As Autodesk.AutoCAD.DatabaseServices.ObjectId In Poly3D
                        Dim v3d As Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d = DirectCast(Trans1.GetObject _
                                (vId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d)
                        Poly2D.AddVertexAt(Index_poly, New Point2d(v3d.Position.X, v3d.Position.Y), 0, 0, 0)
                        Index_poly = Index_poly + 1
                    Next
                    Poly2D.Elevation = 0
                End If


                Dim pT1 As New Point3d(x1, y1, 0)
                Dim pT2 As New Point3d(x2, y2, 0)
                Dim CL1 As New Point3d
                CL1 = Poly2D.GetClosestPointTo(pT1, Vector3d.ZAxis, False)

                Dim CL2 As New Point3d
                CL2 = Poly2D.GetClosestPointTo(pT2, Vector3d.ZAxis, False)

                Dim CH1 As Double = Poly2D.GetDistAtPoint(CL1)
                Dim CH2 As Double = Poly2D.GetDistAtPoint(CL2)
                Dim Line1 As Line
                If CH1 > CH2 Then
                    Line1 = New Line(New Point3d(x2, y2, 0), New Point3d(x1, y1, 0))
                Else
                    Line1 = New Line(New Point3d(x1, y1, 0), New Point3d(x2, y2, 0))
                End If

                Dim Point1 As New Point3d
                If TypeOf (CurveCL) Is Polyline Then
                    Point1 = Poly2D.GetPointAtDist(Station1 * CSF)
                End If
                If TypeOf (CurveCL) Is Polyline3d Then
                    Point1 = Poly3D.GetPointAtDist(Station1 * CSF)
                End If

                Dim Point2 As New Point3d
                If TypeOf (CurveCL) Is Polyline Then
                    Point2 = Poly2D.GetPointAtDist(Station2 * CSF)
                End If
                If TypeOf (CurveCL) Is Polyline3d Then
                    Point2 = Poly3D.GetPointAtDist(Station2 * CSF)
                End If

                Dim PL1 As New Point3d
                PL1 = Line1.GetClosestPointTo(Point1, Vector3d.ZAxis, True)
                Dim PL2 As New Point3d
                PL2 = Line1.GetClosestPointTo(Point2, Vector3d.ZAxis, True)

                Return New Point3d(PL1.X, PL1.Y, 0).GetVectorTo(New Point3d(PL2.X, PL2.Y, 0)).Length

                Trans1.Commit()
            End Using
        End Using

    End Function

    Private Sub Button_DRAW_engineering_Band_Click(sender As Object, e As EventArgs) Handles Button_DRAW_engineering_Band.Click, Button_DRAW_engineering_Band_CANADA.Click
        If IsNumeric(TextBox_viewport_Width.Text) = False Then
            MsgBox("please specify the viewport width")
            Exit Sub
        End If
        If IsNumeric(TextBox_viewport_SCALE.Text) = False Then
            MsgBox("please specify the viewport scale")
            Exit Sub
        End If

        Dim Band_Separation As Double

        If IsNumeric(TextBox_BAND_SPACING.Text) = False Then
            MsgBox("please specify the viewport scale")
            Exit Sub
        Else
            Band_Separation = CDbl(TextBox_BAND_SPACING.Text)
        End If

        If Freeze_operations = False Then
            Freeze_operations = True
            Dim Round1 As Integer = 0
            If IsNumeric(TextBox_dec1.Text) = True Then
                Round1 = CInt(TextBox_dec1.Text)
            End If

            If IsNothing(Data_table_compiled) = False Then
                If Data_table_compiled.Rows.Count > 0 Then
                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
                    Dim Is_Canada As Boolean = False

                    If IsNothing(Data_table_elbows) = False Then
                        If Data_table_elbows.Columns.Contains("CANADA") = True Then
                            Is_Canada = True
                        End If
                    End If
                    If IsNothing(Data_table_materials) = False Then
                        If Data_table_materials.Columns.Contains("CANADA") = True Then
                            Is_Canada = True
                        End If
                    End If


                    Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Creaza_layer("NO PLOT", 40, "NO PLOT", False)
                            Trans1.Commit()
                        End Using

                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                            BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)


                            If Not ComboBox_blocks_MAT.Text = "" Then
                                Dim Deltax_total As Double = 0
                                Dim Page_previous As Integer = 0
                                If IsNumeric(TextBox_X_MS.Text) = True And IsNumeric(TextBox_Y_MS.Text) = True Then
                                    Start_point = New Point3d(CDbl(TextBox_X_MS.Text), CDbl(TextBox_Y_MS.Text), 0)
                                End If


                                Dim Colectie_nume_atr0 As New Specialized.StringCollection
                                Dim Colectie_valori0 As New Specialized.StringCollection
                                Dim BR_mat As BlockReference = InsertBlock_with_multiple_atributes("", ComboBox_blocks_MAT.Text, New Point3d(0, 0, 0), 1, BTrecord, "0", Colectie_nume_atr0, Colectie_valori0)
                                Dim Min_Dist_mat As Double = Get_distance1_block(BR_mat)

                                Dim Min_Dist_elb As Double = 0
                                If Not ComboBox_blocks_ELBOWS_AS_MAT.Text = "" Then
                                    Dim BR_elb As BlockReference = InsertBlock_with_multiple_atributes("", ComboBox_blocks_ELBOWS_AS_MAT.Text, New Point3d(0, 0, 0), 1, BTrecord, "0", Colectie_nume_atr0, Colectie_valori0)
                                    Min_Dist_elb = Get_distance1_block(BR_elb)
                                End If


                                Dim Viewport_ms_length As Double = CDbl(TextBox_viewport_Width.Text) * CDbl(TextBox_viewport_SCALE.Text)
                                Dim Viewport_ms_height As Double = CDbl(TextBox_viewport_Height.Text) * CDbl(TextBox_viewport_SCALE.Text)

                                For i = 0 To Data_table_compiled.Rows.Count - 1
                                    If IsDBNull(Data_table_compiled.Rows(i).Item("ELLBOW")) = False And
                                       IsDBNull(Data_table_compiled.Rows(i).Item("PAGE")) = False And
                                        IsDBNull(Data_table_compiled.Rows(i).Item("DELTAX")) = False Then
                                        Dim IS_ELBOW As Boolean = Data_table_compiled.Rows(i).Item("ELLBOW")
                                        Dim Page1 As Integer = Data_table_compiled.Rows(i).Item("PAGE")
                                        Dim DeltaX As Double = Data_table_compiled.Rows(i).Item("DELTAX")



                                        If Not Page1 = Page_previous Then

                                            For j = 0 To i
                                                Dim Page2 As Integer = Data_table_compiled.Rows(j).Item("PAGE")
                                                If Page_previous = Page2 Then
                                                    Dim Scale_factor As Double = Stretch_scale_factor * Viewport_ms_length / Deltax_total

                                                    If Scale_factor < 1 Then
                                                        Data_table_compiled.Rows(j).Item("DELTAX") = Data_table_compiled.Rows(j).Item("DELTAX") * Scale_factor
                                                        Data_table_compiled.Rows(j).Item("BAND_LENGTH") = Deltax_total * Scale_factor
                                                    Else

                                                        Data_table_compiled.Rows(j).Item("BAND_LENGTH") = Deltax_total
                                                    End If


                                                End If
                                            Next
                                            Deltax_total = 0
                                        End If
                                        Page_previous = Page1

                                        If IS_ELBOW = True Then
                                            If DeltaX < Min_Dist_elb Then
                                                DeltaX = Min_Dist_elb
                                                Data_table_compiled.Rows(i).Item("DELTAX") = DeltaX
                                            End If
                                        Else
                                            If DeltaX < Min_Dist_mat Then
                                                DeltaX = Min_Dist_mat
                                                Data_table_compiled.Rows(i).Item("DELTAX") = DeltaX
                                            End If
                                        End If

                                        Deltax_total = Deltax_total + DeltaX

                                        If i = Data_table_compiled.Rows.Count - 1 Then
                                            For j = 0 To i
                                                Dim Page2 As Integer = Data_table_compiled.Rows(j).Item("PAGE")
                                                If Page1 = Page2 Then
                                                    Dim Scale_factor As Double = 0.95 * Viewport_ms_length / Deltax_total

                                                    If Scale_factor < 1 Then
                                                        Data_table_compiled.Rows(j).Item("DELTAX") = Data_table_compiled.Rows(j).Item("DELTAX") * Scale_factor
                                                        Data_table_compiled.Rows(j).Item("BAND_LENGTH") = Deltax_total * Scale_factor
                                                    Else

                                                        Data_table_compiled.Rows(j).Item("BAND_LENGTH") = Deltax_total
                                                    End If
                                                End If
                                            Next
                                        End If
                                    End If
                                Next

                                Page_previous = 0
                                Deltax_total = 0
                                'Add_to_clipboard_Data_table(Data_table_compiled)

                                For i = 0 To Data_table_compiled.Rows.Count - 1
                                    If IsDBNull(Data_table_compiled.Rows(i).Item("STA1")) = False And
                                        IsDBNull(Data_table_compiled.Rows(i).Item("STA2")) = False And
                                        IsDBNull(Data_table_compiled.Rows(i).Item("X1")) = False And
                                        IsDBNull(Data_table_compiled.Rows(i).Item("Y1")) = False And
                                        IsDBNull(Data_table_compiled.Rows(i).Item("X2")) = False And
                                        IsDBNull(Data_table_compiled.Rows(i).Item("Y2")) = False And
                                        IsDBNull(Data_table_compiled.Rows(i).Item("M1")) = False And
                                        IsDBNull(Data_table_compiled.Rows(i).Item("M2")) = False And
                                        IsDBNull(Data_table_compiled.Rows(i).Item("ELLBOW")) = False And
                                        IsDBNull(Data_table_compiled.Rows(i).Item("PAGE")) = False And
                                        IsDBNull(Data_table_compiled.Rows(i).Item("DELTAX")) = False And
                                        IsDBNull(Data_table_compiled.Rows(i).Item("BAND_LENGTH")) = False Then

                                        Dim Sta1 As Double = Round(Data_table_compiled.Rows(i).Item("STA1"), Round1)
                                        Dim Sta2 As Double = Round(Data_table_compiled.Rows(i).Item("STA2"), Round1)

                                        Dim Lungime_matchline_viewport As Double = Data_table_compiled.Rows(i).Item("ML_LEN")

                                        Dim M1 As Double = Data_table_compiled.Rows(i).Item("M1")
                                        Dim M2 As Double = Data_table_compiled.Rows(i).Item("M2")
                                        Dim X1 As Double = Data_table_compiled.Rows(i).Item("X1")
                                        Dim Y1 As Double = Data_table_compiled.Rows(i).Item("Y1")
                                        Dim X2 As Double = Data_table_compiled.Rows(i).Item("X2")
                                        Dim Y2 As Double = Data_table_compiled.Rows(i).Item("Y2")

                                        Dim IS_ELBOW As Boolean = Data_table_compiled.Rows(i).Item("ELLBOW")
                                        Dim Page1 As Integer = Data_table_compiled.Rows(i).Item("PAGE")

                                        Dim Mat As String = "XXX"
                                        If IsDBNull(Data_table_compiled.Rows(i).Item("MAT")) = False Then
                                            Mat = Data_table_compiled.Rows(i).Item("MAT")
                                        End If

                                        Dim DeltaX As Double = Data_table_compiled.Rows(i).Item("DELTAX")
                                        Dim Ellbow_ID As String = ""
                                        Dim Ellbow_Descr As String = ""

                                        If IsDBNull(Data_table_compiled.Rows(i).Item("ELLBOW_ID")) = False Then
                                            Ellbow_ID = Data_table_compiled.Rows(i).Item("ELLBOW_ID")
                                        End If

                                        If IsDBNull(Data_table_compiled.Rows(i).Item("ELLBOW_DESCR")) = False Then
                                            Ellbow_Descr = Data_table_compiled.Rows(i).Item("ELLBOW_DESCR")
                                        End If

                                        Dim Sta_middle As Double = 0
                                        If IsDBNull(Data_table_compiled.Rows(i).Item("STA")) = False Then
                                            Sta_middle = Data_table_compiled.Rows(i).Item("STA")
                                        End If


                                        Dim Band_Page_length As Double = Data_table_compiled.Rows(i).Item("BAND_LENGTH")

                                        Dim Center_band_amount As Double = Viewport_ms_length / 2 - Band_Page_length / 2
                                        'xxx()



                                        If Not Page1 = Page_previous Then
                                            Deltax_total = 0

                                            Dim Band_text As New DBText
                                            Band_text.Layer = "NO PLOT"
                                            If RadioButton_left_right.Checked = True Then
                                                Band_text.Justify = AttachmentPoint.MiddleRight
                                                Band_text.AlignmentPoint = New Point3d(Start_point.X + Center_band_amount - 75, Start_point.Y - (Page1 - 1) * Band_Separation, 0)
                                            Else
                                                Band_text.Justify = AttachmentPoint.MiddleLeft
                                                Band_text.AlignmentPoint = New Point3d(Start_point.X - Center_band_amount + 75, Start_point.Y - (Page1 - 1) * Band_Separation, 0)
                                            End If
                                            Band_text.TextString = CStr(Page1)
                                            Band_text.Height = 50
                                            BTrecord.AppendEntity(Band_text)
                                            Trans1.AddNewlyCreatedDBObject(Band_text, True)

                                            Dim Viewport_matchline As New Polyline
                                            Viewport_matchline.Layer = "NO PLOT"
                                            Viewport_matchline.Closed = True
                                            Viewport_matchline.ColorIndex = 1
                                            BTrecord.AppendEntity(Viewport_matchline)
                                            Trans1.AddNewlyCreatedDBObject(Viewport_matchline, True)

                                            If RadioButton_left_right.Checked = True Then
                                                Dim Pt1 As New Point2d(Start_point.X + Viewport_ms_length / 2 - Lungime_matchline_viewport / 2, Start_point.Y - (Page1 - 1) * Band_Separation + Viewport_ms_height)
                                                Dim Pt2 As New Point2d(Start_point.X + Viewport_ms_length / 2 + Lungime_matchline_viewport / 2, Start_point.Y - (Page1 - 1) * Band_Separation + Viewport_ms_height)
                                                Dim Pt3 As New Point2d(Start_point.X + Viewport_ms_length / 2 + Lungime_matchline_viewport / 2, Start_point.Y - (Page1 - 1) * Band_Separation)
                                                Dim Pt4 As New Point2d(Start_point.X + Viewport_ms_length / 2 - Lungime_matchline_viewport / 2, Start_point.Y - (Page1 - 1) * Band_Separation)
                                                Viewport_matchline.AddVertexAt(0, Pt1, 0, 0, 0)
                                                Viewport_matchline.AddVertexAt(1, Pt2, 0, 0, 0)
                                                Viewport_matchline.AddVertexAt(2, Pt3, 0, 0, 0)
                                                Viewport_matchline.AddVertexAt(3, Pt4, 0, 0, 0)
                                            Else
                                                Dim Pt1 As New Point2d(Start_point.X - Viewport_ms_length / 2 + Lungime_matchline_viewport / 2, Start_point.Y - (Page1 - 1) * Band_Separation + Viewport_ms_height)
                                                Dim Pt2 As New Point2d(Start_point.X - Viewport_ms_length / 2 - Lungime_matchline_viewport / 2, Start_point.Y - (Page1 - 1) * Band_Separation + Viewport_ms_height)
                                                Dim Pt3 As New Point2d(Start_point.X - Viewport_ms_length / 2 - Lungime_matchline_viewport / 2, Start_point.Y - (Page1 - 1) * Band_Separation)
                                                Dim Pt4 As New Point2d(Start_point.X - Viewport_ms_length / 2 + Lungime_matchline_viewport / 2, Start_point.Y - (Page1 - 1) * Band_Separation)
                                                Viewport_matchline.AddVertexAt(0, Pt1, 0, 0, 0)
                                                Viewport_matchline.AddVertexAt(1, Pt2, 0, 0, 0)
                                                Viewport_matchline.AddVertexAt(2, Pt3, 0, 0, 0)
                                                Viewport_matchline.AddVertexAt(3, Pt4, 0, 0, 0)
                                            End If


                                            Dim Viewport_big As New Polyline
                                            Viewport_big.Layer = "NO PLOT"
                                            Viewport_big.Closed = True
                                            Viewport_big.ColorIndex = 3
                                            BTrecord.AppendEntity(Viewport_big)
                                            Trans1.AddNewlyCreatedDBObject(Viewport_big, True)

                                            If RadioButton_left_right.Checked = True Then
                                                Dim Pt1 As New Point2d(Start_point.X, Start_point.Y - (Page1 - 1) * Band_Separation + Viewport_ms_height)
                                                Dim Pt2 As New Point2d(Start_point.X + Viewport_ms_length, Start_point.Y - (Page1 - 1) * Band_Separation + Viewport_ms_height)
                                                Dim Pt3 As New Point2d(Start_point.X + Viewport_ms_length, Start_point.Y - (Page1 - 1) * Band_Separation)
                                                Dim Pt4 As New Point2d(Start_point.X, Start_point.Y - (Page1 - 1) * Band_Separation)
                                                Viewport_big.AddVertexAt(0, Pt1, 0, 0, 0)
                                                Viewport_big.AddVertexAt(1, Pt2, 0, 0, 0)
                                                Viewport_big.AddVertexAt(2, Pt3, 0, 0, 0)
                                                Viewport_big.AddVertexAt(3, Pt4, 0, 0, 0)
                                            Else
                                                Dim Pt1 As New Point2d(Start_point.X, Start_point.Y - (Page1 - 1) * Band_Separation + Viewport_ms_height)
                                                Dim Pt2 As New Point2d(Start_point.X - Viewport_ms_length, Start_point.Y - (Page1 - 1) * Band_Separation + Viewport_ms_height)
                                                Dim Pt3 As New Point2d(Start_point.X - Viewport_ms_length, Start_point.Y - (Page1 - 1) * Band_Separation)
                                                Dim Pt4 As New Point2d(Start_point.X, Start_point.Y - (Page1 - 1) * Band_Separation)
                                                Viewport_big.AddVertexAt(0, Pt1, 0, 0, 0)
                                                Viewport_big.AddVertexAt(1, Pt2, 0, 0, 0)
                                                Viewport_big.AddVertexAt(2, Pt3, 0, 0, 0)
                                                Viewport_big.AddVertexAt(3, Pt4, 0, 0, 0)
                                            End If
                                        End If

                                        Page_previous = Page1
                                        Dim Pt_ins As Point3d

                                        If RadioButton_left_right.Checked = True Then
                                            Pt_ins = New Point3d(Start_point.X + Center_band_amount + Deltax_total, Start_point.Y - (Page1 - 1) * Band_Separation, 0)
                                        Else
                                            Pt_ins = New Point3d(Start_point.X - Center_band_amount - Deltax_total, Start_point.Y - (Page1 - 1) * Band_Separation, 0)
                                        End If

                                        Dim Nume_block As String
                                        Dim Nume_visibilitate As String = ""

                                        Dim Colectie_nume_atr As New Specialized.StringCollection
                                        Dim Colectie_valori As New Specialized.StringCollection
                                        If IS_ELBOW = True Then
                                            Nume_block = ComboBox_blocks_ELBOWS_AS_MAT.Text
                                            If Not ComboBox_STA1_att_ELBOW_AS_MAT.Text = "" Then
                                                Colectie_nume_atr.Add(ComboBox_STA1_att_ELBOW_AS_MAT.Text)
                                                If Is_Canada = False Then
                                                    Colectie_valori.Add(Get_chainage_feet_from_double(Sta1 + Get_equation_value(Sta1), Round1))
                                                Else
                                                    Colectie_valori.Add(Get_chainage_from_double(Sta1 + Get_equation_value(Sta1), Round1))
                                                End If

                                            End If

                                            If Not ComboBox_STA2_att_ELBOW_AS_MAT.Text = "" Then
                                                Colectie_nume_atr.Add(ComboBox_STA2_att_ELBOW_AS_MAT.Text)
                                                If Is_Canada = False Then
                                                    Colectie_valori.Add(Get_chainage_feet_from_double(Sta2 + Get_equation_value(Sta2), Round1))
                                                Else
                                                    Colectie_valori.Add(Get_chainage_from_double(Sta2 + Get_equation_value(Sta2), Round1))
                                                End If

                                            End If

                                            If Not ComboBox_station_middle_att_ELBOW_AS_MAT.Text = "" Then
                                                Colectie_nume_atr.Add(ComboBox_station_middle_att_ELBOW_AS_MAT.Text)
                                                If Is_Canada = False Then
                                                    Colectie_valori.Add(Get_chainage_feet_from_double(Sta_middle + Get_equation_value(Sta_middle), Round1))
                                                Else
                                                    Colectie_valori.Add(Get_chainage_from_double(Sta_middle + Get_equation_value(Sta_middle), Round1))
                                                End If

                                            End If


                                            If Not ComboBox_EXTRA_att_ELBOW_AS_MAT.Text = "" Then
                                                Colectie_nume_atr.Add(ComboBox_EXTRA_att_ELBOW_AS_MAT.Text)
                                                If Is_Canada = False Then
                                                    Colectie_valori.Add(Get_chainage_feet_from_double(Sta2 + Get_equation_value(Sta2), Round1))
                                                Else
                                                    Colectie_valori.Add(Get_chainage_from_double(Sta2 + Get_equation_value(Sta2), Round1))
                                                End If
                                            End If
                                            If Not ComboBox_MAT_ELBOW_AS_MAT.Text = "" Then
                                                Colectie_nume_atr.Add(ComboBox_MAT_ELBOW_AS_MAT.Text)
                                                Colectie_valori.Add(Mat)
                                            End If

                                            If Not ComboBox_LEN_att_ELBOW_AS_MAT.Text = "" Then
                                                Colectie_nume_atr.Add(ComboBox_LEN_att_ELBOW_AS_MAT.Text)


                                                If Is_Canada = False Then
                                                    Colectie_valori.Add(Get_String_Rounded(Abs(Round(Sta2, Round1) - Round(Sta1, Round1)), Round1) & "'")
                                                Else
                                                    Colectie_valori.Add(Get_String_Rounded(Abs(Round(Sta2, Round1) - Round(Sta1, Round1)), Round1))
                                                End If
                                            End If

                                            If Is_Canada = True And Not Ellbow_Descr = "" Then
                                                If Not ComboBox_descr_ELBOW_AS_MAT.Text = "" Then
                                                    Colectie_nume_atr.Add(ComboBox_descr_ELBOW_AS_MAT.Text)
                                                    Colectie_valori.Add(Ellbow_Descr)
                                                End If
                                            End If

                                            If Is_Canada = True And Not Ellbow_ID = "" Then
                                                If Not ComboBox_dwg_ID_ELBOW_AS_MAT.Text = "" Then
                                                    Colectie_nume_atr.Add(ComboBox_dwg_ID_ELBOW_AS_MAT.Text)
                                                    Colectie_valori.Add(Ellbow_ID)
                                                End If
                                            End If
                                        Else

                                            Nume_block = ComboBox_blocks_MAT.Text
                                            If Not ComboBox_mat_STA1.Text = "" Then
                                                Colectie_nume_atr.Add(ComboBox_mat_STA1.Text)
                                                If Is_Canada = False Then
                                                    Colectie_valori.Add(Get_chainage_feet_from_double(Sta1 + Get_equation_value(Sta1), Round1))
                                                Else
                                                    Colectie_valori.Add(Get_chainage_from_double(Sta1 + Get_equation_value(Sta1), Round1))
                                                End If
                                            End If
                                            If Not ComboBox_mat_STA1_duplicate.Text = "" Then
                                                Colectie_nume_atr.Add(ComboBox_mat_STA1_duplicate.Text)
                                                If Is_Canada = False Then
                                                    Colectie_valori.Add(Get_chainage_feet_from_double(Sta1 + Get_equation_value(Sta1), Round1))
                                                Else
                                                    Colectie_valori.Add(Get_chainage_from_double(Sta1 + Get_equation_value(Sta1), Round1))
                                                End If
                                            End If
                                            If Not ComboBox_mat_STA2.Text = "" Then
                                                Colectie_nume_atr.Add(ComboBox_mat_STA2.Text)
                                                If Is_Canada = False Then
                                                    Colectie_valori.Add(Get_chainage_feet_from_double(Sta2 + Get_equation_value(Sta2), Round1))
                                                Else
                                                    Colectie_valori.Add(Get_chainage_from_double(Sta2 + Get_equation_value(Sta2), Round1))
                                                End If
                                            End If
                                            If Not ComboBox_mat_STA2_duplicate.Text = "" Then
                                                Colectie_nume_atr.Add(ComboBox_mat_STA2_duplicate.Text)
                                                If Is_Canada = False Then
                                                    Colectie_valori.Add(Get_chainage_feet_from_double(Sta2 + Get_equation_value(Sta2), Round1))
                                                Else
                                                    Colectie_valori.Add(Get_chainage_from_double(Sta2 + Get_equation_value(Sta2), Round1))
                                                End If
                                            End If
                                            If Not ComboBox_mat_mat.Text = "" Then
                                                Colectie_nume_atr.Add(ComboBox_mat_mat.Text)
                                                Colectie_valori.Add(Mat)
                                                Select Case Mat
                                                    Case "1"
                                                        Nume_visibilitate = "Mat-1"
                                                    Case Else
                                                        Nume_visibilitate = "Heavy Wall"


                                                End Select

                                            End If
                                            If Not ComboBox_mat_len.Text = "" Then
                                                Colectie_nume_atr.Add(ComboBox_mat_len.Text)
                                                If Is_Canada = False Then
                                                    Colectie_valori.Add(Get_String_Rounded(Abs(Round(Sta2, Round1) - Round(Sta1, Round1)), Round1) & "'")
                                                Else
                                                    Colectie_valori.Add(Get_String_Rounded(Abs(Round(Sta2, Round1) - Round(Sta1, Round1)), Round1))
                                                End If

                                            End If
                                        End If

                                        Dim BR1 As BlockReference
                                        If Not Nume_block = "" Then
                                            BR1 = InsertBlock_with_multiple_atributes("", Nume_block, Pt_ins, 1, BTrecord, "0", Colectie_nume_atr, Colectie_valori)
                                            Stretch_block(BR1, "Distance1", DeltaX)


                                            If Nume_block = ComboBox_blocks_MAT.Text Then
                                                If Not Nume_visibilitate = "" Then
                                                    change_visibility_block(BR1, Nume_visibilitate)
                                                End If
                                            End If
                                            Deltax_total = Deltax_total + DeltaX
                                        End If



                                        If Is_Canada = True Then
                                            If IsNothing(Data_table_transitions) = False Then
                                                If Data_table_transitions.Rows.Count > 0 Then
                                                    If Not ComboBox_blocks_transition_weld_canada.Text = "" Then
                                                        For j = 0 To Data_table_transitions.Rows.Count - 1
                                                            Dim Tw1 As Double = Round(Data_table_transitions.Rows(j).Item("STA"), Round1)
                                                            If Round(Sta1, Round1) = Tw1 Then
                                                                Dim BR_TW As BlockReference = InsertBlock_with_multiple_atributes("", ComboBox_blocks_transition_weld_canada.Text, Pt_ins, 1, BTrecord, "0", Colectie_nume_atr, Colectie_valori)
                                                            End If
                                                        Next
                                                    End If
                                                End If
                                            End If
                                        End If

                                        If IsNothing(Data_table_buoyancy_canada) = False Then
                                            If Data_table_buoyancy_canada.Rows.Count > 0 Then
                                                For k = 0 To Data_table_buoyancy_canada.Rows.Count - 1
                                                    Dim B_Start As Double = Data_table_buoyancy_canada.Rows(k).Item("CHAINAGE_START")
                                                    Dim B_End As Double = Data_table_buoyancy_canada.Rows(k).Item("CHAINAGE_END")
                                                    Dim B_count1 As Integer = Data_table_buoyancy_canada.Rows(k).Item("COUNT")

                                                    If B_Start >= Sta1 And B_Start < Sta2 Then
                                                        Dim Length_of_block As Double = DeltaX
                                                        Dim Dist_from_start As Double = B_Start - Sta1
                                                        Dim X_from_start As Double = Dist_from_start * Length_of_block / (Sta2 - Sta1)
                                                        If RadioButton_left_right.Checked = True Then
                                                            Data_table_buoyancy_canada.Rows(k).Item("X") = Pt_ins.X + X_from_start
                                                        Else
                                                            Data_table_buoyancy_canada.Rows(k).Item("X") = Pt_ins.X - X_from_start
                                                        End If
                                                        Data_table_buoyancy_canada.Rows(k).Item("Y") = Pt_ins.Y

                                                    End If

                                                    If B_End >= Sta1 And B_End <= Sta2 Then
                                                        Dim Length_of_block As Double = DeltaX
                                                        Dim Dist_from_start As Double = B_End - Sta1
                                                        Dim X_from_start As Double = Dist_from_start * Length_of_block / (Sta2 - Sta1)

                                                        Dim Stretch1 As Double
                                                        If RadioButton_left_right.Checked = True Then
                                                            Stretch1 = Pt_ins.X + X_from_start
                                                        Else
                                                            Stretch1 = Pt_ins.X - X_from_start
                                                        End If
                                                        If IsDBNull(Data_table_buoyancy_canada.Rows(k).Item("STRETCH")) = True Then
                                                            Data_table_buoyancy_canada.Rows(k).Item("STRETCH") = Abs(Stretch1 - Data_table_buoyancy_canada.Rows(k).Item("X"))
                                                        End If
                                                    End If
                                                Next
                                            End If ' asta e de la  If Data_table_buoyancy.Rows.Count > 0
                                        End If ' asta e de la If IsNothing(Data_table_buoyancy) = False

                                        If IsNothing(Data_table_pipes_Canada) = False Then
                                            If Data_table_pipes_Canada.Rows.Count > 0 Then
                                                For k = 0 To Data_table_pipes_Canada.Rows.Count - 1
                                                    Dim P_chainage As Double = Data_table_pipes_Canada.Rows(k).Item("CHAINAGE")
                                                    If P_chainage >= Sta1 And P_chainage < Sta2 Then
                                                        Dim Length_of_block As Double = DeltaX
                                                        Dim Dist_from_start As Double = P_chainage - Sta1
                                                        Dim X_from_start As Double = Dist_from_start * Length_of_block / (Sta2 - Sta1)

                                                        If IsDBNull(Data_table_pipes_Canada.Rows(k).Item("X")) = True Then
                                                            If RadioButton_left_right.Checked = True Then
                                                                Data_table_pipes_Canada.Rows(k).Item("X") = Pt_ins.X + X_from_start
                                                            Else
                                                                Data_table_pipes_Canada.Rows(k).Item("X") = Pt_ins.X - X_from_start
                                                            End If
                                                            Data_table_pipes_Canada.Rows(k).Item("Y") = Pt_ins.Y
                                                        End If

                                                    End If
                                                Next
                                            End If ' asta e de la  If Data_table_pipes.Rows.Count > 0
                                        End If ' asta e de la If IsNothing(Data_table_pipes) = False


                                        If IsNothing(Data_table_water_canada) = False Then
                                            If Data_table_water_canada.Rows.Count > 0 Then
                                                For k = 0 To Data_table_water_canada.Rows.Count - 1
                                                    Dim W_chainage As Double = Data_table_water_canada.Rows(k).Item("CHAINAGE")
                                                    If W_chainage >= Sta1 And W_chainage < Sta2 Then
                                                        Dim Length_of_block As Double = DeltaX
                                                        Dim Dist_from_start As Double = W_chainage - Sta1
                                                        Dim X_from_start As Double = Dist_from_start * Length_of_block / (Sta2 - Sta1)

                                                        If IsDBNull(Data_table_water_canada.Rows(k).Item("X")) = True Then
                                                            If RadioButton_left_right.Checked = True Then
                                                                Data_table_water_canada.Rows(k).Item("X") = Pt_ins.X + X_from_start
                                                            Else
                                                                Data_table_water_canada.Rows(k).Item("X") = Pt_ins.X - X_from_start
                                                            End If
                                                            Data_table_water_canada.Rows(k).Item("Y") = Pt_ins.Y
                                                        End If

                                                    End If
                                                Next
                                            End If ' asta e de la  If Data_table_water.Rows.Count > 0
                                        End If ' asta e de la If IsNothing(Data_table_water) = False

                                        If IsNothing(Data_table_cathodic_canada) = False Then
                                            If Data_table_cathodic_canada.Rows.Count > 0 Then
                                                For k = 0 To Data_table_cathodic_canada.Rows.Count - 1
                                                    Dim W_chainage As Double = Data_table_cathodic_canada.Rows(k).Item("CHAINAGE")
                                                    If W_chainage >= Sta1 And W_chainage < Sta2 Then
                                                        Dim Length_of_block As Double = DeltaX
                                                        Dim Dist_from_start As Double = W_chainage - Sta1
                                                        Dim X_from_start As Double = Dist_from_start * Length_of_block / (Sta2 - Sta1)

                                                        If IsDBNull(Data_table_cathodic_canada.Rows(k).Item("X")) = True Then
                                                            If RadioButton_left_right.Checked = True Then
                                                                Data_table_cathodic_canada.Rows(k).Item("X") = Pt_ins.X + X_from_start
                                                            Else
                                                                Data_table_cathodic_canada.Rows(k).Item("X") = Pt_ins.X - X_from_start
                                                            End If
                                                            Data_table_cathodic_canada.Rows(k).Item("Y") = Pt_ins.Y
                                                        End If

                                                    End If
                                                Next
                                            End If ' asta e de la  If Data_table_cathodic.Rows.Count > 0
                                        End If ' asta e de la If IsNothing(Data_table_cathodic) = False


                                        If IsNothing(Data_table_fitting) = False Then
                                            If Data_table_fitting.Rows.Count > 0 Then
                                                For k = 0 To Data_table_fitting.Rows.Count - 1
                                                    Dim W_chainage As Double = Data_table_fitting.Rows(k).Item("STATION")
                                                    If W_chainage >= Sta1 And W_chainage < Sta2 Then
                                                        Dim Length_of_block As Double = DeltaX
                                                        Dim Dist_from_start As Double = W_chainage - Sta1
                                                        Dim X_from_start As Double = Dist_from_start * Length_of_block / (Sta2 - Sta1)

                                                        If IsDBNull(Data_table_fitting.Rows(k).Item("X")) = True Then
                                                            If RadioButton_left_right.Checked = True Then
                                                                Data_table_fitting.Rows(k).Item("X") = Pt_ins.X + X_from_start
                                                            Else
                                                                Data_table_fitting.Rows(k).Item("X") = Pt_ins.X - X_from_start
                                                            End If
                                                            Data_table_fitting.Rows(k).Item("Y") = Pt_ins.Y
                                                        End If

                                                    End If
                                                Next
                                            End If ' asta e de la  If Data_table_cathodic.Rows.Count > 0
                                        End If ' asta e de la If IsNothing(Data_table_cathodic) = False

                                        If IsNothing(Data_table_cl_crossing) = False Then
                                            If Data_table_cl_crossing.Rows.Count > 0 Then
                                                For k = 0 To Data_table_cl_crossing.Rows.Count - 1
                                                    Dim W_chainage As Double = Data_table_cl_crossing.Rows(k).Item("STATION")
                                                    If W_chainage >= Sta1 And W_chainage < Sta2 Then
                                                        Dim Length_of_block As Double = DeltaX
                                                        Dim Dist_from_start As Double = W_chainage - Sta1
                                                        Dim X_from_start As Double = Dist_from_start * Length_of_block / (Sta2 - Sta1)

                                                        If IsDBNull(Data_table_cl_crossing.Rows(k).Item("X")) = True Then
                                                            If RadioButton_left_right.Checked = True Then
                                                                Data_table_cl_crossing.Rows(k).Item("X") = Pt_ins.X + X_from_start
                                                            Else
                                                                Data_table_cl_crossing.Rows(k).Item("X") = Pt_ins.X - X_from_start
                                                            End If
                                                            Data_table_cl_crossing.Rows(k).Item("Y") = Pt_ins.Y
                                                        End If

                                                    End If
                                                Next
                                            End If ' asta e de la  If Data_table_cathodic.Rows.Count > 0
                                        End If ' asta e de la If IsNothing(Data_table_cathodic) = False

                                        If IsNothing(Data_table_cad_weld) = False Then
                                            If Data_table_cad_weld.Rows.Count > 0 Then
                                                For k = 0 To Data_table_cad_weld.Rows.Count - 1
                                                    Dim W_chainage As Double = Data_table_cad_weld.Rows(k).Item("STATION")
                                                    If W_chainage >= Sta1 And W_chainage < Sta2 Then
                                                        Dim Length_of_block As Double = DeltaX
                                                        Dim Dist_from_start As Double = W_chainage - Sta1
                                                        Dim X_from_start As Double = Dist_from_start * Length_of_block / (Sta2 - Sta1)

                                                        If IsDBNull(Data_table_cad_weld.Rows(k).Item("X")) = True Then
                                                            If RadioButton_left_right.Checked = True Then
                                                                Data_table_cad_weld.Rows(k).Item("X") = Pt_ins.X + X_from_start
                                                            Else
                                                                Data_table_cad_weld.Rows(k).Item("X") = Pt_ins.X - X_from_start
                                                            End If
                                                            Data_table_cad_weld.Rows(k).Item("Y") = Pt_ins.Y
                                                        End If

                                                    End If
                                                Next
                                            End If ' asta e de la  If Data_table_cathodic.Rows.Count > 0
                                        End If ' asta e de la If IsNothing(Data_table_cathodic) = False

                                        If IsNothing(Data_table_river_weights) = False Then
                                            If Data_table_river_weights.Rows.Count > 0 Then
                                                For k = 0 To Data_table_river_weights.Rows.Count - 1
                                                    Dim rW_chainage As Double = Data_table_river_weights.Rows(k).Item("STATION")
                                                    If rW_chainage >= Sta1 And rW_chainage < Sta2 Then
                                                        Dim Length_of_block As Double = DeltaX
                                                        Dim Dist_from_start As Double = rW_chainage - Sta1
                                                        Dim X_from_start As Double = Dist_from_start * Length_of_block / (Sta2 - Sta1)

                                                        If IsDBNull(Data_table_river_weights.Rows(k).Item("X")) = True Then
                                                            If RadioButton_left_right.Checked = True Then
                                                                Data_table_river_weights.Rows(k).Item("X") = Pt_ins.X + X_from_start
                                                            Else
                                                                Data_table_river_weights.Rows(k).Item("X") = Pt_ins.X - X_from_start
                                                            End If
                                                            Data_table_river_weights.Rows(k).Item("Y") = Pt_ins.Y
                                                        End If

                                                    End If
                                                Next
                                            End If ' asta e de la  If Data_table_river_weights.Rows.Count > 0
                                        End If ' asta e de la If IsNothing(Data_table_river_weights) = False



                                    End If 'asta e de la If IsDBNull(Data_table_compiled.Rows(i).Item("STA1")) = False
                                Next ' ASTA e de la  For i = 0 To Data_table_compiled.Rows.Count - 1

                                If IsNothing(Data_table_buoyancy_canada) = False Then
                                    If Data_table_buoyancy_canada.Rows.Count > 0 Then
                                        For k = 0 To Data_table_buoyancy_canada.Rows.Count - 1

                                            Dim Nume_visibilitate As String = ""
                                            If IsDBNull(Data_table_buoyancy_canada.Rows(k).Item("CHAINAGE_START")) = False And
                                                IsDBNull(Data_table_buoyancy_canada.Rows(k).Item("CHAINAGE_END")) = False And
                                                IsDBNull(Data_table_buoyancy_canada.Rows(k).Item("COUNT")) = False And
                                                IsDBNull(Data_table_buoyancy_canada.Rows(k).Item("X")) = False And
                                                IsDBNull(Data_table_buoyancy_canada.Rows(k).Item("Y")) = False And
                                                IsDBNull(Data_table_buoyancy_canada.Rows(k).Item("STRETCH")) = False Then

                                                Dim B_Start As Double = Data_table_buoyancy_canada.Rows(k).Item("CHAINAGE_START")
                                                Dim B_End As Double = Data_table_buoyancy_canada.Rows(k).Item("CHAINAGE_END")
                                                Dim B_count1 As Integer = Data_table_buoyancy_canada.Rows(k).Item("COUNT")

                                                Dim B_spacing As String = ""
                                                If IsDBNull(Data_table_buoyancy_canada.Rows(k).Item("SPACING")) = False Then
                                                    Dim B_spacing1 As Double = Data_table_buoyancy_canada.Rows(k).Item("SPACING")
                                                    B_spacing = Get_String_Rounded(B_spacing1, Round1)
                                                End If


                                                Dim B_TYPE As String = ""
                                                If IsDBNull(Data_table_buoyancy_canada.Rows(k).Item("DESCRIPTION")) = False Then
                                                    B_TYPE = Data_table_buoyancy_canada.Rows(k).Item("DESCRIPTION")
                                                End If



                                                Dim Pt_ins As New Point3d(Data_table_buoyancy_canada.Rows(k).Item("X"), Data_table_buoyancy_canada.Rows(k).Item("Y"), 0)
                                                Dim B_stretch As Double = Data_table_buoyancy_canada.Rows(k).Item("STRETCH")




                                                Dim Nume_block As String = ""

                                                If B_count1 = 1 Then

                                                    If Not ComboBox_blocks_pipe_crossings_canada.Text = "" Then
                                                        Nume_block = ComboBox_blocks_pipe_crossings_canada.Text
                                                    End If
                                                End If


                                                If B_count1 = 2 Then

                                                    If B_TYPE.ToUpper.Contains("SCREW") = True Then
                                                        If Not ComboBox_blocks_pipe_crossings_canada.Text = "" Then
                                                            Nume_block = ComboBox_blocks_pipe_crossings_canada.Text
                                                        End If
                                                    Else

                                                        If Not ComboBox_blocks_screw_anchor_multiple_canada.Text = "" Then
                                                            Nume_block = ComboBox_blocks_screw_anchor_multiple_canada.Text
                                                        End If
                                                    End If

                                                End If


                                                If B_count1 > 2 Then

                                                    If Not ComboBox_blocks_screw_anchor_multiple_canada.Text = "" Then
                                                        Nume_block = ComboBox_blocks_screw_anchor_multiple_canada.Text
                                                    End If
                                                End If


                                                If Not Nume_block = "" Then
                                                    Dim Colectie_nume_atr As New Specialized.StringCollection
                                                    Dim Colectie_valori As New Specialized.StringCollection

                                                    If Not ComboBox_att_screw_anch_mult_1.Text = "" Then
                                                        Colectie_nume_atr.Add(ComboBox_att_screw_anch_mult_1.Text)
                                                        If Is_Canada = False Then
                                                            Colectie_valori.Add(Get_chainage_feet_from_double(B_Start + Get_equation_value(B_Start), Round1))
                                                        Else
                                                            Colectie_valori.Add(Get_chainage_from_double(B_Start + Get_equation_value(B_Start), Round1))
                                                        End If
                                                    End If

                                                    If Not ComboBox_att_screw_anch_mult_2.Text = "" Then
                                                        Colectie_nume_atr.Add(ComboBox_att_screw_anch_mult_2.Text)
                                                        If Is_Canada = False Then
                                                            Colectie_valori.Add(Get_chainage_feet_from_double(B_End + Get_equation_value(B_End), Round1))
                                                        Else
                                                            Colectie_valori.Add(Get_chainage_from_double(B_End + Get_equation_value(B_End), Round1))
                                                        End If
                                                    End If

                                                    If Not ComboBox_att_screw_anch_mult_3.Text = "" Then
                                                        Colectie_nume_atr.Add(ComboBox_att_screw_anch_mult_3.Text)
                                                        Dim Extra1 As String = ""
                                                        If B_TYPE.ToUpper.Contains("SCREW") = True Then
                                                            Extra1 = " SA"
                                                            Nume_visibilitate = "Screw anchors"
                                                        ElseIf B_TYPE.ToUpper.Contains("RIVER") = True Then
                                                            Extra1 = " RW"
                                                            Nume_visibilitate = "River weights"
                                                        Else
                                                            Extra1 = " GFW"
                                                            Nume_visibilitate = "Bag Weights"
                                                        End If
                                                        Colectie_valori.Add(B_count1 & Extra1)
                                                    End If

                                                    If Not ComboBox_att_screw_anch_mult_4.Text = "" Then
                                                        Colectie_nume_atr.Add(ComboBox_att_screw_anch_mult_4.Text)
                                                        Dim Extra1 As String = " C/C"
                                                        Colectie_valori.Add(B_spacing & Extra1)
                                                    End If



                                                    If B_count1 = 1 Or B_count1 = 2 Then
                                                        If Not ComboBox_att_pipe1.Text = "" Then
                                                            Colectie_nume_atr.Add(ComboBox_att_pipe1.Text)
                                                            If Is_Canada = False Then
                                                                Colectie_valori.Add(Get_chainage_feet_from_double(B_Start + Get_equation_value(B_Start), Round1))
                                                            Else
                                                                Colectie_valori.Add(Get_chainage_from_double(B_Start + Get_equation_value(B_Start), Round1))
                                                            End If
                                                        End If

                                                        If Not ComboBox_att_pipe2.Text = "" Then
                                                            Colectie_nume_atr.Add(ComboBox_att_pipe2.Text)
                                                            Colectie_valori.Add(B_TYPE)
                                                        End If
                                                    End If


                                                    Dim BR1 As BlockReference = InsertBlock_with_multiple_atributes("", Nume_block, Pt_ins, 1, BTrecord, "0", Colectie_nume_atr, Colectie_valori)

                                                    If Not Nume_visibilitate = "" Then
                                                        change_visibility_block(BR1, Nume_visibilitate)
                                                    End If
                                                    If Nume_block = ComboBox_blocks_screw_anchor_multiple_canada.Text Then
                                                        Stretch_block(BR1, "Distance1", B_stretch)
                                                    End If

                                                    If B_count1 = 2 And B_TYPE.ToUpper.Contains("SCREW") = True Then
                                                        If Not ComboBox_att_pipe1.Text = "" Then

                                                            For q = 0 To Colectie_nume_atr.Count - 1
                                                                Colectie_nume_atr.RemoveAt(0)
                                                            Next

                                                            For q = 0 To Colectie_valori.Count - 1
                                                                Colectie_valori.RemoveAt(0)
                                                            Next

                                                            Colectie_nume_atr.Add(ComboBox_att_pipe1.Text)
                                                            If Is_Canada = False Then
                                                                Colectie_valori.Add(Get_chainage_feet_from_double(B_End + Get_equation_value(B_End), Round1))
                                                            Else
                                                                Colectie_valori.Add(Get_chainage_from_double(B_End + Get_equation_value(B_End), Round1))
                                                            End If
                                                        End If

                                                        If Not ComboBox_att_pipe2.Text = "" Then
                                                            Colectie_nume_atr.Add(ComboBox_att_pipe2.Text)
                                                            Colectie_valori.Add(B_TYPE)
                                                        End If
                                                        If RadioButton_left_right.Checked = True Then
                                                            Pt_ins = New Point3d(Data_table_buoyancy_canada.Rows(k).Item("X") + B_stretch, Data_table_buoyancy_canada.Rows(k).Item("Y"), 0)
                                                        Else
                                                            Pt_ins = New Point3d(Data_table_buoyancy_canada.Rows(k).Item("X") - B_stretch, Data_table_buoyancy_canada.Rows(k).Item("Y"), 0)
                                                        End If
                                                        Dim BR2 As BlockReference = InsertBlock_with_multiple_atributes("", Nume_block, Pt_ins, 1, BTrecord, "0", Colectie_nume_atr, Colectie_valori)
                                                        change_visibility_block(BR2, "SCREW ANCHOR")
                                                    End If




                                                End If 'ASTA E DE LA If Not Nume_block = "" 

                                            End If ' ASTA E DE LA IF ISDBNULL

                                        Next ' ASTA E DE LA For k = 0 To Data_table_buoyancy.Rows.Count - 1


                                    End If ' asta e de la  If Data_table_buoyancy.Rows.Count > 0
                                End If ' asta e de la If IsNothing(Data_table_buoyancy) = False

                                If IsNothing(Data_table_pipes_Canada) = False Then
                                    If Data_table_pipes_Canada.Rows.Count > 0 Then
                                        Dim Nume_visibilitate As String = ""
                                        For k = 0 To Data_table_pipes_Canada.Rows.Count - 1
                                            If IsDBNull(Data_table_pipes_Canada.Rows(k).Item("CHAINAGE")) = False Then
                                                Dim P_Chainage As Double = Data_table_pipes_Canada.Rows(k).Item("CHAINAGE")
                                                Dim Pt_ins As New Point3d(Data_table_pipes_Canada.Rows(k).Item("X"), Data_table_pipes_Canada.Rows(k).Item("Y"), 0)

                                                Dim P_Description As String = ""
                                                If IsDBNull(Data_table_pipes_Canada.Rows(k).Item("DESCRIPTION")) = False Then
                                                    P_Description = Data_table_pipes_Canada.Rows(k).Item("DESCRIPTION")
                                                End If

                                                Dim P_clearance As String = ""
                                                If IsDBNull(Data_table_pipes_Canada.Rows(k).Item("CLEARANCE")) = False Then
                                                    P_clearance = Data_table_pipes_Canada.Rows(k).Item("CLEARANCE")
                                                End If

                                                Dim P_cover As String = ""
                                                If IsDBNull(Data_table_pipes_Canada.Rows(k).Item("COVER")) = False Then
                                                    P_cover = Data_table_pipes_Canada.Rows(k).Item("COVER")
                                                End If

                                                Dim P_DWG_ID As String = ""
                                                If IsDBNull(Data_table_pipes_Canada.Rows(k).Item("DWG_ID")) = False Then
                                                    P_DWG_ID = Data_table_pipes_Canada.Rows(k).Item("DWG_ID")
                                                End If

                                                Dim Nume_block As String = ""
                                                If Not ComboBox_blocks_pipe_crossings_canada.Text = "" Then
                                                    Nume_block = ComboBox_blocks_pipe_crossings_canada.Text
                                                End If

                                                If Not Nume_block = "" Then
                                                    Dim Colectie_nume_atr As New Specialized.StringCollection
                                                    Dim Colectie_valori As New Specialized.StringCollection

                                                    If Not ComboBox_att_pipe1.Text = "" Then
                                                        Colectie_nume_atr.Add(ComboBox_att_pipe1.Text)
                                                        If Is_Canada = False Then
                                                            Colectie_valori.Add(Get_chainage_feet_from_double(P_Chainage + Get_equation_value(P_Chainage), Round1))
                                                        Else
                                                            Colectie_valori.Add(Get_chainage_from_double(P_Chainage + Get_equation_value(P_Chainage), Round1))
                                                        End If
                                                    End If

                                                    If Not ComboBox_att_pipe2.Text = "" And Not P_Description = "" Then
                                                        Colectie_nume_atr.Add(ComboBox_att_pipe2.Text)
                                                        Colectie_valori.Add(P_Description)
                                                        If P_Description.ToUpper.Contains("CABLE") = True Then
                                                            Nume_visibilitate = "CABLE"
                                                        Else
                                                            Nume_visibilitate = "PIPELINE"
                                                        End If
                                                    End If

                                                    If Not ComboBox_att_pipe4.Text = "" And Not P_clearance = "" Then
                                                        Colectie_nume_atr.Add(ComboBox_att_pipe4.Text)
                                                        Colectie_valori.Add(P_clearance)
                                                    End If

                                                    If Not ComboBox_att_pipe3.Text = "" And Not P_cover = "" Then
                                                        Colectie_nume_atr.Add(ComboBox_att_pipe3.Text)
                                                        Colectie_valori.Add(P_cover)
                                                    End If

                                                    If Not ComboBox_att_pipe5.Text = "" And Not P_DWG_ID = "" Then
                                                        Colectie_nume_atr.Add(ComboBox_att_pipe5.Text)
                                                        Colectie_valori.Add(P_DWG_ID)
                                                    End If

                                                    Dim BR1 As BlockReference = InsertBlock_with_multiple_atributes("", Nume_block, Pt_ins, 1, BTrecord, "0", Colectie_nume_atr, Colectie_valori)

                                                    If Not Nume_visibilitate = "" Then
                                                        change_visibility_block(BR1, Nume_visibilitate)
                                                    End If


                                                End If 'ASTA E DE LA If Not Nume_block = "" 

                                            End If ' ASTA E DE LA IF ISDBNULL

                                        Next ' ASTA E DE LA For k = 0 To Data_table_pipes.Rows.Count - 1
                                    End If ' asta e de la  If Data_table_pipes.Rows.Count > 0
                                End If ' asta e de la If IsNothing(Data_table_pipes) = False


                                If IsNothing(Data_table_water_canada) = False Then
                                    If Data_table_water_canada.Rows.Count > 0 Then
                                        For k = 0 To Data_table_water_canada.Rows.Count - 1
                                            Dim Nume_visibilitate As String = "GENERAL"

                                            If IsDBNull(Data_table_water_canada.Rows(k).Item("CHAINAGE")) = False Then
                                                Dim W_Chainage As Double = Data_table_water_canada.Rows(k).Item("CHAINAGE")
                                                Dim Pt_ins As New Point3d(Data_table_water_canada.Rows(k).Item("X"), Data_table_water_canada.Rows(k).Item("Y"), 0)

                                                Dim W_Description As String = ""
                                                If IsDBNull(Data_table_water_canada.Rows(k).Item("DESCRIPTION")) = False Then
                                                    W_Description = Data_table_water_canada.Rows(k).Item("DESCRIPTION")
                                                End If

                                                Dim W_cover As String = ""
                                                If IsDBNull(Data_table_water_canada.Rows(k).Item("COVER")) = False Then
                                                    W_cover = Data_table_water_canada.Rows(k).Item("COVER")
                                                End If

                                                Dim W_DWG_ID As String = ""
                                                If IsDBNull(Data_table_water_canada.Rows(k).Item("DWG_ID")) = False Then
                                                    W_DWG_ID = Data_table_water_canada.Rows(k).Item("DWG_ID")
                                                End If

                                                Dim Nume_block As String = ""
                                                If Not ComboBox_blocks_pipe_crossings_canada.Text = "" Then
                                                    Nume_block = ComboBox_blocks_pipe_crossings_canada.Text
                                                End If

                                                If Not Nume_block = "" Then
                                                    Dim Colectie_nume_atr As New Specialized.StringCollection
                                                    Dim Colectie_valori As New Specialized.StringCollection

                                                    If Not ComboBox_att_pipe1.Text = "" Then
                                                        Colectie_nume_atr.Add(ComboBox_att_pipe1.Text)
                                                        If Is_Canada = False Then
                                                            Colectie_valori.Add(Get_chainage_feet_from_double(W_Chainage + Get_equation_value(W_Chainage), Round1))
                                                        Else
                                                            Colectie_valori.Add(Get_chainage_from_double(W_Chainage + Get_equation_value(W_Chainage), Round1))
                                                        End If
                                                    End If

                                                    If Not ComboBox_att_pipe2.Text = "" And Not W_Description = "" Then
                                                        Colectie_nume_atr.Add(ComboBox_att_pipe2.Text)
                                                        Colectie_valori.Add(W_Description)
                                                    End If

                                                    If Not ComboBox_att_pipe3.Text = "" And Not W_cover = "" Then
                                                        Colectie_nume_atr.Add(ComboBox_att_pipe3.Text)
                                                        Colectie_valori.Add(W_cover)
                                                    End If

                                                    If Not ComboBox_att_pipe5.Text = "" And Not W_DWG_ID = "" Then
                                                        Colectie_nume_atr.Add(ComboBox_att_pipe5.Text)
                                                        Colectie_valori.Add(W_DWG_ID)
                                                    End If

                                                    Dim BR1 As BlockReference = InsertBlock_with_multiple_atributes("", Nume_block, Pt_ins, 1, BTrecord, "0", Colectie_nume_atr, Colectie_valori)
                                                    If Not Nume_visibilitate = "" Then
                                                        change_visibility_block(BR1, Nume_visibilitate)
                                                    End If
                                                End If 'ASTA E DE LA If Not Nume_block = "" 

                                            End If ' ASTA E DE LA IF ISDBNULL

                                        Next ' ASTA E DE LA For k = 0 To Data_table_water.Rows.Count - 1
                                    End If ' asta e de la  If Data_table_water.Rows.Count > 0
                                End If ' asta e de la If IsNothing(Data_table_water) = False

                                If IsNothing(Data_table_cathodic_canada) = False Then
                                    If Data_table_cathodic_canada.Rows.Count > 0 Then
                                        For k = 0 To Data_table_cathodic_canada.Rows.Count - 1
                                            Dim Nume_visibilitate As String = "CP"

                                            If IsDBNull(Data_table_cathodic_canada.Rows(k).Item("CHAINAGE")) = False Then
                                                Dim W_Chainage As Double = Data_table_cathodic_canada.Rows(k).Item("CHAINAGE")
                                                Dim Pt_ins As New Point3d(Data_table_cathodic_canada.Rows(k).Item("X"), Data_table_cathodic_canada.Rows(k).Item("Y"), 0)

                                                Dim W_Description As String = ""
                                                If IsDBNull(Data_table_cathodic_canada.Rows(k).Item("DESCRIPTION")) = False Then
                                                    W_Description = Data_table_cathodic_canada.Rows(k).Item("DESCRIPTION")
                                                End If


                                                Dim Nume_block As String = ""
                                                If Not ComboBox_blocks_pipe_crossings_canada.Text = "" Then
                                                    Nume_block = ComboBox_blocks_pipe_crossings_canada.Text
                                                End If

                                                If Not Nume_block = "" Then
                                                    Dim Colectie_nume_atr As New Specialized.StringCollection
                                                    Dim Colectie_valori As New Specialized.StringCollection

                                                    If Not ComboBox_att_pipe1.Text = "" Then
                                                        Colectie_nume_atr.Add(ComboBox_att_pipe1.Text)
                                                        If Is_Canada = False Then
                                                            Colectie_valori.Add(Get_chainage_feet_from_double(W_Chainage + Get_equation_value(W_Chainage), Round1))
                                                        Else
                                                            Colectie_valori.Add(Get_chainage_from_double(W_Chainage + Get_equation_value(W_Chainage), Round1))
                                                        End If
                                                    End If

                                                    If Not ComboBox_att_pipe2.Text = "" And Not W_Description = "" Then
                                                        Colectie_nume_atr.Add(ComboBox_att_pipe2.Text)
                                                        Colectie_valori.Add(W_Description)
                                                    End If

                                                    Dim BR1 As BlockReference = InsertBlock_with_multiple_atributes("", Nume_block, Pt_ins, 1, BTrecord, "0", Colectie_nume_atr, Colectie_valori)
                                                    If Not Nume_visibilitate = "" Then
                                                        change_visibility_block(BR1, Nume_visibilitate)
                                                    End If
                                                End If 'ASTA E DE LA If Not Nume_block = "" 

                                            End If ' ASTA E DE LA IF ISDBNULL

                                        Next ' ASTA E DE LA For k = 0 To Data_table_cathodic.Rows.Count - 1
                                    End If ' asta e de la  If Data_table_cathodic.Rows.Count > 0
                                End If ' asta e de la If IsNothing(Data_table_cathodic) = False

                                If IsNothing(Data_table_fitting) = False Then
                                    If Data_table_fitting.Rows.Count > 0 Then
                                        For k = 0 To Data_table_fitting.Rows.Count - 1


                                            If IsDBNull(Data_table_fitting.Rows(k).Item("STATION")) = False And IsDBNull(Data_table_fitting.Rows(k).Item("X")) = False And IsDBNull(Data_table_fitting.Rows(k).Item("Y")) = False Then
                                                Dim Fitting_Chainage As Double = Data_table_fitting.Rows(k).Item("STATION")

                                                Dim Pt_ins As New Point3d(Data_table_fitting.Rows(k).Item("X"), Data_table_fitting.Rows(k).Item("Y"), 0)

                                                Dim Fitting_Description As String = ""
                                                If IsDBNull(Data_table_fitting.Rows(k).Item("DESCRIPTION")) = False Then
                                                    Fitting_Description = Data_table_fitting.Rows(k).Item("DESCRIPTION")
                                                End If

                                                Dim Fitting_Material As String = ""
                                                If IsDBNull(Data_table_fitting.Rows(k).Item("MATERIAL")) = False Then
                                                    Fitting_Material = Data_table_fitting.Rows(k).Item("MATERIAL")
                                                End If

                                                Dim Nume_block As String = ""
                                                If Not ComboBox_blocks_fitting.Text = "" Then
                                                    Nume_block = ComboBox_blocks_fitting.Text
                                                End If

                                                If Not Nume_block = "" Then
                                                    Dim Colectie_nume_atr As New Specialized.StringCollection
                                                    Dim Colectie_valori As New Specialized.StringCollection

                                                    If Not ComboBox_STA_att_fitting.Text = "" Then
                                                        Colectie_nume_atr.Add(ComboBox_STA_att_fitting.Text)
                                                        If Is_Canada = False Then
                                                            Colectie_valori.Add(Get_chainage_feet_from_double(Fitting_Chainage + Get_equation_value(Fitting_Chainage), Round1))
                                                        Else
                                                            Colectie_valori.Add(Get_chainage_from_double(Fitting_Chainage + Get_equation_value(Fitting_Chainage), Round1))
                                                        End If
                                                    End If

                                                    If Not ComboBox_DESCRIPTION_att_fitting.Text = "" And Not Fitting_Description = "" Then
                                                        Colectie_nume_atr.Add(ComboBox_DESCRIPTION_att_fitting.Text)
                                                        Colectie_valori.Add(Fitting_Description)
                                                    End If
                                                    If Not ComboBox_MAT_att_fitting.Text = "" And Not Fitting_Material = "" Then
                                                        Colectie_nume_atr.Add(ComboBox_MAT_att_fitting.Text)
                                                        Colectie_valori.Add(Fitting_Material)
                                                    End If


                                                    Dim BR1 As BlockReference = InsertBlock_with_multiple_atributes("", Nume_block, Pt_ins, 1, BTrecord, "0", Colectie_nume_atr, Colectie_valori)

                                                End If 'ASTA E DE LA If Not Nume_block = "" 

                                            End If ' ASTA E DE LA IF ISDBNULL

                                        Next ' ASTA E DE LA For k = 0 To Data_table_fitting.Rows.Count - 1
                                    End If ' asta e de la  If Data_table_fitting.Rows.Count > 0
                                End If

                                If IsNothing(Data_table_cl_crossing) = False Then
                                    If Data_table_cl_crossing.Rows.Count > 0 Then
                                        For k = 0 To Data_table_cl_crossing.Rows.Count - 1


                                            If IsDBNull(Data_table_cl_crossing.Rows(k).Item("STATION")) = False And IsDBNull(Data_table_cl_crossing.Rows(k).Item("X")) = False And IsDBNull(Data_table_cl_crossing.Rows(k).Item("Y")) = False Then
                                                Dim CL_Chainage As Double = Data_table_cl_crossing.Rows(k).Item("STATION")

                                                Dim Pt_ins As New Point3d(Data_table_cl_crossing.Rows(k).Item("X"), Data_table_cl_crossing.Rows(k).Item("Y"), 0)

                                                Dim CL_Description As String = ""
                                                If IsDBNull(Data_table_cl_crossing.Rows(k).Item("DESCRIPTION")) = False Then
                                                    CL_Description = Data_table_cl_crossing.Rows(k).Item("DESCRIPTION")
                                                End If



                                                Dim Nume_block As String = ""
                                                If Not ComboBox_blocks_CL_crossing.Text = "" Then
                                                    Nume_block = ComboBox_blocks_CL_crossing.Text
                                                End If

                                                If Not Nume_block = "" Then
                                                    Dim Colectie_nume_atr As New Specialized.StringCollection
                                                    Dim Colectie_valori As New Specialized.StringCollection

                                                    If Not ComboBox_STA_att_CL_crossing.Text = "" Then
                                                        Colectie_nume_atr.Add(ComboBox_STA_att_CL_crossing.Text)
                                                        If Is_Canada = False Then
                                                            Colectie_valori.Add(Get_chainage_feet_from_double(CL_Chainage + Get_equation_value(CL_Chainage), Round1))
                                                        Else
                                                            Colectie_valori.Add(Get_chainage_from_double(CL_Chainage + Get_equation_value(CL_Chainage), Round1))
                                                        End If
                                                    End If

                                                    If Not ComboBox_DESCRIPTION_att_CL_crossing.Text = "" And Not CL_Description = "" Then
                                                        Colectie_nume_atr.Add(ComboBox_DESCRIPTION_att_CL_crossing.Text)
                                                        Colectie_valori.Add(CL_Description)
                                                    End If



                                                    Dim BR1 As BlockReference = InsertBlock_with_multiple_atributes("", Nume_block, Pt_ins, 1, BTrecord, "0", Colectie_nume_atr, Colectie_valori)

                                                End If 'ASTA E DE LA If Not Nume_block = "" 

                                            End If ' ASTA E DE LA IF ISDBNULL

                                        Next ' ASTA E DE LA For k = 0 To Data_table_cl_crossing.Rows.Count - 1
                                    End If ' asta e de la  If Data_table_cl_crossing.Rows.Count > 0
                                End If

                                If IsNothing(Data_table_cad_weld) = False Then
                                    If Data_table_cad_weld.Rows.Count > 0 Then
                                        For k = 0 To Data_table_cad_weld.Rows.Count - 1


                                            If IsDBNull(Data_table_cad_weld.Rows(k).Item("STATION")) = False And IsDBNull(Data_table_cad_weld.Rows(k).Item("X")) = False And IsDBNull(Data_table_cad_weld.Rows(k).Item("Y")) = False Then
                                                Dim CW_station As Double = Data_table_cad_weld.Rows(k).Item("STATION")

                                                Dim Pt_ins As New Point3d(Data_table_cad_weld.Rows(k).Item("X"), Data_table_cad_weld.Rows(k).Item("Y"), 0)

                                                Dim Nume_block As String = ""
                                                If Not ComboBox_blocks_cad_weld.Text = "" Then
                                                    Nume_block = ComboBox_blocks_cad_weld.Text
                                                End If

                                                If Not Nume_block = "" Then
                                                    Dim Colectie_nume_atr As New Specialized.StringCollection
                                                    Dim Colectie_valori As New Specialized.StringCollection

                                                    If Not ComboBox_STA_att_cad_weld.Text = "" Then
                                                        Colectie_nume_atr.Add(ComboBox_STA_att_cad_weld.Text)
                                                        If Is_Canada = False Then
                                                            Colectie_valori.Add(Get_chainage_feet_from_double(CW_station + Get_equation_value(CW_station), Round1))
                                                        Else
                                                            Colectie_valori.Add(Get_chainage_from_double(CW_station + Get_equation_value(CW_station), Round1))
                                                        End If
                                                    End If



                                                    Dim BR1 As BlockReference = InsertBlock_with_multiple_atributes("", Nume_block, Pt_ins, 1, BTrecord, "0", Colectie_nume_atr, Colectie_valori)

                                                End If 'ASTA E DE LA If Not Nume_block = "" 

                                            End If ' ASTA E DE LA IF ISDBNULL

                                        Next ' ASTA E DE LA For k = 0 To Data_table_cad_weld.Rows.Count - 1
                                    End If ' asta e de la  If Data_table_cad_weld.Rows.Count > 0
                                End If


                                If IsNothing(Data_table_river_weights) = False Then
                                    If Data_table_river_weights.Rows.Count > 0 Then
                                        For k = 0 To Data_table_river_weights.Rows.Count - 1


                                            If IsDBNull(Data_table_river_weights.Rows(k).Item("STATION")) = False And IsDBNull(Data_table_river_weights.Rows(k).Item("X")) = False And IsDBNull(Data_table_river_weights.Rows(k).Item("Y")) = False Then
                                                Dim RW_station As Double = Data_table_river_weights.Rows(k).Item("STATION")

                                                Dim Pt_ins As New Point3d(Data_table_river_weights.Rows(k).Item("X"), Data_table_river_weights.Rows(k).Item("Y"), 0)


                                                Dim Nume_block As String = ""
                                                If Not ComboBox_blocks_RIVER_WEIGHTS.Text = "" Then
                                                    Nume_block = ComboBox_blocks_RIVER_WEIGHTS.Text
                                                End If

                                                If Not Nume_block = "" Then
                                                    Dim Colectie_nume_atr As New Specialized.StringCollection
                                                    Dim Colectie_valori As New Specialized.StringCollection

                                                    If Not ComboBox_STA_att_river_weights.Text = "" Then
                                                        Colectie_nume_atr.Add(ComboBox_STA_att_river_weights.Text)
                                                        If Is_Canada = False Then
                                                            Colectie_valori.Add(Get_chainage_feet_from_double(RW_station + Get_equation_value(RW_station), Round1))
                                                        Else
                                                            Colectie_valori.Add(Get_chainage_from_double(RW_station + Get_equation_value(RW_station), Round1))
                                                        End If
                                                    End If



                                                    Dim BR1 As BlockReference = InsertBlock_with_multiple_atributes("", Nume_block, Pt_ins, 1, BTrecord, "0", Colectie_nume_atr, Colectie_valori)

                                                End If 'ASTA E DE LA If Not Nume_block = "" 

                                            End If ' ASTA E DE LA IF ISDBNULL

                                        Next ' ASTA E DE LA For k = 0 To Data_table_river_weights.Rows.Count - 1
                                    End If ' asta e de la  If Data_table_river_weights.Rows.Count > 0
                                End If

                            End If ' asta e de la If Not ComboBox_blocks_MAT.Text = "" 

                            Trans1.Commit()
                        End Using
                    End Using
                End If
            End If
            MsgBox("Done")
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_pick_POSITION_ZERO_Click(sender As Object, e As EventArgs) Handles Button_pick_POSITION_ZERO.Click
        If Freeze_operations = False Then
            Freeze_operations = True
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



                        Freeze_operations = False

                    End Using
                End Using

            Catch ex As Exception
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Bands inserted")
            Freeze_operations = False
        End If

    End Sub

    Private Sub Button_pick_viewport_corner_Click(sender As Object, e As EventArgs) Handles Button_pick_viewport_corner.Click
        If Freeze_operations = False Then
            Freeze_operations = True

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



                        Freeze_operations = False

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

    Private Sub Button_draw_viewport_Click(sender As Object, e As EventArgs) Handles Button_draw_Viewport.Click


        Try
            If IsNumeric(TextBox_page.Text) = False Then
                MsgBox("Please specify the page!")
                Exit Sub
            End If
            If IsNumeric(TextBox_viewport_SCALE.Text) = False Then
                MsgBox("Please specify the viewport scale!")
                Exit Sub
            End If
            If IsNumeric(TextBox_viewport_Height.Text) = False Then
                MsgBox("Please specify the viewport height!")
                Exit Sub
            End If
            If IsNumeric(TextBox_viewport_Width.Text) = False Then
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

            Dim Scale1 As Double = 1 / CDbl(TextBox_viewport_SCALE.Text)

            Dim Spacing1 As Double = CDbl(TextBox_BAND_SPACING.Text)

            Dim H1 As Double = CDbl(TextBox_viewport_Height.Text)
            Dim W1 As Double = CDbl(TextBox_viewport_Width.Text)

            Dim x_MS As Double = CDbl(TextBox_X_MS.Text)
            Dim y_MS As Double = CDbl(TextBox_Y_MS.Text)

            Dim x_pS As Double = CDbl(TextBox_X_PS.Text)
            Dim y_PS As Double = CDbl(TextBox_Y_PS.Text)

            Dim DeltaY As Double
            If IsNumeric(TextBox_shift_viewport_y.Text) = True Then
                DeltaY = CDbl(TextBox_shift_viewport_y.Text)
            End If

            Dim DeltaX As Double
            If IsNumeric(TextBox_shift_viewport_X.Text) = True Then
                DeltaX = CDbl(TextBox_shift_viewport_X.Text)
            End If

            If Scale1 <= 0 Or Page1 <= 0 Or Spacing1 <= 0 Or H1 <= 0 Or W1 <= 0 Then
                MsgBox("Negative values not allowed")
                Exit Sub
            End If

            If Freeze_operations = False Then
                Freeze_operations = True

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

                        Dim Point_target As New Point3d

                        Dim Line_len As Double = 0


                        Line_len = Abs(CDbl(TextBox_viewport_Height.Text))




                        If RadioButton_left_right_viewport.Checked = True Then
                            Point_target = New Point3d(x_MS + (W1 / 2) / Scale1 + DeltaX / Scale1, y_MS - Spacing1 * (Page1 - 1) + DeltaY / Scale1 + Line_len / 2, 0)
                        Else
                            Point_target = New Point3d(x_MS - (W1 / 2) / Scale1 + DeltaX / Scale1, y_MS - Spacing1 * (Page1 - 1) + DeltaY / Scale1 + Line_len / 2, 0)
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

            End If

        Catch ex As Exception
            MsgBox(ex.Message)
            Freeze_operations = False
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
        Freeze_operations = False

    End Sub

    Private Sub Button_draw_multiple_viewports_Click(sender As Object, e As EventArgs) Handles Button_draw_multiple_viewports.Click


        Try

            If IsNumeric(TextBox_viewport_SCALE.Text) = False Then
                MsgBox("Please specify the viewport scale!")
                Exit Sub
            End If
            If IsNumeric(TextBox_viewport_Height.Text) = False Then
                MsgBox("Please specify the viewport height!")
                Exit Sub
            End If
            If IsNumeric(TextBox_viewport_Width.Text) = False Then
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

            If IsNumeric(TextBox_Page_start.Text) = False Then
                MsgBox("Please specify the start page!")
                Exit Sub
            End If

            If IsNumeric(TextBox_page_end.Text) = False Then
                MsgBox("Please specify the end page!")
                Exit Sub
            End If


            If IsNumeric(TextBox_layout_start.Text) = False Then
                MsgBox("Please specify the layout_start page!")
                Exit Sub
            End If

            Dim Scale1 As Double = 1 / CDbl(TextBox_viewport_SCALE.Text)

            Dim Spacing1 As Double = CDbl(TextBox_BAND_SPACING.Text)

            Dim H1 As Double = CDbl(TextBox_viewport_Height.Text)
            Dim W1 As Double = CDbl(TextBox_viewport_Width.Text)

            Dim x_MS As Double = CDbl(TextBox_X_MS.Text)
            Dim y_MS As Double = CDbl(TextBox_Y_MS.Text)

            Dim x_pS As Double = CDbl(TextBox_X_PS.Text)
            Dim y_PS As Double = CDbl(TextBox_Y_PS.Text)

            Dim DeltaY As Double
            If IsNumeric(TextBox_shift_viewport_y.Text) = True Then
                DeltaY = CDbl(TextBox_shift_viewport_y.Text)
            End If

            Dim DeltaX As Double
            If IsNumeric(TextBox_shift_viewport_X.Text) = True Then
                DeltaX = CDbl(TextBox_shift_viewport_X.Text)
            End If

            If Scale1 <= 0 Or Spacing1 <= 0 Or H1 <= 0 Or W1 <= 0 Then
                MsgBox("Negative values not allowed")
                Freeze_operations = False
                Exit Sub
            End If




            If Freeze_operations = False Then
                Freeze_operations = True


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





                        Dim Page_start As Integer = CInt(TextBox_Page_start.Text)
                        Dim Page_end As Integer = CInt(TextBox_page_end.Text)
                        Dim Layout_start As Integer = CInt(TextBox_layout_start.Text)


                        Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current


                        Dim Layoutdict As DBDictionary

                        Layoutdict = Trans1.GetObject(ThisDrawing.Database.LayoutDictionaryId, OpenMode.ForRead)
                        Dim nr_layouts As Integer
                        nr_layouts = Layoutdict.Count


                        Dim INDEX_ORDER As String = "INDEX_ORDER"
                        Dim LAYOUT_NAME As String = "LAYOUT_NAME"

                        Dim Data_table As New System.Data.DataTable
                        Data_table.Columns.Add(INDEX_ORDER, GetType(Integer))
                        Data_table.Columns.Add(LAYOUT_NAME, GetType(String))
                        Dim Index_datatable As Integer = 0
                        For Each entry As DBDictionaryEntry In Layoutdict
                            Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForRead)
                            If Not Layout1.TabOrder = 0 Then
                                Data_table.Rows.Add()
                                Data_table.Rows(Index_datatable).Item(INDEX_ORDER) = Layout1.TabOrder
                                Data_table.Rows(Index_datatable).Item(LAYOUT_NAME) = Layout1.LayoutName
                                Index_datatable = Index_datatable + 1
                            End If
                        Next

                        Data_table = Sort_data_table(Data_table, INDEX_ORDER)

                        Add_to_clipboard_Data_table(Data_table)



                        Dim Tilemode1 As Integer = Application.GetSystemVariable("TILEMODE")
                        Dim CVport1 As Integer = Application.GetSystemVariable("CVPORT")

                        If Tilemode1 = 0 Then
                            If Not CVport1 = 1 Then
                                Editor1.SwitchToPaperSpace()
                            End If
                        Else
                            Application.SetSystemVariable("TILEMODE", 0)
                        End If

                        Dim Number_of_pages As Integer = Page_end - Page_start
                        Dim Band_index As Integer = Page_start

                        If Data_table.Rows.Count >= Number_of_pages + 1 Then

                            For i = Layout_start To Layout_start + Number_of_pages
                                Dim Nume_layout As String = ""
                                For j = 0 To Data_table.Rows.Count - 1
                                    If Data_table.Rows(j).Item(INDEX_ORDER) = i Then
                                        Nume_layout = Data_table.Rows(j).Item(LAYOUT_NAME)
                                        Exit For
                                    End If
                                Next
                                If Not Nume_layout = "" Then
                                    LayoutManager1.CurrentLayout = Nume_layout

                                    Dim BTrecordPS As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                    BTrecordPS = Trans1.GetObject(BlockTable_data1(BlockTableRecord.PaperSpace), OpenMode.ForWrite)

                                    Dim Point_target As New Point3d

                                    Dim Line_len As Double = 0
                                    'xxx()

                                    Line_len = Abs(CDbl(TextBox_viewport_Height.Text))




                                    If RadioButton_left_right_viewport.Checked = True Then
                                        Point_target = New Point3d(x_MS + (W1 / 2) / Scale1 + DeltaX / Scale1, y_MS - Spacing1 * (Band_index - 1) + DeltaY / Scale1 + Line_len / 2, 0)
                                    Else
                                        Point_target = New Point3d(x_MS - (W1 / 2) / Scale1 + DeltaX / Scale1, y_MS - Spacing1 * (Band_index - 1) + DeltaY / Scale1 + Line_len / 2, 0)
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
                                    Band_index = Band_index + 1
                                End If

                            Next
                        Else
                            MsgBox("There are " & Data_table.Rows.Count & " layouts and you specified " & Number_of_pages + 1 & "pages")

                        End If
                        ' Dim Layout2 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(Nume_nou), OpenMode.ForRead)



                        '


                        'BTrecordPS = Trans1.GetObject(BlockTable_data1(BlockTableRecord.PaperSpace), OpenMode.ForWrite)








                        Trans1.Commit()

                    End Using
                End Using
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Freeze_operations = False
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
        Freeze_operations = False

    End Sub

    Private Sub Button_asBuilt_Load_materials_ID_Click(sender As Object, e As EventArgs) Handles Button_asBuilt_Load_materials.Click
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

        If Freeze_operations = False Then
            Freeze_operations = True

            Try

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Dim Start1 As Integer = 0
                Dim End1 As Integer = 0
                If IsNumeric(TextBox_read_start_row.Text) = True Then
                    Start1 = CInt(TextBox_read_start_row.Text)
                End If
                If IsNumeric(TextBox_read_end_row.Text) = True Then
                    End1 = CInt(TextBox_read_end_row.Text)
                End If

                If End1 = 0 Then
                    MsgBox("End row specified incorrectly")
                    Freeze_operations = False
                    Exit Sub
                End If

                If Start1 = 0 Then
                    MsgBox("Start row specified incorrectly")
                    Freeze_operations = False
                    Exit Sub
                End If
                If End1 < Start1 Then
                    MsgBox("Start row bigger than end row")
                    Freeze_operations = False
                    Exit Sub
                End If
                Dim Column_Station As String = ""
                Column_Station = TextBox_ASBUILT_pipe_ID.Text.ToUpper

                Dim Column_mat_no As String = ""
                Column_mat_no = TextBox_ASBUILT_material_number.Text.ToUpper

                If Column_Station = "" Then
                    Freeze_operations = False
                    MsgBox("No station column specified")
                    Exit Sub
                End If

                If Column_mat_no = "" Then
                    Freeze_operations = False
                    MsgBox("No Pipe_ID column specified")
                    Exit Sub
                End If


                Data_table_materials = New System.Data.DataTable
                Data_table_materials.Columns.Add("STA1", GetType(Double))
                Data_table_materials.Columns.Add("STA2", GetType(Double))
                Data_table_materials.Columns.Add("MAT", GetType(String))

                Dim Index_dt As Double = 0
                Dim Material_previous As String = "a"
                Dim Station_prev As Double = 0
                For i = Start1 To End1
                    Dim Station1 As String = W1.Range(Column_Station & i).Value2
                    Dim Material_no As String = W1.Range(Column_mat_no & i).Value2



                    If IsNumeric(Station1) = True And Not Material_no = "" Then
                        If Round(CDbl(Station1), 0) = 6249 Then
                            Dim Debug As String = Station1
                        End If

                        If i = Start1 Then
                            Station_prev = CDbl(Station1)
                        End If

                        If Not Material_no = Material_previous Then
                            If Not i = Start1 Then
                                Data_table_materials.Rows(Index_dt - 1).Item("STA2") = Station_prev
                            End If


                            Data_table_materials.Rows.Add()
                            Data_table_materials.Rows(Index_dt).Item("STA1") = Station_prev
                            Data_table_materials.Rows(Index_dt).Item("MAT") = Material_no
                            Data_table_materials.Rows(Index_dt).Item("STA2") = CDbl(Station1)
                            Index_dt = Index_dt + 1

                            Material_previous = Material_no

                        End If



                        Station_prev = CDbl(Station1)

                        If i = End1 Then
                            Data_table_materials.Rows(Index_dt - 1).Item("STA2") = Station_prev
                        End If

                    End If


                Next

                'MsgBox(Data_table_Centerline.Rows.Count)
                Add_to_clipboard_Data_table(Data_table_materials)

            Catch ex As Exception
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "DONE - " & Data_table_materials.Rows.Count & " materials loaded")
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_read_bends_Click(sender As Object, e As EventArgs) Handles Button_read_bends_canada.Click
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
        If Freeze_operations = False Then
            Freeze_operations = True

            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
            W1 = Get_active_worksheet_from_Excel()
            Dim Start1 As Integer = 0
            Dim End1 As Integer = 0
            If IsNumeric(TextBox_tools_start_row.Text) = True Then
                Start1 = CInt(TextBox_tools_start_row.Text)
            End If
            If IsNumeric(TextBox_tools_end_row.Text) = True Then
                End1 = CInt(TextBox_tools_end_row.Text)
            End If

            If End1 = 0 Then
                MsgBox("End row specified incorrectly")
                Freeze_operations = False
                Exit Sub
            End If

            If Start1 = 0 Then
                MsgBox("Start row specified incorrectly")
                Freeze_operations = False
                Exit Sub
            End If
            If End1 < Start1 Then
                MsgBox("Start row bigger than end row")
                Freeze_operations = False
                Exit Sub
            End If
            Dim Column_chainage As String = ""
            Column_chainage = TextBox_Chainage_tools.Text.ToUpper

            Dim Column_bend As String = ""
            Column_bend = TextBox_bend_tools.Text.ToUpper


            Dim Column_weld As String = ""
            Column_weld = TextBox_weld_tools.Text.ToUpper

            If Column_chainage = "" Then
                Freeze_operations = False
                MsgBox("No chainage column specified")
                Exit Sub
            End If


            If Column_bend = "" Then
                Freeze_operations = False
                MsgBox("No bend column specified")
                Exit Sub
            End If

            If Column_weld = "" Then
                Freeze_operations = False
                MsgBox("No weld column specified")
                Exit Sub
            End If


            Dim WELD_BEFORE_MINUS1 As String = "WELD_BEFORE_MINUS_1"
            Dim WELD_AFTER_PLUS1 As String = "WELD_AFTER_PLUS_1"
            Dim WELD_BEFORE As String = "WELD_BEFORE"
            Dim WELD_AFTER As String = "WELD_AFTER"
            Dim WELD_DESCRIPTION_BEFORE As String = "WELD_DESCRIPTION_BEFORE"
            Dim WELD_DESCRIPTION_AFTER As String = "WELD_DESCRIPTION_AFTER"
            Dim WELD_DESCRIPTION_BEFORE_MINUS1 As String = "WELD_DESCRIPTION_BEFORE_MINUS_1"
            Dim WELD_DESCRIPTION_AFTER_PLUS1 As String = "WELD_DESCRIPTION_AFTER_PLUS_1"


            Data_table_bends = New System.Data.DataTable
            Data_table_bends.Columns.Add("CHAINAGE", GetType(Double))
            Data_table_bends.Columns.Add("DESCRIPTION", GetType(String))


            Data_table_bends.Columns.Add(WELD_BEFORE_MINUS1, GetType(Double))
            Data_table_bends.Columns.Add(WELD_BEFORE, GetType(Double))
            Data_table_bends.Columns.Add(WELD_AFTER, GetType(Double))
            Data_table_bends.Columns.Add(WELD_AFTER_PLUS1, GetType(Double))

            Data_table_bends.Columns.Add(WELD_DESCRIPTION_BEFORE_MINUS1, GetType(String))
            Data_table_bends.Columns.Add(WELD_DESCRIPTION_BEFORE, GetType(String))
            Data_table_bends.Columns.Add(WELD_DESCRIPTION_AFTER, GetType(String))
            Data_table_bends.Columns.Add(WELD_DESCRIPTION_AFTER_PLUS1, GetType(String))

            Dim Data_table_welds As New System.Data.DataTable
            Data_table_welds.Columns.Add("CHAINAGE", GetType(Double))
            Data_table_welds.Columns.Add("WELD", GetType(String))

            Dim Index_dt As Double = 0
            Dim Index_dtw As Double = 0

            Try
                For i = Start1 To End1
                    Dim Chainage1 As String = W1.Range(Column_chainage & i).Value2
                    Dim Bend1 As String = W1.Range(Column_bend & i).Value2
                    Dim Weld1 As String = W1.Range(Column_weld & i).Value2
                    If IsNothing(Chainage1) = False Then
                        If IsNumeric(Chainage1) = True Then
                            If Not Bend1 = "" Then
                                Data_table_bends.Rows.Add()
                                Data_table_bends.Rows(Index_dt).Item("CHAINAGE") = CDbl(Chainage1)
                                Data_table_bends.Rows(Index_dt).Item("DESCRIPTION") = Bend1
                                Index_dt = Index_dt + 1

                            Else
                                If Not Weld1 = "" Then

                                    Data_table_welds.Rows.Add()
                                    Data_table_welds.Rows(Index_dtw).Item("CHAINAGE") = CDbl(Chainage1)
                                    Data_table_welds.Rows(Index_dtw).Item("WELD") = Weld1
                                    Index_dtw = Index_dtw + 1
                                End If
                            End If



                        End If
                    End If


                Next

                Data_table_welds = Sort_data_table(Data_table_welds, "CHAINAGE")

                If Data_table_bends.Rows.Count > 0 And Data_table_welds.Rows.Count > 0 Then
                    For i = 0 To Data_table_bends.Rows.Count - 1
                        Dim Chainage1 As Double = Data_table_bends.Rows(i).Item("CHAINAGE")
                        For j = 0 To Data_table_welds.Rows.Count - 2
                            Dim ChainWeld1 As Double = Data_table_welds.Rows(j).Item("CHAINAGE")
                            Dim ChainWeld2 As Double = Data_table_welds.Rows(j + 1).Item("CHAINAGE")
                            Dim Descr1 As String = Data_table_welds.Rows(j).Item("WELD")
                            Dim Descr2 As String = Data_table_welds.Rows(j + 1).Item("WELD")

                            If Chainage1 > ChainWeld1 And Chainage1 < ChainWeld2 Then

                                Data_table_bends.Rows(i).Item(WELD_BEFORE) = ChainWeld1
                                Data_table_bends.Rows(i).Item(WELD_DESCRIPTION_BEFORE) = Descr1

                                Data_table_bends.Rows(i).Item(WELD_AFTER) = ChainWeld2
                                Data_table_bends.Rows(i).Item(WELD_DESCRIPTION_AFTER) = Descr2

                                If j > 1 Then
                                    For k = 1 To j
                                        Dim ChainWeld1_1 As Double = Data_table_welds.Rows(j - k).Item("CHAINAGE")
                                        Dim Descr1_1 As String = Data_table_welds.Rows(j - k).Item("WELD")
                                        If Not Descr1_1.ToUpper = "SW" Then
                                            Data_table_bends.Rows(i).Item(WELD_BEFORE_MINUS1) = ChainWeld1_1
                                            Data_table_bends.Rows(i).Item(WELD_DESCRIPTION_BEFORE_MINUS1) = Descr1_1
                                            Exit For
                                        End If
                                    Next
                                End If

                                If j + 1 < Data_table_welds.Rows.Count Then
                                    For k = j + 1 To Data_table_welds.Rows.Count - 1
                                        Dim ChainWeld1_1 As Double = Data_table_welds.Rows(k).Item("CHAINAGE")
                                        Dim Descr1_1 As String = Data_table_welds.Rows(k).Item("WELD")
                                        If Not Descr1_1.ToUpper = "SW" Then
                                            Data_table_bends.Rows(i).Item(WELD_AFTER_PLUS1) = ChainWeld1_1
                                            Data_table_bends.Rows(i).Item(WELD_DESCRIPTION_AFTER_PLUS1) = Descr1_1
                                            Exit For
                                        End If
                                    Next
                                End If


                                Exit For
                            End If


                        Next
                    Next

                    Data_table_bends = Sort_data_table(Data_table_bends, "CHAINAGE")
                    Add_to_clipboard_Data_table(Data_table_bends)

                End If

            Catch ex As System.SystemException
                MsgBox(ex.Message)

            End Try


            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "DONE")
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_load_materials_from_excel_as_built_canada_Click(sender As Object, e As EventArgs) Handles Button_load_materials_from_excel_as_built_canada.Click


        If Freeze_operations = False Then
            Freeze_operations = True

            Try

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Dim Start1 As Integer = 0
                Dim End1 As Integer = 0
                If IsNumeric(TextBox_ASBUILT_canada_row_start.Text) = True Then
                    Start1 = CInt(TextBox_ASBUILT_canada_row_start.Text)
                End If
                If IsNumeric(TextBox_ASBUILT_canada_row_end.Text) = True Then
                    End1 = CInt(TextBox_ASBUILT_canada_row_end.Text)
                End If

                If End1 = 0 Then
                    MsgBox("End row specified incorrectly")
                    Freeze_operations = False
                    Exit Sub
                End If

                If Start1 = 0 Then
                    MsgBox("Start row specified incorrectly")
                    Freeze_operations = False
                    Exit Sub
                End If
                If End1 < Start1 Then
                    MsgBox("Start row bigger than end row")
                    Freeze_operations = False
                    Exit Sub
                End If
                Dim Column_chainage As String = ""
                Column_chainage = TextBox_chainage_mat_as_built_canada.Text.ToUpper

                Dim Column_mat As String = ""
                Column_mat = TextBox_material_as_built_canada.Text.ToUpper

                If Column_chainage = "" Or Column_mat = "" Then
                    Freeze_operations = False
                    Exit Sub
                End If

                Dim Data_table_materials1 As New System.Data.DataTable
                Data_table_materials1.Columns.Add("STA1", GetType(Double))
                Data_table_materials1.Columns.Add("MAT", GetType(String))

                Data_table_materials = New System.Data.DataTable
                Data_table_materials.Columns.Add("STA1", GetType(Double))
                Data_table_materials.Columns.Add("STA2", GetType(Double))
                Data_table_materials.Columns.Add("MAT", GetType(String))
                Data_table_materials.Columns.Add("CANADA", GetType(String))


                Dim Index_data_table As Double = 0



                For i = Start1 To End1
                    Dim Station_string1 As String = W1.Range(Column_chainage & i).Value2
                    Dim Material As String = W1.Range(Column_mat & i).Value2
                    If IsNumeric(Station_string1) = True And Not Material = "" Then
                        Data_table_materials1.Rows.Add()
                        Data_table_materials1.Rows(Index_data_table).Item("STA1") = CDbl(Station_string1)
                        Data_table_materials1.Rows(Index_data_table).Item("MAT") = Material
                        Index_data_table = Index_data_table + 1
                    End If

                Next


                Data_table_materials1 = Sort_data_table(Data_table_materials1, "STA1")

                Index_data_table = 0

                Dim Mat_previous As String = ""

                If Data_table_materials1.Rows.Count > 1 Then
                    Dim Material0 As String = Data_table_materials1.Rows(0).Item("MAT")
                    Data_table_materials.Rows.Add()
                    Data_table_materials.Rows(Index_data_table).Item("STA1") = Round(Data_table_materials1.Rows(0).Item("STA1"), 2)
                    Data_table_materials.Rows(Index_data_table).Item("MAT") = Material0
                    Index_data_table = Index_data_table + 1
                    Mat_previous = Material0

                    For i = 1 To Data_table_materials1.Rows.Count - 1
                        Dim Material As String = Data_table_materials1.Rows(i).Item("MAT")
                        Dim Chainage1 As Double = Round(Data_table_materials1.Rows(i).Item("STA1"), 2)

                        If Not Mat_previous = Material Then
                            Data_table_materials.Rows.Add()
                            Data_table_materials.Rows(Index_data_table - 1).Item("STA2") = Chainage1
                            Data_table_materials.Rows(Index_data_table).Item("STA1") = Chainage1
                            Data_table_materials.Rows(Index_data_table).Item("MAT") = Material
                            Index_data_table = Index_data_table + 1
                            Mat_previous = Material
                        Else
                            If i = Data_table_materials1.Rows.Count - 1 Then
                                Data_table_materials.Rows(Index_data_table - 1).Item("STA2") = Chainage1
                            End If
                        End If
                    Next

                End If

                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Material table has " & Data_table_materials.Rows.Count & " rows")

                'MsgBox(Data_table_Centerline.Rows.Count)
                Add_to_clipboard_Data_table(Data_table_materials)


            Catch ex As Exception
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try

            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_LOAD_AS_BUILT_CANADA_ELBOW_Click(sender As Object, e As EventArgs) Handles Button_LOAD_AS_BUILT_CANADA_ELBOW.Click

        If Freeze_operations = False Then
            Freeze_operations = True

            Try

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Dim Start1 As Integer = 0
                Dim End1 As Integer = 0
                If IsNumeric(TextBox_ASBUILT_canada_row_start.Text) = True Then
                    Start1 = CInt(TextBox_ASBUILT_canada_row_start.Text)
                End If
                If IsNumeric(TextBox_ASBUILT_canada_row_end.Text) = True Then
                    End1 = CInt(TextBox_ASBUILT_canada_row_end.Text)
                End If

                If End1 = 0 Then
                    MsgBox("End row specified incorrectly")
                    Freeze_operations = False
                    Exit Sub
                End If

                If Start1 = 0 Then
                    MsgBox("Start row specified incorrectly")
                    Freeze_operations = False
                    Exit Sub
                End If
                If End1 < Start1 Then
                    MsgBox("Start row bigger than end row")
                    Freeze_operations = False
                    Exit Sub
                End If
                Dim Column_sta1 As String = ""
                Column_sta1 = TextBox_St_Start_ellb_can.Text.ToUpper
                Dim Column_sta2 As String = ""
                Column_sta2 = TextBox_St_End_elb_can.Text.ToUpper

                Dim Column_sta As String = ""
                Column_sta = TextBox_St_middle_elb_can.Text.ToUpper


                Dim Column_descr1 As String = ""
                Column_descr1 = TextBox_Descr_elb_can.Text.ToUpper
                Dim Column_dwg_ID As String = ""
                Column_dwg_ID = TextBox_dwg_id_elb_can.Text.ToUpper

                If Column_sta1 = "" Or Column_sta2 = "" Then
                    Freeze_operations = False
                    Exit Sub
                End If





                Data_table_elbows = New System.Data.DataTable
                Data_table_elbows.Columns.Add("STA1", GetType(Double))
                Data_table_elbows.Columns.Add("STA2", GetType(Double))
                Data_table_elbows.Columns.Add("STA", GetType(Double))
                Data_table_elbows.Columns.Add("DESCR", GetType(String))
                Data_table_elbows.Columns.Add("ID", GetType(String))
                Data_table_elbows.Columns.Add("CANADA", GetType(String))


                Dim Index_data_table As Double



                For i = Start1 To End1
                    Dim Station_string As String = W1.Range(Column_sta & i).Value2
                    Dim Station_string1 As String = W1.Range(Column_sta1 & i).Value2
                    Dim Station_string2 As String = W1.Range(Column_sta2 & i).Value2
                    Dim Description1 As String = W1.Range(Column_descr1 & i).Value2
                    Dim DWG_ID As String = W1.Range(Column_dwg_ID & i).Value2

                    If IsNumeric(Station_string1) = True And IsNumeric(Station_string2) = True Then
                        Data_table_elbows.Rows.Add()
                        Data_table_elbows.Rows(Index_data_table).Item("STA1") = CDbl(Station_string1)
                        Data_table_elbows.Rows(Index_data_table).Item("STA2") = CDbl(Station_string2)
                        If IsNumeric(Station_string) = True Then
                            Data_table_elbows.Rows(Index_data_table).Item("STA") = CDbl(Station_string)
                        End If

                        If Not Description1 = "" Then
                            Data_table_elbows.Rows(Index_data_table).Item("DESCR") = Description1
                        End If
                        If Not DWG_ID = "" Then
                            Data_table_elbows.Rows(Index_data_table).Item("ID") = DWG_ID
                        End If
                        Index_data_table = Index_data_table + 1
                    End If



                Next


                Data_table_elbows = Sort_data_table(Data_table_elbows, "STA1")


                Add_to_clipboard_Data_table(Data_table_elbows)
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & Data_table_elbows.Rows.Count & " elbows loaded")

            Catch ex As Exception
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try

            Freeze_operations = False
        End If
    End Sub

    Private Sub RadioButton_ETC_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton_ETC.CheckedChanged, RadioButton_thornbury.CheckedChanged, RadioButton_asbuilt_eng_Etc.CheckedChanged
        If RadioButton_ETC.Checked = True Then
            TextBox_BAND_SPACING.Text = "1200"
            TextBox_viewport_SCALE.Text = "1"
            TextBox_viewport_Height.Text = "700"
            TextBox_viewport_Width.Text = "6751.3049"
            TextBox_X_MS.Text = "50000"
            TextBox_Y_MS.Text = "-1000"
            TextBox_X_PS.Text = "-373.6951"
            TextBox_Y_PS.Text = "1071.6521"
            TextBox_shift_viewport_X.Text = "-137"
            TextBox_shift_viewport_y.Text = "-417.6491"
            RadioButton_left_right.Checked = True
            RadioButton_left_right_viewport.Checked = True
            Stretch_scale_factor = 1

            Label_scale_1_to.Text = "Scale 1" & Chr(34) & " to"
            Label_inches.Text = "'"
            TextBox_CSF.Text = "1"
        ElseIf RadioButton_thornbury.Checked = True Then
            TextBox_BAND_SPACING.Text = "1500"
            TextBox_viewport_SCALE.Text = "7.5"
            TextBox_viewport_Height.Text = "108.9"
            TextBox_viewport_Width.Text = "919.48"
            TextBox_X_MS.Text = "0"
            TextBox_Y_MS.Text = "-1000"
            TextBox_X_PS.Text = "65.53"
            TextBox_Y_PS.Text = "126.82"
            TextBox_shift_viewport_X.Text = "0"
            TextBox_shift_viewport_y.Text = "-3"
            TextBox_CSF.Text = "0.999578"
            RadioButton_right_left.Checked = True
            RadioButton_Right_left_viewport.Checked = True
            Stretch_scale_factor = 0.97

            Label_scale_1_to.Text = "Scale 1000:"
            Label_inches.Text = "x 1000"
        ElseIf RadioButton_asbuilt_eng_Etc.Checked = True Then
            TextBox_BAND_SPACING.Text = "1500"
            TextBox_viewport_SCALE.Text = "1"
            TextBox_viewport_Height.Text = "700"
            TextBox_viewport_Width.Text = "5880"
            TextBox_X_MS.Text = "50000"
            TextBox_Y_MS.Text = "-1000"
            TextBox_X_PS.Text = "1150.8442"
            TextBox_Y_PS.Text = "1071.6521"
            TextBox_shift_viewport_X.Text = "0"
            TextBox_shift_viewport_y.Text = "-417.65"
            RadioButton_left_right.Checked = True
            RadioButton_left_right_viewport.Checked = True
            Stretch_scale_factor = 0.97
            TextBox_CSF.Text = "1"
            Label_scale_1_to.Text = "Scale 1" & Chr(34) & " to"
            Label_inches.Text = "'"
        End If
    End Sub

    Private Sub Button_load_transitions_from_excel_as_built_canada_Click(sender As Object, e As EventArgs) Handles Button_load_transitions_from_excel_as_built_canada.Click

        If Freeze_operations = False Then
            Freeze_operations = True

            Try

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Dim Start1 As Integer = 0
                Dim End1 As Integer = 0
                If IsNumeric(TextBox_ASBUILT_canada_row_start.Text) = True Then
                    Start1 = CInt(TextBox_ASBUILT_canada_row_start.Text)
                End If
                If IsNumeric(TextBox_ASBUILT_canada_row_end.Text) = True Then
                    End1 = CInt(TextBox_ASBUILT_canada_row_end.Text)
                End If

                If End1 = 0 Then
                    MsgBox("End row specified incorrectly")
                    Freeze_operations = False
                    Exit Sub
                End If

                If Start1 = 0 Then
                    MsgBox("Start row specified incorrectly")
                    Freeze_operations = False
                    Exit Sub
                End If
                If End1 < Start1 Then
                    MsgBox("Start row bigger than end row")
                    Freeze_operations = False
                    Exit Sub
                End If
                Dim Column_chainage As String = ""
                Column_chainage = TextBox_transition_chainage.Text.ToUpper

                Dim Column_description As String = ""
                Column_description = TextBox_transition_description.Text.ToUpper

                If Column_chainage = "" Then
                    Freeze_operations = False
                    Exit Sub
                End If

                Data_table_transitions = New System.Data.DataTable
                Data_table_transitions.Columns.Add("STA", GetType(Double))
                Data_table_transitions.Columns.Add("DESCRIPTION", GetType(String))

                Dim Index_data_table As Double = 0

                For i = Start1 To End1
                    Dim Station_string1 As String = W1.Range(Column_chainage & i).Value2
                    Dim Description As String = W1.Range(Column_description & i).Value2
                    If IsNothing(Station_string1) = False And IsNothing(Description) = False Then
                        If IsNumeric(Station_string1) = True And Description.ToUpper.Contains("TRANSITION") = True And Description.ToUpper.Contains("WELD") = True Then
                            Data_table_transitions.Rows.Add()
                            Data_table_transitions.Rows(Index_data_table).Item("STA") = CDbl(Station_string1)
                            Data_table_transitions.Rows(Index_data_table).Item("DESCRIPTION") = Description
                            Index_data_table = Index_data_table + 1
                        End If
                    End If

                Next

                Data_table_transitions = Sort_data_table(Data_table_transitions, "STA")

                'MsgBox(Data_table_Centerline.Rows.Count)
                Add_to_clipboard_Data_table(Data_table_transitions)
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & Data_table_transitions.Rows.Count & " weld transitions loaded")


            Catch ex As Exception
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try

            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_read_excel_CL_Click(sender As Object, e As EventArgs) Handles Button_read_excel_CL_canada.Click
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
        If Freeze_operations = False Then
            Freeze_operations = True

            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
            W1 = Get_active_worksheet_from_Excel()
            Dim Start1 As Integer = 0
            Dim End1 As Integer = 0
            If IsNumeric(TextBox_tools_start_row.Text) = True Then
                Start1 = CInt(TextBox_tools_start_row.Text)
            End If
            If IsNumeric(TextBox_tools_end_row.Text) = True Then
                End1 = CInt(TextBox_tools_end_row.Text)
            End If

            If End1 = 0 Then
                MsgBox("End row specified incorrectly")
                Freeze_operations = False
                Exit Sub
            End If

            If Start1 = 0 Then
                MsgBox("Start row specified incorrectly")
                Freeze_operations = False
                Exit Sub
            End If
            If End1 < Start1 Then
                MsgBox("Start row bigger than end row")
                Freeze_operations = False
                Exit Sub
            End If
            Dim Column_chainage As String = ""
            Column_chainage = TextBox_chainage_cl_tools.Text.ToUpper

            Dim Column_description As String = ""
            Column_description = TextBox_description_CL_tools.Text.ToUpper


            If Column_chainage = "" Then
                Freeze_operations = False
                MsgBox("No chainage column specified")
                Exit Sub
            End If


            If Column_description = "" Then
                Freeze_operations = False
                MsgBox("No description column specified")
                Exit Sub
            End If



            Dim Data_table_crossing_CL As New System.Data.DataTable
            Data_table_crossing_CL.Columns.Add("CHAINAGE", GetType(Double))
            Data_table_crossing_CL.Columns.Add("DESCRIPTION", GetType(String))

            Dim Index_dt As Double = 0


            Try
                For i = Start1 To End1
                    Dim Chainage1 As String = W1.Range(Column_chainage & i).Value2
                    Dim Description1 As String = W1.Range(Column_description & i).Value2

                    If IsNothing(Chainage1) = False And IsNothing(Description1) = False Then
                        If IsNumeric(Chainage1) = True Then
                            If Description1.ToUpper.Contains("CL") = True Then
                                Data_table_crossing_CL.Rows.Add()
                                Data_table_crossing_CL.Rows(Index_dt).Item("CHAINAGE") = CDbl(Chainage1)
                                Data_table_crossing_CL.Rows(Index_dt).Item("DESCRIPTION") = Description1
                                Index_dt = Index_dt + 1

                            End If



                        End If
                    End If


                Next

                Data_table_crossing_CL = Sort_data_table(Data_table_crossing_CL, "CHAINAGE")

                Add_to_clipboard_Data_table(Data_table_crossing_CL)

                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & Data_table_crossing_CL.Rows.Count & " descriptions containing CL copied to clipboard")

            Catch ex As System.SystemException
                MsgBox(ex.Message)

            End Try


            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "DONE")
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_read_buoyancy_Click(sender As Object, e As EventArgs) Handles Button_read_buoyancy_canada.Click

        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
        If Freeze_operations = False Then
            Freeze_operations = True

            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
            W1 = Get_active_worksheet_from_Excel()
            Dim Start1 As Integer = 0
            Dim End1 As Integer = 0
            If IsNumeric(TextBox_tools_start_row.Text) = True Then
                Start1 = CInt(TextBox_tools_start_row.Text)
            End If
            If IsNumeric(TextBox_tools_end_row.Text) = True Then
                End1 = CInt(TextBox_tools_end_row.Text)
            End If

            If End1 = 0 Then
                MsgBox("End row specified incorrectly")
                Freeze_operations = False
                Exit Sub
            End If

            If Start1 = 0 Then
                MsgBox("Start row specified incorrectly")
                Freeze_operations = False
                Exit Sub
            End If
            If End1 < Start1 Then
                MsgBox("Start row bigger than end row")
                Freeze_operations = False
                Exit Sub
            End If

            Dim Column_chainage As String = ""
            Column_chainage = TextBox_bu_chainage.Text.ToUpper

            Dim Column_Descr As String = ""
            Column_Descr = TextBox_bu_descr.Text.ToUpper


            Dim Column_count As String = ""
            Column_count = TextBox_bu_count.Text.ToUpper

            Dim Column_spacing As String = ""
            Column_spacing = TextBox_bu_spacing.Text.ToUpper

            If Column_chainage = "" Then
                Freeze_operations = False
                MsgBox("No chainage column specified")
                Exit Sub
            End If


            If Column_Descr = "" Then
                Freeze_operations = False
                MsgBox("No description column specified")
                Exit Sub
            End If

            If Column_count = "" Then
                Freeze_operations = False
                MsgBox("No count column specified")
                Exit Sub
            End If

            If Column_spacing = "" Then
                Freeze_operations = False
                MsgBox("No spacing column specified")
                Exit Sub
            End If


            Data_table_buoyancy_canada_without_matchlines = New System.Data.DataTable
            Data_table_buoyancy_canada_without_matchlines.Columns.Add("CHAINAGE_START", GetType(Double))
            Data_table_buoyancy_canada_without_matchlines.Columns.Add("CHAINAGE_END", GetType(Double))
            Data_table_buoyancy_canada_without_matchlines.Columns.Add("COUNT", GetType(Integer))
            Data_table_buoyancy_canada_without_matchlines.Columns.Add("SPACING", GetType(Double))
            Data_table_buoyancy_canada_without_matchlines.Columns.Add("DESCRIPTION", GetType(String))

            Dim Index_dt As Double = 0

            Dim Interval_start As Boolean = True

            Try
                For i = Start1 To End1
                    Dim Chainage1 As String = W1.Range(Column_chainage & i).Value2
                    Dim Descr1 As String = ""
                    Try
                        Descr1 = W1.Range(Column_Descr & i).Value2
                    Catch ex As System.SystemException

                    End Try

                    Dim Count1 As String = ""
                    Try
                        Count1 = W1.Range(Column_count & i).Value2
                    Catch ex As System.SystemException

                    End Try
                    Dim Spacing1 As String = ""

                    Try
                        Spacing1 = W1.Range(Column_spacing & i).Value2
                    Catch ex As System.SystemException

                    End Try


                    If IsNothing(Chainage1) = False Then
                        If IsNumeric(Chainage1) = True Then




                            If Interval_start = True Then
                                Data_table_buoyancy_canada_without_matchlines.Rows.Add()
                                Data_table_buoyancy_canada_without_matchlines.Rows(Index_dt).Item("CHAINAGE_START") = CDbl(Chainage1)


                                If IsNumeric(Count1) = True Then
                                    Data_table_buoyancy_canada_without_matchlines.Rows(Index_dt).Item("COUNT") = CInt(Count1)
                                End If



                                If IsNumeric(Spacing1) = True Then
                                    Data_table_buoyancy_canada_without_matchlines.Rows(Index_dt).Item("SPACING") = CDbl(Spacing1)
                                End If



                                If Not Descr1 = "" Then
                                    Data_table_buoyancy_canada_without_matchlines.Rows(Index_dt).Item("DESCRIPTION") = Descr1
                                End If


                                If IsNumeric(Count1) = True Then
                                    If CInt(Count1) = 1 Then
                                        Data_table_buoyancy_canada_without_matchlines.Rows(Index_dt).Item("SPACING") = 0
                                        Data_table_buoyancy_canada_without_matchlines.Rows(Index_dt).Item("CHAINAGE_END") = CDbl(Chainage1)
                                        Index_dt = Index_dt + 1
                                    Else
                                        Interval_start = False
                                        GoTo 123
                                    End If
                                Else
                                    Data_table_buoyancy_canada_without_matchlines.Rows(Index_dt).Item("COUNT") = 1
                                    Data_table_buoyancy_canada_without_matchlines.Rows(Index_dt).Item("SPACING") = 0
                                    Data_table_buoyancy_canada_without_matchlines.Rows(Index_dt).Item("CHAINAGE_END") = CDbl(Chainage1)
                                    Index_dt = Index_dt + 1
                                End If


                            End If

                            If Interval_start = False Then
                                Data_table_buoyancy_canada_without_matchlines.Rows(Index_dt).Item("CHAINAGE_END") = CDbl(Chainage1)
                                Index_dt = Index_dt + 1
                                Interval_start = True
                            End If

123:

                        End If
                    End If


                Next
                If IsNothing(Data_table_buoyancy_canada_without_matchlines) = False Then
                    Data_table_buoyancy_canada_without_matchlines = Sort_data_table(Data_table_buoyancy_canada_without_matchlines, "CHAINAGE_START")
                    Add_to_clipboard_Data_table(Data_table_buoyancy_canada_without_matchlines)

                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & Data_table_buoyancy_canada_without_matchlines.Rows.Count & " buoyancy items copied to clipboard")
                Else
                    MsgBox("no buoyancy data loaded")
                End If

            Catch ex As System.SystemException
                MsgBox(ex.Message)

            End Try


            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_load_Buoyancy_canada_Click(sender As Object, e As EventArgs) Handles Button_load_Buoyancy_canada.Click

        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
        If Freeze_operations = False Then
            Freeze_operations = True

            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
            W1 = Get_active_worksheet_from_Excel()
            Dim Start1 As Integer = 0
            Dim End1 As Integer = 0
            If IsNumeric(TextBox_ASBUILT_canada_row_start.Text) = True Then
                Start1 = CInt(TextBox_ASBUILT_canada_row_start.Text)
            End If
            If IsNumeric(TextBox_ASBUILT_canada_row_end.Text) = True Then
                End1 = CInt(TextBox_ASBUILT_canada_row_end.Text)
            End If
            If End1 = 0 Then
                MsgBox("End row specified incorrectly")
                Freeze_operations = False
                Exit Sub
            End If

            If Start1 = 0 Then
                MsgBox("Start row specified incorrectly")
                Freeze_operations = False
                Exit Sub
            End If
            If End1 < Start1 Then
                MsgBox("Start row bigger than end row")
                Freeze_operations = False
                Exit Sub
            End If

            Dim Column_chainage_start As String = ""
            Column_chainage_start = TextBox_asbuilt_bu_Chain_start.Text.ToUpper
            Dim Column_chainage_end As String = ""
            Column_chainage_end = TextBox_asbuilt_bu_Chain_end.Text.ToUpper


            Dim Column_Descr As String = ""
            Column_Descr = TextBox_asbuilt_bu_description.Text.ToUpper


            Dim Column_count As String = ""
            Column_count = TextBox_asbuilt_bu_count.Text.ToUpper

            Dim Column_spacing As String = ""
            Column_spacing = TextBox_asbuilt_bu_spacing.Text.ToUpper

            If Column_chainage_start = "" Or Column_chainage_end = "" Then
                Freeze_operations = False
                MsgBox("No chainage column specified")
                Exit Sub
            End If


            If Column_Descr = "" Then
                Freeze_operations = False
                MsgBox("No description column specified")
                Exit Sub
            End If

            If Column_count = "" Then
                Freeze_operations = False
                MsgBox("No count column specified")
                Exit Sub
            End If

            If Column_spacing = "" Then
                Freeze_operations = False
                MsgBox("No spacing column specified")
                Exit Sub
            End If


            Data_table_buoyancy_canada_without_matchlines = New System.Data.DataTable
            Data_table_buoyancy_canada_without_matchlines.Columns.Add("CHAINAGE_START", GetType(Double))
            Data_table_buoyancy_canada_without_matchlines.Columns.Add("CHAINAGE_END", GetType(Double))
            Data_table_buoyancy_canada_without_matchlines.Columns.Add("COUNT", GetType(Integer))
            Data_table_buoyancy_canada_without_matchlines.Columns.Add("SPACING", GetType(Double))
            Data_table_buoyancy_canada_without_matchlines.Columns.Add("DESCRIPTION", GetType(String))

            Dim Index_dt As Double = 0


            Try
                For i = Start1 To End1
                    Dim Chainage1 As String = W1.Range(Column_chainage_start & i).Value2
                    Dim Chainage2 As String = W1.Range(Column_chainage_end & i).Value2
                    Dim Descr1 As String = ""
                    Try
                        Descr1 = W1.Range(Column_Descr & i).Value2
                    Catch ex As System.SystemException

                    End Try

                    Dim Count1 As String = ""
                    Try
                        Count1 = W1.Range(Column_count & i).Value2
                    Catch ex As System.SystemException

                    End Try
                    Dim Spacing1 As String = ""

                    Try
                        Spacing1 = W1.Range(Column_spacing & i).Value2
                    Catch ex As System.SystemException

                    End Try


                    If IsNothing(Chainage1) = False And IsNothing(Chainage2) = False Then
                        If IsNumeric(Chainage1) = True And IsNumeric(Chainage2) = True Then





                            Data_table_buoyancy_canada_without_matchlines.Rows.Add()
                            Data_table_buoyancy_canada_without_matchlines.Rows(Index_dt).Item("CHAINAGE_START") = CDbl(Chainage1)
                            Data_table_buoyancy_canada_without_matchlines.Rows(Index_dt).Item("CHAINAGE_END") = CDbl(Chainage2)

                            If IsNumeric(Count1) = True Then
                                Data_table_buoyancy_canada_without_matchlines.Rows(Index_dt).Item("COUNT") = CInt(Count1)
                            End If



                            If IsNumeric(Spacing1) = True Then
                                Data_table_buoyancy_canada_without_matchlines.Rows(Index_dt).Item("SPACING") = CDbl(Spacing1)
                            End If



                            If Not Descr1 = "" Then
                                Data_table_buoyancy_canada_without_matchlines.Rows(Index_dt).Item("DESCRIPTION") = Descr1
                            End If
                            Index_dt = Index_dt + 1
                        End If
                    End If


                Next

                If IsNothing(Data_table_buoyancy_canada_without_matchlines) = False Then
                    Data_table_buoyancy_canada_without_matchlines = Sort_data_table(Data_table_buoyancy_canada_without_matchlines, "CHAINAGE_START")
                    Add_to_clipboard_Data_table(Data_table_buoyancy_canada_without_matchlines)

                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & Data_table_buoyancy_canada_without_matchlines.Rows.Count & " buoyancy items copied to clipboard")
                Else
                    MsgBox("no buoyancy data loaded")
                End If
            Catch ex As System.SystemException
                MsgBox(ex.Message)

            End Try


            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_load_pipes_canada_Click(sender As Object, e As EventArgs) Handles Button_load_pipes_canada.Click

        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
        If Freeze_operations = False Then
            Freeze_operations = True

            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
            W1 = Get_active_worksheet_from_Excel()
            Dim Start1 As Integer = 0
            Dim End1 As Integer = 0
            If IsNumeric(TextBox_ASBUILT_canada_row_start.Text) = True Then
                Start1 = CInt(TextBox_ASBUILT_canada_row_start.Text)
            End If
            If IsNumeric(TextBox_ASBUILT_canada_row_end.Text) = True Then
                End1 = CInt(TextBox_ASBUILT_canada_row_end.Text)
            End If
            If End1 = 0 Then
                MsgBox("End row specified incorrectly")
                Freeze_operations = False
                Exit Sub
            End If

            If Start1 = 0 Then
                MsgBox("Start row specified incorrectly")
                Freeze_operations = False
                Exit Sub
            End If
            If End1 < Start1 Then
                MsgBox("Start row bigger than end row")
                Freeze_operations = False
                Exit Sub
            End If

            Dim Column_chainage As String = ""
            Column_chainage = TextBox_pipes_chainage.Text.ToUpper

            Dim Column_Descr1 As String = ""
            Column_Descr1 = TextBox_pipes_descr1.Text.ToUpper

            Dim Column_Descr2 As String = ""
            Column_Descr2 = TextBox_pipes_descr2.Text.ToUpper

            Dim Column_clearance As String = ""
            Column_clearance = TextBox_pipes_clearance.Text.ToUpper

            Dim Column_cover As String = ""
            Column_cover = TextBox_pipes_cover.Text.ToUpper



            If Column_chainage = "" Then
                Freeze_operations = False
                MsgBox("No chainage column specified")
                Exit Sub
            End If



            Dim Column_dwg_ID As String = ""
            Column_dwg_ID = TextBox_Pipes_DWG_ID.Text.ToUpper




            Data_table_pipes_Canada = New System.Data.DataTable
            Data_table_pipes_Canada.Columns.Add("CHAINAGE", GetType(Double))
            Data_table_pipes_Canada.Columns.Add("DESCRIPTION", GetType(String))
            Data_table_pipes_Canada.Columns.Add("CLEARANCE", GetType(String))
            Data_table_pipes_Canada.Columns.Add("COVER", GetType(String))
            Data_table_pipes_Canada.Columns.Add("DWG_ID", GetType(String))
            Data_table_pipes_Canada.Columns.Add("X", GetType(Double))
            Data_table_pipes_Canada.Columns.Add("Y", GetType(Double))

            Dim Index_dt As Double = 0


            Try
                For i = Start1 To End1


                    Dim Chainage1 As String = ""
                    Try
                        Chainage1 = W1.Range(Column_chainage & i).Value2
                    Catch ex As System.SystemException

                    End Try

                    Dim Descr1 As String = ""
                    Try
                        Descr1 = W1.Range(Column_Descr1 & i).Value2
                    Catch ex As System.SystemException

                    End Try

                    Dim Descr2 As String = ""
                    Try
                        Descr2 = W1.Range(Column_Descr2 & i).Value2
                    Catch ex As System.SystemException

                    End Try

                    Dim Clearance1 As String = ""
                    Try
                        Clearance1 = W1.Range(Column_clearance & i).Value2
                    Catch ex As System.SystemException

                    End Try
                    Dim Cover1 As String = ""

                    Try
                        Cover1 = W1.Range(Column_cover & i).Value2
                    Catch ex As System.SystemException

                    End Try

                    Dim DwgId1 As String = ""

                    Try
                        DwgId1 = W1.Range(Column_dwg_ID & i).Value2
                    Catch ex As System.SystemException

                    End Try


                    If IsNumeric(Chainage1) = True Then

                        Data_table_pipes_Canada.Rows.Add()
                        Data_table_pipes_Canada.Rows(Index_dt).Item("CHAINAGE") = CDbl(Chainage1)

                        Dim Descriptie As String = ""

                        If Not Descr1 = "" And Not Descr2 = "" Then
                            Descriptie = Descr1 & vbCrLf & Descr2
                        End If

                        If Not Descr1 = "" And Descr2 = "" Then
                            Descriptie = Descr1
                        End If

                        If Not Descr2 = "" And Descr1 = "" Then
                            Descriptie = Descr2
                        End If

                        If Not Descriptie = "" Then
                            Data_table_pipes_Canada.Rows(Index_dt).Item("DESCRIPTION") = Descriptie
                        End If

                        If Not Clearance1 = "" Then
                            Data_table_pipes_Canada.Rows(Index_dt).Item("CLEARANCE") = Clearance1
                        End If

                        If Not Cover1 = "" Then
                            Data_table_pipes_Canada.Rows(Index_dt).Item("COVER") = Cover1
                        End If

                        If Not DwgId1 = "" Then
                            Data_table_pipes_Canada.Rows(Index_dt).Item("DWG_ID") = DwgId1
                        End If

                        Index_dt = Index_dt + 1
                    End If

                Next

                If IsNothing(Data_table_pipes_Canada) = False Then
                    Data_table_pipes_Canada = Sort_data_table(Data_table_pipes_Canada, "CHAINAGE")
                    Add_to_clipboard_Data_table(Data_table_pipes_Canada)

                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & Data_table_pipes_Canada.Rows.Count & " pipe items copied to clipboard")
                Else
                    MsgBox("no pipe data loaded")
                End If
            Catch ex As System.SystemException
                MsgBox(ex.Message)

            End Try


            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_load_water_canada_Click(sender As Object, e As EventArgs) Handles Button_load_water_canada.Click


        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
        If Freeze_operations = False Then
            Freeze_operations = True

            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
            W1 = Get_active_worksheet_from_Excel()
            Dim Start1 As Integer = 0
            Dim End1 As Integer = 0
            If IsNumeric(TextBox_ASBUILT_canada_row_start.Text) = True Then
                Start1 = CInt(TextBox_ASBUILT_canada_row_start.Text)
            End If
            If IsNumeric(TextBox_ASBUILT_canada_row_end.Text) = True Then
                End1 = CInt(TextBox_ASBUILT_canada_row_end.Text)
            End If
            If End1 = 0 Then
                MsgBox("End row specified incorrectly")
                Freeze_operations = False
                Exit Sub
            End If

            If Start1 = 0 Then
                MsgBox("Start row specified incorrectly")
                Freeze_operations = False
                Exit Sub
            End If
            If End1 < Start1 Then
                MsgBox("Start row bigger than end row")
                Freeze_operations = False
                Exit Sub
            End If

            Dim Column_chainage As String = ""
            Column_chainage = TextBox_water_chainage.Text.ToUpper

            Dim Column_Descr1 As String = ""
            Column_Descr1 = TextBox_water_descr1.Text.ToUpper

            Dim Column_Descr2 As String = ""
            Column_Descr2 = TextBox_water_descr2.Text.ToUpper

            Dim Column_Descr3 As String = ""
            Column_Descr3 = TextBox_water_descr3.Text.ToUpper

            Dim Column_cover As String = ""
            Column_cover = TextBox_WATER_cover.Text.ToUpper



            If Column_chainage = "" Then
                Freeze_operations = False
                MsgBox("No chainage column specified")
                Exit Sub
            End If




            Dim Column_dwg_ID As String = ""
            Column_dwg_ID = TextBox_water_dwg_id.Text.ToUpper




            Data_table_water_canada = New System.Data.DataTable
            Data_table_water_canada.Columns.Add("CHAINAGE", GetType(Double))
            Data_table_water_canada.Columns.Add("DESCRIPTION", GetType(String))
            Data_table_water_canada.Columns.Add("COVER", GetType(String))
            Data_table_water_canada.Columns.Add("DWG_ID", GetType(String))
            Data_table_water_canada.Columns.Add("X", GetType(Double))
            Data_table_water_canada.Columns.Add("Y", GetType(Double))

            Dim Index_dt As Double = 0


            Try
                For i = Start1 To End1


                    Dim Chainage1 As String = ""
                    Try
                        Chainage1 = W1.Range(Column_chainage & i).Value2
                    Catch ex As System.SystemException

                    End Try

                    Dim Descr1 As String = ""
                    Try
                        Descr1 = W1.Range(Column_Descr1 & i).Value2
                    Catch ex As System.SystemException

                    End Try

                    Dim Descr2 As String = ""
                    Try
                        Descr2 = W1.Range(Column_Descr2 & i).Value2
                    Catch ex As System.SystemException

                    End Try


                    Dim Descr3 As String = ""
                    Try
                        Descr3 = W1.Range(Column_Descr3 & i).Value2
                    Catch ex As System.SystemException

                    End Try

                    Dim Cover1 As String = ""

                    Try
                        Cover1 = W1.Range(Column_cover & i).Value2
                    Catch ex As System.SystemException

                    End Try

                    Dim DwgId1 As String = ""

                    Try
                        DwgId1 = W1.Range(Column_dwg_ID & i).Value2
                    Catch ex As System.SystemException

                    End Try


                    If IsNumeric(Chainage1) = True Then

                        Data_table_water_canada.Rows.Add()
                        Data_table_water_canada.Rows(Index_dt).Item("CHAINAGE") = CDbl(Chainage1)

                        Dim Descriptie As String = ""

                        If Not Descr1 = "" And Not Descr2 = "" And Not Descr3 = "" Then
                            Descriptie = Descr1 & vbCrLf & Descr2 & vbCrLf & Descr3
                        End If

                        If Not Descr1 = "" And Descr2 = "" And Descr3 = "" Then
                            Descriptie = Descr1
                        End If

                        If Not Descr2 = "" And Descr1 = "" And Descr3 = "" Then
                            Descriptie = Descr2
                        End If

                        If Not Descr3 = "" And Descr1 = "" And Descr2 = "" Then
                            Descriptie = Descr3
                        End If

                        If Not Descr1 = "" And Not Descr2 = "" And Descr3 = "" Then
                            Descriptie = Descr1 & vbCrLf & Descr2
                        End If

                        If Not Descr1 = "" And Descr2 = "" And Not Descr3 = "" Then
                            Descriptie = Descr1 & vbCrLf & Descr3
                        End If

                        If Descr1 = "" And Not Descr2 = "" And Not Descr3 = "" Then
                            Descriptie = Descr2 & vbCrLf & Descr3
                        End If




                        If Not Descriptie = "" Then
                            Data_table_water_canada.Rows(Index_dt).Item("DESCRIPTION") = Descriptie
                        End If

                        If Not Cover1 = "" Then
                            Data_table_water_canada.Rows(Index_dt).Item("COVER") = Cover1
                        End If

                        If Not DwgId1 = "" Then
                            Data_table_water_canada.Rows(Index_dt).Item("DWG_ID") = DwgId1
                        End If

                        Index_dt = Index_dt + 1
                    End If

                Next

                If IsNothing(Data_table_water_canada) = False Then
                    Data_table_water_canada = Sort_data_table(Data_table_water_canada, "CHAINAGE")
                    Add_to_clipboard_Data_table(Data_table_water_canada)

                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & Data_table_water_canada.Rows.Count & " water items copied to clipboard")
                Else
                    MsgBox("no water data loaded")
                End If
            Catch ex As System.SystemException
                MsgBox(ex.Message)

            End Try


            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_load_cathodic_Click(sender As Object, e As EventArgs) Handles Button_load_CP_Canada.Click


        If Freeze_operations = False Then
            Freeze_operations = True

            Try

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Dim Start1 As Integer = 0
                Dim End1 As Integer = 0
                If IsNumeric(TextBox_ASBUILT_canada_row_start.Text) = True Then
                    Start1 = CInt(TextBox_ASBUILT_canada_row_start.Text)
                End If
                If IsNumeric(TextBox_ASBUILT_canada_row_end.Text) = True Then
                    End1 = CInt(TextBox_ASBUILT_canada_row_end.Text)
                End If

                If End1 = 0 Then
                    MsgBox("End row specified incorrectly")
                    Freeze_operations = False
                    Exit Sub
                End If

                If Start1 = 0 Then
                    MsgBox("Start row specified incorrectly")
                    Freeze_operations = False
                    Exit Sub
                End If
                If End1 < Start1 Then
                    MsgBox("Start row bigger than end row")
                    Freeze_operations = False
                    Exit Sub
                End If
                Dim Column_chainage As String = ""
                Column_chainage = TextBox_CP_Chainage.Text.ToUpper

                Dim Column_description As String = ""
                Column_description = TextBox_CP_descr.Text.ToUpper



                Data_table_cathodic_canada = New System.Data.DataTable
                Data_table_cathodic_canada.Columns.Add("CHAINAGE", GetType(Double))
                Data_table_cathodic_canada.Columns.Add("DESCRIPTION", GetType(String))
                Data_table_cathodic_canada.Columns.Add("X", GetType(Double))
                Data_table_cathodic_canada.Columns.Add("Y", GetType(Double))



                Dim Index_data_table As Double = 0

                For i = Start1 To End1
                    Dim Station_string1 As String = W1.Range(Column_chainage & i).Value2
                    Dim Description As String = W1.Range(Column_description & i).Value2
                    If IsNothing(Station_string1) = False Then
                        If IsNumeric(Station_string1) = True Then
                            Data_table_cathodic_canada.Rows.Add()
                            Data_table_cathodic_canada.Rows(Index_data_table).Item("CHAINAGE") = CDbl(Station_string1)
                            If IsNothing(Description) = False Then
                                If Not Description = "" Then
                                    Data_table_cathodic_canada.Rows(Index_data_table).Item("DESCRIPTION") = Description
                                End If

                            End If

                            Index_data_table = Index_data_table + 1
                        End If
                    End If

                Next

                Data_table_cathodic_canada = Sort_data_table(Data_table_cathodic_canada, "CHAINAGE")

                'MsgBox(Data_table_Centerline.Rows.Count)
                Add_to_clipboard_Data_table(Data_table_cathodic_canada)
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & Data_table_cathodic_canada.Rows.Count & " CP loaded")


            Catch ex As Exception
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try

            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_Read_pt_output_chainage_Click(sender As Object, e As EventArgs) Handles Button_Read_pt_output_chainage.Click

        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
        If Freeze_operations = False Then
            Freeze_operations = True

            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
            W1 = Get_active_worksheet_from_Excel()
            Dim Start1 As Integer = 0
            Dim End1 As Integer = 0
            If IsNumeric(TextBox_tools_start_row.Text) = True Then
                Start1 = CInt(TextBox_tools_start_row.Text)
            End If
            If IsNumeric(TextBox_tools_end_row.Text) = True Then
                End1 = CInt(TextBox_tools_end_row.Text)
            End If

            If End1 = 0 Then
                MsgBox("End row specified incorrectly")
                Freeze_operations = False
                Exit Sub
            End If

            If Start1 = 0 Then
                MsgBox("Start row specified incorrectly")
                Freeze_operations = False
                Exit Sub
            End If
            If End1 < Start1 Then
                MsgBox("Start row bigger than end row")
                Freeze_operations = False
                Exit Sub
            End If


            Dim Column_N As String = ""
            Column_N = TextBox_PT_N.Text.ToUpper


            Dim Column_E As String = ""
            Column_E = TextBox_PT_E.Text.ToUpper

            Dim Column_station As String = ""
            Column_station = TextBox_PT_CHAINAGE.Text.ToUpper


            If Column_N = "" Then
                Freeze_operations = False
                MsgBox("No N column specified")
                Exit Sub
            End If

            If Column_E = "" Then
                Freeze_operations = False
                MsgBox("No E column specified")
                Exit Sub
            End If

            If Column_station = "" Then
                Freeze_operations = False
                MsgBox("No station column specified")
                Exit Sub
            End If



            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
            Editor1 = ThisDrawing.Editor
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

            Try

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


                                Dim Poly3d As Polyline3d
                                Dim Poly2D As Polyline
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
                                    Poly2D = New Polyline
                                    Dim Index2d As Double = 0
                                    For Each ObjId As ObjectId In Poly3d
                                        Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, OpenMode.ForRead)
                                        Poly2D.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)
                                        Index2d = Index2d + 1
                                    Next
                                    Poly2D.Elevation = 0
                                End If

                                If TypeOf Ent1 Is Polyline Then
                                    Poly2D = Ent1
                                End If

                                If IsNothing(Poly2D) = False Then
                                    For i = Start1 To End1

                                        Dim X As Double
                                        Dim Xstring As String = W1.Range(Column_E & i).Value

                                        Dim Y As Double
                                        Dim Ystring As String = W1.Range(Column_N & i).Value

                                        If IsNumeric(Xstring) = True And IsNumeric(Ystring) = True Then
                                            X = CDbl(Xstring)
                                            Y = CDbl(Ystring)

                                            Dim Point_on_poly2d As New Point3d

                                            Point_on_poly2d = Poly2D.GetClosestPointTo(New Point3d(X, Y, 0), Vector3d.ZAxis, False)

                                            If IsNothing(Poly3d) = False Then
                                                Dim Dist1 As Double = Round(Poly3d.GetDistanceAtParameter(Poly2D.GetParameterAtPoint(Point_on_poly2d)), 3)

                                                W1.Range(Column_station & i).Value = Dist1

                                                Dim new_leader As New MLeader
                                                new_leader = Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly2d,
                                                                                                    "Station = " & Get_String_Rounded(Dist1, 3) _
                                                                                                   , 1, 0.2, 0.5, 5, 5)
                                            Else
                                                Dim Dist1 As Double = Round(Poly2D.GetDistAtPoint(Point_on_poly2d), 3)

                                                W1.Range(Column_station & i).Value = Dist1

                                                Dim new_leader As New MLeader
                                                new_leader = Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly2d,
                                                                                                    "Station = " & Get_String_Rounded(Dist1, 3) _
                                                                                                   , 1, 0.2, 0.5, 5, 5)
                                            End If





                                        End If

                                    Next
                                    MsgBox("Done")
                                Else
                                    MsgBox("No Polyline")
                                End If







                                Editor1.Regen()
                                Trans1.Commit()
                            End Using
                        End If
                    End If
                End Using










            Catch ex As System.SystemException
                MsgBox(ex.Message)

            End Try


            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_read_pipe_tally_Click(sender As Object, e As EventArgs) Handles Button_read_pipe_tally.Click
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

        If Freeze_operations = False Then
            Freeze_operations = True

            Try

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Dim Start1 As Integer = 0
                Dim End1 As Integer = 0
                If IsNumeric(TextBox_read_start_row.Text) = True Then
                    Start1 = CInt(TextBox_read_start_row.Text)
                End If
                If IsNumeric(TextBox_read_end_row.Text) = True Then
                    End1 = CInt(TextBox_read_end_row.Text)
                End If

                If End1 = 0 Then
                    MsgBox("End row specified incorrectly")
                    Freeze_operations = False
                    Exit Sub
                End If

                If Start1 = 0 Then
                    MsgBox("Start row specified incorrectly")
                    Freeze_operations = False
                    Exit Sub
                End If
                If End1 < Start1 Then
                    MsgBox("Start row bigger than end row")
                    Freeze_operations = False
                    Exit Sub
                End If
                Dim Column_pipe_ID As String = "A"
                Column_pipe_ID = TextBox_read_pipeID.Text.ToUpper

                Dim Column_wt As String = "A"
                Column_wt = TextBox_read_Wall_thickness.Text.ToUpper



                Dim Column_grade As String = "A"
                Column_grade = TextBox_read_grade.Text.ToUpper
                Dim Column_Pipe_diam As String = "A"
                Column_Pipe_diam = TextBox_read_pipe_diam.Text.ToUpper
                Dim Column_coating As String = "A"
                Column_coating = TextBox_read_pipe_coating.Text.ToUpper

                Dim Column_joint As String = "A"
                Column_joint = TextBox_read_joint_number.Text.ToUpper



                Data_table_read_pipe_tally = New System.Data.DataTable
                Data_table_read_pipe_tally.Columns.Add("PIPE_ID", GetType(String))
                Data_table_read_pipe_tally.Columns.Add("JOINT_NUMBER", GetType(String))
                Data_table_read_pipe_tally.Columns.Add("WT", GetType(String))
                Data_table_read_pipe_tally.Columns.Add("GRADE", GetType(String))
                Data_table_read_pipe_tally.Columns.Add("PIPE_DIAM", GetType(String))
                Data_table_read_pipe_tally.Columns.Add("COATING", GetType(String))

                Dim Index_dt As Double = 0

                For i = Start1 To End1
                    Dim Pipe_id1 As String = ""

                    Dim WT As String = ""
                    Dim Grade As String = ""
                    Dim Pipe_diam As String = ""
                    Dim Pipe_coating As String = ""
                    Dim Joint As String = ""

                    Try
                        Grade = W1.Range(Column_grade & i).Value2
                        Pipe_diam = W1.Range(Column_Pipe_diam & i).Value2
                        WT = W1.Range(Column_wt & i).Value2
                        Joint = W1.Range(Column_joint & i).Value2
                        Pipe_id1 = W1.Range(Column_pipe_ID & i).Value2
                        Pipe_coating = W1.Range(Column_coating & i).Value2
                    Catch ex As System.SystemException

                    End Try

                    If Not Pipe_id1 = "" Then
                        Data_table_read_pipe_tally.Rows.Add()
                        Data_table_read_pipe_tally.Rows(Index_dt).Item("PIPE_ID") = Pipe_id1

                        If Not WT = "" Then Data_table_read_pipe_tally.Rows(Index_dt).Item("WT") = WT
                        If Not Grade = "" Then Data_table_read_pipe_tally.Rows(Index_dt).Item("GRADE") = Grade
                        If Not Pipe_diam = "" Then Data_table_read_pipe_tally.Rows(Index_dt).Item("PIPE_DIAM") = Pipe_diam
                        If Not Pipe_coating = "" Then Data_table_read_pipe_tally.Rows(Index_dt).Item("COATING") = Pipe_coating
                        If Not Joint = "" Then Data_table_read_pipe_tally.Rows(Index_dt).Item("JOINT_NUMBER") = Joint


                        Index_dt = Index_dt + 1
                    End If
                Next

                'MsgBox(Data_table_Centerline.Rows.Count)
                Add_to_clipboard_Data_table(Data_table_read_pipe_tally)

            Catch ex As Exception
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "DONE")
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_read_all_points_Click(sender As Object, e As EventArgs) Handles Button_read_all_points.Click

        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

        If Freeze_operations = False Then
            Freeze_operations = True

            Try

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Dim Start1 As Integer = 0
                Dim End1 As Integer = 0
                If IsNumeric(TextBox_read_start_row.Text) = True Then
                    Start1 = CInt(TextBox_read_start_row.Text)
                End If
                If IsNumeric(TextBox_read_end_row.Text) = True Then
                    End1 = CInt(TextBox_read_end_row.Text)
                End If

                If End1 = 0 Then
                    MsgBox("End row specified incorrectly")
                    Freeze_operations = False
                    Exit Sub
                End If

                If Start1 = 0 Then
                    MsgBox("Start row specified incorrectly")
                    Freeze_operations = False
                    Exit Sub
                End If
                If End1 < Start1 Then
                    MsgBox("Start row bigger than end row")
                    Freeze_operations = False
                    Exit Sub
                End If

                Dim Column_point_number As String = "A"
                Column_point_number = TextBox_read_all_pt_point_no.Text.ToUpper

                Dim Column_description As String = "A"
                Column_description = TextBox_read_all_pt_descr.Text.ToUpper

                Dim Column_Notes_start As String = "A"
                Column_Notes_start = TextBox_asb_start.Text.ToUpper

                Dim Column_Notes_end As String = "A"
                Column_Notes_end = TextBox_asb_end.Text.ToUpper

                Data_table_read_all_points = New System.Data.DataTable
                Data_table_read_all_points.Columns.Add("POINT_NUMBER", GetType(String))
                Data_table_read_all_points.Columns.Add("DESCRIPTION", GetType(String))
                Data_table_read_all_points.Columns.Add("NOTES", GetType(String))


                Dim Index_dt As Double = 0

                For i = Start1 To End1
                    Dim Point_number As String = ""
                    Dim Descr1 As String = ""
                    Dim Notes As String = ""



                    Try
                        For j = W1.Range(Column_Notes_start & "1").Column To W1.Range(Column_Notes_end & "1").Column
                            Dim Note1 As String = W1.Cells(i, j).Value2
                            If Not Note1 = "" Then
                                If Notes = "" Then
                                    Notes = Note1
                                Else
                                    Notes = Notes & " " & Note1
                                End If
                            End If
                        Next
                        Descr1 = W1.Range(Column_description & i).Value2
                        Point_number = W1.Range(Column_point_number & i).Value2

                    Catch ex As System.SystemException

                    End Try

                    If Not Point_number = "" Then
                        Data_table_read_all_points.Rows.Add()
                        Data_table_read_all_points.Rows(Index_dt).Item("POINT_NUMBER") = Point_number
                        If Not Descr1 = "" Then Data_table_read_all_points.Rows(Index_dt).Item("DESCRIPTION") = Descr1
                        If Not Notes = "" Then Data_table_read_all_points.Rows(Index_dt).Item("NOTES") = Notes
                        Index_dt = Index_dt + 1
                    End If
                Next

                'MsgBox(Data_table_Centerline.Rows.Count)
                Add_to_clipboard_Data_table(Data_table_read_all_points)

            Catch ex As Exception
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "DONE")
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_tools_read_materials_Click(sender As Object, e As EventArgs) Handles Button_tools_read_materials.Click

        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
        If IsNothing(Data_table_read_pipe_tally) = True Then
            MsgBox("no pipe tally loaded")
            Exit Sub
        End If

        If Data_table_read_pipe_tally.Rows.Count = 0 Then
            MsgBox("no pipe tally loaded")
            Exit Sub
        End If

        If IsNothing(Data_table_read_all_points) = True Then
            MsgBox("no all points loaded")
            Exit Sub
        End If

        If Data_table_read_all_points.Rows.Count = 0 Then
            MsgBox("no all points loaded")
            Exit Sub
        End If

        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Dim Start1 As Integer = 0
                Dim End1 As Integer = 0
                If IsNumeric(TextBox_read_start_row.Text) = True Then
                    Start1 = CInt(TextBox_read_start_row.Text)
                End If
                If IsNumeric(TextBox_read_end_row.Text) = True Then
                    End1 = CInt(TextBox_read_end_row.Text)
                End If

                If End1 = 0 Then
                    MsgBox("End row specified incorrectly")
                    Freeze_operations = False
                    Exit Sub
                End If

                If Start1 = 0 Then
                    MsgBox("Start row specified incorrectly")
                    Freeze_operations = False
                    Exit Sub
                End If
                If End1 < Start1 Then
                    MsgBox("Start row bigger than end row")
                    Freeze_operations = False
                    Exit Sub
                End If

                Dim Column_pipe_ID As String = "A"
                Column_pipe_ID = TextBox_read_hmm_back.Text.ToUpper

                Dim Column_station As String = "A"
                Column_station = TextBox_read_mat_station.Text.ToUpper

                Dim Column_pt_no As String = "A"
                Column_pt_no = TextBox_read_mat_Pt_number.Text.ToUpper




                Data_table_read_materials = New System.Data.DataTable
                Data_table_read_materials.Columns.Add("PIPE_ID", GetType(String))
                Data_table_read_materials.Columns.Add("POINT_NO", GetType(String))
                Data_table_read_materials.Columns.Add("STATION", GetType(Double))
                Data_table_read_materials.Columns.Add("DESCRIPTION", GetType(String))
                Data_table_read_materials.Columns.Add("DESCRIPTION_PT_NO", GetType(String))


                For i = Start1 To End1
                    Dim Pipe_id1 As String = ""

                    Dim Station As String = ""
                    Dim Pt_no As String = ""

                    Try
                        Pipe_id1 = W1.Range(Column_pipe_ID & i).Value2
                        Station = W1.Range(Column_station & i).Value2
                        Pt_no = W1.Range(Column_pt_no & i).Value2

                    Catch ex As System.SystemException

                    End Try

                    If Not Pipe_id1 = "" And IsNumeric(Station) = True Then
                        Data_table_read_materials.Rows.Add()
                        Dim Index_dt As Double = Data_table_read_materials.Rows.Count - 1
                        Data_table_read_materials.Rows(Index_dt).Item("PIPE_ID") = Pipe_id1
                        Data_table_read_materials.Rows(Index_dt).Item("POINT_NO") = Pt_no

                        Data_table_read_materials.Rows(Index_dt).Item("STATION") = CDbl(Station)
                        For j = 0 To Data_table_read_pipe_tally.Rows.Count - 1
                            If IsDBNull(Data_table_read_pipe_tally.Rows(j).Item("PIPE_ID")) = False Then
                                Dim Pipe_Id_tally As String = Data_table_read_pipe_tally.Rows(j).Item("PIPE_ID")
                                If Pipe_Id_tally.ToUpper = Pipe_id1.ToUpper Then
                                    Dim Description As String = ""

                                    Dim Pipe_diam As String = ""
                                    Dim WT As String = ""
                                    Dim Grade As String = ""

                                    Dim Pipe_coating As String = ""
                                    Dim Joint As String = ""




                                    If IsDBNull(Data_table_read_pipe_tally.Rows(j).Item("JOINT_NUMBER")) = False Then
                                        Joint = Data_table_read_pipe_tally.Rows(j).Item("JOINT_NUMBER")
                                    End If

                                    If IsDBNull(Data_table_read_pipe_tally.Rows(j).Item("WT")) = False Then
                                        WT = Data_table_read_pipe_tally.Rows(j).Item("WT")
                                    End If

                                    If IsDBNull(Data_table_read_pipe_tally.Rows(j).Item("GRADE")) = False Then
                                        Grade = Data_table_read_pipe_tally.Rows(j).Item("GRADE")
                                    End If
                                    If IsDBNull(Data_table_read_pipe_tally.Rows(j).Item("PIPE_DIAM")) = False Then
                                        Pipe_diam = Data_table_read_pipe_tally.Rows(j).Item("PIPE_DIAM")
                                    End If
                                    If IsDBNull(Data_table_read_pipe_tally.Rows(j).Item("COATING")) = False Then
                                        Pipe_coating = Data_table_read_pipe_tally.Rows(j).Item("COATING")
                                    End If

                                    Select Case Joint.ToUpper
                                        Case "FITTING"
                                            Description = Joint
                                            If Not Pipe_diam = "" Then
                                                Description = Description & " " & Pipe_diam
                                            End If
                                            If Not WT = "" Then
                                                Description = Description & " " & WT
                                            End If
                                            If Not Grade = "" Then
                                                Description = Description & " " & Grade
                                            End If
                                            If Not Pipe_coating = "" Then
                                                Description = Description & " " & Pipe_coating
                                            End If

                                        Case Else


                                            If Not Pipe_diam = "" Then
                                                Description = Pipe_diam
                                            End If

                                            If Not WT = "" Then
                                                If Not Description = "" Then
                                                    Description = Description & " " & WT
                                                Else
                                                    Description = WT
                                                End If

                                            End If
                                            If Not Grade = "" Then
                                                If Not Description = "" Then
                                                    Description = Description & " " & Grade
                                                Else
                                                    Description = Grade
                                                End If

                                            End If
                                            If Not Pipe_coating = "" Then
                                                If Not Description = "" Then
                                                    Description = Description & " " & Pipe_coating
                                                Else
                                                    Description = Pipe_coating
                                                End If

                                            End If
                                    End Select

                                    If Not Description = "" Then
                                        Data_table_read_materials.Rows(Index_dt).Item("DESCRIPTION") = Description
                                    End If

                                    Exit For

                                End If
                            End If
                        Next


                        For k = 0 To Data_table_read_all_points.Rows.Count - 1
                            If IsDBNull(Data_table_read_all_points.Rows(k).Item("POINT_NUMBER")) = False Then
                                Dim PTNO As String = Data_table_read_all_points.Rows(k).Item("POINT_NUMBER")
                                If PTNO.ToUpper = Pt_no Then
                                    Dim Description_pt_no As String = ""





                                    If IsDBNull(Data_table_read_all_points.Rows(k).Item("DESCRIPTION")) = False Then
                                        Description_pt_no = Data_table_read_all_points.Rows(k).Item("DESCRIPTION")
                                    End If

                                    If IsDBNull(Data_table_read_all_points.Rows(k).Item("NOTES")) = False Then
                                        If Description_pt_no = "" Then
                                            Description_pt_no = Data_table_read_all_points.Rows(k).Item("NOTES")
                                        Else
                                            Description_pt_no = Description_pt_no & " " & Data_table_read_all_points.Rows(k).Item("NOTES")
                                        End If

                                    End If



                                    If Not Description_pt_no = "" Then
                                        Data_table_read_materials.Rows(Index_dt).Item("DESCRIPTION_PT_NO") = Description_pt_no
                                    End If

                                    Exit For

                                End If
                            End If
                        Next


                    End If

                    If Pipe_id1 = "" And IsNumeric(Station) = True And Not Pt_no = "" Then
                        Data_table_read_materials.Rows.Add()
                        Dim Index_dt As Double = Data_table_read_materials.Rows.Count - 1
                        Data_table_read_materials.Rows(Index_dt).Item("POINT_NO") = Pt_no

                        Data_table_read_materials.Rows(Index_dt).Item("STATION") = CDbl(Station)



                        For k = 0 To Data_table_read_all_points.Rows.Count - 1
                            If IsDBNull(Data_table_read_all_points.Rows(k).Item("POINT_NUMBER")) = False Then
                                Dim PTNO As String = Data_table_read_all_points.Rows(k).Item("POINT_NUMBER")
                                If PTNO.ToUpper = Pt_no Then
                                    Dim Description_pt_no As String = ""





                                    If IsDBNull(Data_table_read_all_points.Rows(k).Item("DESCRIPTION")) = False Then
                                        Description_pt_no = Data_table_read_all_points.Rows(k).Item("DESCRIPTION")
                                    End If

                                    If IsDBNull(Data_table_read_all_points.Rows(k).Item("NOTES")) = False Then
                                        If Description_pt_no = "" Then
                                            Description_pt_no = Data_table_read_all_points.Rows(k).Item("NOTES")
                                        Else
                                            Description_pt_no = Description_pt_no & " " & Data_table_read_all_points.Rows(k).Item("NOTES")
                                        End If

                                    End If



                                    If Not Description_pt_no = "" Then
                                        Data_table_read_materials.Rows(Index_dt).Item("DESCRIPTION_PT_NO") = Description_pt_no
                                    End If

                                    Exit For

                                End If
                            End If
                        Next


                    End If

                Next

                'MsgBox(Data_table_Centerline.Rows.Count)
                Add_to_clipboard_Data_table(Data_table_read_materials)

            Catch ex As Exception
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "DONE")
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_read_fitting_Click(sender As Object, e As EventArgs) Handles Button_read_fitting.Click
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
        If IsNothing(Data_table_read_all_points) = True Then
            MsgBox("no all points table loaded")
            Exit Sub
        End If

        If Data_table_read_all_points.Rows.Count = 0 Then
            MsgBox("no all points table loaded")
            Exit Sub
        End If


        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Dim Start1 As Integer = 0
                Dim End1 As Integer = 0
                If IsNumeric(TextBox_read_start_row.Text) = True Then
                    Start1 = CInt(TextBox_read_start_row.Text)
                End If
                If IsNumeric(TextBox_read_end_row.Text) = True Then
                    End1 = CInt(TextBox_read_end_row.Text)
                End If

                If End1 = 0 Then
                    MsgBox("End row specified incorrectly")
                    Freeze_operations = False
                    Exit Sub
                End If

                If Start1 = 0 Then
                    MsgBox("Start row specified incorrectly")
                    Freeze_operations = False
                    Exit Sub
                End If
                If End1 < Start1 Then
                    MsgBox("Start row bigger than end row")
                    Freeze_operations = False
                    Exit Sub
                End If

                Dim Column_point_no As String = "A"
                Column_point_no = TextBox_read_fitting_point_no.Text.ToUpper

                Dim Column_station As String = "A"
                Column_station = TextBox_read_fitting_sta.Text.ToUpper



                Dim Description As String = "DESCRIPTION"

                Data_table_read_fitting = New System.Data.DataTable
                Data_table_read_fitting.Columns.Add("POINT_NUMBER", GetType(String))
                Data_table_read_fitting.Columns.Add("STATION", GetType(Double))
                Data_table_read_fitting.Columns.Add(Description, GetType(String))




                Dim Index_dt As Double = 0

                For i = Start1 To End1
                    Dim Point_no As String = ""
                    Dim Station As String = ""


                    Try
                        Point_no = W1.Range(Column_point_no & i).Value2
                        Station = W1.Range(Column_station & i).Value2

                    Catch ex As System.SystemException

                    End Try

                    If Not Point_no = "" And IsNumeric(Station) = True Then


                        For k = 0 To Data_table_read_all_points.Rows.Count - 1
                            If IsDBNull(Data_table_read_all_points.Rows(k).Item("POINT_NUMBER")) = False Then
                                Dim Point_number_ALL_pt As String = Data_table_read_all_points.Rows(k).Item("POINT_NUMBER")
                                If Point_number_ALL_pt.ToUpper = Point_no.ToUpper Then

                                    Dim Notes As String = ""
                                    Dim Descr_all_pt As String = ""

                                    If IsDBNull(Data_table_read_all_points.Rows(k).Item("DESCRIPTION")) = False Then
                                        Descr_all_pt = Data_table_read_all_points.Rows(k).Item("DESCRIPTION")
                                    End If

                                    If IsDBNull(Data_table_read_all_points.Rows(k).Item("NOTES")) = False Then
                                        Notes = Data_table_read_all_points.Rows(k).Item("NOTES")
                                    End If

                                    'INDUCTION HOT BEND
                                    Dim Add_it_to_table As Boolean = False

                                    If Notes.ToUpper.Contains("BEND") = True Then
                                        If Notes.ToUpper.Contains("COLD") = False Then
                                            Add_it_to_table = True
                                        End If
                                    Else
                                        Add_it_to_table = True
                                    End If

                                    If Add_it_to_table = True Then

                                        Data_table_read_fitting.Rows.Add()
                                        Data_table_read_fitting.Rows(Index_dt).Item("POINT_NUMBER") = Point_no
                                        Data_table_read_fitting.Rows(Index_dt).Item("STATION") = CDbl(Station)
                                        Data_table_read_fitting.Rows(Index_dt).Item(Description) = Notes


                                        Index_dt = Index_dt + 1



                                    End If

                                    Exit For

                                End If
                            End If
                        Next


                    End If
                Next

                'MsgBox(Data_table_Centerline.Rows.Count)
                Add_to_clipboard_Data_table(Data_table_read_fitting)

            Catch ex As Exception
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "DONE")
            MsgBox("Done")
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_load_fitting_Click(sender As Object, e As EventArgs) Handles Button_load_fitting.Click

        If Freeze_operations = False Then
            Freeze_operations = True

            Try

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Dim Start1 As Integer = 0
                Dim End1 As Integer = 0
                If IsNumeric(TextBox_ROW_START.Text) = True Then
                    Start1 = CInt(TextBox_ROW_START.Text)
                End If
                If IsNumeric(TextBox_ROW_END.Text) = True Then
                    End1 = CInt(TextBox_ROW_END.Text)
                End If

                If End1 = 0 Then
                    MsgBox("End row specified incorrectly")
                    Freeze_operations = False
                    Exit Sub
                End If

                If Start1 = 0 Then
                    MsgBox("Start row specified incorrectly")
                    Freeze_operations = False
                    Exit Sub
                End If
                If End1 < Start1 Then
                    MsgBox("Start row bigger than end row")
                    Freeze_operations = False
                    Exit Sub
                End If
                Dim Column_station As String = ""
                Column_station = TextBox_fitting_station.Text.ToUpper

                Dim Column_description As String = ""
                Column_description = TextBox_fitting_description.Text.ToUpper
                Dim Column_MATERIAL As String = ""
                Column_MATERIAL = TextBox_fitting_material.Text.ToUpper


                Data_table_fitting = New System.Data.DataTable
                Data_table_fitting.Columns.Add("STATION", GetType(Double))
                Data_table_fitting.Columns.Add("DESCRIPTION", GetType(String))
                Data_table_fitting.Columns.Add("MATERIAL", GetType(String))
                Data_table_fitting.Columns.Add("X", GetType(Double))
                Data_table_fitting.Columns.Add("Y", GetType(Double))


                Dim Index_data_table As Double = 0

                For i = Start1 To End1
                    Dim Station_string1 As String = W1.Range(Column_station & i).Value2
                    Dim Description As String = W1.Range(Column_description & i).Value2
                    Dim Material As String = W1.Range(Column_MATERIAL & i).Value2
                    If IsNothing(Station_string1) = False Then
                        If IsNumeric(Station_string1) = True Then
                            Data_table_fitting.Rows.Add()
                            Data_table_fitting.Rows(Index_data_table).Item("STATION") = CDbl(Station_string1)
                            If IsNothing(Description) = False Then
                                If Not Description = "" Then
                                    Data_table_fitting.Rows(Index_data_table).Item("DESCRIPTION") = Description
                                End If

                            End If

                            If IsNothing(Material) = False Then
                                If Not Material = "" Then
                                    Data_table_fitting.Rows(Index_data_table).Item("MATERIAL") = Material
                                End If

                            End If

                            Index_data_table = Index_data_table + 1
                        End If
                    End If

                Next

                Data_table_fitting = Sort_data_table(Data_table_fitting, "STATION")

                'MsgBox(Data_table_Centerline.Rows.Count)
                Add_to_clipboard_Data_table(Data_table_fitting)
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & Data_table_fitting.Rows.Count & " fittings loaded")


            Catch ex As Exception
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try

            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_redefine_fittings_Click(sender As Object, e As EventArgs) Handles Button_redefine_fittings.Click



        If Freeze_operations = False Then
            Freeze_operations = True
            Dim Round1 As Integer = 0
            If IsNumeric(TextBox_dec1.Text) = True Then
                Round1 = CInt(TextBox_dec1.Text)
            End If
            Try

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Dim Start1 As Integer = 0
                Dim End1 As Integer = 0
                If IsNumeric(TextBox_ROW_START.Text) = True Then
                    Start1 = CInt(TextBox_ROW_START.Text)
                End If
                If IsNumeric(TextBox_ROW_END.Text) = True Then
                    End1 = CInt(TextBox_ROW_END.Text)
                End If

                If End1 = 0 Then
                    MsgBox("End row specified incorrectly")
                    Freeze_operations = False
                    Exit Sub
                End If

                If Start1 = 0 Then
                    MsgBox("Start row specified incorrectly")
                    Freeze_operations = False
                    Exit Sub
                End If
                If End1 < Start1 Then
                    MsgBox("Start row bigger than end row")
                    Freeze_operations = False
                    Exit Sub
                End If
                Dim Column_station As String = ""
                Column_station = TextBox_fitting_station.Text.ToUpper

                Dim Column_description As String = ""
                Column_description = TextBox_fitting_description.Text.ToUpper
                Dim Column_MATERIAL As String = ""
                Column_MATERIAL = TextBox_fitting_material.Text.ToUpper


                Data_table_fitting = New System.Data.DataTable
                Data_table_fitting.Columns.Add("STATION", GetType(Double))
                Data_table_fitting.Columns.Add("DESCRIPTION", GetType(String))
                Data_table_fitting.Columns.Add("MATERIAL", GetType(String))
                Data_table_fitting.Columns.Add("X", GetType(Double))
                Data_table_fitting.Columns.Add("Y", GetType(Double))


                Dim Index_data_table As Double = 0

                For i = Start1 To End1
                    Dim Station_string1 As String = W1.Range(Column_station & i).Value2
                    Dim Description As String = W1.Range(Column_description & i).Value2
                    Dim Material As String = W1.Range(Column_MATERIAL & i).Value2
                    If IsNothing(Station_string1) = False Then
                        If IsNumeric(Station_string1) = True Then
                            Data_table_fitting.Rows.Add()
                            Data_table_fitting.Rows(Index_data_table).Item("STATION") = CDbl(Station_string1)
                            If IsNothing(Description) = False Then
                                If Not Description = "" Then
                                    Data_table_fitting.Rows(Index_data_table).Item("DESCRIPTION") = Description
                                End If

                            End If

                            If IsNothing(Material) = False Then
                                If Not Material = "" Then
                                    Data_table_fitting.Rows(Index_data_table).Item("MATERIAL") = Material
                                End If

                            End If

                            Index_data_table = Index_data_table + 1
                        End If
                    End If

                Next

                Data_table_fitting = Sort_data_table(Data_table_fitting, "STATION")

                'MsgBox(Data_table_Centerline.Rows.Count)
                Add_to_clipboard_Data_table(Data_table_fitting)
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & Data_table_fitting.Rows.Count & " fittings loaded")

                Dim Empty_array() As ObjectId

                Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor

                Editor1.SetImpliedSelection(Empty_array)

                Using lock As DocumentLock = ThisDrawing.LockDocument

                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Try


                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
1234:

                            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                            Object_Prompt.MessageForAdding = vbLf & "Select engineering bands:"

                            Object_Prompt.SingleOnly = False

                            Rezultat1 = Editor1.GetSelection(Object_Prompt)



                            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Editor1.WriteMessage(vbLf & "Command:")
                                Freeze_operations = False
                                Exit Sub
                            End If



                            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                If IsNothing(Rezultat1) = False Then
                                    If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                        If IsNothing(Rezultat1) = False Then
                                            For i = 0 To Rezultat1.Value.Count - 1
                                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                                Obj1 = Rezultat1.Value.Item(i)
                                                Dim Ent1 As Entity
                                                Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                                                If TypeOf Ent1 Is BlockReference Then
                                                    Dim Block1 As BlockReference = Ent1
                                                    Dim Block_name As String

                                                    Dim BlockTrec As BlockTableRecord = Nothing
                                                    If Block1.IsDynamicBlock = True Then
                                                        BlockTrec = Trans1.GetObject(Block1.DynamicBlockTableRecord, OpenMode.ForRead)
                                                        Block_name = BlockTrec.Name
                                                    Else
                                                        BlockTrec = Trans1.GetObject(Block1.BlockTableRecord, OpenMode.ForRead)
                                                        Block_name = BlockTrec.Name
                                                    End If


                                                    If Block_name = ComboBox_blocks_fitting.Text Then

                                                        Dim Station As Double = -1



                                                        If Block1.AttributeCollection.Count > 0 Then
                                                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection


                                                            For Each id In attColl
                                                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                                                If attref.Tag = ComboBox_STA_att_fitting.Text Then
                                                                    Dim Continut As String = attref.TextString
                                                                    If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                                        Station = Round(CDbl(Replace(Continut, "+", "")), Round1)
                                                                    End If
                                                                End If
                                                            Next

                                                            If Not Station = -1 Then
                                                                Dim Station_excel As Double = -1
                                                                Dim New_material As String = ""
                                                                Dim New_description As String = ""

                                                                For j = 0 To Data_table_fitting.Rows.Count - 1
                                                                    If IsDBNull(Data_table_fitting.Rows(j).Item("STATION")) = False Then
                                                                        Station_excel = Data_table_fitting.Rows(j).Item("STATION")
                                                                        If Round(Station_excel, Round1) = Round(Station, Round1) Then
                                                                            If IsDBNull(Data_table_fitting.Rows(j).Item("MATERIAL")) = False Then
                                                                                New_material = Data_table_fitting.Rows(j).Item("MATERIAL")
                                                                            End If
                                                                            If IsDBNull(Data_table_fitting.Rows(j).Item("DESCRIPTION")) = False Then
                                                                                New_description = Data_table_fitting.Rows(j).Item("DESCRIPTION")
                                                                            End If
                                                                            If Not New_description = "" Or Not New_material = "" Then
                                                                                Block1.UpgradeOpen()
                                                                                For Each id In attColl
                                                                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForWrite)

                                                                                    If Not New_material = "" Then
                                                                                        If attref.Tag = ComboBox_MAT_att_fitting.Text Then
                                                                                            attref.TextString = New_material
                                                                                        End If
                                                                                    End If

                                                                                    If Not New_description = "" Then
                                                                                        If attref.Tag = ComboBox_DESCRIPTION_att_fitting.Text Then
                                                                                            attref.TextString = New_description
                                                                                        End If
                                                                                    End If
                                                                                Next
                                                                            End If
                                                                            Exit For
                                                                        End If

                                                                    End If
                                                                Next
                                                            End If
                                                        End If

                                                    End If

                                                End If


                                            Next
                                            Trans1.Commit()


                                        End If
                                    End If
                                End If
                            End If

                        End Using



                        Editor1.SetImpliedSelection(Empty_array)
                        Editor1.WriteMessage(vbLf & "Command:")

                    Catch ex As Exception
                        Editor1.SetImpliedSelection(Empty_array)
                        Editor1.WriteMessage(vbLf & "Command:")
                        Freeze_operations = False
                        MsgBox(ex.Message)
                    End Try

                End Using


            Catch ex As Exception
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try

            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_load_CL_Click(sender As Object, e As EventArgs) Handles Button_load_CL.Click
        If Freeze_operations = False Then
            Freeze_operations = True

            Try

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Dim Start1 As Integer = 0
                Dim End1 As Integer = 0
                If IsNumeric(TextBox_ROW_START.Text) = True Then
                    Start1 = CInt(TextBox_ROW_START.Text)
                End If
                If IsNumeric(TextBox_ROW_END.Text) = True Then
                    End1 = CInt(TextBox_ROW_END.Text)
                End If

                If End1 = 0 Then
                    MsgBox("End row specified incorrectly")
                    Freeze_operations = False
                    Exit Sub
                End If

                If Start1 = 0 Then
                    MsgBox("Start row specified incorrectly")
                    Freeze_operations = False
                    Exit Sub
                End If
                If End1 < Start1 Then
                    MsgBox("Start row bigger than end row")
                    Freeze_operations = False
                    Exit Sub
                End If
                Dim Column_station As String = ""
                Column_station = TextBox_CL_station.Text.ToUpper

                Dim Column_description As String = ""
                Column_description = TextBox_CL_description.Text.ToUpper



                Data_table_cl_crossing = New System.Data.DataTable
                Data_table_cl_crossing.Columns.Add("STATION", GetType(Double))
                Data_table_cl_crossing.Columns.Add("DESCRIPTION", GetType(String))
                Data_table_cl_crossing.Columns.Add("X", GetType(Double))
                Data_table_cl_crossing.Columns.Add("Y", GetType(Double))



                Dim Index_data_table As Double = 0

                For i = Start1 To End1
                    Dim Station_string1 As String = W1.Range(Column_station & i).Value2
                    Dim Description As String = W1.Range(Column_description & i).Value2
                    If IsNothing(Station_string1) = False Then
                        If IsNumeric(Station_string1) = True Then
                            Data_table_cl_crossing.Rows.Add()
                            Data_table_cl_crossing.Rows(Index_data_table).Item("STATION") = CDbl(Station_string1)
                            If IsNothing(Description) = False Then
                                If Not Description = "" Then
                                    Data_table_cl_crossing.Rows(Index_data_table).Item("DESCRIPTION") = Description
                                End If

                            End If

                            Index_data_table = Index_data_table + 1
                        End If
                    End If

                Next

                Data_table_cl_crossing = Sort_data_table(Data_table_cl_crossing, "STATION")

                'MsgBox(Data_table_Centerline.Rows.Count)
                Add_to_clipboard_Data_table(Data_table_cl_crossing)
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & Data_table_cl_crossing.Rows.Count & " cl crossings loaded")


            Catch ex As Exception
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try

            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_load_cad_weld_Click(sender As Object, e As EventArgs) Handles Button_load_cad_weld.Click
        If Freeze_operations = False Then
            Freeze_operations = True

            Try

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Dim Start1 As Integer = 0
                Dim End1 As Integer = 0
                If IsNumeric(TextBox_ROW_START.Text) = True Then
                    Start1 = CInt(TextBox_ROW_START.Text)
                End If
                If IsNumeric(TextBox_ROW_END.Text) = True Then
                    End1 = CInt(TextBox_ROW_END.Text)
                End If

                If End1 = 0 Then
                    MsgBox("End row specified incorrectly")
                    Freeze_operations = False
                    Exit Sub
                End If

                If Start1 = 0 Then
                    MsgBox("Start row specified incorrectly")
                    Freeze_operations = False
                    Exit Sub
                End If
                If End1 < Start1 Then
                    MsgBox("Start row bigger than end row")
                    Freeze_operations = False
                    Exit Sub
                End If
                Dim Column_station As String = ""
                Column_station = TextBox_cathodic_station.Text.ToUpper




                Data_table_cad_weld = New System.Data.DataTable
                Data_table_cad_weld.Columns.Add("STATION", GetType(Double))
                Data_table_cad_weld.Columns.Add("X", GetType(Double))
                Data_table_cad_weld.Columns.Add("Y", GetType(Double))




                Dim Index_data_table As Double = 0

                For i = Start1 To End1
                    Dim Station_string1 As String = W1.Range(Column_station & i).Value2

                    If IsNothing(Station_string1) = False Then
                        If IsNumeric(Station_string1) = True Then
                            Data_table_cad_weld.Rows.Add()
                            Data_table_cad_weld.Rows(Index_data_table).Item("STATION") = CDbl(Station_string1)


                            Index_data_table = Index_data_table + 1
                        End If
                    End If

                Next

                Data_table_cad_weld = Sort_data_table(Data_table_cad_weld, "STATION")

                'MsgBox(Data_table_Centerline.Rows.Count)
                Add_to_clipboard_Data_table(Data_table_cad_weld)
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & Data_table_cad_weld.Rows.Count & " cad welds loaded")


            Catch ex As Exception
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try

            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_load_river_weights_Click(sender As Object, e As EventArgs) Handles Button_load_river_weights.Click
        If Freeze_operations = False Then
            Freeze_operations = True

            Try

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Dim Start1 As Integer = 0
                Dim End1 As Integer = 0
                If IsNumeric(TextBox_ROW_START.Text) = True Then
                    Start1 = CInt(TextBox_ROW_START.Text)
                End If
                If IsNumeric(TextBox_ROW_END.Text) = True Then
                    End1 = CInt(TextBox_ROW_END.Text)
                End If

                If End1 = 0 Then
                    MsgBox("End row specified incorrectly")
                    Freeze_operations = False
                    Exit Sub
                End If

                If Start1 = 0 Then
                    MsgBox("Start row specified incorrectly")
                    Freeze_operations = False
                    Exit Sub
                End If
                If End1 < Start1 Then
                    MsgBox("Start row bigger than end row")
                    Freeze_operations = False
                    Exit Sub
                End If
                Dim Column_station As String = ""
                Column_station = TextBox_river_weights_station.Text.ToUpper




                Data_table_river_weights = New System.Data.DataTable
                Data_table_river_weights.Columns.Add("STATION", GetType(Double))
                Data_table_river_weights.Columns.Add("X", GetType(Double))
                Data_table_river_weights.Columns.Add("Y", GetType(Double))



                Dim Index_data_table As Double = 0

                For i = Start1 To End1
                    Dim Station_string1 As String = W1.Range(Column_station & i).Value2

                    If IsNothing(Station_string1) = False Then
                        If IsNumeric(Station_string1) = True Then
                            Data_table_river_weights.Rows.Add()
                            Data_table_river_weights.Rows(Index_data_table).Item("STATION") = CDbl(Station_string1)


                            Index_data_table = Index_data_table + 1
                        End If
                    End If

                Next

                Data_table_river_weights = Sort_data_table(Data_table_river_weights, "STATION")

                'MsgBox(Data_table_Centerline.Rows.Count)
                Add_to_clipboard_Data_table(Data_table_river_weights)
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & Data_table_river_weights.Rows.Count & " river weights loaded")


            Catch ex As Exception
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try

            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_pick_all_Click(sender As Object, e As EventArgs) Handles Button_pick_all.Click
        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Round1 As Integer = 0
            If IsNumeric(TextBox_dec1.Text) = True Then
                Round1 = CInt(TextBox_dec1.Text)
            End If

            Dim Empty_array() As ObjectId

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor

            Dim Index_data_table As Integer = 0

            If IsNothing(Data_table_material_count) = True Then
                Data_table_material_count = New System.Data.DataTable
                Data_table_material_count.Columns.Add("STA1", GetType(Double))
                Data_table_material_count.Columns.Add("STA2", GetType(Double))
                Data_table_material_count.Columns.Add("MAT", GetType(String))
                Data_table_material_count.Columns.Add("AMOUNT_IN_BLOCK", GetType(Double))
                Data_table_material_count.Columns.Add("CALCULATED_AMOUNT", GetType(Double))
            Else
                Index_data_table = Data_table_material_count.Rows.Count
            End If

            Editor1.SetImpliedSelection(Empty_array)

            Using lock As DocumentLock = ThisDrawing.LockDocument

                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                Try


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
1234:

                        Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                        Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_Prompt.MessageForAdding = vbLf & "Select Blocks:"

                        Object_Prompt.SingleOnly = False

                        Rezultat1 = Editor1.GetSelection(Object_Prompt)



                        If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Editor1.WriteMessage(vbLf & "Command:")
                            Freeze_operations = False
                            Exit Sub
                        End If



                        If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            If IsNothing(Rezultat1) = False Then
                                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                    If IsNothing(Rezultat1) = False Then
                                        For i = 0 To Rezultat1.Value.Count - 1
                                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                            Obj1 = Rezultat1.Value.Item(i)
                                            Dim Ent1 As Entity
                                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                                            If TypeOf Ent1 Is BlockReference Then
                                                Dim Block1 As BlockReference = Ent1
                                                Dim Block_name As String

                                                Dim BlockTrec As BlockTableRecord = Nothing
                                                If Block1.IsDynamicBlock = True Then
                                                    BlockTrec = Trans1.GetObject(Block1.DynamicBlockTableRecord, OpenMode.ForRead)
                                                    Block_name = BlockTrec.Name
                                                Else
                                                    BlockTrec = Trans1.GetObject(Block1.BlockTableRecord, OpenMode.ForRead)
                                                    Block_name = BlockTrec.Name
                                                End If



                                                Dim Material As String = ""
                                                Dim Station1 As Double = 0
                                                Dim Station2 As Double = 0
                                                Dim Length1 As Double = 0

                                                If Block1.AttributeCollection.Count > 0 Then
                                                    Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection

                                                    Dim Has_mat As Boolean = False
                                                    Dim Has_sta1 As Boolean = False
                                                    Dim Has_sta2 As Boolean = False
                                                    Dim Has_len As Boolean = False
                                                    Dim Has_station1 As Boolean = False
                                                    Dim Has_station2 As Boolean = False
                                                    Dim Has_station3 As Boolean = False
                                                    Dim Has_station4 As Boolean = False

                                                    For Each id In attColl
                                                        Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)

                                                        If attref.Tag = ComboBox_mc_MAT.Text Then
                                                            Has_mat = True
                                                        End If

                                                        If attref.Tag = ComboBox_mc_STA1.Text Then
                                                            Has_sta1 = True
                                                        End If

                                                        If attref.Tag = ComboBox_mc_STA2.Text Then
                                                            Has_sta2 = True
                                                        End If

                                                        If attref.Tag = ComboBox_mc_LEN.Text Then
                                                            Has_len = True
                                                        End If
                                                        If attref.Tag = ComboBox_atr_count1.Text And Block_name = ComboBox_blocks_count1.Text Then
                                                            Has_station1 = True
                                                        End If
                                                        If attref.Tag = ComboBox_atr_count2.Text And Block_name = ComboBox_blocks_count2.Text Then
                                                            Has_station2 = True
                                                        End If
                                                        If attref.Tag = ComboBox_atr_count3.Text And Block_name = ComboBox_blocks_count3.Text Then
                                                            Has_station3 = True
                                                        End If
                                                        If attref.Tag = ComboBox_atr_count4.Text And Block_name = ComboBox_blocks_count4.Text Then
                                                            Has_station4 = True
                                                        End If
                                                    Next

                                                    If Has_mat = True And Has_sta1 = True And Has_sta2 = True And Has_len = True Then


                                                        For Each id In attColl
                                                            Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)

                                                            If attref.Tag = ComboBox_mc_MAT.Text Then
                                                                Material = attref.TextString
                                                            End If

                                                            If attref.Tag = ComboBox_mc_STA1.Text Then
                                                                Dim Continut As String = attref.TextString
                                                                If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                                    Station1 = Round(CDbl(Replace(Continut, "+", "")), Round1)
                                                                End If
                                                            End If

                                                            If attref.Tag = ComboBox_mc_STA2.Text Then
                                                                Dim Continut As String = attref.TextString
                                                                If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                                    Station2 = Round(CDbl(Replace(Continut, "+", "")), Round1)
                                                                End If
                                                            End If

                                                            If attref.Tag = ComboBox_mc_LEN.Text Then
                                                                Dim Continut As String = Replace(attref.TextString, "'", "")
                                                                If IsNumeric(Replace(Continut, "'", "")) = True Then
                                                                    Length1 = Round(CDbl(Continut), Round1)
                                                                End If

                                                            End If
                                                        Next

                                                        Dim len_calc As Double = Abs(Station1 - Station2)

                                                        If CheckBox_use_len.Checked = True Then
                                                            len_calc = Length1
                                                        End If

                                                        Data_table_material_count.Rows.Add()
                                                        Data_table_material_count.Rows(Index_data_table).Item("STA1") = Station1
                                                        Data_table_material_count.Rows(Index_data_table).Item("STA2") = Station2
                                                        Data_table_material_count.Rows(Index_data_table).Item("MAT") = Material
                                                        Data_table_material_count.Rows(Index_data_table).Item("AMOUNT_IN_BLOCK") = Length1
                                                        Data_table_material_count.Rows(Index_data_table).Item("CALCULATED_AMOUNT") = len_calc
                                                        If CheckBox_fix_length_display_value.Checked = True Then
                                                            If Not Length1 = len_calc Then
                                                                Block1.UpgradeOpen()
                                                                For Each id In attColl
                                                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForWrite)
                                                                    If attref.Tag = ComboBox_mc_LEN.Text Then
                                                                        attref.TextString = Get_String_Rounded(len_calc, Round1) & "'"
                                                                        Exit For
                                                                    End If
                                                                Next
                                                            End If
                                                        End If
                                                        Index_data_table = Index_data_table + 1

                                                    End If

                                                    If Has_station1 = True Then
                                                        For Each id In attColl
                                                            Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                                            If attref.Tag = ComboBox_atr_count1.Text Then
                                                                Dim Continut As String = attref.TextString
                                                                If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                                    Station1 = Round(CDbl(Replace(Continut, "+", "")), Round1)
                                                                End If
                                                            End If
                                                            If attref.Tag = ComboBox_atr_count1_MAT.Text Then
                                                                Material = attref.TextString
                                                            End If
                                                        Next

                                                        If Not Station1 = 0 Then
                                                            Data_table_material_count.Rows.Add()
                                                            Data_table_material_count.Rows(Index_data_table).Item("STA1") = Station1
                                                            Data_table_material_count.Rows(Index_data_table).Item("STA2") = Station1
                                                            Data_table_material_count.Rows(Index_data_table).Item("MAT") = Material
                                                            Data_table_material_count.Rows(Index_data_table).Item("AMOUNT_IN_BLOCK") = 1
                                                            Data_table_material_count.Rows(Index_data_table).Item("CALCULATED_AMOUNT") = 1
                                                            Index_data_table = Index_data_table + 1
                                                        End If

                                                    End If


                                                    If Has_station2 = True Then
                                                        For Each id In attColl
                                                            Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                                            If attref.Tag = ComboBox_atr_count2.Text Then
                                                                Dim Continut As String = attref.TextString
                                                                If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                                    Station1 = Round(CDbl(Replace(Continut, "+", "")), Round1)
                                                                End If
                                                            End If
                                                        Next

                                                        If Not Station1 = 0 Then
                                                            Data_table_material_count.Rows.Add()
                                                            Data_table_material_count.Rows(Index_data_table).Item("STA1") = Station1
                                                            Data_table_material_count.Rows(Index_data_table).Item("STA2") = Station1
                                                            Data_table_material_count.Rows(Index_data_table).Item("MAT") = Block_name
                                                            Data_table_material_count.Rows(Index_data_table).Item("AMOUNT_IN_BLOCK") = 1
                                                            Data_table_material_count.Rows(Index_data_table).Item("CALCULATED_AMOUNT") = 1
                                                            Index_data_table = Index_data_table + 1
                                                        End If

                                                    End If



                                                    If Has_station3 = True Then
                                                        For Each id In attColl
                                                            Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                                            If attref.Tag = ComboBox_atr_count3.Text Then
                                                                Dim Continut As String = attref.TextString
                                                                If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                                    Station1 = Round(CDbl(Replace(Continut, "+", "")), Round1)
                                                                End If
                                                            End If
                                                        Next

                                                        If Not Station1 = 0 Then
                                                            Data_table_material_count.Rows.Add()
                                                            Data_table_material_count.Rows(Index_data_table).Item("STA1") = Station1
                                                            Data_table_material_count.Rows(Index_data_table).Item("STA2") = Station1
                                                            Data_table_material_count.Rows(Index_data_table).Item("MAT") = Block_name
                                                            Data_table_material_count.Rows(Index_data_table).Item("AMOUNT_IN_BLOCK") = 1
                                                            Data_table_material_count.Rows(Index_data_table).Item("CALCULATED_AMOUNT") = 1
                                                            Index_data_table = Index_data_table + 1
                                                        End If

                                                    End If



                                                    If Has_station4 = True Then
                                                        For Each id In attColl
                                                            Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                                            If attref.Tag = ComboBox_atr_count4.Text Then
                                                                Dim Continut As String = attref.TextString
                                                                If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                                    Station1 = Round(CDbl(Replace(Continut, "+", "")), Round1)
                                                                End If
                                                            End If
                                                        Next

                                                        If Not Station1 = 0 Then
                                                            Data_table_material_count.Rows.Add()
                                                            Data_table_material_count.Rows(Index_data_table).Item("STA1") = Station1
                                                            Data_table_material_count.Rows(Index_data_table).Item("STA2") = Station1
                                                            Data_table_material_count.Rows(Index_data_table).Item("MAT") = Block_name
                                                            Data_table_material_count.Rows(Index_data_table).Item("AMOUNT_IN_BLOCK") = 1
                                                            Data_table_material_count.Rows(Index_data_table).Item("CALCULATED_AMOUNT") = 1
                                                            Index_data_table = Index_data_table + 1
                                                        End If

                                                    End If


                                                End If

                                            End If


                                        Next
                                        Trans1.Commit()

                                        If IsNothing(Data_table_material_count) = False Then

                                            Data_table_material_count = Sort_data_table(Data_table_material_count, "STA1")
                                            DataGridView_components.DataSource = Data_table_material_count
                                            Add_to_clipboard_Data_table(Data_table_material_count)

                                        End If
                                    End If
                                End If
                            End If
                        End If

                    End Using



                    Editor1.SetImpliedSelection(Empty_array)
                    Editor1.WriteMessage(vbLf & "Command:")

                Catch ex As Exception
                    Editor1.SetImpliedSelection(Empty_array)
                    Editor1.WriteMessage(vbLf & "Command:")
                    Freeze_operations = False
                    MsgBox(ex.Message)
                End Try

            End Using
            Freeze_operations = False

            Button_calculate_totals_Click(sender, e)

        End If
    End Sub

    Private Sub Button_clear_table_Click(sender As Object, e As EventArgs) Handles Button_clear_table.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                If IsNothing(Data_table_material_count) = False Then
                    Data_table_material_count.Clear()
                End If
                If IsNothing(Data_table_material_results) = False Then
                    Data_table_material_results.Clear()
                End If
            Catch ex As System.SystemException
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_calculate_totals_Click(sender As Object, e As EventArgs) Handles Button_calculate_totals.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Dim Round1 As Integer = 0
            If IsNumeric(TextBox_dec1.Text) = True Then
                Round1 = CInt(TextBox_dec1.Text)
            End If
            Try
                If IsNothing(Data_table_material_count) = False Then


                    Data_table_material_results = New System.Data.DataTable
                    Data_table_material_results.Columns.Add("MATERIAL", GetType(String))
                    Data_table_material_results.Columns.Add("QTY", GetType(Double))

                    Dim Material_list As System.Collections.Generic.IEnumerable(Of String) = Data_table_material_count.AsEnumerable().Select(Function(x) x.Field(Of String)("MAT")).Distinct()
                    Dim Quantity As Double = 0
                    For Each Material As String In Material_list
                        Quantity = Data_table_material_count.AsEnumerable().Where(Function(x) x.Field(Of String)("MAT") = Material).Sum(Function(x) x.Field(Of Double)("CALCULATED_AMOUNT"))
                        Data_table_material_results.Rows.Add(Material, Quantity)
                    Next

                    Data_table_material_results = Sort_data_table(Data_table_material_results, "MATERIAL")

                    Dim Data_table_material_results_numbers As New System.Data.DataTable
                    Data_table_material_results_numbers.Columns.Add("MATERIAL", GetType(Integer))
                    Data_table_material_results_numbers.Columns.Add("QTY", GetType(Double))

                    Dim Index1 As Integer = 0
                    For i = 0 To Data_table_material_results.Rows.Count - 1
                        If IsDBNull(Data_table_material_results.Rows(i).Item(0)) = False Then
                            Dim Mat As String = Data_table_material_results.Rows(i).Item(0)
                            If IsNumeric(Mat) = True Then
                                Data_table_material_results_numbers.Rows.Add()
                                Data_table_material_results_numbers.Rows(Index1).Item(0) = CInt(Mat)
                                If IsDBNull(Data_table_material_results.Rows(i).Item(1)) = False Then
                                    Dim qty As Double = Data_table_material_results.Rows(i).Item(1)
                                    Data_table_material_results_numbers.Rows(Index1).Item(1) = qty
                                End If
                                Index1 = Index1 + 1
                            End If
                        End If
                    Next

                    If Data_table_material_results_numbers.Rows.Count > 0 Then
                        Data_table_material_results_numbers = Sort_data_table(Data_table_material_results_numbers, "MATERIAL")
                        For i = 0 To Data_table_material_results_numbers.Rows.Count - 1
                            If IsDBNull(Data_table_material_results_numbers.Rows(i).Item(0)) = False Then
                                Dim Mat As Integer = Data_table_material_results_numbers.Rows(i).Item(0)
                                Data_table_material_results.Rows(i).Item(0) = Mat.ToString
                                If IsDBNull(Data_table_material_results_numbers.Rows(i).Item(1)) = False Then
                                    Dim qty As Double = Data_table_material_results_numbers.Rows(i).Item(1)
                                    Data_table_material_results.Rows(i).Item(1) = qty
                                End If
                            End If
                        Next
                    End If


                    DataGridView_results.DataSource = Data_table_material_results
                End If
            Catch ex As System.SystemException
                MsgBox(ex.Message)
            End Try

            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_transfer_to_excel_Click(sender As Object, e As EventArgs) Handles Button_transfer_to_excel.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Dim Round1 As Integer = 0
            If IsNumeric(TextBox_dec1.Text) = True Then
                Round1 = CInt(TextBox_dec1.Text)
            End If
            Try
                If IsNothing(Data_table_material_results) = False Then
                    If Data_table_material_results.Rows.Count > 0 Then
                        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                        W1 = Get_active_worksheet_from_Excel_with_error()
                        Dim Row1 As Integer = 0
                        Dim Band1 As String = ""

                        If IsNumeric(TextBox_current_row_excel.Text) = True Then
                            Row1 = CInt(TextBox_current_row_excel.Text)
                            If Row1 < 1 Then
                                Freeze_operations = False
                                MsgBox("No excel row specified")
                                Exit Sub
                            End If
                        Else
                            Freeze_operations = False
                            MsgBox("No excel row specified")
                            Exit Sub
                        End If

                        If Not TextBox_band_name.Text = "" Then
                            Band1 = TextBox_band_name.Text
                        End If

                        Dim Column_band As String = "A"


                        If Row1 = 1 Then
                            W1.Range(Column_band & Row1).Value2 = "Band number"
                            Row1 = 2
                        End If

                        Dim Number_of_columns As Integer = Data_table_material_results.Rows.Count

                        W1.Rows(Row1).ClearContents()


                        Dim Col_start As Integer = 2

                        For i = 0 To Data_table_material_results.Rows.Count - 1

                            If IsDBNull(Data_table_material_results.Rows(i).Item(0)) = False Then



                                Dim Mat As String = Data_table_material_results.Rows(i).Item(0)
                                Dim Mat_no As Integer = 0
                                If IsNumeric(Mat) = True Then Mat_no = CInt(Mat)


                                Dim Inserted As Boolean = False

                                For j = Col_start To 100
                                    Dim Val_excel_mat_j As String = W1.Cells(1, j).value2
                                    If Not Val_excel_mat_j = "" Then

                                        If IsNumeric(Val_excel_mat_j) = True Then
                                            If Mat_no = CInt(Val_excel_mat_j) Then
                                                If IsDBNull(Data_table_material_results.Rows(i).Item(1)) = False Then
                                                    W1.Cells(Row1, j).Value2 = Data_table_material_results.Rows(i).Item(1)
                                                End If
                                                Inserted = True
                                                Exit For
                                            End If
                                        Else
                                            If Mat = Val_excel_mat_j Then
                                                If IsDBNull(Data_table_material_results.Rows(i).Item(1)) = False Then
                                                    W1.Cells(Row1, j).Value2 = Data_table_material_results.Rows(i).Item(1)
                                                End If
                                                Inserted = True
                                                Exit For
                                            End If
                                        End If

                                    Else
                                        If Inserted = False Then

                                            W1.Cells(1, j).Value2 = Mat
                                            W1.Cells(Row1, j).Value2 = Data_table_material_results.Rows(i).Item(1)
                                        End If

                                        Exit For
                                    End If

                                Next

                                W1.Cells(Row1, 1).Value2 = Band1


                            End If


                        Next






                        If IsNumeric(Band1) = True Then
                            TextBox_band_name.Text = (CInt(Band1) + 1).ToString
                        End If

                        Row1 = Row1 + 1
                        TextBox_current_row_excel.Text = Row1.ToString
                    End If
                End If

            Catch ex As System.SystemException
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try

            Freeze_operations = False
            If CheckBox_clear_after_transfer.Checked = True Then
                Button_clear_table_Click(sender, e)
            End If

        End If
    End Sub

    Private Sub Button_add_to_blocks_Click(sender As Object, e As EventArgs) Handles Button_add_to_blocks.Click
        If Freeze_operations = False Then
            Freeze_operations = True


            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

            Dim Empty_array() As ObjectId

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Editor1.SetImpliedSelection(Empty_array)
            Try
                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select block:"

                Object_Prompt.SingleOnly = True

                Rezultat1 = Editor1.GetSelection(Object_Prompt)


                If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.WriteMessage(vbLf & "Command:")
                    Freeze_operations = False
                    Exit Sub
                End If



                Dim Start1 As Integer = 1
                If IsNumeric(TextBox_transfer_row_start.Text) = True Then
                    Start1 = Abs(CInt(TextBox_transfer_row_start.Text))
                End If
                Dim End1 As Integer = 0
                If IsNumeric(TextBox_transfer_row_end.Text) = True Then
                    End1 = Abs(CInt(TextBox_transfer_row_end.Text))
                End If

                If End1 < Start1 Then

                    Editor1.WriteMessage(vbLf & "Command:")
                    Freeze_operations = False
                    Exit Sub

                End If




                Dim Column_start As String = TextBox_block_att_value_column_start.Text.ToUpper
                Dim Column_end As String = TextBox_block_att_value_column_end.Text.ToUpper
                Dim Column_atr_name As String = TextBox_block_att_name_column.Text.ToUpper


                Dim Col_start As Integer = 0
                Dim Col_end As Integer = 0
                Dim Col_atr As Integer = 0

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Nothing


                Try
                    W1 = Get_active_worksheet_from_Excel_with_error()
                    If IsNothing(W1) = False Then
                        Col_start = W1.Range(Column_start & "1").Column
                        Col_end = W1.Range(Column_end & "1").Column
                        Col_atr = W1.Range(Column_atr_name & "1").Column
                    End If

                Catch ex As System.SystemException

                    Editor1.WriteMessage(vbLf & "Command:")
                    Freeze_operations = False
                    Exit Sub

                End Try

                Dim Row_with_layout_names As Integer = 0
                If IsNumeric(TextBox_ROW_layout_NAME.Text) = True Then
                    Row_with_layout_names = CInt(TextBox_ROW_layout_NAME.Text)
                End If


                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                    If IsNothing(Rezultat1) = False Then
                        Using lock1 As DocumentLock = ThisDrawing.LockDocument
                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                                Dim Layoutdict As DBDictionary = Trans1.GetObject(ThisDrawing.Database.LayoutDictionaryId, OpenMode.ForRead)
                                Dim Ent1 As Entity
                                Ent1 = Trans1.GetObject(Rezultat1.Value.Item(0).ObjectId, OpenMode.ForRead)
                                Dim Block_name As String = ""



                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.BlockReference Then
                                    Dim Block1 As BlockReference = TryCast(Ent1, BlockReference)
                                    If IsNothing(Block1) = False Then
                                        If Block1.AttributeCollection.Count > 0 Then
                                            Dim BlockTrec As BlockTableRecord = Nothing
                                            If Block1.IsDynamicBlock = True Then
                                                BlockTrec = Trans1.GetObject(Block1.DynamicBlockTableRecord, OpenMode.ForRead)
                                                Block_name = BlockTrec.Name
                                            Else
                                                BlockTrec = Trans1.GetObject(Block1.BlockTableRecord, OpenMode.ForRead)
                                                Block_name = BlockTrec.Name
                                            End If
                                        End If
                                    End If

                                End If

                                Dim Data_table1 As New System.Data.DataTable
                                Data_table1.Columns.Add("ATR", GetType(String))
                                Data_table1.Columns.Add("VALUE", GetType(String))
                                Data_table1.Columns.Add("LAYOUT", GetType(String))
                                Dim Index1 As Integer = 0

                                For j = Col_start To Col_end
                                    Dim Excel_layout As String = W1.Cells(Row_with_layout_names, j).Value2





                                    For i = Start1 To End1
                                        Dim Excel_attribute_name As String = W1.Range(Column_atr_name & i.ToString).Value2
                                        Dim Excel_value As String = W1.Cells(i, j).Value2
                                        If Not Excel_attribute_name = "" And Not Excel_layout = "" Then
                                            Data_table1.Rows.Add()
                                            Data_table1.Rows(Index1).Item("ATR") = Excel_attribute_name
                                            Data_table1.Rows(Index1).Item("VALUE") = Excel_value
                                            Data_table1.Rows(Index1).Item("LAYOUT") = Excel_layout
                                            Index1 = Index1 + 1
                                        End If


                                    Next
                                Next

                                If Data_table1.Rows.Count > 0 Then
                                    Dim Lay_name As String = ""
                                    For i = 0 To Data_table1.Rows.Count - 1
                                        If IsDBNull(Data_table1.Rows(i).Item("LAYOUT")) = False Then
                                            If Not Lay_name = Data_table1.Rows(i).Item("LAYOUT") Then
                                                Lay_name = Data_table1.Rows(i).Item("LAYOUT")

                                                For Each entry As DBDictionaryEntry In Layoutdict
                                                    Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)



                                                    If Not Layout1.TabOrder = 0 And Layout1.LayoutName = Lay_name Then
                                                        LayoutManager1.CurrentLayout = Lay_name

                                                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(Layout1.BlockTableRecordId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                                                        For Each obid As ObjectId In BTrecord
                                                            Dim Block1 As BlockReference = TryCast(Trans1.GetObject(obid, OpenMode.ForRead), BlockReference)
                                                            If IsNothing(Block1) = False Then
                                                                If Block1.Name = Block_name Then
                                                                    If Block1.AttributeCollection.Count > 0 Then

                                                                        For Each id As ObjectId In Block1.AttributeCollection
                                                                            If Not id.IsErased Then
                                                                                Dim attRef As AttributeReference = DirectCast(Trans1.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite), AttributeReference)

                                                                                Dim Tag As String = attRef.Tag
                                                                                For s = 0 To Data_table1.Rows.Count - 1
                                                                                    If IsDBNull(Data_table1.Rows(s).Item("ATR")) = False And IsDBNull(Data_table1.Rows(s).Item("LAYOUT")) = False Then
                                                                                        If Data_table1.Rows(s).Item("LAYOUT") = Lay_name Then
                                                                                            Dim aTR As String = Data_table1.Rows(s).Item("ATR")
                                                                                            Dim Value1 As String = ""

                                                                                            If IsDBNull(Data_table1.Rows(s).Item("VALUE")) = False Then
                                                                                                Value1 = Data_table1.Rows(s).Item("VALUE")
                                                                                            End If


                                                                                            If Tag.ToUpper = aTR.ToUpper Then
                                                                                                If attRef.IsMTextAttribute = False Then
                                                                                                    attRef.TextString = Value1
                                                                                                Else
                                                                                                    attRef.MTextAttribute.Contents = Value1
                                                                                                End If

                                                                                            End If

                                                                                        End If
                                                                                    End If
                                                                                Next
                                                                            End If
                                                                        Next
                                                                    End If
                                                                End If
                                                            End If
                                                        Next
                                                    End If
                                                Next
                                            End If
                                        End If
                                    Next
                                End If













                                Trans1.Commit()

                            End Using
                        End Using
                    End If
                End If



                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
            Catch ex As Exception
                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
            End Try


            Freeze_operations = False
        End If



    End Sub

    Private Sub Button_write_to_excel_Click(sender As Object, e As EventArgs) Handles Button_write_to_excel.Click
        If Freeze_operations = False Then
            Freeze_operations = True

            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
            W1 = Get_active_worksheet_from_Excel()
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

            Dim Empty_array() As ObjectId

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Editor1.SetImpliedSelection(Empty_array)
            Try
                Try
                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                    Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt.MessageForAdding = vbLf & "Select blocks:"

                    Object_Prompt.SingleOnly = False

                    Rezultat1 = Editor1.GetSelection(Object_Prompt)


                    If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Editor1.WriteMessage(vbLf & "Command:")
                        Freeze_operations = False
                        Exit Sub
                    End If



                    Dim Row1 As Integer = 1
                    If IsNumeric(TextBox_transfer_START_ROW.Text) = True Then
                        Row1 = Abs(CInt(TextBox_transfer_START_ROW.Text))
                    End If
                    Dim Column_atr_name As String
                    Column_atr_name = TextBox_COL_ATR_NAME.Text
                    Dim Column_atr_value As String
                    Column_atr_value = TextBox_COL_ATR_VALUE.Text

                    If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        If IsNothing(Rezultat1) = False Then
                            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                                    For i = 0 To Rezultat1.Value.Count - 1



                                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                        Obj1 = Rezultat1.Value.Item(i)
                                        Dim Ent1 As Entity
                                        Ent1 = Trans1.GetObject(Obj1.ObjectId, OpenMode.ForRead)

                                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.BlockReference Then
                                            Dim Block1 As BlockReference = TryCast(Ent1, BlockReference)
                                            If Block1.AttributeCollection.Count > 0 Then
                                                W1.Range(Column_atr_name & Row1).Value2 = Block1.Name
                                                Row1 = Row1 + 1
                                                For Each id As ObjectId In Block1.AttributeCollection
                                                    If Not id.IsErased Then
                                                        Dim attRef As AttributeReference = DirectCast(Trans1.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), AttributeReference)
                                                        Dim Continut As String = attRef.TextString
                                                        Dim Tag As String = attRef.Tag
                                                        W1.Range(Column_atr_value & Row1).Value2 = Continut
                                                        W1.Range(Column_atr_name & Row1).Value2 = Tag
                                                        Row1 = Row1 + 1
                                                    End If
                                                Next


                                            End If


                                        End If



                                    Next
                                    Trans1.Commit()

                                End Using
                            End Using
                        End If
                    End If



                    Editor1.SetImpliedSelection(Empty_array)
                    Editor1.WriteMessage(vbLf & "Command:")
                Catch ex As Exception
                    Editor1.SetImpliedSelection(Empty_array)
                    Editor1.WriteMessage(vbLf & "Command:")
                    MsgBox(ex.Message)
                End Try
            Catch EX As System.Runtime.InteropServices.COMException
                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
                MsgBox(EX.Message)
            End Try

            Freeze_operations = False
        End If

    End Sub

    Private Sub Button_create_atribute_list_in_excel_Click(sender As Object, e As EventArgs) Handles Button_create_atribute_list_in_excel.Click
        If Freeze_operations = False Then
            Freeze_operations = True

            Try
                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Dim Column_band_number As String = TextBox_column_band_number.Text.ToUpper
                Dim Column_material_start As String = TextBox_column_mat_start.Text.ToUpper
                Dim Column_material_end As String = TextBox_column_mat_end.Text.ToUpper
                Dim Row_mat_start As Integer = CInt(TextBox_row_mat_count_start.Text)
                Dim Row_mat_end As Integer = CInt(TextBox_row_mat_count_end.Text)
                Dim Column_legend_mat_number As String = TextBox_column_mat_number.Text.ToUpper
                Dim Column_legend_mat_description As String = TextBox_column_mat_description.Text.ToUpper
                Dim Column_legend_mat_SUFFIX As String = TextBox_count_suffix_column.Text.ToUpper

                Dim Row_legend_start As Integer = CInt(TextBox_row_legend_start.Text)
                Dim Row_legend_end As Integer = CInt(TextBox_row_legend_end.Text)
                Dim Col_start As Integer = W1.Range(Column_material_start & "1").Column
                Dim Col_end As Integer = W1.Range(Column_material_end & "1").Column
                Dim First_cell_address As String = TextBox_first_cell_address.Text
                Dim Col_wr As Integer = W1.Range(First_cell_address).Column
                Dim Row_wr As Integer = W1.Range(First_cell_address).Row + 2

                For i = Row_mat_start + 1 To Row_mat_end
                    For j = Col_start To Col_end
                        Dim Qty_per_sheet As String = W1.Cells(i, j).value2
                        Dim Material_number As String = W1.Cells(Row_mat_start, j).value2
                        Dim Suffix1 As String = ""
                        If Not Qty_per_sheet = "" Then
                            For k = Row_legend_start To Row_legend_end
                                Dim Material_number_legend As String = W1.Range(Column_legend_mat_number & k).Value2
                                Dim Material_description_legend As String = W1.Range(Column_legend_mat_description & k).Value2


                                If Material_number = Material_number_legend Then
                                    Suffix1 = W1.Range(Column_legend_mat_SUFFIX & k).Value2
                                    If IsNumeric(Material_number) = True Then
                                        W1.Cells(Row_wr - 2, Col_wr).value2 = Material_number
                                        W1.Cells(Row_wr - 1, Col_wr).value2 = Material_description_legend
                                    Else
                                        W1.Cells(Row_wr - 1, Col_wr).value2 = Material_description_legend
                                    End If
                                End If
                            Next

                            W1.Cells(Row_wr, Col_wr).value2 = Qty_per_sheet & Suffix1
                            Row_wr = Row_wr + 3
                        End If
                    Next
                    Col_wr = Col_wr + 1
                    Row_wr = W1.Range(First_cell_address).Row + 2
                Next

            Catch ex As SystemException
                MsgBox(ex.Message)
            End Try

            Freeze_operations = False
        End If


    End Sub

    Private Sub Button_embed_cover_in_excel_Click(sender As Object, e As EventArgs) Handles Button_embed_cover_in_excel.Click
        Try
            Dim Start1 As Integer = CInt(TextBox_cover_row_start.Text)
            Dim End1 As Integer = CInt(TextBox_cover_row_end.Text)
            Dim Col_Sta1 = TextBox_col_sta1.Text.ToUpper
            Dim Col_Sta2 = TextBox_col_sta2.Text.ToUpper
            Dim Col_Sta12 = TextBox_col_sta12.Text.ToUpper
            Dim Col_Sta22 = TextBox_col_sta22.Text.ToUpper
            Dim Col_cvr1 = TextBox_col_cover1.Text.ToUpper
            Dim Col_cvr21 = TextBox_col_cover21.Text.ToUpper
            Dim Col_descr1 = TextBox_col_descr1.Text.ToUpper
            Dim Col_descr21 = TextBox_col_descr21.Text.ToUpper

            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
            W1 = Get_active_worksheet_from_Excel()

            Dim def_cvr As Integer = 3
            Dim max_L As Double = 310560.8
            Dim List1 As New List(Of Integer)
            List1.Add(def_cvr)
            Dim List2 As New List(Of Integer)
            List2.Add(def_cvr)

            Dim idt As Integer = 0
            Dim DT1(idt) As System.Data.DataTable
            DT1(idt) = New System.Data.DataTable
            DT1(idt).Columns.Add("STA1", GetType(Double))
            DT1(idt).Columns.Add("STA2", GetType(Double))
            DT1(idt).Columns.Add("CVR", GetType(Double))
            DT1(idt).Columns.Add("DESCRIPTION", GetType(String))
            DT1(idt).Rows.Add()
            DT1(idt).Rows(0).Item("STA1") = 0
            DT1(idt).Rows(0).Item("STA2") = max_L
            DT1(idt).Rows(0).Item("CVR") = def_cvr
            DT1(idt).Rows(0).Item("DESCRIPTION") = "default"
            idt = idt + 1

            For ii = Start1 To End1
                Dim Sta1 As Double = W1.Range(Col_Sta1 & ii).Value2
                Dim Sta2 As Double = W1.Range(Col_Sta2 & ii).Value2
                Dim Cvr1 As Double = W1.Range(Col_cvr1 & ii).Value2
                Dim Descr1 As String = W1.Range(Col_descr1 & ii).Value2
                If List1.Contains(Cvr1) = False Then
                    List1.Add(Cvr1)
                    List2.Add(Cvr1)
                    ReDim Preserve DT1(idt)
                    DT1(idt) = New System.Data.DataTable
                    DT1(idt).Columns.Add("STA1", GetType(Double))
                    DT1(idt).Columns.Add("STA2", GetType(Double))
                    DT1(idt).Columns.Add("CVR", GetType(Double))
                    DT1(idt).Columns.Add("DESCRIPTION", GetType(String))
                    idt = idt + 1
                End If

                Dim Idx As Integer = List1.IndexOf(Cvr1)

                DT1(Idx).Rows.Add()
                DT1(Idx).Rows(DT1(Idx).Rows.Count - 1).Item("STA1") = Sta1
                DT1(Idx).Rows(DT1(Idx).Rows.Count - 1).Item("STA2") = Sta2
                DT1(Idx).Rows(DT1(Idx).Rows.Count - 1).Item("CVR") = Cvr1
                DT1(Idx).Rows(DT1(Idx).Rows.Count - 1).Item("DESCRIPTION") = Descr1
            Next

            Dim DT2 = New System.Data.DataTable
            DT2.Columns.Add("STA1", GetType(Double))
            DT2.Columns.Add("STA2", GetType(Double))
            DT2.Columns.Add("CVR", GetType(Double))
            DT2.Columns.Add("DESCRIPTION", GetType(String))

            List2.Sort()



            For Index_list = 0 To List2.Count - 1
                Dim Idx As Integer = List1.IndexOf(List2(Index_list))
                For j = 0 To DT1(Idx).Rows.Count - 1

                    Dim Sta1 As Double = DT1(Idx).Rows(j).Item("STA1")
                    Dim Sta2 As Double = DT1(Idx).Rows(j).Item("STA2")
                    Dim Cvr1 As Double = DT1(Idx).Rows(j).Item("CVR")
                    Dim Descr1 As String = DT1(Idx).Rows(j).Item("DESCRIPTION")

                    If DT2.Rows.Count > 0 Then
                        Dim Rc As Integer = DT2.Rows.Count
                        For i = 0 To Rc - 1
                            Dim Sta01 As Double = DT2.Rows(i).Item("STA1")
                            Dim Sta02 As Double = DT2.Rows(i).Item("STA2")
                            Dim Cvr01 As Double = DT2.Rows(i).Item("CVR")
                            Dim Descr01 As String = DT2.Rows(i).Item("DESCRIPTION")

                            If Cvr1 > Cvr01 Then
                                If Sta1 > Sta01 And Sta2 < Sta02 Then
                                    DT2.Rows(i).Item("STA2") = Sta1
                                    Dim row1 As System.Data.DataRow
                                    row1 = DT2.NewRow()
                                    row1("STA1") = Sta1
                                    row1("STA2") = Sta2
                                    row1("CVR") = Cvr1
                                    row1("DESCRIPTION") = Descr1
                                    DT2.Rows.InsertAt(row1, i + 1)
                                    row1 = DT2.NewRow()
                                    row1("STA1") = Sta2
                                    row1("STA2") = Sta02
                                    row1("CVR") = Cvr01
                                    row1("DESCRIPTION") = Descr01
                                    DT2.Rows.InsertAt(row1, i + 2)
                                    Rc = Rc + 2
                                    Exit For
                                End If
                                If Sta1 > Sta01 And Sta2 >= Sta02 And Sta1 <= Sta02 Then
                                    DT2.Rows(i).Item("STA2") = Sta1
                                    Sta02 = Sta1
                                    Dim row1 As System.Data.DataRow
                                    row1 = DT2.NewRow()
                                    row1("STA1") = Sta1
                                    row1("STA2") = Sta2
                                    row1("CVR") = Cvr1
                                    row1("DESCRIPTION") = Descr1
                                    DT2.Rows.InsertAt(row1, i + 1)
                                    Rc = Rc + 1
                                    For k = DT2.Rows.Count - 1 To i + 1 Step -1
                                        Dim Stax1 As Double = DT2.Rows(k).Item("STA1")
                                        Dim Stax2 As Double = DT2.Rows(k).Item("STA2")
                                        Dim Cvrx As Double = DT2.Rows(k).Item("CVR")
                                        If Stax1 > Sta1 And Stax2 < Sta2 Then
                                            If Cvrx <= Cvr1 Then
                                                DT2.Rows(k).Delete()
                                                Rc = Rc - 1
                                            End If
                                        End If

                                        If Stax1 < Sta2 And Stax2 > Sta2 And Stax1 > Sta1 Then
                                            DT2.Rows(k).Item("STA1") = Sta2

                                        End If

                                    Next
                                    Exit For


                                End If
                            End If

                            If Cvr1 = Cvr01 Then
                                If Sta1 > Sta01 And Sta2 > Sta02 And Sta1 < Sta02 Then
                                    DT2.Rows(i).Item("STA2") = Sta2
                                    If i < DT2.Rows.Count - 1 Then
                                        DT2.Rows(i + 1).Item("STA1") = Sta2
                                    End If

                                End If


                            End If


                        Next

                    Else
                        DT2.Rows.Add()
                        DT2.Rows(0).Item("STA1") = Sta1
                        DT2.Rows(0).Item("STA2") = Sta2
                        DT2.Rows(0).Item("CVR") = Cvr1
                        DT2.Rows(0).Item("DESCRIPTION") = Descr1
                    End If





                Next



            Next

            For i = 0 To DT2.Rows.Count - 1
                W1.Range(Col_Sta12 & i + 2).Value2 = DT2.Rows(i).Item("STA1")
                W1.Range(Col_Sta22 & i + 2).Value2 = DT2.Rows(i).Item("STA2")
                W1.Range(Col_cvr21 & i + 2).Value2 = DT2.Rows(i).Item("CVR")
                W1.Range(Col_descr21 & i + 2).Value2 = DT2.Rows(i).Item("DESCRIPTION")
            Next

            MsgBox("DONE")

        Catch ex As System.Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_embed_design_in_excel_Click(sender As Object, e As EventArgs) Handles Button_embed_design_in_excel.Click

        Try
            Dim Start1 As Integer = CInt(TextBox_cover_row_start.Text)
            Dim End1 As Integer = CInt(TextBox_cover_row_end.Text)
            Dim Col_Sta1 = TextBox_col_sta1.Text.ToUpper
            Dim Col_Sta2 = TextBox_col_sta2.Text.ToUpper
            Dim Col_Sta12 = TextBox_col_sta12.Text.ToUpper
            Dim Col_Sta22 = TextBox_col_sta22.Text.ToUpper
            Dim Col_design1 = TextBox_col_cover1.Text.ToUpper
            Dim Col_design21 = TextBox_col_cover21.Text.ToUpper
            Dim Col_descr1 = TextBox_col_descr1.Text.ToUpper
            Dim Col_descr21 = TextBox_col_descr21.Text.ToUpper

            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
            W1 = Get_active_worksheet_from_Excel()

            Dim def_design As Double = 0.72
            Dim max_L As Double = 310560.8
            Dim List1 As New List(Of Double)
            List1.Add(def_design)
            Dim List2 As New List(Of Double)
            List2.Add(def_design)

            Dim idt As Integer = 0
            Dim DT1(idt) As System.Data.DataTable
            DT1(idt) = New System.Data.DataTable
            DT1(idt).Columns.Add("STA1", GetType(Double))
            DT1(idt).Columns.Add("STA2", GetType(Double))
            DT1(idt).Columns.Add("DESIGN", GetType(Double))
            DT1(idt).Columns.Add("DESCRIPTION", GetType(String))
            DT1(idt).Rows.Add()
            DT1(idt).Rows(0).Item("STA1") = 0
            DT1(idt).Rows(0).Item("STA2") = max_L
            DT1(idt).Rows(0).Item("DESIGN") = def_design
            DT1(idt).Rows(0).Item("DESCRIPTION") = "default"
            idt = idt + 1

            For ii = Start1 To End1
                Dim Sta1 As Double = W1.Range(Col_Sta1 & ii).Value2
                Dim Sta2 As Double = W1.Range(Col_Sta2 & ii).Value2
                Dim design1 As Double = W1.Range(Col_design1 & ii).Value2
                Dim Descr1 As String = W1.Range(Col_descr1 & ii).Value2
                If List1.Contains(design1) = False Then
                    List1.Add(design1)
                    List2.Add(design1)
                    ReDim Preserve DT1(idt)
                    DT1(idt) = New System.Data.DataTable
                    DT1(idt).Columns.Add("STA1", GetType(Double))
                    DT1(idt).Columns.Add("STA2", GetType(Double))
                    DT1(idt).Columns.Add("DESIGN", GetType(Double))
                    DT1(idt).Columns.Add("DESCRIPTION", GetType(String))
                    idt = idt + 1
                End If

                Dim Idx As Integer = List1.IndexOf(design1)

                DT1(Idx).Rows.Add()
                DT1(Idx).Rows(DT1(Idx).Rows.Count - 1).Item("STA1") = Sta1
                DT1(Idx).Rows(DT1(Idx).Rows.Count - 1).Item("STA2") = Sta2
                DT1(Idx).Rows(DT1(Idx).Rows.Count - 1).Item("DESIGN") = design1
                DT1(Idx).Rows(DT1(Idx).Rows.Count - 1).Item("DESCRIPTION") = Descr1
            Next

            Dim DT2 = New System.Data.DataTable
            DT2.Columns.Add("STA1", GetType(Double))
            DT2.Columns.Add("STA2", GetType(Double))
            DT2.Columns.Add("DESIGN", GetType(Double))
            DT2.Columns.Add("DESCRIPTION", GetType(String))

            List2.Sort()



            For Index_list = List2.Count - 1 To 0 Step -1
                Dim Idx As Integer = List1.IndexOf(List2(Index_list))
                For j = 0 To DT1(Idx).Rows.Count - 1

                    Dim Sta1 As Double = DT1(Idx).Rows(j).Item("STA1")
                    Dim Sta2 As Double = DT1(Idx).Rows(j).Item("STA2")
                    Dim design1 As Double = DT1(Idx).Rows(j).Item("DESIGN")
                    Dim Descr1 As String = DT1(Idx).Rows(j).Item("DESCRIPTION")

                    If DT2.Rows.Count > 0 Then
                        Dim Rc As Integer = DT2.Rows.Count
                        For i = 0 To Rc - 1
                            Dim Sta01 As Double = DT2.Rows(i).Item("STA1")
                            Dim Sta02 As Double = DT2.Rows(i).Item("STA2")
                            Dim design01 As Double = DT2.Rows(i).Item("DESIGN")
                            Dim Descr01 As String = DT2.Rows(i).Item("DESCRIPTION")

                            If design1 < design01 Then
                                If Sta1 > Sta01 And Sta2 < Sta02 Then
                                    DT2.Rows(i).Item("STA2") = Sta1
                                    Dim row1 As System.Data.DataRow
                                    row1 = DT2.NewRow()
                                    row1("STA1") = Sta1
                                    row1("STA2") = Sta2
                                    row1("DESIGN") = design1
                                    row1("DESCRIPTION") = Descr1
                                    DT2.Rows.InsertAt(row1, i + 1)
                                    row1 = DT2.NewRow()
                                    row1("STA1") = Sta2
                                    row1("STA2") = Sta02
                                    row1("DESIGN") = design01
                                    row1("DESCRIPTION") = Descr01
                                    DT2.Rows.InsertAt(row1, i + 2)
                                    Rc = Rc + 2
                                    Exit For
                                End If
                                If Sta1 > Sta01 And Sta2 >= Sta02 And Sta1 <= Sta02 Then
                                    DT2.Rows(i).Item("STA2") = Sta1
                                    Sta02 = Sta1
                                    Dim row1 As System.Data.DataRow
                                    row1 = DT2.NewRow()
                                    row1("STA1") = Sta1
                                    row1("STA2") = Sta2
                                    row1("DESIGN") = design1
                                    row1("DESCRIPTION") = Descr1
                                    DT2.Rows.InsertAt(row1, i + 1)
                                    Rc = Rc + 1
                                    For k = DT2.Rows.Count - 1 To i + 1 Step -1
                                        Dim Stax1 As Double = DT2.Rows(k).Item("STA1")
                                        Dim Stax2 As Double = DT2.Rows(k).Item("STA2")
                                        Dim designx As Double = DT2.Rows(k).Item("DESIGN")
                                        If Stax1 > Sta1 And Stax2 < Sta2 Then
                                            If designx <= design1 Then
                                                DT2.Rows(k).Delete()
                                                Rc = Rc - 1
                                            End If
                                        End If

                                        If Stax1 < Sta2 And Stax2 > Sta2 And Stax1 > Sta1 Then
                                            DT2.Rows(k).Item("STA1") = Sta2

                                        End If

                                    Next
                                    Exit For


                                End If
                            End If

                            If design1 = design01 Then
                                If Sta1 > Sta01 And Sta2 > Sta02 And Sta1 < Sta02 Then
                                    DT2.Rows(i).Item("STA2") = Sta2
                                    If i < DT2.Rows.Count - 1 Then
                                        DT2.Rows(i + 1).Item("STA1") = Sta2
                                    End If

                                End If


                            End If


                        Next

                    Else
                        DT2.Rows.Add()
                        DT2.Rows(0).Item("STA1") = Sta1
                        DT2.Rows(0).Item("STA2") = Sta2
                        DT2.Rows(0).Item("DESIGN") = design1
                        DT2.Rows(0).Item("DESCRIPTION") = Descr1
                    End If





                Next



            Next

            For i = 0 To DT2.Rows.Count - 1
                W1.Range(Col_Sta12 & i + 2).Value2 = DT2.Rows(i).Item("STA1")
                W1.Range(Col_Sta22 & i + 2).Value2 = DT2.Rows(i).Item("STA2")
                W1.Range(Col_design21 & i + 2).Value2 = DT2.Rows(i).Item("DESIGN")
                W1.Range(Col_descr21 & i + 2).Value2 = DT2.Rows(i).Item("DESCRIPTION")
            Next

            MsgBox("DONE")

        Catch ex As System.Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_adjust_bands_Click(sender As Object, e As EventArgs) Handles Button_adjust_bands.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Try

                Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
                ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                Editor1.SetImpliedSelection(Empty_array)

                Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim Result_point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                        Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Please pick the first matchline:")
                        PP1.AllowNone = False
                        Result_point1 = Editor1.GetPoint(PP1)
                        If Result_point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                            Freeze_operations = False
                            Exit Sub
                        End If

                        Dim Result_point2 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                        Dim PP2 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Please pick the second matchline:")
                        PP2.AllowNone = False
                        PP2.UseBasePoint = True
                        PP2.BasePoint = Result_point1.Value

                        Result_point2 = Editor1.GetPoint(PP2)
                        If Result_point2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                            Freeze_operations = False
                            Exit Sub
                        End If



                        Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                        Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_Prompt.MessageForAdding = vbLf & "Select blocks:"

                        Object_Prompt.SingleOnly = False

                        Rezultat1 = Editor1.GetSelection(Object_Prompt)


                        If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Editor1.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If

                        Dim Dt1 As New System.Data.DataTable
                        Dt1.Columns.Add("STA1", GetType(Double))
                        Dt1.Columns.Add("STA2", GetType(Double))
                        Dt1.Columns.Add("NO_LETTER", GetType(Integer))
                        Dt1.Columns.Add("OBJID", GetType(ObjectId))
                        Dt1.Columns.Add("PT_INS", GetType(Point3d))
                        Dt1.Columns.Add("STRETCH", GetType(Double))
                        Dim Ri As Integer = 0

                        For i = 0 To Rezultat1.Value.Count - 1
                            Dim block1 As BlockReference = TryCast(Trans1.GetObject(Rezultat1.Value(i).ObjectId, OpenMode.ForRead), BlockReference)
                            If IsNothing(block1) = False Then
                                If block1.IsDynamicBlock = True Then
                                    Dim Col_atr As AttributeCollection = block1.AttributeCollection
                                    If Col_atr.Count > 0 Then
                                        Dim Sta1 As Double = 0
                                        Dim Sta2 As Double = 0
                                        Dim No_letter As Integer = 0
                                        For Each id In Col_atr
                                            Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                            Dim Continut As String = attref.TextString
                                            Select Case attref.Tag.ToUpper

                                                Case "STA1"
                                                    If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                        Sta1 = CDbl(Replace(Continut, "+", ""))
                                                    End If
                                                Case "STA1"
                                                    If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                        Sta2 = CDbl(Replace(Continut, "+", ""))
                                                    End If
                                                Case Else
                                                    Dim No1 As Integer = Continut.Length
                                                    If No1 > No_letter Then No_letter = No1
                                            End Select
                                        Next

                                        If Not Sta1 = 0 Or Not Sta2 = 0 Then
                                            Dt1.Rows.Add()
                                            Dt1.Rows(Ri).Item("STA1") = Sta1
                                            Dt1.Rows(Ri).Item("STA2") = Sta2
                                            Dt1.Rows(Ri).Item("OBJID") = block1.ObjectId
                                            Dt1.Rows(Ri).Item("NO_LETTER") = No_letter
                                            Ri = Ri + 1
                                        End If

                                    End If
                                End If
                            End If
                        Next
                        If IsNothing(Dt1) = False Then
                            If Dt1.Rows.Count > 0 Then
                                Dt1 = Sort_data_table(Dt1, "STA1")
                                Dim Total_distance As Double = Abs(Result_point1.Value.X - Result_point2.Value.X)
                                Dim Match1 As Double = Dt1.Rows(0).Item("STA1")
                                Dim Match2 As Double = Dt1.Rows(Dt1.Rows.Count - 1).Item("STA2")
                                Dim stretch_factor As Double = Total_distance / (Match2 - Match1)
                                For i = 0 To Dt1.Rows.Count - 1
                                    Dim Block1 As BlockReference = Trans1.GetObject(Dt1.Rows(i).Item("OBJID"), OpenMode.ForWrite)
                                    Dim x0 = Result_point1.Value.X
                                    Dim Sta1 As Double = Dt1.Rows(i).Item("STA1")
                                    Dim Sta2 As Double = Dt1.Rows(i).Item("STA2")
                                    Block1.Position = New Point3d(x0 + (Sta1 - Match1) * stretch_factor, Block1.Position.Y, 0)
                                    Stretch_block(Block1, "Distance1", (Sta2 - Sta1) * stretch_factor)
                                    Dt1.Rows(i).Item("STRETCH") = (Sta2 - Sta1) * stretch_factor
                                    Dt1.Rows(i).Item("PT_INS") = Block1.Position
                                Next

                                For i = 0 To Dt1.Rows.Count - 1
                                    Dim Block1 As BlockReference = Trans1.GetObject(Dt1.Rows(i).Item("OBJID"), OpenMode.ForWrite)
                                    Dim Letter_dist0 As Double = 16.6 * Dt1.Rows(i).Item("NO_LETTER")
                                    Dim Stretch0 As Double = Dt1.Rows(i).Item("STRETCH")
                                    Dim Sta01 As Double = Dt1.Rows(i).Item("STA1")
                                    Dim Sta02 As Double = Dt1.Rows(i).Item("STA2")
                                    If Letter_dist0 > Stretch0 Then
                                        For j = 0 To Dt1.Rows.Count - 1
                                            If Not j = i Then
                                                Dim Sta11 As Double = Dt1.Rows(j).Item("STA1")
                                                Dim Sta12 As Double = Dt1.Rows(j).Item("STA2")
                                                Dim Letter_dist1 As Double = 16.6 * Dt1.Rows(j).Item("NO_LETTER")
                                                Dim Stretch1 As Double = Dt1.Rows(j).Item("STRETCH")
                                                If Sta01 >= Sta11 And Sta02 <= Sta12 Then


                                                End If


                                            End If
                                        Next
                                    End If





                                Next


                            End If
                        End If
                        Trans1.Commit()
                    End Using
                End Using



            Catch ex As System.Exception
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub
End Class