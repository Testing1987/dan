Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports ACSMCOMPONENTS20Lib
Public Class Platt_Generator_form

    Public Shared Vw_scale As Double
    Public Shared Rotatie As Double
    Public Shared Poly_parc_for_jig As Polyline
    Public Shared pt_CERC As Point3d
    Dim New_plat_name_column As String = "PLAT_NAME"
    Dim Settings_file As String
    Dim Layer_name_Main_Viewport As String = "VP"
    Dim Layer_name_locus_VP As String = "VP_Locus"
    Dim Layer_name_Blocks As String = "TEXT"
    Dim Layer_name_text As String = "TEXT"
    Dim Layer_name_no_plot1 As String = "NO PLOT"
    Dim Layer_poly_parcela_new As String = "E_Bdy_Property"
    Dim Layer_hatch_locus As String = "locus_hatch"

    Dim LinetypeScale_Poly_parcela_new As Double = 0.5
    Dim Rotatie_originala As Double
    Dim Vw_height As Double = 4.2
    Dim Vw_width As Double = 7
    Dim Vw_CenX As Double = 8.5
    Dim Vw_CenY As Double = 7.4
    Dim VwL_height As Double = 4.2
    Dim VwL_width As Double = 7
    Dim VwL_CenX As Double = 8.5
    Dim VwL_CenY As Double = 7.4
    Dim County_pt As New Point3d(9, -3, 0)
    Dim State_pt As New Point3d(9, -3 - 2 * 0.08, 0)
    Dim Town_pt As New Point3d(9, -3 - 4 * 0.08, 0)
    Dim Owner_pt As New Point3d(9, -3 - 6 * 0.08, 0)
    Dim Line_list_pt As New Point3d(9, -3 - 8 * 0.08, 0)
    Dim HMMID_pt As New Point3d(9, -3 - 10 * 0.08, 0)
    Dim Diameter_pt As New Point3d(9, -3 - 12 * 0.08, 0)
    Dim Name_pt As New Point3d(9, -3 - 14 * 0.08, 0)
    Dim Segment_pt As New Point3d(9, -3 - 16 * 0.08, 0)

    Dim VW_TARGET_X As String = "VW_TARGET_X"
    Dim VW_TARGET_Y As String = "VW_TARGET_Y"
    Dim VW_TWIST As String = "VW_TWIST"
    Dim VW_CUST_SCALE As String = "VW_CUST_SCALE"
    Dim VW_NEW_NAME As String = "VW_NEW_NAME"
    Dim Obj_type As String = "OBJ_TYPE"

    Dim Punct_insertie_north_arrow As New Point3d(2.7722595, 8.9239174, 0)
    Dim Punct_insertie_Lnorth_arrow As New Point3d(2.7722595, 8.9239174, 0)
    Dim Colectie_ID As Specialized.StringCollection
    Dim Viewport_loaded As Boolean = False
    Dim Scale_factor As Double = 0
    Dim View_rotation As Double = 0
    Dim Text_height As Double = 0.08

    Dim Station_at_point_layer As String = "Text_Stationing_Wksp"
    Dim Centerline_stationing_layer As String = "Text_Stationing_CL"
    Dim Bearing_and_distance_layer As String = "Text_Dims_Bearing"
    Dim Tie_distance_layer As String = "Text_Dims_Ties"
    Dim Arc_leader_polyline_layer As String = "Text"


    Dim Data_table_layers As System.Data.DataTable





    Dim Freeze_operations As Boolean = False


    Private Sub Platt_Generator_form_Load(sender As Object, e As EventArgs) Handles Me.Load
        HScrollBar_rotate.Minimum = -180
        HScrollBar_rotate.Maximum = 180
        HScrollBar_rotate.Value = 0
        Settings_file = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\plat_settings.csv"
        Dim index_DT As Integer = 0
        Data_table_layers = New System.Data.DataTable
        Data_table_layers.Columns.Add("VIEWPORT_TYPE", GetType(String))
        Data_table_layers.Columns.Add("LAYER_NAME", GetType(String))
        Data_table_layers.Columns.Add("COLOR_INDEX", GetType(String))
        Data_table_layers.Columns.Add("THAW_FREEZE", GetType(String))


        Try

            If System.IO.File.Exists(Settings_file) = True Then
                Using Reader1 As New System.IO.StreamReader(Settings_file)
                    Dim Line1 As String
                    While Reader1.Peek > 0
                        Line1 = Reader1.ReadLine
                        If Line1.Contains(",") = True Then
                            Dim Array1() As String = Line1.Split(",")

                            Dim Line_no As String = Array1(0)
                            If Array1.Length >= 3 Then
                                Dim Value1 As String = Array1(2)
                                Select Case Line_no
                                    Case 1
                                        TextBox_Output_Directory.Text = Value1
                                    Case 2
                                        TextBox_xref_model_space.Text = Value1
                                    Case 3
                                        TextBox_dwt_template.Text = Value1
                                    Case 4
                                        TextBox_sheet_set_template.Text = Value1
                                    Case 5
                                        TextBox_north_arrow_Big_X.Text = Value1
                                    Case 6
                                        TextBox_north_arrow_Big_y.Text = Value1
                                    Case 7
                                        TextBox_main_viewport_height.Text = Value1
                                    Case 8
                                        TextBox_main_viewport_width.Text = Value1
                                    Case 9
                                        TextBox_main_viewport_center_X.Text = Value1
                                    Case 10
                                        TextBox_main_viewport_center_Y.Text = Value1
                                    Case 11
                                        TextBox_north_arrow.Text = Value1
                                    Case "11A"
                                        If Array1.Length >= 4 Then
                                            With ComboBox_SHEET_SET_scale_main
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                    Case 12
                                        If Value1 = "YES" Then
                                            CheckBox_add_locus.Checked = True
                                        Else
                                            CheckBox_add_locus.Checked = False
                                        End If
                                    Case 13
                                        TextBox_north_arrow_small_X.Text = Value1
                                    Case 14
                                        TextBox_north_arrow_small_Y.Text = Value1
                                    Case 15
                                        TextBox_locus_viewport_height.Text = Value1
                                    Case 16
                                        TextBox_locus_viewport_width.Text = Value1
                                    Case 17
                                        TextBox_locus_viewport_center_X.Text = Value1
                                    Case 18
                                        TextBox_locus_viewport_center_Y.Text = Value1
                                    Case 19
                                        TextBox_north_arrow_locus.Text = Value1
                                    Case 20
                                        If Array1.Length >= 4 Then
                                            With ComboBox_SHEET_SET_scale_locus
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                    Case 21
                                        With ComboBox_plat_name
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 22
                                        With ComboBox_state
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_Sheet_Set_state
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If

                                    Case 23
                                        With ComboBox_county
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_Sheet_Set_county
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                    Case 24
                                        With ComboBox_town
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_Sheet_Set_town
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                    Case 25
                                        With ComboBox_linelist
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_Sheet_Set_linelist
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                    Case 26
                                        With ComboBox_HMM_ID
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_Sheet_Set_hmm_ID
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                    Case 27
                                        With ComboBox_MBL
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_Sheet_Set_MBL
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                    Case 28
                                        With ComboBox_deed_page
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_Sheet_Set_deed_page
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                    Case 29
                                        With ComboBox_APN
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_Sheet_Set_APN
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                    Case 30
                                        With ComboBox_access_road_length
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_Sheet_Set_access_road_length
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 5 Then
                                            With TextBox_round_access_road_length
                                                .Text = Array1(4)
                                            End With
                                        End If
                                    Case 31
                                        With ComboBox_Section_TWP_Range
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_Sheet_Set_sec_twp_range
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                    Case 32
                                        With ComboBox_owner
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_Sheet_Set_owner
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                    Case 33
                                        With ComboBox_crossing_length
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_Sheet_Set_crossing_len_FT
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 5 Then
                                            With TextBox_round_crossing_length_ft
                                                .Text = Array1(4)
                                            End With
                                        End If
                                    Case "33A"
                                        If Array1.Length >= 4 Then
                                            With ComboBox_Sheet_Set_crossing_len_rod
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 5 Then
                                            With TextBox_round_crossing_length_rod
                                                .Text = Array1(4)
                                            End With
                                        End If

                                    Case 34
                                        If Value1 = "YES" Then
                                            CheckBox_USE_ned_for_naming.Checked = True
                                        Else
                                            CheckBox_USE_ned_for_naming.Checked = False
                                        End If
                                    Case 35
                                        With ComboBox_pipe_diameter
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_Sheet_Set_pipe_DIAMETER
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                    Case 36
                                        With ComboBox_pipe_name
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_Sheet_Set_pipe_name
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                    Case 37
                                        With ComboBox_pipe_segment
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_Sheet_Set_pipe_segment
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                    Case 38
                                        With ComboBox_area_EX_E
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_Sheet_Set_ex_e
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 5 Then
                                            With TextBox_round_EX_E
                                                .Text = Array1(4)
                                            End With
                                        End If
                                    Case 39
                                        With ComboBox_area_P_E
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_Sheet_Set_P_E
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 5 Then
                                            With TextBox_round_P_E
                                                .Text = Array1(4)
                                            End With
                                        End If
                                    Case 40
                                        With ComboBox_area_TWS
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_Sheet_Set_TWS
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 5 Then
                                            With TextBox_round_TWS
                                                .Text = Array1(4)
                                            End With
                                        End If
                                    Case 41
                                        With ComboBox_area_ATWS
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_Sheet_Set_atws
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 5 Then
                                            With TextBox_round_ATWS
                                                .Text = Array1(4)
                                            End With
                                        End If
                                    Case 42
                                        With ComboBox_area_A_R
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_Sheet_Set_a_R
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 5 Then
                                            With TextBox_round_A_R
                                                .Text = Array1(4)
                                            End With
                                        End If
                                    Case 43
                                        With ComboBox_area_WARE
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_Sheet_Set_ware
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 5 Then
                                            With TextBox_round_WARE
                                                .Text = Array1(4)
                                            End With
                                        End If
                                    Case 44
                                        With ComboBox_area_TWS_ABD
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_Sheet_Set_tws_abd
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 5 Then
                                            With TextBox_round_TWS_ABD
                                                .Text = Array1(4)
                                            End With
                                        End If
                                    Case 45
                                        With ComboBox_area_TWS_PD
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_Sheet_Set_tws_PD
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 5 Then
                                            With TextBox_round_TWS_PD
                                                .Text = Array1(4)
                                            End With
                                        End If
                                    Case 46
                                        If Value1 = "YES" Then
                                            CheckBox_convert_sqft_to_acres.Checked = True
                                        Else
                                            CheckBox_convert_sqft_to_acres.Checked = False
                                        End If
                                    Case 47
                                        If Value1 = "YES" Then
                                            CheckBox_ignore_deflections_less_0_5.Checked = True
                                        Else
                                            CheckBox_ignore_deflections_less_0_5.Checked = False
                                        End If
                                    Case 48
                                        If Value1 = "YES" Then
                                            CheckBox_add_bearing_and_distances.Checked = True
                                        Else
                                            CheckBox_add_bearing_and_distances.Checked = False
                                        End If
                                    Case 49
                                        If Value1 = "YES" Then
                                            CheckBox_display_stations.Checked = True
                                        Else
                                            CheckBox_display_stations.Checked = False
                                        End If
                                    Case 50
                                        If Value1 = "YES" Then
                                            CheckBox_rotate_to_north.Checked = True
                                        Else
                                            CheckBox_rotate_to_north.Checked = False
                                        End If
                                    Case 122
                                        With ComboBox_state_update
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_set_state_update
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If

                                        If Array1.Length >= 8 Then TextBox_prefix_state_update.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_state_update.Text = Array1(8)


                                    Case 123
                                        With ComboBox_county_update
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_set_county_update
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_county_update.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_county_update.Text = Array1(8)
                                    Case 124
                                        With ComboBox_town_update
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_set_town_update
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_town_update.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_town_update.Text = Array1(8)
                                    Case 125
                                        With ComboBox_linelist_update
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_set_linelist_update
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_linelist_update.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_linelist_update.Text = Array1(8)
                                    Case 126
                                        With ComboBox_HMM_ID_update
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_set_HMMID_update
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                    Case 127
                                        With ComboBox_MBL_update
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_set_MBL_update
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_mbl_update.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_mbl_update.Text = Array1(8)
                                    Case 128
                                        With ComboBox_deed_page_update
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_set_deed_page_update
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_deed_page_update.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_deed_page_update.Text = Array1(8)
                                    Case 129
                                        With ComboBox_APN_update
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_set_apn_update
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_apn_update.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_apn_update.Text = Array1(8)
                                    Case 130
                                        With ComboBox_access_road_length_update
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_set_access_road_length_update
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 5 Then
                                            With TextBox_round_access_road_length_update
                                                .Text = Array1(4)
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_a_r_length_update.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_a_r_length_update.Text = Array1(8)
                                    Case 131
                                        With ComboBox_Section_TWP_Range_update
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        With ComboBox_sheet_set_sec_twp_range_update
                                            .Items.Add(Array1(3))
                                            .SelectedIndex = .Items.IndexOf(Array1(3))
                                        End With
                                        If Array1.Length >= 8 Then TextBox_prefix_sec_twp_rge_update.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_sec_twp_rge_update.Text = Array1(8)
                                    Case 132
                                        With ComboBox_owner_update
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_set_owner_update
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_owner_update.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_owner_update.Text = Array1(8)
                                    Case 133
                                        With ComboBox_crossing_length_update
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_set_crosing_length_ft_update
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 5 Then
                                            With TextBox_round_crossing_length_ft_update
                                                .Text = Array1(4)
                                            End With
                                        End If

                                        If Array1.Length >= 8 Then TextBox_prefix_crossing_length_ft_update.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_crossing_length_ft_update.Text = Array1(8)

                                    Case "133A"
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_set_crosing_length_rod_update
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 5 Then
                                            With TextBox_round_crossing_length_rod_update
                                                .Text = Array1(4)
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_crossing_length_rod_update.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_crossing_length_rod_update.Text = Array1(8)

                                    Case 138
                                        With ComboBox_area_EX_E_update
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_set_ex_e_update
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 5 Then
                                            With TextBox_round_EX_E_update
                                                .Text = Array1(4)
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_AREA_EX_E_update.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_AREA_EX_E_update.Text = Array1(8)
                                    Case 139
                                        With ComboBox_area_P_E_update
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_set_p_e_update
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 5 Then
                                            With TextBox_round_P_E_update
                                                .Text = Array1(4)
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_AREA_P_E_update.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_AREA_P_E_update.Text = Array1(8)
                                    Case 140
                                        With ComboBox_area_TWS_update
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_set_tws_update
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 5 Then
                                            With TextBox_round_TWS_update
                                                .Text = Array1(4)
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_AREA_TWS_update.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_AREA_TWS_update.Text = Array1(8)
                                    Case 141
                                        With ComboBox_area_ATWS_update
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_set_atws_update
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 5 Then
                                            With TextBox_round_ATWS_update
                                                .Text = Array1(4)
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_area_ATWS_update.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_area_ATWS_update.Text = Array1(8)
                                    Case 142
                                        With ComboBox_area_A_R_update
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_set_a_r_update
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 5 Then
                                            With TextBox_round_A_R_update
                                                .Text = Array1(4)
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_AREA_A_R_update.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_AREA_A_R_update.Text = Array1(8)
                                    Case 143
                                        With ComboBox_area_WARE_update
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_set_ware_update
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 5 Then
                                            With TextBox_round_WARE_update
                                                .Text = Array1(4)
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_AREA_WARE_update.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_AREA_WARE_update.Text = Array1(8)
                                    Case 144
                                        With ComboBox_area_TWS_ABD_update
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_set_tws_ABD_update
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 5 Then
                                            With TextBox_round_TWS_ABD_update
                                                .Text = Array1(4)
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_AREA_TWS_ABD_update.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_AREA_TWS_ABD_update.Text = Array1(8)
                                    Case 145
                                        With ComboBox_area_TWS_PD_update
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_set_TWS_PD_update
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 5 Then
                                            With TextBox_round_TWS_PD_update
                                                .Text = Array1(4)
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_AREA_TWS_PD_update.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_AREA_TWS_PD_update.Text = Array1(8)
                                    Case 146
                                        If Value1 = "YES" Then
                                            CheckBox_convert_sqft_to_acres_UPDATE.Checked = True
                                        Else
                                            CheckBox_convert_sqft_to_acres_UPDATE.Checked = False
                                        End If

                                    Case 150
                                        With ComboBox_shape_user1
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_user1
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_user1.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_user1.Text = Array1(8)

                                    Case 151
                                        With ComboBox_shape_user2
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_user2
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_user2.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_user2.Text = Array1(8)
                                    Case 152
                                        With ComboBox_shape_user3
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_user3
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_user3.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_user3.Text = Array1(8)
                                    Case 153
                                        With ComboBox_shape_user4
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_user4
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_user4.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_user4.Text = Array1(8)
                                    Case 154
                                        With ComboBox_shape_user5
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_user5
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_user5.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_user5.Text = Array1(8)
                                    Case 155
                                        With ComboBox_shape_user6
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_user6
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_user6.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_user6.Text = Array1(8)
                                    Case 156
                                        With ComboBox_shape_user7
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_user7
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_user7.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_user7.Text = Array1(8)
                                    Case 157
                                        With ComboBox_shape_user8
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_user8
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_user8.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_user8.Text = Array1(8)
                                    Case 158
                                        With ComboBox_shape_user9
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_user9
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_user9.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_user9.Text = Array1(8)
                                    Case 159
                                        With ComboBox_shape_user10
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_user10
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_user10.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_user10.Text = Array1(8)
                                    Case 160
                                        With ComboBox_shape_user11
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_user11
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_user11.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_user11.Text = Array1(8)
                                    Case 161
                                        With ComboBox_shape_user12
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_user12
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_user12.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_user12.Text = Array1(8)
                                    Case 162
                                        With ComboBox_shape_user13
                                            .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                        If Array1.Length >= 4 Then
                                            With ComboBox_sheet_user13
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                        End If
                                        If Array1.Length >= 8 Then TextBox_prefix_user13.Text = Array1(7)
                                        If Array1.Length >= 9 Then TextBox_suffix_user13.Text = Array1(8)
                                    Case 1000 To 1999
                                        If Array1.Length >= 5 Then
                                            Dim Viewport_type As String = ""
                                            If Array1(1).ToUpper = "LOCUS" Then
                                                Viewport_type = "LOCUS"
                                            End If
                                            If Array1(1).ToUpper = "MAIN" Then
                                                Viewport_type = "MAIN"
                                            End If
                                            If Not Viewport_type = "" Then
                                                Dim Thaw_Freeze As String = ""

                                                If Array1(4).ToUpper = "THAW" Then
                                                    Thaw_Freeze = "THAW"
                                                End If
                                                If Array1(4).ToUpper = "FROZEN" Then
                                                    Thaw_Freeze = "FROZEN"
                                                End If
                                                If Not Thaw_Freeze = "" Then
                                                    Data_table_layers.Rows.Add()
                                                    Data_table_layers.Rows(index_DT).Item("VIEWPORT_TYPE") = Viewport_type
                                                    Data_table_layers.Rows(index_DT).Item("THAW_FREEZE") = Thaw_Freeze
                                                    Data_table_layers.Rows(index_DT).Item("LAYER_NAME") = Array1(2)
                                                    If IsNumeric(Array1(3)) = True Then
                                                        Data_table_layers.Rows(index_DT).Item("COLOR_INDEX") = Array1(3)
                                                    End If
                                                    index_DT = index_DT + 1
                                                End If
                                            End If

                                        End If
                                End Select
                            End If
                        End If
                    End While

                End Using
            Else
                'TextBox_xref_model_space.Text = "G:\KinderMorgan\339501_NEDProject\DataProd\DWG\Segment_A\A_Host_Basefile_Plats.dwg"
                'TextBox_dwt_template.Text = "C:\Users\pop70694\Documents\Work Files\2015-10-19 plat generator error\danF.dwt" '  "G:\KinderMorgan\339501_NEDProject\DataProd\Work\Drafting\Land_Plats\1_Platt_Generator\Platt_Generator_Temp_19F.dwt"
                'TextBox_sheet_set_template.Text = "C:\Users\pop70694\Documents\Work Files\2015-10-19 plat generator error\Hancock_plats.dst" ' "G:\KinderMorgan\339501_NEDProject\DataProd\Work\Drafting\Land_Plats\1_Platt_Generator\Platt_Sheet_Set_19F.dst"
                'TextBox_Output_Directory.Text = "C:\Users\pop70694\Documents\Work Files\" ' "G:\KinderMorgan\339501_NEDProject\DataProd\Work\Drafting\Land_Plats\Segment_K\"
            End If

        Catch ex As System.Exception
            MsgBox(ex.Message)
        End Try


        Incarca_existing_Blocks_with_attributes_to_combobox(ComboBox_blocks)

        If ComboBox_blocks.Items.Contains("Owner_Linelist") = True Then
            ComboBox_blocks.SelectedIndex = ComboBox_blocks.Items.IndexOf("Owner_Linelist")
        Else
            ComboBox_blocks.SelectedIndex = 0
        End If

        Incarca_existing_Atributes_to_combobox(ComboBox_blocks.Text, ComboBox_bl_atr1)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks.Text, ComboBox_bl_atr2)
        If ComboBox_bl_atr1.Items.Count > 1 Then
            ComboBox_bl_atr1.SelectedIndex = 1
        End If
        If ComboBox_bl_atr2.Items.Count > 2 Then
            ComboBox_bl_atr2.SelectedIndex = 2
        End If
    End Sub

    Private Sub Button_read_settings_Click(sender As Object, e As EventArgs) Handles Button_read_settings.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Dim index_DT As Integer = 0
            Data_table_layers = New System.Data.DataTable
            Data_table_layers.Columns.Add("VIEWPORT_TYPE", GetType(String))
            Data_table_layers.Columns.Add("LAYER_NAME", GetType(String))
            Data_table_layers.Columns.Add("COLOR_INDEX", GetType(String))
            Data_table_layers.Columns.Add("THAW_FREEZE", GetType(String))
            Try
                If System.IO.File.Exists(Settings_file) = True Then
                    Using Reader1 As New System.IO.StreamReader(Settings_file)
                        Dim Line1 As String

                        While Reader1.Peek > 0
                            Line1 = Reader1.ReadLine
                            If Line1.Contains(",") = True Then
                                Dim Array1() As String = Line1.Split(",")
                                Dim Line_no As String = Array1(0)

                                If Array1.Length >= 3 Then
                                    Dim Value1 As String = Array1(2)
                                    Select Case Line_no
                                        Case 1
                                            TextBox_Output_Directory.Text = Value1
                                        Case 2
                                            TextBox_xref_model_space.Text = Value1
                                        Case 3
                                            TextBox_dwt_template.Text = Value1
                                        Case 4
                                            TextBox_sheet_set_template.Text = Value1
                                        Case 5
                                            TextBox_north_arrow_Big_X.Text = Value1
                                        Case 6
                                            TextBox_north_arrow_Big_y.Text = Value1
                                        Case 7
                                            TextBox_main_viewport_height.Text = Value1
                                        Case 8
                                            TextBox_main_viewport_width.Text = Value1
                                        Case 9
                                            TextBox_main_viewport_center_X.Text = Value1
                                        Case 10
                                            TextBox_main_viewport_center_Y.Text = Value1
                                        Case 11
                                            TextBox_north_arrow.Text = Value1
                                        Case "11A"
                                            If Array1.Length >= 4 Then
                                                With ComboBox_SHEET_SET_scale_main
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                        Case 12
                                            If Value1 = "YES" Then
                                                CheckBox_add_locus.Checked = True
                                            Else
                                                CheckBox_add_locus.Checked = False
                                            End If
                                        Case 13
                                            TextBox_north_arrow_small_X.Text = Value1
                                        Case 14
                                            TextBox_north_arrow_small_Y.Text = Value1
                                        Case 15
                                            TextBox_locus_viewport_height.Text = Value1
                                        Case 16
                                            TextBox_locus_viewport_width.Text = Value1
                                        Case 17
                                            TextBox_locus_viewport_center_X.Text = Value1
                                        Case 18
                                            TextBox_locus_viewport_center_Y.Text = Value1
                                        Case 19
                                            TextBox_north_arrow_locus.Text = Value1
                                        Case 20
                                            If Array1.Length >= 4 Then
                                                With ComboBox_SHEET_SET_scale_locus
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If

                                        Case 21
                                            With ComboBox_plat_name
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                        Case 22
                                            With ComboBox_state
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_Sheet_Set_state
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                        Case "22A"
                                            If Value1 = "YES" Then
                                                CheckBox_state_abreviation.Checked = True
                                            Else
                                                CheckBox_state_abreviation.Checked = False
                                            End If
                                        Case 23
                                            With ComboBox_county
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_Sheet_Set_county
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                        Case 24
                                            With ComboBox_town
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_Sheet_Set_town
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                        Case 25
                                            With ComboBox_linelist
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_Sheet_Set_linelist
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                        Case 26
                                            With ComboBox_HMM_ID
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_Sheet_Set_hmm_ID
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                        Case 27
                                            With ComboBox_MBL
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_Sheet_Set_MBL
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                        Case 28
                                            With ComboBox_deed_page
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_Sheet_Set_deed_page
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                        Case 29
                                            With ComboBox_APN
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_Sheet_Set_APN
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                        Case 30
                                            With ComboBox_access_road_length
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_Sheet_Set_access_road_length
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 5 Then
                                                With TextBox_round_access_road_length
                                                    .Text = Array1(4)
                                                End With
                                            End If
                                        Case 31
                                            With ComboBox_Section_TWP_Range
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_Sheet_Set_sec_twp_range
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                        Case 32
                                            With ComboBox_owner
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_Sheet_Set_owner
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                        Case 33
                                            With ComboBox_crossing_length
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_Sheet_Set_crossing_len_FT
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 5 Then
                                                With TextBox_round_crossing_length_ft
                                                    .Text = Array1(4)
                                                End With
                                            End If
                                        Case "33A"
                                            If Array1.Length >= 4 Then
                                                With ComboBox_Sheet_Set_crossing_len_rod
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 5 Then
                                                With TextBox_round_crossing_length_rod
                                                    .Text = Array1(4)
                                                End With
                                            End If
                                        Case 34
                                            If Value1 = "YES" Then
                                                CheckBox_USE_ned_for_naming.Checked = True
                                            Else
                                                CheckBox_USE_ned_for_naming.Checked = False
                                            End If
                                        Case 35
                                            With ComboBox_pipe_diameter
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_Sheet_Set_pipe_DIAMETER
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                        Case 36
                                            With ComboBox_pipe_name
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_Sheet_Set_pipe_name
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                        Case 37
                                            With ComboBox_pipe_segment
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_Sheet_Set_pipe_segment
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                        Case 38
                                            With ComboBox_area_EX_E
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_Sheet_Set_ex_e
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 5 Then
                                                With TextBox_round_EX_E
                                                    .Text = Array1(4)
                                                End With
                                            End If
                                        Case 39
                                            With ComboBox_area_P_E
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_Sheet_Set_P_E
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 5 Then
                                                With TextBox_round_P_E
                                                    .Text = Array1(4)
                                                End With
                                            End If
                                        Case 40
                                            With ComboBox_area_TWS
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_Sheet_Set_TWS
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 5 Then
                                                With TextBox_round_TWS
                                                    .Text = Array1(4)
                                                End With
                                            End If
                                        Case 41
                                            With ComboBox_area_ATWS
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_Sheet_Set_atws
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 5 Then
                                                With TextBox_round_ATWS
                                                    .Text = Array1(4)
                                                End With
                                            End If
                                        Case 42
                                            With ComboBox_area_A_R
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_Sheet_Set_a_R
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 5 Then
                                                With TextBox_round_A_R
                                                    .Text = Array1(4)
                                                End With
                                            End If
                                        Case 43
                                            With ComboBox_area_WARE
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_Sheet_Set_ware
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 5 Then
                                                With TextBox_round_WARE
                                                    .Text = Array1(4)
                                                End With
                                            End If
                                        Case 44
                                            With ComboBox_area_TWS_ABD
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_Sheet_Set_tws_abd
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 5 Then
                                                With TextBox_round_TWS_ABD
                                                    .Text = Array1(4)
                                                End With
                                            End If
                                        Case 45
                                            With ComboBox_area_TWS_PD
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_Sheet_Set_tws_PD
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 5 Then
                                                With TextBox_round_TWS_PD
                                                    .Text = Array1(4)
                                                End With
                                            End If
                                        Case 46
                                            If Value1 = "YES" Then
                                                CheckBox_convert_sqft_to_acres.Checked = True
                                            Else
                                                CheckBox_convert_sqft_to_acres.Checked = False
                                            End If
                                        Case 47
                                            If Value1 = "YES" Then
                                                CheckBox_ignore_deflections_less_0_5.Checked = True
                                            Else
                                                CheckBox_ignore_deflections_less_0_5.Checked = False
                                            End If
                                        Case 48
                                            If Value1 = "YES" Then
                                                CheckBox_add_bearing_and_distances.Checked = True
                                            Else
                                                CheckBox_add_bearing_and_distances.Checked = False
                                            End If
                                        Case 49
                                            If Value1 = "YES" Then
                                                CheckBox_display_stations.Checked = True
                                            Else
                                                CheckBox_display_stations.Checked = False
                                            End If
                                        Case 50
                                            If Value1 = "YES" Then
                                                CheckBox_rotate_to_north.Checked = True
                                            Else
                                                CheckBox_rotate_to_north.Checked = False
                                            End If
                                        Case 122
                                            With ComboBox_state_update
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_set_state_update
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If

                                            If Array1.Length >= 8 Then TextBox_prefix_state_update.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_state_update.Text = Array1(8)
                                        Case "122A"
                                            If Value1 = "YES" Then
                                                CheckBox_state_abreviation_UPDATE.Checked = True
                                            Else
                                                CheckBox_state_abreviation_UPDATE.Checked = False
                                            End If

                                        Case 123
                                            With ComboBox_county_update
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_set_county_update
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_county_update.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_county_update.Text = Array1(8)
                                        Case 124
                                            With ComboBox_town_update
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_set_town_update
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_town_update.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_town_update.Text = Array1(8)
                                        Case 125
                                            With ComboBox_linelist_update
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_set_linelist_update
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_linelist_update.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_linelist_update.Text = Array1(8)
                                        Case 126
                                            With ComboBox_HMM_ID_update
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_set_HMMID_update
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                        Case 127
                                            With ComboBox_MBL_update
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_set_MBL_update
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_mbl_update.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_mbl_update.Text = Array1(8)
                                        Case 128
                                            With ComboBox_deed_page_update
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_set_deed_page_update
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_deed_page_update.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_deed_page_update.Text = Array1(8)
                                        Case 129
                                            With ComboBox_APN_update
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_set_apn_update
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_apn_update.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_apn_update.Text = Array1(8)
                                        Case 130
                                            With ComboBox_access_road_length_update
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_set_access_road_length_update
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 5 Then
                                                With TextBox_round_access_road_length_update
                                                    .Text = Array1(4)
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_a_r_length_update.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_a_r_length_update.Text = Array1(8)
                                        Case 131
                                            With ComboBox_Section_TWP_Range_update
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            With ComboBox_sheet_set_sec_twp_range_update
                                                .Items.Add(Array1(3))
                                                .SelectedIndex = .Items.IndexOf(Array1(3))
                                            End With
                                            If Array1.Length >= 8 Then TextBox_prefix_sec_twp_rge_update.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_sec_twp_rge_update.Text = Array1(8)
                                        Case 132
                                            With ComboBox_owner_update
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_set_owner_update
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_owner_update.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_owner_update.Text = Array1(8)
                                        Case 133
                                            With ComboBox_crossing_length_update
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_set_crosing_length_ft_update
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 5 Then
                                                With TextBox_round_crossing_length_ft_update
                                                    .Text = Array1(4)
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_crossing_length_ft_update.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_crossing_length_ft_update.Text = Array1(8)
                                        Case "133A"
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_set_crosing_length_rod_update
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 5 Then
                                                With TextBox_round_crossing_length_rod_update
                                                    .Text = Array1(4)
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_crossing_length_rod_update.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_crossing_length_rod_update.Text = Array1(8)
                                        Case 138
                                            With ComboBox_area_EX_E_update
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_set_ex_e_update
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 5 Then
                                                With TextBox_round_EX_E_update
                                                    .Text = Array1(4)
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_AREA_EX_E_update.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_AREA_EX_E_update.Text = Array1(8)
                                        Case 139
                                            With ComboBox_area_P_E_update
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_set_p_e_update
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 5 Then
                                                With TextBox_round_P_E_update
                                                    .Text = Array1(4)
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_AREA_P_E_update.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_AREA_P_E_update.Text = Array1(8)
                                        Case 140
                                            With ComboBox_area_TWS_update
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_set_tws_update
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 5 Then
                                                With TextBox_round_TWS_update
                                                    .Text = Array1(4)
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_AREA_TWS_update.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_AREA_TWS_update.Text = Array1(8)
                                        Case 141
                                            With ComboBox_area_ATWS_update
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_set_atws_update
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 5 Then
                                                With TextBox_round_ATWS_update
                                                    .Text = Array1(4)
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_area_ATWS_update.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_area_ATWS_update.Text = Array1(8)
                                        Case 142
                                            With ComboBox_area_A_R_update
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_set_a_r_update
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 5 Then
                                                With TextBox_round_A_R_update
                                                    .Text = Array1(4)
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_AREA_A_R_update.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_AREA_A_R_update.Text = Array1(8)
                                        Case 143
                                            With ComboBox_area_WARE_update
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_set_ware_update
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 5 Then
                                                With TextBox_round_WARE_update
                                                    .Text = Array1(4)
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_AREA_WARE_update.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_AREA_WARE_update.Text = Array1(8)
                                        Case 144
                                            With ComboBox_area_TWS_ABD_update
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_set_tws_ABD_update
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 5 Then
                                                With TextBox_round_TWS_ABD_update
                                                    .Text = Array1(4)
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_AREA_TWS_ABD_update.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_AREA_TWS_ABD_update.Text = Array1(8)
                                        Case 145
                                            With ComboBox_area_TWS_PD_update
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_set_TWS_PD_update
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 5 Then
                                                With TextBox_round_TWS_PD_update
                                                    .Text = Array1(4)
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_AREA_TWS_PD_update.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_AREA_TWS_PD_update.Text = Array1(8)
                                        Case 146
                                            If Value1 = "YES" Then
                                                CheckBox_convert_sqft_to_acres_UPDATE.Checked = True
                                            Else
                                                CheckBox_convert_sqft_to_acres_UPDATE.Checked = False
                                            End If

                                        Case 150
                                            With ComboBox_shape_user1
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_user1
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_user1.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_user1.Text = Array1(8)

                                        Case 151
                                            With ComboBox_shape_user2
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_user2
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_user2.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_user2.Text = Array1(8)
                                        Case 152
                                            With ComboBox_shape_user3
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_user3
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_user3.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_user3.Text = Array1(8)
                                        Case 153
                                            With ComboBox_shape_user4
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_user4
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_user4.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_user4.Text = Array1(8)
                                        Case 154
                                            With ComboBox_shape_user5
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_user5
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_user5.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_user5.Text = Array1(8)
                                        Case 155
                                            With ComboBox_shape_user6
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_user6
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_user6.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_user6.Text = Array1(8)
                                        Case 156
                                            With ComboBox_shape_user7
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_user7
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_user7.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_user7.Text = Array1(8)
                                        Case 157
                                            With ComboBox_shape_user8
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_user8
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_user8.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_user8.Text = Array1(8)
                                        Case 158
                                            With ComboBox_shape_user9
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_user9
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_user9.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_user9.Text = Array1(8)
                                        Case 159
                                            With ComboBox_shape_user10
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_user10
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_user10.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_user10.Text = Array1(8)
                                        Case 160
                                            With ComboBox_shape_user11
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_user11
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_user11.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_user11.Text = Array1(8)
                                        Case 161
                                            With ComboBox_shape_user12
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_user12
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_user12.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_user12.Text = Array1(8)
                                        Case 162
                                            With ComboBox_shape_user13
                                                .Items.Add(Value1)
                                                .SelectedIndex = .Items.IndexOf(Value1)
                                            End With
                                            If Array1.Length >= 4 Then
                                                With ComboBox_sheet_user13
                                                    .Items.Add(Array1(3))
                                                    .SelectedIndex = .Items.IndexOf(Array1(3))
                                                End With
                                            End If
                                            If Array1.Length >= 8 Then TextBox_prefix_user13.Text = Array1(7)
                                            If Array1.Length >= 9 Then TextBox_suffix_user13.Text = Array1(8)
                                        Case 1000 To 1999
                                            If Array1.Length >= 5 Then
                                                Dim Viewport_type As String = ""
                                                If Array1(1).ToUpper = "LOCUS" Then
                                                    Viewport_type = "LOCUS"
                                                End If
                                                If Array1(1).ToUpper = "MAIN" Then
                                                    Viewport_type = "MAIN"
                                                End If
                                                If Not Viewport_type = "" Then
                                                    Dim Thaw_Freeze As String = ""

                                                    If Array1(4).ToUpper = "THAW" Then
                                                        Thaw_Freeze = "THAW"
                                                    End If
                                                    If Array1(4).ToUpper = "FROZEN" Then
                                                        Thaw_Freeze = "FROZEN"
                                                    End If
                                                    If Not Thaw_Freeze = "" Then
                                                        Data_table_layers.Rows.Add()
                                                        Data_table_layers.Rows(index_DT).Item("VIEWPORT_TYPE") = Viewport_type
                                                        Data_table_layers.Rows(index_DT).Item("THAW_FREEZE") = Thaw_Freeze
                                                        Data_table_layers.Rows(index_DT).Item("LAYER_NAME") = Array1(2)
                                                        If IsNumeric(Array1(3)) = True Then
                                                            Data_table_layers.Rows(index_DT).Item("COLOR_INDEX") = Array1(3)
                                                        End If
                                                        index_DT = index_DT + 1
                                                    End If
                                                End If

                                            End If

                                    End Select
                                End If
                            End If
                        End While

                    End Using

                Else
                    MsgBox("There is no file on the desktop - no settings have been loaded")

                End If



            Catch ex As System.Exception
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_create_settings_Click(sender As Object, e As EventArgs) Handles Button_create_settings.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                If System.IO.File.Exists(Settings_file) = True Then
                    IO.File.Delete(Settings_file)
                End If
                Using Fs As IO.FileStream = IO.File.Create(Settings_file)

                End Using


                Using sw As IO.StreamWriter = New IO.StreamWriter(Settings_file)
                    sw.Write("ITEM,DESCRIPTION,VALUE (OBJECT DATA),SHEET_SET_FIELD,NO DEC,EMPTY,EMPTY,PREFIX,SUFFIX")
                    sw.Write(vbCrLf & "1,Output Folder," & TextBox_Output_Directory.Text)
                    sw.Write(vbCrLf & "2,XREF (Basemap)," & TextBox_xref_model_space.Text)
                    sw.Write(vbCrLf & "3,DWT template," & TextBox_dwt_template.Text)
                    sw.Write(vbCrLf & "4,DST sheet set manager template," & TextBox_sheet_set_template.Text)
                    sw.Write(vbCrLf & "5,North Arrow X Coordinate (Main Viewport)," & TextBox_north_arrow_Big_X.Text)
                    sw.Write(vbCrLf & "6,North Arrow Y Coordinate (Main Viewport)," & TextBox_north_arrow_Big_y.Text)
                    sw.Write(vbCrLf & "7,Main Viewport Height," & TextBox_main_viewport_height.Text)
                    sw.Write(vbCrLf & "8,Main Viewport Width," & TextBox_main_viewport_width.Text)
                    sw.Write(vbCrLf & "9,Main Viewport Center X Coordinate," & TextBox_main_viewport_center_X.Text)
                    sw.Write(vbCrLf & "10,Main Viewport Center Y Coordinate," & TextBox_main_viewport_center_Y.Text)
                    sw.Write(vbCrLf & "11,North Arrow Autocad Block name (Main Viewport)," & TextBox_north_arrow.Text)
                    If Not ComboBox_SHEET_SET_scale_locus.Text = "" Then
                        sw.Write(vbCrLf & "11A,Sheet Set Field - Main Viewport Scale,," & ComboBox_SHEET_SET_scale_main.Text)
                    End If


                    Dim YesNO As String

                    If CheckBox_add_locus.Checked = True Then
                        YesNO = "YES"
                    Else
                        YesNO = "NO"
                    End If

                    sw.Write(vbCrLf & "12,Create Locus Viewport," & YesNO)

                    If CheckBox_add_locus.Checked = True Then
                        sw.Write(vbCrLf & "13,North Arrow X Coordinate (Locus Viewport)," & TextBox_north_arrow_small_X.Text)
                        sw.Write(vbCrLf & "14,North Arrow Y Coordinate (Locus Viewport)," & TextBox_north_arrow_small_Y.Text)
                        sw.Write(vbCrLf & "15,Locus Viewport Height," & TextBox_locus_viewport_height.Text)
                        sw.Write(vbCrLf & "16,Locus Viewport Width," & TextBox_locus_viewport_width.Text)
                        sw.Write(vbCrLf & "17,Locus Viewport Center X Coordinate," & TextBox_locus_viewport_center_X.Text)
                        sw.Write(vbCrLf & "18,Locus Viewport Center Y Coordinate," & TextBox_locus_viewport_center_Y.Text)
                        sw.Write(vbCrLf & "19,North Arrow Autocad Block name (Locus Viewport)," & TextBox_north_arrow_locus.Text)
                        If Not ComboBox_SHEET_SET_scale_locus.Text = "" Then
                            sw.Write(vbCrLf & "20,Sheet Set Field - Locus Viewport Scale,," & ComboBox_SHEET_SET_scale_locus.Text)
                        End If
                    End If


                    If Not ComboBox_plat_name.Text = "" Then sw.Write(vbCrLf & "21,PLAT NAME (Object Data)," & ComboBox_plat_name.Text)

                    With ComboBox_state
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_Sheet_Set_state.Text = "" Then
                                SS_field = ComboBox_Sheet_Set_state.Text
                            End If
                            sw.Write(vbCrLf & "22,State (Object Data)," & .Text & "," & SS_field)
                        End If
                    End With

                    Dim Abr1 As String

                    If CheckBox_state_abreviation.Checked = True Then
                        Abr1 = "YES"
                    Else
                        Abr1 = "NO"
                    End If

                    sw.Write(vbCrLf & "22A,Use State abbreviations," & Abr1)


                    With ComboBox_county
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_Sheet_Set_county.Text = "" Then
                                SS_field = ComboBox_Sheet_Set_county.Text
                            End If
                            sw.Write(vbCrLf & "23,County (Object Data)," & .Text & "," & SS_field)
                        End If
                    End With

                    With ComboBox_town
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_Sheet_Set_town.Text = "" Then
                                SS_field = ComboBox_Sheet_Set_town.Text
                            End If
                            sw.Write(vbCrLf & "24,Town (Object Data)," & .Text & "," & SS_field)
                        End If
                    End With

                    With ComboBox_linelist
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_Sheet_Set_linelist.Text = "" Then
                                SS_field = ComboBox_Sheet_Set_linelist.Text
                            End If
                            sw.Write(vbCrLf & "25,Linelist (Object Data)," & .Text & "," & SS_field)
                        End If
                    End With

                    With ComboBox_HMM_ID
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_Sheet_Set_hmm_ID.Text = "" Then
                                SS_field = ComboBox_Sheet_Set_hmm_ID.Text
                            End If
                            sw.Write(vbCrLf & "26,Parcel ID (Object Data)," & .Text & "," & SS_field)
                        End If
                    End With

                    With ComboBox_MBL
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_Sheet_Set_MBL.Text = "" Then
                                SS_field = ComboBox_Sheet_Set_MBL.Text
                            End If
                            sw.Write(vbCrLf & "27,MBL (Object Data)," & .Text & "," & SS_field)
                        End If
                    End With

                    With ComboBox_deed_page
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_Sheet_Set_deed_page.Text = "" Then
                                SS_field = ComboBox_Sheet_Set_deed_page.Text
                            End If
                            sw.Write(vbCrLf & "28,DEED BOOK PAGE (Object Data)," & .Text & "," & SS_field)
                        End If
                    End With

                    With ComboBox_APN
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_Sheet_Set_APN.Text = "" Then
                                SS_field = ComboBox_Sheet_Set_APN.Text
                            End If
                            sw.Write(vbCrLf & "29,APN (Object Data)," & .Text & "," & SS_field)
                        End If
                    End With

                    With ComboBox_access_road_length
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_Sheet_Set_access_road_length.Text = "" Then
                                SS_field = ComboBox_Sheet_Set_access_road_length.Text
                            End If
                            Dim Round1 As Integer = 2
                            If IsNumeric(TextBox_round_access_road_length.Text) = True Then
                                Round1 = Abs(CInt(TextBox_round_access_road_length.Text))
                            End If
                            sw.Write(vbCrLf & "30,Access road length (Object Data)," & .Text & "," & SS_field & "," & Round1.ToString)
                        End If
                    End With

                    With ComboBox_Section_TWP_Range
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_Sheet_Set_sec_twp_range.Text = "" Then
                                SS_field = ComboBox_Sheet_Set_sec_twp_range.Text
                            End If
                            sw.Write(vbCrLf & "31,SECTION - TWP - RANGE (Object Data)," & .Text & "," & SS_field)
                        End If
                    End With

                    With ComboBox_owner
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_Sheet_Set_owner.Text = "" Then
                                SS_field = ComboBox_Sheet_Set_owner.Text
                            End If
                            sw.Write(vbCrLf & "32,OWNER (Object Data)," & .Text & "," & SS_field)
                        End If
                    End With

                    With ComboBox_crossing_length
                        If Not .Text = "" Then
                            Dim SS_field1 As String = .Text
                            If Not ComboBox_Sheet_Set_crossing_len_FT.Text = "" Then
                                SS_field1 = ComboBox_Sheet_Set_crossing_len_FT.Text
                            End If
                            Dim Round1 As Integer = 2
                            If IsNumeric(TextBox_round_crossing_length_ft.Text) = True Then
                                Round1 = Abs(CInt(TextBox_round_crossing_length_ft.Text))
                            End If

                            Dim SS_field2 As String = ""
                            If Not ComboBox_Sheet_Set_crossing_len_rod.Text = "" Then
                                SS_field2 = ComboBox_Sheet_Set_crossing_len_rod.Text
                            End If
                            Dim Round2 As Integer = 2
                            If IsNumeric(TextBox_round_crossing_length_rod.Text) = True Then
                                Round2 = Abs(CInt(TextBox_round_crossing_length_rod.Text))
                            End If

                            sw.Write(vbCrLf & "33,CL CROSSING LENGTH (Object Data)," & .Text & "," & SS_field1 & "," & Round1.ToString)
                            sw.Write(vbCrLf & "33A,CL CROSSING LENGTH (SHEET SET FIELD RODS),," & SS_field2 & "," & Round2.ToString)
                        End If
                    End With


                    Dim nedYesNO As String

                    If CheckBox_USE_ned_for_naming.Checked = True Then
                        nedYesNO = "YES"
                    Else
                        nedYesNO = "NO"
                    End If

                    sw.Write(vbCrLf & "34,Use NED for naming," & nedYesNO)

                    With ComboBox_pipe_diameter
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_Sheet_Set_pipe_DIAMETER.Text = "" Then
                                SS_field = ComboBox_Sheet_Set_pipe_DIAMETER.Text
                            End If
                            sw.Write(vbCrLf & "35,PIPE DIAMETER (Object Data)," & .Text & "," & SS_field)
                        End If
                    End With

                    With ComboBox_pipe_name
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_Sheet_Set_pipe_name.Text = "" Then
                                SS_field = ComboBox_Sheet_Set_pipe_name.Text
                            End If
                            sw.Write(vbCrLf & "36,PIPE name (Object Data)," & .Text & "," & SS_field)
                        End If
                    End With

                    With ComboBox_pipe_segment
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_Sheet_Set_pipe_segment.Text = "" Then
                                SS_field = ComboBox_Sheet_Set_pipe_segment.Text
                            End If
                            sw.Write(vbCrLf & "37,PIPE segment (Object Data)," & .Text & "," & SS_field)
                        End If
                    End With

                    With ComboBox_area_EX_E
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_Sheet_Set_ex_e.Text = "" Then
                                SS_field = ComboBox_Sheet_Set_ex_e.Text
                            End If
                            Dim Round1 As Integer = 2
                            If IsNumeric(TextBox_round_EX_E.Text) = True Then
                                Round1 = Abs(CInt(TextBox_round_EX_E.Text))
                            End If

                            sw.Write(vbCrLf & "38,AREA Existing Easement (Object Data)," & .Text & "," & SS_field & "," & Round1)
                        End If
                    End With

                    With ComboBox_area_P_E
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_Sheet_Set_P_E.Text = "" Then
                                SS_field = ComboBox_Sheet_Set_P_E.Text
                            End If
                            Dim Round1 As Integer = 2
                            If IsNumeric(TextBox_round_P_E.Text) = True Then
                                Round1 = Abs(CInt(TextBox_round_P_E.Text))
                            End If

                            sw.Write(vbCrLf & "39,AREA Permanent Easement (Object Data)," & .Text & "," & SS_field & "," & Round1)
                        End If
                    End With

                    With ComboBox_area_TWS
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_Sheet_Set_TWS.Text = "" Then
                                SS_field = ComboBox_Sheet_Set_TWS.Text
                            End If
                            Dim Round1 As Integer = 2
                            If IsNumeric(TextBox_round_TWS.Text) = True Then
                                Round1 = Abs(CInt(TextBox_round_TWS.Text))
                            End If

                            sw.Write(vbCrLf & "40,AREA Temporary Work Space (Object Data)," & .Text & "," & SS_field & "," & Round1)
                        End If
                    End With

                    With ComboBox_area_ATWS
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_Sheet_Set_atws.Text = "" Then
                                SS_field = ComboBox_Sheet_Set_atws.Text
                            End If
                            Dim Round1 As Integer = 2
                            If IsNumeric(TextBox_round_ATWS.Text) = True Then
                                Round1 = Abs(CInt(TextBox_round_ATWS.Text))
                            End If

                            sw.Write(vbCrLf & "41,AREA Additional Temporary Work Space (Object Data)," & .Text & "," & SS_field & "," & Round1)
                        End If
                    End With

                    With ComboBox_area_A_R
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_Sheet_Set_a_R.Text = "" Then
                                SS_field = ComboBox_Sheet_Set_a_R.Text
                            End If
                            Dim Round1 As Integer = 2
                            If IsNumeric(TextBox_round_A_R.Text) = True Then
                                Round1 = Abs(CInt(TextBox_round_A_R.Text))
                            End If

                            sw.Write(vbCrLf & "42,AREA Access Road (Object Data)," & .Text & "," & SS_field & "," & Round1)
                        End If
                    End With

                    With ComboBox_area_WARE
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_Sheet_Set_ware.Text = "" Then
                                SS_field = ComboBox_Sheet_Set_ware.Text
                            End If
                            Dim Round1 As Integer = 2
                            If IsNumeric(TextBox_round_WARE.Text) = True Then
                                Round1 = Abs(CInt(TextBox_round_WARE.Text))
                            End If

                            sw.Write(vbCrLf & "43,AREA Wareyard (Object Data)," & .Text & "," & SS_field & "," & Round1)
                        End If
                    End With

                    With ComboBox_area_TWS_ABD
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_Sheet_Set_tws_abd.Text = "" Then
                                SS_field = ComboBox_Sheet_Set_tws_abd.Text
                            End If
                            Dim Round1 As Integer = 2
                            If IsNumeric(TextBox_round_TWS_ABD.Text) = True Then
                                Round1 = Abs(CInt(TextBox_round_TWS_ABD.Text))
                            End If

                            sw.Write(vbCrLf & "44,AREA Temporary Workspace Abandoned (Object Data)," & .Text & "," & SS_field & "," & Round1)
                        End If
                    End With

                    With ComboBox_area_TWS_PD
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_Sheet_Set_tws_PD.Text = "" Then
                                SS_field = ComboBox_Sheet_Set_tws_PD.Text
                            End If
                            Dim Round1 As Integer = 2
                            If IsNumeric(TextBox_round_TWS_PD.Text) = True Then
                                Round1 = Abs(CInt(TextBox_round_TWS_PD.Text))
                            End If

                            sw.Write(vbCrLf & "45,AREA Temporary Workspace Previously Disturbed (Object Data)," & .Text & "," & SS_field & "," & Round1)
                        End If
                    End With


                    Dim convert1 As String

                    If CheckBox_convert_sqft_to_acres.Checked = True Then
                        convert1 = "YES"
                    Else
                        convert1 = "NO"
                    End If

                    sw.Write(vbCrLf & "46,Convert SQFT to acres," & convert1)







                    If CheckBox_ignore_deflections_less_0_5.Checked = False Then
                        sw.Write(vbCrLf & "47,Ignore points with deflection less than 0.5,NO")
                    Else
                        sw.Write(vbCrLf & "47,Ignore points with deflection less than 0.5,YES")
                    End If

                    If CheckBox_add_bearing_and_distances.Checked = False Then
                        sw.Write(vbCrLf & "48,Add Bearing and distances along CL,NO")
                    Else
                        sw.Write(vbCrLf & "48,Add Bearing and distances along CL|,YES")
                    End If

                    If CheckBox_display_stations.Checked = False Then
                        sw.Write(vbCrLf & "49,Display stations along CL,NO")
                    Else
                        sw.Write(vbCrLf & "49,Display stations along CL,YES")
                    End If

                    If CheckBox_rotate_to_north.Checked = False Then
                        sw.Write(vbCrLf & "50,Set View to North,NO")
                    Else
                        sw.Write(vbCrLf & "50,Set View to North,YES")
                    End If

                    If Not Station_at_point_layer = "" Then
                        sw.Write(vbCrLf & "51,Station at Point Layer," & Station_at_point_layer)
                    End If

                    If Not Centerline_stationing_layer = "" Then
                        sw.Write(vbCrLf & "52,Centerline Stationing Layer," & Centerline_stationing_layer)
                    End If

                    If Not Bearing_and_distance_layer = "" Then
                        sw.Write(vbCrLf & "53,Bearing and Distance Layer," & Bearing_and_distance_layer)
                    End If
                    If Not Tie_distance_layer = "" Then
                        sw.Write(vbCrLf & "54,Tie Distance Layer," & Tie_distance_layer)
                    End If
                    If Not Arc_leader_polyline_layer = "" Then
                        sw.Write(vbCrLf & "55,Arc Leader Polyline Layer," & Arc_leader_polyline_layer)
                    End If
                    If Not Layer_name_Main_Viewport = "" Then
                        sw.Write(vbCrLf & "56,Layer name Main Viewport," & Layer_name_Main_Viewport)
                    End If

                    If Not Layer_name_locus_VP = "" Then
                        sw.Write(vbCrLf & "57,Layer name Locus Viewport," & Layer_name_locus_VP)
                    End If

                    If Not Layer_name_Blocks = "" Then
                        sw.Write(vbCrLf & "58,Layer name for Blocks," & Layer_name_Blocks)
                    End If

                    If Not Layer_name_text = "" Then
                        sw.Write(vbCrLf & "59,Layer name for Text," & Layer_name_text)
                    End If

                    If Not Layer_poly_parcela_new = "" Then
                        sw.Write(vbCrLf & "60,Layer name for new parcel," & Layer_poly_parcela_new)
                    End If

                    If Not Layer_hatch_locus = "" Then
                        sw.Write(vbCrLf & "61,Layer name for Locus hatch," & Layer_hatch_locus)
                    End If


                    With ComboBox_state_update
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_set_state_update.Text = "" Then
                                SS_field = ComboBox_sheet_set_state_update.Text
                            End If
                            Dim Suffix As String = TextBox_suffix_state_update.Text
                            Dim Prefix As String = TextBox_prefix_state_update.Text
                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix
                                End If
                            End If
                            sw.Write(vbCrLf & "122,State UPDATE (Object Data)," & .Text & "," & SS_field & Extra1)
                        End If
                    End With


                    Dim Abr2 As String
                    If CheckBox_state_abreviation_UPDATE.Checked = True Then
                        Abr2 = "YES"
                    Else
                        Abr2 = "NO"
                    End If

                    sw.Write(vbCrLf & "122A,Use State abbreviations UPDATE," & Abr2)

                    With ComboBox_county_update
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_set_county_update.Text = "" Then
                                SS_field = ComboBox_sheet_set_county_update.Text
                            End If
                            Dim Suffix As String = TextBox_suffix_county_update.Text
                            Dim Prefix As String = TextBox_prefix_county_update.Text
                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix
                                End If
                            End If
                            sw.Write(vbCrLf & "123,County UPDATE (Object Data)," & .Text & "," & SS_field & Extra1)
                        End If
                    End With

                    With ComboBox_town_update
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_set_town_update.Text = "" Then
                                SS_field = ComboBox_sheet_set_town_update.Text
                            End If
                            Dim Suffix As String = TextBox_suffix_town_update.Text
                            Dim Prefix As String = TextBox_prefix_town_update.Text
                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix
                                End If
                            End If
                            sw.Write(vbCrLf & "124,Town UPDATE (Object Data)," & .Text & "," & SS_field & Extra1)
                        End If
                    End With

                    With ComboBox_linelist_update
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_set_linelist_update.Text = "" Then
                                SS_field = ComboBox_sheet_set_linelist_update.Text
                            End If
                            Dim Suffix As String = TextBox_suffix_linelist_update.Text
                            Dim Prefix As String = TextBox_prefix_linelist_update.Text
                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix
                                End If
                            End If
                            sw.Write(vbCrLf & "125,Linelist update (Object Data)," & .Text & "," & SS_field & Extra1)
                        End If
                    End With

                    With ComboBox_HMM_ID_update
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_set_HMMID_update.Text = "" Then
                                SS_field = ComboBox_sheet_set_HMMID_update.Text
                            End If


                            sw.Write(vbCrLf & "126,Parcel ID update (Object Data)," & .Text & "," & SS_field)
                        End If
                    End With

                    With ComboBox_MBL_update
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_set_MBL_update.Text = "" Then
                                SS_field = ComboBox_sheet_set_MBL_update.Text
                            End If
                            Dim Suffix As String = TextBox_suffix_mbl_update.Text
                            Dim Prefix As String = TextBox_prefix_mbl_update.Text
                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix
                                End If
                            End If

                            sw.Write(vbCrLf & "127,MBL update (Object Data)," & .Text & "," & SS_field & Extra1)
                        End If
                    End With

                    With ComboBox_deed_page_update
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_set_deed_page_update.Text = "" Then
                                SS_field = ComboBox_sheet_set_deed_page_update.Text
                            End If
                            Dim Suffix As String = TextBox_suffix_deed_page_update.Text
                            Dim Prefix As String = TextBox_prefix_deed_page_update.Text
                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix
                                End If
                            End If
                            sw.Write(vbCrLf & "128,DEED BOOK PAGE update (Object Data)," & .Text & "," & SS_field & Extra1)
                        End If
                    End With

                    With ComboBox_APN_update
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_set_apn_update.Text = "" Then
                                SS_field = ComboBox_sheet_set_apn_update.Text
                            End If
                            Dim Suffix As String = TextBox_suffix_apn_update.Text
                            Dim Prefix As String = TextBox_prefix_apn_update.Text
                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix
                                End If
                            End If
                            sw.Write(vbCrLf & "129,APN update (Object Data)," & .Text & "," & SS_field & Extra1)
                        End If
                    End With

                    With ComboBox_access_road_length_update
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_set_access_road_length_update.Text = "" Then
                                SS_field = ComboBox_sheet_set_access_road_length_update.Text
                            End If
                            Dim Round1 As Integer = 2
                            If IsNumeric(TextBox_round_access_road_length_update.Text) = True Then
                                Round1 = Abs(CInt(TextBox_round_access_road_length_update.Text))
                            End If

                            Dim Suffix As String = TextBox_suffix_a_r_length_update.Text
                            Dim Prefix As String = TextBox_prefix_a_r_length_update.Text

                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,," & Prefix
                                End If
                            End If

                            sw.Write(vbCrLf & "130,Access road length update (Object Data)," & .Text & "," & SS_field & "," & Round1.ToString & Extra1)
                        End If
                    End With

                    With ComboBox_Section_TWP_Range_update
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_set_sec_twp_range_update.Text = "" Then
                                SS_field = ComboBox_sheet_set_sec_twp_range_update.Text
                            End If


                            Dim Suffix As String = TextBox_suffix_sec_twp_rge_update.Text
                            Dim Prefix As String = TextBox_prefix_sec_twp_rge_update.Text

                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix
                                End If
                            End If

                            sw.Write(vbCrLf & "131,SECTION - TWP - RANGE update (Object Data)," & .Text & "," & SS_field & Extra1)
                        End If
                    End With

                    With ComboBox_owner_update
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_set_owner_update.Text = "" Then
                                SS_field = ComboBox_sheet_set_owner_update.Text
                            End If



                            Dim Suffix As String = TextBox_suffix_owner_update.Text
                            Dim Prefix As String = TextBox_prefix_owner_update.Text

                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix
                                End If
                            End If

                            sw.Write(vbCrLf & "132,OWNER update (Object Data)," & .Text & "," & SS_field & Extra1)
                        End If
                    End With

                    With ComboBox_crossing_length_update
                        If Not .Text = "" Then
                            Dim SS_field1 As String = .Text
                            If Not ComboBox_sheet_set_crosing_length_ft_update.Text = "" Then
                                SS_field1 = ComboBox_sheet_set_crosing_length_ft_update.Text
                            End If
                            Dim Round1 As Integer = 2
                            If IsNumeric(TextBox_round_crossing_length_ft_update.Text) = True Then
                                Round1 = Abs(CInt(TextBox_round_crossing_length_ft_update.Text))
                            End If

                            Dim SS_field2 As String = ""
                            If Not ComboBox_sheet_set_crosing_length_rod_update.Text = "" Then
                                SS_field2 = ComboBox_sheet_set_crosing_length_rod_update.Text
                            End If
                            Dim Round2 As Integer = 2
                            If IsNumeric(TextBox_round_crossing_length_rod_update.Text) = True Then
                                Round2 = Abs(CInt(TextBox_round_crossing_length_rod_update.Text))
                            End If

                            Dim Suffix1 As String = TextBox_suffix_crossing_length_ft_update.Text
                            Dim Prefix1 As String = TextBox_prefix_crossing_length_ft_update.Text

                            Dim Extra1 As String = ""
                            If Not Suffix1 = "" Or Not Prefix1 = "" Then
                                If Not Suffix1 = "" And Not Prefix1 = "" Then
                                    Extra1 = "," & Prefix1 & "," & Suffix1
                                ElseIf Not Suffix1 = "" Then
                                    Extra1 = ",," & Suffix1
                                ElseIf Not Prefix1 = "" Then
                                    Extra1 = "," & Prefix1
                                End If
                            End If

                            sw.Write(vbCrLf & "133,CL CROSSING LENGTH update (Object Data)," & .Text & "," & SS_field1 & "," & Round1.ToString _
                                     & ",," & Extra1)
                            Dim Suffix2 As String = TextBox_suffix_crossing_length_rod_update.Text
                            Dim Prefix2 As String = TextBox_prefix_crossing_length_rod_update.Text

                            Dim Extra2 As String = ""
                            If Not Suffix2 = "" Or Not Prefix2 = "" Then
                                If Not Suffix2 = "" And Not Prefix2 = "" Then
                                    Extra2 = "," & Prefix2 & "," & Suffix2
                                ElseIf Not Suffix2 = "" Then
                                    Extra2 = ",," & Suffix2
                                ElseIf Not Prefix2 = "" Then
                                    Extra2 = "," & Prefix2
                                End If
                            End If

                            sw.Write(vbCrLf & "133A,CL CROSSING LENGTH update (SHEET SET FIELD rods),," & SS_field2 & "," & Round2.ToString & ",," & Extra2)
                        End If
                    End With

                    With ComboBox_area_EX_E_update
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_set_ex_e_update.Text = "" Then
                                SS_field = ComboBox_sheet_set_ex_e_update.Text
                            End If
                            Dim Round1 As Integer = 2
                            If IsNumeric(TextBox_round_EX_E_update.Text) = True Then
                                Round1 = Abs(CInt(TextBox_round_EX_E_update.Text))
                            End If

                            Dim Suffix As String = TextBox_suffix_AREA_EX_E_update.Text
                            Dim Prefix As String = TextBox_prefix_AREA_EX_E_update.Text

                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,," & Prefix
                                End If
                            End If

                            sw.Write(vbCrLf & "138,AREA Existing Easement UPDATE (Object Data)," & .Text & "," & SS_field & "," & Round1 & Extra1)
                        End If
                    End With

                    With ComboBox_area_P_E_update
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_set_p_e_update.Text = "" Then
                                SS_field = ComboBox_sheet_set_p_e_update.Text
                            End If
                            Dim Round1 As Integer = 2
                            If IsNumeric(TextBox_round_P_E_update.Text) = True Then
                                Round1 = Abs(CInt(TextBox_round_P_E_update.Text))
                            End If

                            Dim Suffix As String = TextBox_suffix_AREA_P_E_update.Text
                            Dim Prefix As String = TextBox_prefix_AREA_P_E_update.Text

                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,," & Prefix
                                End If
                            End If

                            sw.Write(vbCrLf & "139,AREA Permanent Easement UPDATE (Object Data)," & .Text & "," & SS_field & "," & Round1 & Extra1)
                        End If
                    End With

                    With ComboBox_area_TWS_update
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_set_tws_update.Text = "" Then
                                SS_field = ComboBox_sheet_set_tws_update.Text
                            End If
                            Dim Round1 As Integer = 2
                            If IsNumeric(TextBox_round_TWS_update.Text) = True Then
                                Round1 = Abs(CInt(TextBox_round_TWS_update.Text))
                            End If

                            Dim Suffix As String = TextBox_suffix_AREA_TWS_update.Text
                            Dim Prefix As String = TextBox_prefix_AREA_TWS_update.Text

                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,," & Prefix
                                End If
                            End If

                            sw.Write(vbCrLf & "140,AREA Temporary Work Space UPDATE (Object Data)," & .Text & "," & SS_field & "," & Round1 & Extra1)
                        End If
                    End With

                    With ComboBox_area_ATWS_update
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_set_atws_update.Text = "" Then
                                SS_field = ComboBox_sheet_set_atws_update.Text
                            End If
                            Dim Round1 As Integer = 2
                            If IsNumeric(TextBox_round_ATWS_update.Text) = True Then
                                Round1 = Abs(CInt(TextBox_round_ATWS_update.Text))
                            End If

                            Dim Suffix As String = TextBox_suffix_AREA_TWS_update.Text
                            Dim Prefix As String = TextBox_prefix_AREA_TWS_update.Text

                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,," & Prefix
                                End If
                            End If

                            sw.Write(vbCrLf & "141,AREA Additional Temporary Work Space UPDATE (Object Data)," & .Text & "," & SS_field & "," & Round1 & Extra1)
                        End If
                    End With

                    With ComboBox_area_A_R_update
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_set_a_r_update.Text = "" Then
                                SS_field = ComboBox_sheet_set_a_r_update.Text
                            End If
                            Dim Round1 As Integer = 2
                            If IsNumeric(TextBox_round_A_R_update.Text) = True Then
                                Round1 = Abs(CInt(TextBox_round_A_R_update.Text))
                            End If

                            Dim Suffix As String = TextBox_suffix_AREA_A_R_update.Text
                            Dim Prefix As String = TextBox_prefix_AREA_A_R_update.Text

                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,," & Prefix
                                End If
                            End If

                            sw.Write(vbCrLf & "142,AREA Access Road UPDATE (Object Data)," & .Text & "," & SS_field & "," & Round1 & Extra1)
                        End If
                    End With

                    With ComboBox_area_WARE_update
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_set_ware_update.Text = "" Then
                                SS_field = ComboBox_sheet_set_ware_update.Text
                            End If
                            Dim Round1 As Integer = 2
                            If IsNumeric(TextBox_round_WARE_update.Text) = True Then
                                Round1 = Abs(CInt(TextBox_round_WARE_update.Text))
                            End If

                            Dim Suffix As String = TextBox_suffix_AREA_WARE_update.Text
                            Dim Prefix As String = TextBox_prefix_AREA_WARE_update.Text

                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,," & Prefix
                                End If
                            End If

                            sw.Write(vbCrLf & "143,AREA Wareyard UPDATE (Object Data)," & .Text & "," & SS_field & "," & Round1 & Extra1)
                        End If
                    End With

                    With ComboBox_area_TWS_ABD_update
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_set_tws_ABD_update.Text = "" Then
                                SS_field = ComboBox_sheet_set_tws_ABD_update.Text
                            End If
                            Dim Round1 As Integer = 2
                            If IsNumeric(TextBox_round_TWS_ABD_update.Text) = True Then
                                Round1 = Abs(CInt(TextBox_round_TWS_ABD_update.Text))
                            End If

                            Dim Suffix As String = TextBox_suffix_AREA_TWS_ABD_update.Text
                            Dim Prefix As String = TextBox_prefix_AREA_TWS_ABD_update.Text

                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,," & Prefix
                                End If
                            End If

                            sw.Write(vbCrLf & "144,AREA Temporary Workspace Abandoned UPDATE (Object Data)," & .Text & "," & SS_field & "," & Round1 & Extra1)
                        End If
                    End With

                    With ComboBox_area_TWS_PD_update
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_set_TWS_PD_update.Text = "" Then
                                SS_field = ComboBox_sheet_set_TWS_PD_update.Text
                            End If
                            Dim Round1 As Integer = 2
                            If IsNumeric(TextBox_round_TWS_PD_update.Text) = True Then
                                Round1 = Abs(CInt(TextBox_round_TWS_PD_update.Text))
                            End If


                            Dim Suffix As String = TextBox_suffix_AREA_TWS_PD_update.Text
                            Dim Prefix As String = TextBox_prefix_AREA_TWS_PD_update.Text

                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,," & Prefix
                                End If
                            End If

                            sw.Write(vbCrLf & "145,AREA Temporary Workspace Previously Disturbed update (Object Data)," & .Text & "," & SS_field & "," & Round1 & Extra1)
                        End If
                    End With


                    Dim convert2 As String

                    If CheckBox_convert_sqft_to_acres_UPDATE.Checked = True Then
                        convert2 = "YES"
                    Else
                        convert2 = "NO"
                    End If

                    sw.Write(vbCrLf & "146,Convert SQFT to acres," & convert2)

                    With ComboBox_shape_user1
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_user1.Text = "" Then
                                SS_field = ComboBox_sheet_user1.Text
                            End If
                            Dim Suffix As String = TextBox_suffix_user1.Text
                            Dim Prefix As String = TextBox_prefix_user1.Text

                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix
                                End If
                            End If

                            sw.Write(vbCrLf & "150,USER1 update (Object Data)," & .Text & "," & SS_field & Extra1)
                        End If
                    End With

                    With ComboBox_shape_user2
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_user2.Text = "" Then
                                SS_field = ComboBox_sheet_user2.Text
                            End If
                            Dim Suffix As String = TextBox_suffix_user2.Text
                            Dim Prefix As String = TextBox_prefix_user2.Text

                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix
                                End If
                            End If
                            sw.Write(vbCrLf & "151,USER2 update (Object Data)," & .Text & "," & SS_field & Extra1)
                        End If
                    End With
                    With ComboBox_shape_user3
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_user3.Text = "" Then
                                SS_field = ComboBox_sheet_user3.Text
                            End If
                            Dim Suffix As String = TextBox_suffix_user3.Text
                            Dim Prefix As String = TextBox_prefix_user3.Text

                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix
                                End If
                            End If
                            sw.Write(vbCrLf & "152,USER3 update (Object Data)," & .Text & "," & SS_field & Extra1)
                        End If
                    End With
                    With ComboBox_shape_user4
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_user4.Text = "" Then
                                SS_field = ComboBox_sheet_user4.Text
                            End If

                            Dim Suffix As String = TextBox_suffix_user4.Text
                            Dim Prefix As String = TextBox_prefix_user4.Text

                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix
                                End If
                            End If

                            sw.Write(vbCrLf & "153,USER4 update (Object Data)," & .Text & "," & SS_field & Extra1)
                        End If
                    End With


                    With ComboBox_shape_user5
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_user5.Text = "" Then
                                SS_field = ComboBox_sheet_user5.Text
                            End If

                            Dim Suffix As String = TextBox_suffix_user5.Text
                            Dim Prefix As String = TextBox_prefix_user5.Text

                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix
                                End If
                            End If

                            sw.Write(vbCrLf & "154,USER5 update (Object Data)," & .Text & "," & SS_field & Extra1)
                        End If
                    End With

                    With ComboBox_shape_user6
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_user6.Text = "" Then
                                SS_field = ComboBox_sheet_user6.Text
                            End If

                            Dim Suffix As String = TextBox_suffix_user6.Text
                            Dim Prefix As String = TextBox_prefix_user6.Text

                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix
                                End If
                            End If

                            sw.Write(vbCrLf & "155,USER6 update (Object Data)," & .Text & "," & SS_field & Extra1)
                        End If
                    End With

                    With ComboBox_shape_user7
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_user7.Text = "" Then
                                SS_field = ComboBox_sheet_user7.Text
                            End If
                            Dim Suffix As String = TextBox_suffix_user7.Text
                            Dim Prefix As String = TextBox_prefix_user7.Text
                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix
                                End If
                            End If
                            sw.Write(vbCrLf & "156,USER7 update (Object Data)," & .Text & "," & SS_field & Extra1)
                        End If
                    End With

                    With ComboBox_shape_user8
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_user8.Text = "" Then
                                SS_field = ComboBox_sheet_user8.Text
                            End If
                            Dim Suffix As String = TextBox_suffix_user8.Text
                            Dim Prefix As String = TextBox_prefix_user8.Text
                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix
                                End If
                            End If
                            sw.Write(vbCrLf & "157,USER8 update (Object Data)," & .Text & "," & SS_field & Extra1)
                        End If
                    End With

                    With ComboBox_shape_user9
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_user9.Text = "" Then
                                SS_field = ComboBox_sheet_user9.Text
                            End If
                            Dim Suffix As String = TextBox_suffix_user9.Text
                            Dim Prefix As String = TextBox_prefix_user9.Text
                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix
                                End If
                            End If
                            sw.Write(vbCrLf & "158,USER9 update (Object Data)," & .Text & "," & SS_field & Extra1)
                        End If
                    End With

                    With ComboBox_shape_user10
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_user10.Text = "" Then
                                SS_field = ComboBox_sheet_user10.Text
                            End If
                            Dim Suffix As String = TextBox_suffix_user10.Text
                            Dim Prefix As String = TextBox_prefix_user10.Text
                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix
                                End If
                            End If
                            sw.Write(vbCrLf & "159,USER10 update (Object Data)," & .Text & "," & SS_field & Extra1)
                        End If
                    End With

                    With ComboBox_shape_user11
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_user11.Text = "" Then
                                SS_field = ComboBox_sheet_user11.Text
                            End If
                            Dim Suffix As String = TextBox_suffix_user11.Text
                            Dim Prefix As String = TextBox_prefix_user11.Text
                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix
                                End If
                            End If
                            sw.Write(vbCrLf & "160,USER11 update (Object Data)," & .Text & "," & SS_field & Extra1)
                        End If
                    End With

                    With ComboBox_shape_user12
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_user12.Text = "" Then
                                SS_field = ComboBox_sheet_user12.Text
                            End If
                            Dim Suffix As String = TextBox_suffix_user12.Text
                            Dim Prefix As String = TextBox_prefix_user12.Text
                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix
                                End If
                            End If
                            sw.Write(vbCrLf & "161,USER12 update (Object Data)," & .Text & "," & SS_field & Extra1)
                        End If
                    End With

                    With ComboBox_shape_user13
                        If Not .Text = "" Then
                            Dim SS_field As String = .Text
                            If Not ComboBox_sheet_user13.Text = "" Then
                                SS_field = ComboBox_sheet_user13.Text
                            End If
                            Dim Suffix As String = TextBox_suffix_user13.Text
                            Dim Prefix As String = TextBox_prefix_user13.Text
                            Dim Extra1 As String = ""
                            If Not Suffix = "" Or Not Prefix = "" Then
                                If Not Suffix = "" And Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix & "," & Suffix
                                ElseIf Not Suffix = "" Then
                                    Extra1 = ",,,,," & Suffix
                                ElseIf Not Prefix = "" Then
                                    Extra1 = ",,,," & Prefix
                                End If
                            End If
                            sw.Write(vbCrLf & "162,USER13 update (Object Data)," & .Text & "," & SS_field & Extra1)
                        End If
                    End With


                    If Data_table_layers.Rows.Count > 0 Then
                        For i = 0 To Data_table_layers.Rows.Count - 1

                            If IsDBNull(Data_table_layers.Rows(i).Item("VIEWPORT_TYPE")) = False And _
                                IsDBNull(Data_table_layers.Rows(i).Item("LAYER_NAME")) = False And _
                                IsDBNull(Data_table_layers.Rows(i).Item("THAW_FREEZE")) = False Then

                                Dim Viewport_type As String = Data_table_layers.Rows(i).Item("VIEWPORT_TYPE")
                                Dim Nume_layer As String = Data_table_layers.Rows(i).Item("LAYER_NAME")
                                Dim Thaw_freeze As String = Data_table_layers.Rows(i).Item("THAW_FREEZE")
                                Dim Color_index As String = ""

                                If IsDBNull(Data_table_layers.Rows(i).Item("COLOR_INDEX")) = False Then
                                    Color_index = Data_table_layers.Rows(i).Item("COLOR_INDEX").ToString
                                End If
                                sw.Write(vbCrLf & (1000 + i).ToString & "," & Viewport_type & "," & Nume_layer & "," & Color_index & "," & Thaw_freeze)
                            End If



                        Next
                    End If

                End Using

            Catch ex As System.Exception
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_load_object_data_FROM_PARCELS_Click(sender As Object, e As EventArgs) Handles Button_load_object_data_from_parcels.Click

        Dim Empty_array() As ObjectId
        If Freeze_operations = False Then
            Freeze_operations = True


            Dim BaseMap_drawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = BaseMap_drawing.Editor
            Editor1.SetImpliedSelection(Empty_array)
            Try
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()


                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptEntityResult

                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select entity containing object data:")

                Rezultat1 = Editor1.GetEntity(Object_Prompt)

                If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.WriteMessage(vbLf & "Command:")
                    Editor1.SetImpliedSelection(Empty_array)
                    Freeze_operations = False
                    Exit Sub
                End If

                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                    If IsNothing(Rezultat1) = False Then

                        Using Lock_dwg As DocumentLock = BaseMap_drawing.LockDocument
                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = BaseMap_drawing.TransactionManager.StartTransaction
                                Dim Ent1 As Entity = Trans1.GetObject(Rezultat1.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                                Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables
                                Dim Id1 As ObjectId = Ent1.ObjectId

                                Using Records1 As Autodesk.Gis.Map.ObjectData.Records = Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                    If IsNothing(Records1) = False Then
                                        If Records1.Count > 0 Then
                                            With ComboBox_plat_name
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_state
                                                .Items.Clear()
                                                .Text = ""
                                            End With
                                            With ComboBox_county
                                                .Items.Clear()
                                                .Text = ""
                                            End With
                                            With ComboBox_town
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_linelist
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_HMM_ID
                                                .Items.Clear()
                                                .Text = ""
                                            End With
                                            With ComboBox_MBL
                                                .Items.Clear()
                                                .Text = ""
                                            End With
                                            With ComboBox_deed_page
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_APN
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_access_road_length
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_Section_TWP_Range
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_owner
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_crossing_length
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_area_EX_E
                                                .Items.Clear()
                                                .Text = ""
                                            End With
                                            With ComboBox_area_P_E
                                                .Items.Clear()
                                                .Text = ""
                                            End With
                                            With ComboBox_area_TWS
                                                .Items.Clear()
                                                .Text = ""
                                            End With
                                            With ComboBox_area_ATWS
                                                .Items.Clear()
                                                .Text = ""
                                            End With
                                            With ComboBox_area_A_R
                                                .Items.Clear()
                                                .Text = ""
                                            End With
                                            With ComboBox_area_WARE
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_area_TWS_ABD
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_area_TWS_PD
                                                .Items.Clear()
                                                .Text = ""
                                            End With





                                            Dim Record1 As Autodesk.Gis.Map.ObjectData.Record


                                            For Each Record1 In Records1
                                                Dim Tabla1 As Autodesk.Gis.Map.ObjectData.Table
                                                Tabla1 = Tables1(Record1.TableName)


                                                Dim Field_defs1 As Autodesk.Gis.Map.ObjectData.FieldDefinitions
                                                Field_defs1 = Tabla1.FieldDefinitions



                                                For i = 0 To Record1.Count - 1
                                                    Dim Field_def1 As Autodesk.Gis.Map.ObjectData.FieldDefinition
                                                    Field_def1 = Field_defs1(i)
                                                    With ComboBox_plat_name
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With

                                                    With ComboBox_state
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With
                                                    With ComboBox_county
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With
                                                    With ComboBox_town
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With



                                                    With ComboBox_linelist
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With
                                                    With ComboBox_HMM_ID
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With
                                                    With ComboBox_MBL
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With


                                                    With ComboBox_deed_page
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With

                                                    With ComboBox_APN
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With

                                                    With ComboBox_access_road_length
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With

                                                    With ComboBox_Section_TWP_Range
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With

                                                    With ComboBox_owner
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With
                                                    With ComboBox_crossing_length
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With

                                                    With ComboBox_area_EX_E
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With
                                                    With ComboBox_area_P_E
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With
                                                    With ComboBox_area_TWS
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With

                                                    With ComboBox_area_ATWS
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With

                                                    With ComboBox_area_A_R
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With
                                                    With ComboBox_area_WARE
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With

                                                    With ComboBox_area_TWS_ABD
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With

                                                    With ComboBox_area_TWS_PD
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With



                                                Next



                                            Next

                                            With ComboBox_plat_name
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("APN2") = True Then
                                                        .SelectedIndex = .Items.IndexOf("APN2")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_county
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("PHYSCOUNTY") = True Then
                                                        .SelectedIndex = .Items.IndexOf("PHYSCOUNTY")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_state
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("PHYSSTATE") = True Then
                                                        .SelectedIndex = .Items.IndexOf("PHYSSTATE")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With


                                            With ComboBox_town
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("SGC_TOWN") = True Then
                                                        .SelectedIndex = .Items.IndexOf("SGC_TOWN")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_owner
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("SGC_OWNER") = True Then
                                                        .SelectedIndex = .Items.IndexOf("SGC_OWNER")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With





                                            With ComboBox_linelist
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("LINELIST") = True Then
                                                        .SelectedIndex = .Items.IndexOf("LINELIST")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_HMM_ID
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("PARCEL_HMM") = True Then
                                                        .SelectedIndex = .Items.IndexOf("PARCEL_HMM")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_MBL
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("SGC_MBL") = True Then
                                                        .SelectedIndex = .Items.IndexOf("SGC_MBL")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_area_ATWS
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("ATWSSQFT") = True Then
                                                        .SelectedIndex = .Items.IndexOf("ATWSSQFT")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_area_EX_E
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("E_EASE_SF") = True Then
                                                        .SelectedIndex = .Items.IndexOf("E_EASE_SF")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_area_P_E
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("PERMSQFT") = True Then
                                                        .SelectedIndex = .Items.IndexOf("PERMSQFT")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_area_TWS
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("TWSSQFT") = True Then
                                                        .SelectedIndex = .Items.IndexOf("TWSSQFT")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_area_A_R
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("ACSSQFT") = True Then
                                                        .SelectedIndex = .Items.IndexOf("ACSSQFT")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With


                                            With ComboBox_area_WARE
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("WARESQFT") = True Then
                                                        .SelectedIndex = .Items.IndexOf("WARESQFT")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_deed_page
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("DEED_BK_PK") = True Then
                                                        .SelectedIndex = .Items.IndexOf("DEED_BK_PK")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_APN
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("APN") = True Then
                                                        .SelectedIndex = .Items.IndexOf("APN")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_crossing_length
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("CL_CROSSFT") = True Then
                                                        .SelectedIndex = .Items.IndexOf("CL_CROSSFT")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_Section_TWP_Range
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("PHYS_S_T_R") = True Then
                                                        .SelectedIndex = .Items.IndexOf("PHYS_S_T_R")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_area_TWS_ABD
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("TWS_AB_SF") = True Then
                                                        .SelectedIndex = .Items.IndexOf("TWS_AB_SF")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_area_TWS_PD
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("TWS_PD_SF") = True Then
                                                        .SelectedIndex = .Items.IndexOf("TWS_PD_SF")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_access_road_length
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("AR_LENGTH") = True Then
                                                        .SelectedIndex = .Items.IndexOf("AR_LENGTH")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                        End If
                                    End If
                                End Using






                            End Using
                        End Using

                    End If
                End If

                MsgBox("DONE")
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

    Private Sub Button_generate_Platt_Click(sender As Object, e As EventArgs) Handles Button_generate_Platt.Click

        If Freeze_operations = False Then
            Freeze_operations = True

            If IsNothing(Colectie_ID) = True Then
                MsgBox("Please read HMM ID excel file")
                Freeze_operations = False
                Exit Sub
            End If
            If Colectie_ID.Count = 0 Then
                MsgBox("Please read HMM ID excel file")
                Freeze_operations = False
                Exit Sub
            End If

            If TextBox_Output_Directory.Text = "" Then
                MsgBox("Please specify the output folder")
                Freeze_operations = False
                Exit Sub
            End If
            If Not Strings.Right(TextBox_Output_Directory.Text, 1) = "\" Then
                TextBox_Output_Directory.Text = TextBox_Output_Directory.Text & "\"
            End If

            If TextBox_dwt_template.Text = "" Then
                MsgBox("Please specify the dwt template file")
                Freeze_operations = False
                Exit Sub
            End If
            If Not Strings.Right(TextBox_dwt_template.Text, 3).ToUpper = "DWT" Then
                MsgBox("Please specify the dwt template file")
                Freeze_operations = False
                Exit Sub
            End If

            If TextBox_sheet_set_template.Text = "" Then
                MsgBox("Please specify the SHEET SET file")
                Freeze_operations = False
                Exit Sub
            End If
            If Not Strings.Right(TextBox_sheet_set_template.Text, 3).ToUpper = "DST" Then
                MsgBox("Please specify the SHEET SET file")
                Freeze_operations = False
                Exit Sub
            End If

            If TextBox_xref_model_space.Text = "" Then
                MsgBox("Please specify the BASEMAP file")
                Freeze_operations = False
                Exit Sub
            End If
            If Not Strings.Right(TextBox_xref_model_space.Text, 3).ToUpper = "DWG" Then
                MsgBox("Please specify the BASEMAP file")
                Freeze_operations = False
                Exit Sub
            End If

            If IsNumeric(TextBox_main_viewport_center_X.Text) = True Then
                Vw_CenX = CDbl(TextBox_main_viewport_center_X.Text)
            End If
            If IsNumeric(TextBox_main_viewport_center_Y.Text) = True Then
                Vw_CenY = CDbl(TextBox_main_viewport_center_Y.Text)
            End If
            If IsNumeric(TextBox_main_viewport_width.Text) = True Then
                Vw_width = CDbl(TextBox_main_viewport_width.Text)
            End If
            If IsNumeric(TextBox_main_viewport_height.Text) = True Then
                Vw_height = CDbl(TextBox_main_viewport_height.Text)
            End If

            If IsNumeric(TextBox_main_viewport_center_X.Text) = True Then
                Vw_CenX = CDbl(TextBox_main_viewport_center_X.Text)
            End If
            If IsNumeric(TextBox_main_viewport_center_Y.Text) = True Then
                Vw_CenY = CDbl(TextBox_main_viewport_center_Y.Text)
            End If
            If IsNumeric(TextBox_main_viewport_width.Text) = True Then
                Vw_width = CDbl(TextBox_main_viewport_width.Text)
            End If
            If IsNumeric(TextBox_main_viewport_height.Text) = True Then
                Vw_height = CDbl(TextBox_main_viewport_height.Text)
            End If



            If IsNumeric(TextBox_locus_viewport_center_X.Text) = True Then
                VwL_CenX = CDbl(TextBox_locus_viewport_center_X.Text)
            End If
            If IsNumeric(TextBox_locus_viewport_center_Y.Text) = True Then
                VwL_CenY = CDbl(TextBox_locus_viewport_center_Y.Text)
            End If
            If IsNumeric(TextBox_locus_viewport_width.Text) = True Then
                VwL_width = CDbl(TextBox_locus_viewport_width.Text)
            End If
            If IsNumeric(TextBox_locus_viewport_height.Text) = True Then
                VwL_height = CDbl(TextBox_locus_viewport_height.Text)
            End If

            Dim Locus_scale As Double = 1 / 2000

            Dim Conversie_SQFT_ACREES As Double = 1 * 2.2956841 / 100000
            If CheckBox_convert_sqft_to_acres.Checked = False Then
                Conversie_SQFT_ACREES = 1
            End If

            Dim Empty_array() As ObjectId

            Dim BaseMap_drawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Dim Poly_CL As Polyline



            If IO.File.Exists(TextBox_xref_model_space.Text) = True Then
                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = BaseMap_drawing.Editor
                Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                Try

                    Dim Data_table_FROM_CL As New System.Data.DataTable
                    If Not ComboBox_pipe_diameter.Text = "" Then Data_table_FROM_CL.Columns.Add(ComboBox_pipe_diameter.Text, GetType(String))
                    If Not ComboBox_pipe_name.Text = "" Then Data_table_FROM_CL.Columns.Add(ComboBox_pipe_name.Text, GetType(String))
                    If Not ComboBox_pipe_segment.Text = "" Then Data_table_FROM_CL.Columns.Add(ComboBox_pipe_segment.Text, GetType(String))

                    Dim Prompt_optionsCL As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select Centerline:")
                    Prompt_optionsCL.SetRejectMessage(vbLf & "You did not selected a polyline")
                    Prompt_optionsCL.AddAllowedClass(GetType(Polyline), True)

                    Dim Rezultat_CL As Autodesk.AutoCAD.EditorInput.PromptEntityResult = Editor1.GetEntity(Prompt_optionsCL)








                    If Rezultat_CL.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        Using lock1 As DocumentLock = BaseMap_drawing.LockDocument
                            Using Trans_basemap As Autodesk.AutoCAD.DatabaseServices.Transaction = BaseMap_drawing.TransactionManager.StartTransaction
                                HostApplicationServices.WorkingDatabase = BaseMap_drawing.Database
                                Creaza_layer(Layer_name_Main_Viewport, 7, "Viewport", False)
                                Creaza_layer(Layer_name_Blocks, 2, "Text and Blocks", True)
                                Poly_CL = Trans_basemap.GetObject(Rezultat_CL.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                                Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables
                                Using Records1 As Autodesk.Gis.Map.ObjectData.Records = Tables1.GetObjectRecords(Convert.ToInt32(0), Poly_CL.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                    If IsNothing(Records1) = False Then
                                        If Records1.Count > 0 Then
                                            For Each Record1 As Autodesk.Gis.Map.ObjectData.Record In Records1
                                                Dim Tabla1 As Autodesk.Gis.Map.ObjectData.Table
                                                Tabla1 = Tables1(Record1.TableName)
                                                Dim Field_defs1 As Autodesk.Gis.Map.ObjectData.FieldDefinitions
                                                Field_defs1 = Tabla1.FieldDefinitions
                                                For j = 0 To Record1.Count - 1
                                                    Dim Field_def1 As Autodesk.Gis.Map.ObjectData.FieldDefinition
                                                    Field_def1 = Field_defs1(j)

                                                    If Field_def1.Name.ToUpper = ComboBox_pipe_diameter.Text Then
                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                        Valoare_record1 = Record1(j)
                                                        If Not Valoare_record1.StrValue = "" Then
                                                            If Data_table_FROM_CL.Rows.Count = 0 Then
                                                                Data_table_FROM_CL.Rows.Add()
                                                            End If
                                                            Data_table_FROM_CL.Rows(0).Item(ComboBox_pipe_diameter.Text) = Valoare_record1.StrValue
                                                        End If
                                                    End If

                                                    If Field_def1.Name.ToUpper = ComboBox_pipe_name.Text Then
                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                        Valoare_record1 = Record1(j)
                                                        If Not Valoare_record1.StrValue = "" Then
                                                            If Data_table_FROM_CL.Rows.Count = 0 Then
                                                                Data_table_FROM_CL.Rows.Add()
                                                            End If
                                                            Data_table_FROM_CL.Rows(0).Item(ComboBox_pipe_name.Text) = Valoare_record1.StrValue
                                                        End If
                                                    End If

                                                    If Field_def1.Name.ToUpper = ComboBox_pipe_segment.Text Then
                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                        Valoare_record1 = Record1(j)
                                                        If Not Valoare_record1.StrValue = "" Then
                                                            If Data_table_FROM_CL.Rows.Count = 0 Then
                                                                Data_table_FROM_CL.Rows.Add()
                                                            End If
                                                            Dim Segment_string As String = Valoare_record1.StrValue
                                                            Segment_string = Replace(Replace(Segment_string, "SEGMENT", ""), " ", "")
                                                            Data_table_FROM_CL.Rows(0).Item(ComboBox_pipe_segment.Text) = Segment_string
                                                        End If
                                                    End If
                                                Next
                                            Next
                                        Else
                                            If MsgBox("No object data attached to the centerline" & vbCrLf & "Do you still want to continue?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                                                Freeze_operations = False
                                                Editor1.SetImpliedSelection(Empty_array)
                                                Editor1.WriteMessage(vbLf & "Command:")
                                                Exit Sub
                                            Else
                                                Data_table_FROM_CL.Rows.Add()
                                                Data_table_FROM_CL.Columns.Add("XXX", GetType(String))
                                                Data_table_FROM_CL.Rows(0).Item(0) = "XXX"

                                            End If
                                        End If
                                    Else
                                        If MsgBox("No object data attached to the centerline" & vbCrLf & "Do you still want to continue?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                                            Freeze_operations = False
                                            Editor1.SetImpliedSelection(Empty_array)
                                            Editor1.WriteMessage(vbLf & "Command:")
                                            Exit Sub
                                        Else
                                            Data_table_FROM_CL.Rows.Add()

                                            Data_table_FROM_CL.Columns.Add("XXX", GetType(String))
                                            Data_table_FROM_CL.Rows(0).Item(0) = "XXX"
                                        End If
                                    End If
                                End Using
                                Trans_basemap.Commit()
                            End Using
                        End Using


                    Else
                        Data_table_FROM_CL.Rows.Add()

                        Data_table_FROM_CL.Columns.Add("XXX", GetType(String))
                        Data_table_FROM_CL.Rows(0).Item(0) = "XXX"
                    End If


                    Dim Data_table_cu_valori As New System.Data.DataTable
                    Dim Index_data_table_valori As Double = 0

                    Data_table_cu_valori.Columns.Add(New_plat_name_column, GetType(String))
                    With ComboBox_state
                        If Not .Text = "" Then
                            If Data_table_cu_valori.Columns.Contains(.Text) = False Then
                                Data_table_cu_valori.Columns.Add(.Text, GetType(String))
                            End If
                        End If
                    End With

                    If Not ComboBox_county.Text = "" Then
                        If Data_table_cu_valori.Columns.Contains(ComboBox_county.Text) = False Then
                            Data_table_cu_valori.Columns.Add(ComboBox_county.Text, GetType(String))
                        End If
                    End If

                    If Not ComboBox_town.Text = "" Then
                        If Data_table_cu_valori.Columns.Contains(ComboBox_town.Text) = False Then
                            Data_table_cu_valori.Columns.Add(ComboBox_town.Text, GetType(String))
                        End If
                    End If

                    If Not ComboBox_linelist.Text = "" Then
                        If Data_table_cu_valori.Columns.Contains(ComboBox_linelist.Text) = False Then
                            Data_table_cu_valori.Columns.Add(ComboBox_linelist.Text, GetType(String))
                        End If
                    End If

                    If Not ComboBox_HMM_ID.Text = "" Then
                        If Data_table_cu_valori.Columns.Contains(ComboBox_HMM_ID.Text) = False Then
                            Data_table_cu_valori.Columns.Add(ComboBox_HMM_ID.Text, GetType(String))
                        End If
                    End If

                    If Not ComboBox_MBL.Text = "" Then
                        If Data_table_cu_valori.Columns.Contains(ComboBox_MBL.Text) = False Then
                            Data_table_cu_valori.Columns.Add(ComboBox_MBL.Text, GetType(String))
                        End If
                    End If

                    If Not ComboBox_deed_page.Text = "" Then
                        If Data_table_cu_valori.Columns.Contains(ComboBox_deed_page.Text) = False Then
                            Data_table_cu_valori.Columns.Add(ComboBox_deed_page.Text, GetType(String))
                        End If
                    End If

                    If Not ComboBox_APN.Text = "" Then
                        If Data_table_cu_valori.Columns.Contains(ComboBox_APN.Text) = False Then
                            Data_table_cu_valori.Columns.Add(ComboBox_APN.Text, GetType(String))
                        End If
                    End If

                    With ComboBox_access_road_length
                        If Not .Text = "" Then
                            If Data_table_cu_valori.Columns.Contains(.Text) = False Then
                                Data_table_cu_valori.Columns.Add(.Text, GetType(Double))
                            End If
                        End If
                    End With


                    If Not ComboBox_Section_TWP_Range.Text = "" Then
                        If Data_table_cu_valori.Columns.Contains(ComboBox_Section_TWP_Range.Text) = False Then
                            Data_table_cu_valori.Columns.Add(ComboBox_Section_TWP_Range.Text, GetType(String))
                        End If
                    End If

                    If Not ComboBox_owner.Text = "" Then
                        If Data_table_cu_valori.Columns.Contains(ComboBox_owner.Text) = False Then
                            Data_table_cu_valori.Columns.Add(ComboBox_owner.Text, GetType(String))
                        End If
                    End If


                    With ComboBox_crossing_length
                        If Not .Text = "" Then
                            If Data_table_cu_valori.Columns.Contains(.Text) = False Then
                                Data_table_cu_valori.Columns.Add(.Text, GetType(Double))
                            End If
                        End If
                    End With


                    If Not ComboBox_area_EX_E.Text = "" Then
                        If Data_table_cu_valori.Columns.Contains(ComboBox_area_EX_E.Text) = False Then
                            Data_table_cu_valori.Columns.Add(ComboBox_area_EX_E.Text, GetType(Double))
                        End If
                    End If

                    If Not ComboBox_area_P_E.Text = "" Then
                        If Data_table_cu_valori.Columns.Contains(ComboBox_area_P_E.Text) = False Then
                            Data_table_cu_valori.Columns.Add(ComboBox_area_P_E.Text, GetType(Double))
                        End If
                    End If

                    If Not ComboBox_area_TWS.Text = "" Then
                        If Data_table_cu_valori.Columns.Contains(ComboBox_area_TWS.Text) = False Then
                            Data_table_cu_valori.Columns.Add(ComboBox_area_TWS.Text, GetType(Double))
                        End If
                    End If

                    If Not ComboBox_area_ATWS.Text = "" Then
                        If Data_table_cu_valori.Columns.Contains(ComboBox_area_ATWS.Text) = False Then
                            Data_table_cu_valori.Columns.Add(ComboBox_area_ATWS.Text, GetType(Double))
                        End If
                    End If

                    If Not ComboBox_area_A_R.Text = "" Then
                        If Data_table_cu_valori.Columns.Contains(ComboBox_area_A_R.Text) = False Then
                            Data_table_cu_valori.Columns.Add(ComboBox_area_A_R.Text, GetType(Double))
                        End If
                    End If

                    If Not ComboBox_area_WARE.Text = "" Then
                        If Data_table_cu_valori.Columns.Contains(ComboBox_area_WARE.Text) = False Then
                            Data_table_cu_valori.Columns.Add(ComboBox_area_WARE.Text, GetType(Double))
                        End If
                    End If

                    With ComboBox_area_TWS_ABD
                        If Not .Text = "" Then
                            If Data_table_cu_valori.Columns.Contains(.Text) = False Then
                                Data_table_cu_valori.Columns.Add(.Text, GetType(Double))
                            End If
                        End If
                    End With

                    With ComboBox_area_TWS_PD
                        If Not .Text = "" Then
                            If Data_table_cu_valori.Columns.Contains(.Text) = False Then
                                Data_table_cu_valori.Columns.Add(.Text, GetType(Double))
                            End If
                        End If
                    End With



                    Data_table_cu_valori.Columns.Add(VW_TARGET_X, GetType(Double))
                    Data_table_cu_valori.Columns.Add(VW_TARGET_Y, GetType(Double))
                    Data_table_cu_valori.Columns.Add(VW_TWIST, GetType(Double))
                    Data_table_cu_valori.Columns.Add(VW_CUST_SCALE, GetType(Double))
                    Data_table_cu_valori.Columns.Add(VW_NEW_NAME, GetType(String))
                    Data_table_cu_valori.Columns.Add(Obj_type, GetType(String))

                    Dim Colectie_Parcele() As Polyline
                    Dim Colectie_Parcele_for_points() As Polyline

                    Dim Index_parcele As Integer = 0
                    Dim Index_parcele_for_points As Integer = 0


                    Using lock1 As DocumentLock = BaseMap_drawing.LockDocument
                        Using Trans_basemap As Autodesk.AutoCAD.DatabaseServices.Transaction = BaseMap_drawing.TransactionManager.StartTransaction
                            HostApplicationServices.WorkingDatabase = BaseMap_drawing.Database
                            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            BlockTable1 = BaseMap_drawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            Dim BTrecord_MS_basemap As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans_basemap.GetObject(BlockTable1(BlockTableRecord.ModelSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                            If Rezultat_CL.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Poly_CL = Trans_basemap.GetObject(Rezultat_CL.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                            Else
                                Poly_CL = New Polyline
                                Poly_CL.AddVertexAt(0, New Point2d(0, 0), 0, 0, 0)
                                Poly_CL.AddVertexAt(1, New Point2d(1, 0), 0, 0, 0)
                            End If

                            Dim view0 As ViewTableRecord

                            Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                            Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables
                            Dim Found_parcel As Boolean = False

                            For Each ObjID1 As ObjectId In BTrecord_MS_basemap
                                Dim Ent1 As Entity
                                Ent1 = TryCast(Trans_basemap.GetObject(ObjID1, OpenMode.ForRead), Entity)
                                If IsNothing(Ent1) = False Then
                                    If TypeOf Ent1 Is Polyline Then
                                        Dim Poly_Parc As Polyline = Ent1
                                        Dim Id1 As ObjectId = Poly_Parc.ObjectId


                                        Using Records1 As Autodesk.Gis.Map.ObjectData.Records = Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                            If IsNothing(Records1) = False Then
                                                If Records1.Count > 0 Then

                                                    Dim HMM_ID As String

                                                    For Each Record1 As Autodesk.Gis.Map.ObjectData.Record In Records1
                                                        Dim Tabla1 As Autodesk.Gis.Map.ObjectData.Table
                                                        Tabla1 = Tables1(Record1.TableName)

                                                        Dim Field_defs1 As Autodesk.Gis.Map.ObjectData.FieldDefinitions
                                                        Field_defs1 = Tabla1.FieldDefinitions

                                                        For j = 0 To Record1.Count - 1
                                                            Dim Field_def1 As Autodesk.Gis.Map.ObjectData.FieldDefinition
                                                            Field_def1 = Field_defs1(j)

                                                            If Not ComboBox_HMM_ID.Text = "" Then
                                                                If Field_def1.Name.ToUpper = ComboBox_HMM_ID.Text Then
                                                                    Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                    Valoare_record1 = Record1(j)
                                                                    HMM_ID = Valoare_record1.StrValue


                                                                    If Colectie_ID.Contains(HMM_ID) = True Then
                                                                        Data_table_cu_valori.Rows.Add()
                                                                        Data_table_cu_valori.Rows(Index_data_table_valori).Item(ComboBox_HMM_ID.Text) = Valoare_record1.StrValue
                                                                        Found_parcel = True
                                                                        Exit For
                                                                    End If

                                                                    If IsNumeric(HMM_ID) = True Then
                                                                        Dim hmmid_int As Integer = CInt(HMM_ID)
                                                                        Dim hmmid_dbl As Double = CDbl(HMM_ID)

                                                                        If CDbl(hmmid_int) - hmmid_dbl = 0 Then
                                                                            If Colectie_ID.Contains(hmmid_int.ToString) = True Then
                                                                                Data_table_cu_valori.Rows.Add()
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(ComboBox_HMM_ID.Text) = Valoare_record1.StrValue
                                                                                Found_parcel = True
                                                                                Exit For
                                                                            End If
                                                                        End If
                                                                    End If

                                                                End If
                                                            End If
                                                        Next

                                                        If Found_parcel = True Then

                                                            ReDim Preserve Colectie_Parcele(Index_parcele)
                                                            Poly_parc_for_jig = Poly_Parc
                                                            Colectie_Parcele(Index_parcele) = Poly_Parc
                                                            Index_parcele = Index_parcele + 1

                                                            For j = 0 To Record1.Count - 1
                                                                Dim Field_def1 As Autodesk.Gis.Map.ObjectData.FieldDefinition
                                                                Field_def1 = Field_defs1(j)



                                                                With ComboBox_state
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If Not Valoare_record1.StrValue = "" Then
                                                                                Dim State1 As String = Valoare_record1.StrValue
                                                                                If CheckBox_state_abreviation.Checked = True Then
                                                                                    Select Case State1.ToUpper
                                                                                        Case "PA"
                                                                                            State1 = "Pennsylvania".ToUpper
                                                                                        Case "NY"
                                                                                            State1 = "New York".ToUpper
                                                                                        Case "MA"
                                                                                            State1 = "Massachusetts".ToUpper
                                                                                        Case "NH"
                                                                                            State1 = "New Hampshire".ToUpper
                                                                                        Case "CT"
                                                                                            State1 = "Connecticut".ToUpper
                                                                                        Case "OH"
                                                                                            State1 = "Ohio".ToUpper
                                                                                    End Select
                                                                                End If

                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = State1
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_county
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If Not Valoare_record1.StrValue = "" Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Valoare_record1.StrValue
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With


                                                                With ComboBox_town
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If Not Valoare_record1.StrValue = "" Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Valoare_record1.StrValue
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_linelist
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If Not Valoare_record1.StrValue = "" Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Valoare_record1.StrValue
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_HMM_ID
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If Not Valoare_record1.StrValue = "" Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Valoare_record1.StrValue
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_MBL
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If Not Valoare_record1.StrValue = "" Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Valoare_record1.StrValue
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_deed_page
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If Not Valoare_record1.StrValue = "" Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Valoare_record1.StrValue
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_APN
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If Not Valoare_record1.StrValue = "" Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Valoare_record1.StrValue
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_access_road_length
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue)
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_Section_TWP_Range
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If Not Valoare_record1.StrValue = "" Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Valoare_record1.StrValue
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_owner
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If Not Valoare_record1.StrValue = "" Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Valoare_record1.StrValue
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_crossing_length
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue)
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_area_EX_E
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_area_P_E
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_area_TWS
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_area_ATWS
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With


                                                                With ComboBox_area_A_R
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_area_WARE
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With


                                                                With ComboBox_area_TWS_ABD
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_area_TWS_PD
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With



                                                                With ComboBox_plat_name
                                                                    If CheckBox_USE_ned_for_naming.Checked = False Then
                                                                        If Not .Text = "" Then
                                                                            If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                                Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                                Valoare_record1 = Record1(j)
                                                                                If Not Valoare_record1.StrValue = "" Then
                                                                                    Data_table_cu_valori.Rows(Index_data_table_valori).Item(New_plat_name_column) = Valoare_record1.StrValue.ToUpper
                                                                                End If
                                                                            End If
                                                                        Else
                                                                            MsgBox("Please specify the  plat name field")
                                                                            Freeze_operations = False
                                                                            Exit Sub
                                                                        End If
                                                                    Else
                                                                        Dim New_plat_name As String = "PLAT_"
                                                                        If Not ComboBox_pipe_segment.Text = "" Then
                                                                            If Data_table_FROM_CL.Columns.Contains(ComboBox_pipe_segment.Text) = True Then
                                                                                If IsDBNull(Data_table_FROM_CL.Rows(0).Item(ComboBox_pipe_segment.Text)) = False Then
                                                                                    New_plat_name = "SEG_" & Data_table_FROM_CL.Rows(0).Item(ComboBox_pipe_segment.Text)
                                                                                Else
                                                                                    New_plat_name = "SEG_xxx"
                                                                                End If
                                                                            Else
                                                                                New_plat_name = "SEG_xxx"
                                                                            End If

                                                                        Else
                                                                            New_plat_name = "SEG_xxx"
                                                                        End If

                                                                        If Not ComboBox_linelist.Text = "" Then
                                                                            If IsDBNull(Data_table_cu_valori.Rows(Index_data_table_valori).Item(ComboBox_linelist.Text)) = False Then
                                                                                New_plat_name = New_plat_name & "_" & extrage_numar_din_text_de_la_sfarsitul_textului(Data_table_cu_valori.Rows(Index_data_table_valori).Item(ComboBox_linelist.Text))
                                                                            Else
                                                                                New_plat_name = New_plat_name & "_xxx"
                                                                            End If
                                                                        Else
                                                                            New_plat_name = New_plat_name & "_xxx"
                                                                        End If

                                                                        Data_table_cu_valori.Rows(Index_data_table_valori).Item(New_plat_name_column) = New_plat_name

                                                                    End If

                                                                End With




                                                            Next





                                                            Dim col_int As New Point3dCollection
                                                            Poly_CL.IntersectWith(Poly_Parc, Intersect.OnBothOperands, col_int, IntPtr.Zero, IntPtr.Zero)
                                                            Dim linie1 As Line
                                                            If col_int.Count > 0 Then
                                                                linie1 = New Line(col_int(0), col_int(col_int.Count - 1))
                                                            Else
                                                                linie1 = New Line(Poly_Parc.StartPoint, Poly_Parc.GetPointAtParameter((Poly_Parc.NumberOfVertices - 1) / 2))
                                                            End If





                                                            Dim Point_pt_viewport As Autodesk.AutoCAD.Geometry.Point3d = linie1.GetPointAtDist(linie1.Length / 2)
                                                            Rotatie = GET_Bearing_rad(linie1.StartPoint.X, linie1.StartPoint.Y, linie1.EndPoint.X, linie1.EndPoint.Y)
                                                            If Rotatie < PI + 45 * PI / 180 And Rotatie > PI - 45 * PI / 180 Then
                                                                Rotatie = Rotatie - PI
                                                            End If



                                                            Dim view1 As ViewTableRecord
                                                            Dim View_Table As ViewTable = Trans_basemap.GetObject(BaseMap_drawing.Database.ViewTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                                            Dim Tilemode1 As Integer = Application.GetSystemVariable("TILEMODE")
                                                            Dim CVport1 As Integer = Application.GetSystemVariable("CVPORT")
                                                            If Tilemode1 = 0 Then
                                                                Application.SetSystemVariable("TILEMODE", 1)
                                                            End If
                                                            Dim View_name As String = "Plat"

                                                            If View_Table.Has(View_name) = False Then
                                                                View_Table.UpgradeOpen()
                                                                view1 = New ViewTableRecord
                                                                view1.CenterPoint = New Point2d(0, 0)
                                                                view1.Target = Point_pt_viewport
                                                                view1.ViewTwist = 2 * PI - Rotatie
                                                                view1.Width = 1.3 * linie1.Length
                                                                view1.Height = 1.3 * linie1.Length / 1
                                                                view1.Name = View_name
                                                                View_Table.Add(view1)
                                                                Trans_basemap.AddNewlyCreatedDBObject(view1, True)
                                                            Else
                                                                view1 = View_Table(View_name).GetObject(OpenMode.ForWrite)
                                                                view1.CenterPoint = New Point2d(0, 0)
                                                                view1.Target = Point_pt_viewport
                                                                view1.ViewTwist = 2 * PI - Rotatie
                                                                view1.Width = 1.3 * linie1.Length
                                                                view1.Height = 1.3 * linie1.Length / 1
                                                            End If




                                                            If View_Table.Has("North") = False Then
                                                                View_Table.UpgradeOpen()
                                                                view0 = New ViewTableRecord
                                                                view0.CenterPoint = New Point2d(0, 0)
                                                                view0.Target = Point_pt_viewport
                                                                view0.ViewTwist = 0
                                                                view0.Width = 1.3 * linie1.Length
                                                                view0.Height = 1.3 * linie1.Length / 1
                                                                view0.Name = "North"
                                                                View_Table.Add(view0)
                                                                Trans_basemap.AddNewlyCreatedDBObject(view0, True)
                                                            Else
                                                                view0 = View_Table("North").GetObject(OpenMode.ForWrite)
                                                                view0.CenterPoint = New Point2d(0, 0)
                                                                view0.Target = Point_pt_viewport
                                                                view0.ViewTwist = 0
                                                                view0.Width = 1.3 * linie1.Length
                                                                view0.Height = 1.3 * linie1.Length / 1
                                                            End If

                                                            If CheckBox_rotate_to_north.Checked = False Then
                                                                BaseMap_drawing.Editor.SetCurrentView(view1)
                                                                Rotatie_originala = Rotatie
                                                            Else
                                                                BaseMap_drawing.Editor.SetCurrentView(view0)
                                                                Rotatie_originala = 0
                                                            End If



1254:
                                                            If RadioButton10.Checked = True Then
                                                                Vw_scale = 1 / 10
                                                            End If

                                                            If RadioButton20.Checked = True Then
                                                                Vw_scale = 1 / 20
                                                            End If

                                                            If RadioButton30.Checked = True Then
                                                                Vw_scale = 1 / 30
                                                            End If

                                                            If RadioButton40.Checked = True Then
                                                                Vw_scale = 1 / 40
                                                            End If

                                                            If RadioButton50.Checked = True Then
                                                                Vw_scale = 1 / 50
                                                            End If

                                                            If RadioButton60.Checked = True Then
                                                                Vw_scale = 1 / 60
                                                            End If

                                                            If RadioButton100.Checked = True Then
                                                                Vw_scale = 1 / 100
                                                            End If

                                                            If RadioButton200.Checked = True Then
                                                                Vw_scale = 1 / 200
                                                            End If

                                                            If RadioButton300.Checked = True Then
                                                                Vw_scale = 1 / 300
                                                            End If

                                                            If RadioButton400.Checked = True Then
                                                                Vw_scale = 1 / 400
                                                            End If

                                                            If RadioButton500.Checked = True Then
                                                                Vw_scale = 1 / 500
                                                            End If

                                                            If RadioButton600.Checked = True Then
                                                                Vw_scale = 1 / 600
                                                            End If

                                                            If RadioButton1000.Checked = True Then
                                                                Vw_scale = 1 / 1000
                                                            End If


                                                            Dim PromptPointRezult1 As Autodesk.AutoCAD.EditorInput.PromptPointResult

                                                            Dim Jig1 As New Jig_rectangle_viewport

                                                            If CheckBox_rotate_to_north.Checked = True Then
                                                                PromptPointRezult1 = Jig1.StartJig(Vw_width, Vw_height, False)
                                                            Else
                                                                PromptPointRezult1 = Jig1.StartJig(Vw_width, Vw_height, True)
                                                            End If




                                                            If IsNothing(PromptPointRezult1) = True Then
                                                                Freeze_operations = False
                                                                Exit Sub
                                                            End If

                                                            If Not PromptPointRezult1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                                                Freeze_operations = False
                                                                Exit Sub
                                                            End If

                                                            Dim pointM As New Point3d
                                                            pointM = PromptPointRezult1.Value



                                                            view1 = View_Table(View_name).GetObject(OpenMode.ForWrite)
                                                            view1.CenterPoint = New Point2d(0, 0)
                                                            view1.Target = pointM
                                                            If CheckBox_rotate_to_north.Checked = True Then
                                                                view1.ViewTwist = 0
                                                            Else
                                                                view1.ViewTwist = 2 * PI - Rotatie
                                                            End If

                                                            view1.Width = 1.3 * Vw_width / Vw_scale
                                                            view1.Height = 1.3 * Vw_width / Vw_scale
                                                            BaseMap_drawing.Editor.SetCurrentView(view1)

                                                            Using Trans_basemap2 As Autodesk.AutoCAD.DatabaseServices.Transaction = BaseMap_drawing.TransactionManager.StartTransaction
                                                                Dim Poly1 As New Autodesk.AutoCAD.DatabaseServices.Polyline
                                                                Poly1.AddVertexAt(0, New Point2d(pointM.X - 0.5 * Vw_width / Vw_scale, pointM.Y - 0.5 * Vw_height / Vw_scale), 0, 0, 0)
                                                                Poly1.AddVertexAt(1, New Point2d(pointM.X + 0.5 * Vw_width / Vw_scale, pointM.Y - 0.5 * Vw_height / Vw_scale), 0, 0, 0)
                                                                Poly1.AddVertexAt(2, New Point2d(pointM.X + 0.5 * Vw_width / Vw_scale, pointM.Y + 0.5 * Vw_height / Vw_scale), 0, 0, 0)
                                                                Poly1.AddVertexAt(3, New Point2d(pointM.X - 0.5 * Vw_width / Vw_scale, pointM.Y + 0.5 * Vw_height / Vw_scale), 0, 0, 0)
                                                                Poly1.Closed = True

                                                                If CheckBox_rotate_to_north.Checked = False Then
                                                                    Poly1.TransformBy(Matrix3d.Rotation(Rotatie, Vector3d.ZAxis, pointM))
                                                                End If

                                                                Poly1.Layer = Layer_name_Main_Viewport
                                                                BTrecord_MS_basemap.AppendEntity(Poly1)
                                                                Trans_basemap2.AddNewlyCreatedDBObject(Poly1, True)
                                                                Trans_basemap2.TransactionManager.QueueForGraphicsFlush()
                                                                Trans_basemap2.Commit()

                                                                Select Case MsgBox("Are you OK with the viewport Scale, Position and Rotation?", vbYesNo)
                                                                    Case MsgBoxResult.No
                                                                        Using Trans_basemap3 As Autodesk.AutoCAD.DatabaseServices.Transaction = BaseMap_drawing.TransactionManager.StartTransaction
                                                                            Trans_basemap3.GetObject(Poly1.ObjectId, OpenMode.ForWrite)
                                                                            Poly1.Erase()
                                                                            Trans_basemap3.Commit()
                                                                        End Using



                                                                        HScrollBar_rotate.Value = 0
                                                                        TextBox_rotation_ammount.Text = "0"
                                                                        BaseMap_drawing.Editor.SetCurrentView(view0)

                                                                        GoTo 1254
                                                                    Case MsgBoxResult.Yes

                                                                        Point_pt_viewport = pointM

                                                                        If CheckBox_rotate_to_north.Checked = False Then
                                                                            Data_table_cu_valori.Rows(Index_data_table_valori).Item(VW_TWIST) = 2 * PI - Rotatie
                                                                        Else
                                                                            Data_table_cu_valori.Rows(Index_data_table_valori).Item(VW_TWIST) = 0
                                                                        End If

                                                                        Data_table_cu_valori.Rows(Index_data_table_valori).Item(VW_CUST_SCALE) = Vw_scale
                                                                        Data_table_cu_valori.Rows(Index_data_table_valori).Item(VW_TARGET_X) = Point_pt_viewport.X
                                                                        Data_table_cu_valori.Rows(Index_data_table_valori).Item(VW_TARGET_Y) = Point_pt_viewport.Y
                                                                        Data_table_cu_valori.Rows(Index_data_table_valori).Item(VW_NEW_NAME) = View_name
                                                                End Select ' msgbox YES



                                                                Trans_basemap2.Dispose()
                                                            End Using


                                                            HScrollBar_rotate.Value = 0
                                                            TextBox_rotation_ammount.Text = "0"


                                                        Else

                                                            Poly_parc_for_jig = Nothing
                                                        End If
                                                    Next

                                                    If Found_parcel = True Then
                                                        Index_data_table_valori = Index_data_table_valori + 1
                                                        Found_parcel = False
                                                    Else
                                                        ' MsgBox("Parcel not found")
                                                    End If


                                                End If
                                            End If
                                        End Using





                                    End If


                                    If TypeOf Ent1 Is DBPoint Then
                                        Dim Point_Parc As DBPoint = Ent1

                                        pt_CERC = Point_Parc.Position

                                        Using Records1 As Autodesk.Gis.Map.ObjectData.Records = Tables1.GetObjectRecords(Convert.ToInt32(0), Point_Parc.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                            If IsNothing(Records1) = False Then
                                                If Records1.Count > 0 Then

                                                    Dim HMM_ID As String

                                                    For Each Record1 As Autodesk.Gis.Map.ObjectData.Record In Records1
                                                        Dim Tabla1 As Autodesk.Gis.Map.ObjectData.Table
                                                        Tabla1 = Tables1(Record1.TableName)

                                                        Dim Field_defs1 As Autodesk.Gis.Map.ObjectData.FieldDefinitions
                                                        Field_defs1 = Tabla1.FieldDefinitions

                                                        For j = 0 To Record1.Count - 1
                                                            Dim Field_def1 As Autodesk.Gis.Map.ObjectData.FieldDefinition
                                                            Field_def1 = Field_defs1(j)

                                                            If Not ComboBox_HMM_ID.Text = "" Then
                                                                If Field_def1.Name.ToUpper = ComboBox_HMM_ID.Text Then
                                                                    Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                    Valoare_record1 = Record1(j)
                                                                    HMM_ID = Valoare_record1.StrValue
                                                                    If Colectie_ID.Contains(HMM_ID) = True Then
                                                                        Data_table_cu_valori.Rows.Add()
                                                                        Data_table_cu_valori.Rows(Index_data_table_valori).Item(ComboBox_HMM_ID.Text) = Valoare_record1.StrValue
                                                                        Found_parcel = True
                                                                        Exit For
                                                                    End If

                                                                    If IsNumeric(HMM_ID) = True Then
                                                                        Dim hmmid_int As Integer = CInt(HMM_ID)
                                                                        Dim hmmid_dbl As Double = CDbl(HMM_ID)

                                                                        If CDbl(hmmid_int) - hmmid_dbl = 0 Then
                                                                            If Colectie_ID.Contains(hmmid_int.ToString) = True Then
                                                                                Data_table_cu_valori.Rows.Add()
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(ComboBox_HMM_ID.Text) = Valoare_record1.StrValue
                                                                                Found_parcel = True
                                                                                Exit For
                                                                            End If
                                                                        End If
                                                                    End If

                                                                End If
                                                            End If
                                                        Next

                                                        If Found_parcel = True Then

                                                            Data_table_cu_valori.Rows(Index_data_table_valori).Item(Obj_type) = "POINT"




                                                            For j = 0 To Record1.Count - 1
                                                                Dim Field_def1 As Autodesk.Gis.Map.ObjectData.FieldDefinition
                                                                Field_def1 = Field_defs1(j)



                                                                With ComboBox_state
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If Not Valoare_record1.StrValue = "" Then
                                                                                Dim State1 As String = Valoare_record1.StrValue
                                                                                If CheckBox_state_abreviation.Checked = True Then
                                                                                    Select Case State1.ToUpper
                                                                                        Case "PA"
                                                                                            State1 = "Pennsylvania".ToUpper
                                                                                        Case "NY"
                                                                                            State1 = "New York".ToUpper
                                                                                        Case "MA"
                                                                                            State1 = "Massachusetts".ToUpper
                                                                                        Case "NH"
                                                                                            State1 = "New Hampshire".ToUpper
                                                                                        Case "CT"
                                                                                            State1 = "Connecticut".ToUpper
                                                                                        Case "OH"
                                                                                            State1 = "Ohio".ToUpper
                                                                                    End Select
                                                                                End If
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = State1
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_county
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If Not Valoare_record1.StrValue = "" Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Valoare_record1.StrValue
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With


                                                                With ComboBox_town
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If Not Valoare_record1.StrValue = "" Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Valoare_record1.StrValue
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_linelist
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If Not Valoare_record1.StrValue = "" Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Valoare_record1.StrValue
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_HMM_ID
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If Not Valoare_record1.StrValue = "" Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Valoare_record1.StrValue
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_MBL
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If Not Valoare_record1.StrValue = "" Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Valoare_record1.StrValue
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_deed_page
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If Not Valoare_record1.StrValue = "" Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Valoare_record1.StrValue
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_APN
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If Not Valoare_record1.StrValue = "" Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Valoare_record1.StrValue
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_access_road_length
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue)
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_Section_TWP_Range
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If Not Valoare_record1.StrValue = "" Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Valoare_record1.StrValue
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_owner
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If Not Valoare_record1.StrValue = "" Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Valoare_record1.StrValue
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_crossing_length
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue)
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_area_EX_E
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_area_P_E
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_area_TWS
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_area_ATWS
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With


                                                                With ComboBox_area_A_R
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_area_WARE
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With


                                                                With ComboBox_area_TWS_ABD
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_area_TWS_PD
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With



                                                                With ComboBox_plat_name
                                                                    If CheckBox_USE_ned_for_naming.Checked = False Then
                                                                        If Not .Text = "" Then
                                                                            If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                                Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                                Valoare_record1 = Record1(j)
                                                                                If Not Valoare_record1.StrValue = "" Then
                                                                                    Data_table_cu_valori.Rows(Index_data_table_valori).Item(New_plat_name_column) = Valoare_record1.StrValue.ToUpper
                                                                                End If
                                                                            End If
                                                                        Else
                                                                            MsgBox("Please specify the  plat name field")
                                                                            Freeze_operations = False
                                                                            Exit Sub
                                                                        End If
                                                                    Else
                                                                        Dim New_plat_name As String = "PLAT_"
                                                                        If Not ComboBox_pipe_segment.Text = "" Then
                                                                            If Data_table_FROM_CL.Columns.Contains(ComboBox_pipe_segment.Text) = True Then
                                                                                If IsDBNull(Data_table_FROM_CL.Rows(0).Item(ComboBox_pipe_segment.Text)) = False Then
                                                                                    New_plat_name = "SEG_" & Data_table_FROM_CL.Rows(0).Item(ComboBox_pipe_segment.Text)
                                                                                Else
                                                                                    New_plat_name = "SEG_xxx"
                                                                                End If
                                                                            Else
                                                                                New_plat_name = "SEG_xxx"
                                                                            End If

                                                                        Else
                                                                            New_plat_name = "SEG_xxx"
                                                                        End If

                                                                        If Not ComboBox_linelist.Text = "" Then
                                                                            If IsDBNull(Data_table_cu_valori.Rows(Index_data_table_valori).Item(ComboBox_linelist.Text)) = False Then
                                                                                New_plat_name = New_plat_name & "_" & extrage_numar_din_text_de_la_sfarsitul_textului(Data_table_cu_valori.Rows(Index_data_table_valori).Item(ComboBox_linelist.Text))
                                                                            Else
                                                                                New_plat_name = New_plat_name & "_xxx"
                                                                            End If
                                                                        Else
                                                                            New_plat_name = New_plat_name & "_xxx"
                                                                        End If

                                                                        Data_table_cu_valori.Rows(Index_data_table_valori).Item(New_plat_name_column) = New_plat_name

                                                                    End If

                                                                End With




                                                            Next











                                                            Dim Point_pt_viewport As Autodesk.AutoCAD.Geometry.Point3d = Point_Parc.Position
                                                            Rotatie = 0




                                                            Dim view1 As ViewTableRecord
                                                            Dim View_Table As ViewTable = Trans_basemap.GetObject(BaseMap_drawing.Database.ViewTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                                            Dim Tilemode1 As Integer = Application.GetSystemVariable("TILEMODE")
                                                            Dim CVport1 As Integer = Application.GetSystemVariable("CVPORT")
                                                            If Tilemode1 = 0 Then
                                                                Application.SetSystemVariable("TILEMODE", 1)
                                                            End If
                                                            Dim View_name As String = "Plat"

                                                            If View_Table.Has(View_name) = False Then
                                                                View_Table.UpgradeOpen()
                                                                view1 = New ViewTableRecord
                                                                view1.CenterPoint = New Point2d(0, 0)
                                                                view1.Target = Point_pt_viewport
                                                                view1.ViewTwist = 2 * PI - Rotatie
                                                                view1.Width = 4000
                                                                view1.Height = 2500
                                                                view1.Name = View_name
                                                                View_Table.Add(view1)
                                                                Trans_basemap.AddNewlyCreatedDBObject(view1, True)
                                                            Else
                                                                view1 = View_Table(View_name).GetObject(OpenMode.ForWrite)
                                                                view1.CenterPoint = New Point2d(0, 0)
                                                                view1.Target = Point_pt_viewport
                                                                view1.ViewTwist = 2 * PI - Rotatie
                                                                view1.Width = 4000
                                                                view1.Height = 2500
                                                            End If




                                                            If View_Table.Has("North") = False Then
                                                                View_Table.UpgradeOpen()
                                                                view0 = New ViewTableRecord
                                                                view0.CenterPoint = New Point2d(0, 0)
                                                                view0.Target = Point_pt_viewport
                                                                view0.ViewTwist = 0
                                                                view0.Width = 4000
                                                                view0.Height = 2500
                                                                view0.Name = "North"
                                                                View_Table.Add(view0)
                                                                Trans_basemap.AddNewlyCreatedDBObject(view0, True)
                                                            Else
                                                                view0 = View_Table("North").GetObject(OpenMode.ForWrite)
                                                                view0.CenterPoint = New Point2d(0, 0)
                                                                view0.Target = Point_pt_viewport
                                                                view0.ViewTwist = 0
                                                                view0.Width = 4000
                                                                view0.Height = 2500
                                                            End If






                                                            BaseMap_drawing.Editor.SetCurrentView(view0)
                                                            Rotatie_originala = 0




1255:
                                                            If RadioButton10.Checked = True Then
                                                                Vw_scale = 1 / 10
                                                            End If

                                                            If RadioButton20.Checked = True Then
                                                                Vw_scale = 1 / 20
                                                            End If

                                                            If RadioButton30.Checked = True Then
                                                                Vw_scale = 1 / 30
                                                            End If

                                                            If RadioButton40.Checked = True Then
                                                                Vw_scale = 1 / 40
                                                            End If

                                                            If RadioButton50.Checked = True Then
                                                                Vw_scale = 1 / 50
                                                            End If

                                                            If RadioButton60.Checked = True Then
                                                                Vw_scale = 1 / 60
                                                            End If

                                                            If RadioButton100.Checked = True Then
                                                                Vw_scale = 1 / 100
                                                            End If

                                                            If RadioButton200.Checked = True Then
                                                                Vw_scale = 1 / 200
                                                            End If

                                                            If RadioButton300.Checked = True Then
                                                                Vw_scale = 1 / 300
                                                            End If

                                                            If RadioButton400.Checked = True Then
                                                                Vw_scale = 1 / 400
                                                            End If

                                                            If RadioButton500.Checked = True Then
                                                                Vw_scale = 1 / 500
                                                            End If

                                                            If RadioButton600.Checked = True Then
                                                                Vw_scale = 1 / 600
                                                            End If

                                                            If RadioButton1000.Checked = True Then
                                                                Vw_scale = 1 / 1000
                                                            End If


                                                            Dim PromptPointRezult1 As Autodesk.AutoCAD.EditorInput.PromptPointResult

                                                            Dim Jig1 As New Jig_rectangle_viewport
                                                            PromptPointRezult1 = Jig1.StartJig(Vw_width, Vw_height, False)





                                                            If IsNothing(PromptPointRezult1) = True Then
                                                                Freeze_operations = False
                                                                Exit Sub
                                                            End If

                                                            If Not PromptPointRezult1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                                                Freeze_operations = False
                                                                Exit Sub
                                                            End If

                                                            Dim pointM As New Point3d
                                                            pointM = PromptPointRezult1.Value




                                                            Dim Rezultat_poly_parcel As Autodesk.AutoCAD.EditorInput.PromptEntityResult

                                                            Dim Object_Promptparcelp As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select the associated parcel:")

                                                            Object_Promptparcelp.SetRejectMessage(vbLf & "Please select a polyline")
                                                            Object_Promptparcelp.AddAllowedClass(GetType(Polyline), True)

                                                            Rezultat_poly_parcel = Editor1.GetEntity(Object_Promptparcelp)




                                                            view1 = View_Table(View_name).GetObject(OpenMode.ForWrite)
                                                            view1.CenterPoint = New Point2d(0, 0)
                                                            view1.Target = pointM
                                                            If CheckBox_rotate_to_north.Checked = True Then
                                                                view1.ViewTwist = 0
                                                            Else
                                                                view1.ViewTwist = 2 * PI - Rotatie
                                                            End If

                                                            view1.Width = 1.3 * Vw_width / Vw_scale
                                                            view1.Height = 1.3 * Vw_width / Vw_scale
                                                            BaseMap_drawing.Editor.SetCurrentView(view1)

                                                            Using Trans_basemap2 As Autodesk.AutoCAD.DatabaseServices.Transaction = BaseMap_drawing.TransactionManager.StartTransaction
                                                                Dim Poly1 As New Autodesk.AutoCAD.DatabaseServices.Polyline
                                                                Poly1.AddVertexAt(0, New Point2d(pointM.X - 0.5 * Vw_width / Vw_scale, pointM.Y - 0.5 * Vw_height / Vw_scale), 0, 0, 0)
                                                                Poly1.AddVertexAt(1, New Point2d(pointM.X + 0.5 * Vw_width / Vw_scale, pointM.Y - 0.5 * Vw_height / Vw_scale), 0, 0, 0)
                                                                Poly1.AddVertexAt(2, New Point2d(pointM.X + 0.5 * Vw_width / Vw_scale, pointM.Y + 0.5 * Vw_height / Vw_scale), 0, 0, 0)
                                                                Poly1.AddVertexAt(3, New Point2d(pointM.X - 0.5 * Vw_width / Vw_scale, pointM.Y + 0.5 * Vw_height / Vw_scale), 0, 0, 0)
                                                                Poly1.Closed = True

                                                                If CheckBox_rotate_to_north.Checked = False Then
                                                                    Poly1.TransformBy(Matrix3d.Rotation(Rotatie, Vector3d.ZAxis, pointM))
                                                                End If

                                                                Poly1.Layer = Layer_name_Main_Viewport
                                                                BTrecord_MS_basemap.AppendEntity(Poly1)
                                                                Trans_basemap2.AddNewlyCreatedDBObject(Poly1, True)
                                                                Trans_basemap2.TransactionManager.QueueForGraphicsFlush()
                                                                Trans_basemap2.Commit()

                                                                Select Case MsgBox("Are you OK with the viewport Scale, Position and Rotation?", vbYesNo)
                                                                    Case MsgBoxResult.No
                                                                        Using Trans_basemap3 As Autodesk.AutoCAD.DatabaseServices.Transaction = BaseMap_drawing.TransactionManager.StartTransaction
                                                                            Trans_basemap3.GetObject(Poly1.ObjectId, OpenMode.ForWrite)
                                                                            Poly1.Erase()
                                                                            Trans_basemap3.Commit()
                                                                        End Using



                                                                        HScrollBar_rotate.Value = 0
                                                                        TextBox_rotation_ammount.Text = "0"
                                                                        BaseMap_drawing.Editor.SetCurrentView(view0)

                                                                        GoTo 1255
                                                                    Case MsgBoxResult.Yes

                                                                        Point_pt_viewport = pointM

                                                                        If CheckBox_rotate_to_north.Checked = False Then
                                                                            Data_table_cu_valori.Rows(Index_data_table_valori).Item(VW_TWIST) = 2 * PI - Rotatie
                                                                        Else
                                                                            Data_table_cu_valori.Rows(Index_data_table_valori).Item(VW_TWIST) = 0
                                                                        End If

                                                                        Data_table_cu_valori.Rows(Index_data_table_valori).Item(VW_CUST_SCALE) = Vw_scale
                                                                        Data_table_cu_valori.Rows(Index_data_table_valori).Item(VW_TARGET_X) = Point_pt_viewport.X
                                                                        Data_table_cu_valori.Rows(Index_data_table_valori).Item(VW_TARGET_Y) = Point_pt_viewport.Y
                                                                        Data_table_cu_valori.Rows(Index_data_table_valori).Item(VW_NEW_NAME) = View_name

                                                                        If Rezultat_poly_parcel.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                                                            Dim PPol As Polyline = Trans_basemap.GetObject(Rezultat_poly_parcel.ObjectId, OpenMode.ForRead)
                                                                            ReDim Preserve Colectie_Parcele_for_points(Index_parcele_for_points)
                                                                            Colectie_Parcele_for_points(Index_parcele_for_points) = PPol
                                                                            Index_parcele_for_points = Index_parcele_for_points + 1

                                                                        End If


                                                                End Select ' msgbox YES



                                                                Trans_basemap2.Dispose()
                                                            End Using


                                                            HScrollBar_rotate.Value = 0
                                                            TextBox_rotation_ammount.Text = "0"


                                                        Else

                                                            Poly_parc_for_jig = Nothing
                                                        End If
                                                    Next

                                                    If Found_parcel = True Then
                                                        Index_data_table_valori = Index_data_table_valori + 1
                                                        Found_parcel = False
                                                    Else
                                                        ' MsgBox("Parcel not found")
                                                    End If


                                                End If
                                            End If
                                        End Using





                                    End If



                                End If


                            Next

                            If IsNothing(view0) = False Then
                                BaseMap_drawing.Editor.SetCurrentView(view0)
                            End If

                            Trans_basemap.Commit()

                        End Using
                    End Using

                    If IsNothing(Data_table_cu_valori) = False Then
                        If Data_table_cu_valori.Rows.Count > 0 Then



                            Dim Index_existent As Integer = 1

                            If System.IO.Directory.Exists(TextBox_Output_Directory.Text) = True And System.IO.File.Exists(TextBox_sheet_set_template.Text) = True Then

                                Dim SheetSet_manager As New AcSmSheetSetMgr
                                Dim SheetSet_database As AcSmDatabase = SheetSet_manager.OpenDatabase(TextBox_sheet_set_template.Text, False)
                                Dim sheetSet As AcSmSheetSet
                                sheetSet = SheetSet_database.GetSheetSet()



                                For s = 0 To Data_table_cu_valori.Rows.Count - 1
                                    Dim New_plat_name As String = ""

                                    If IsDBNull(Data_table_cu_valori.Rows(s).Item(New_plat_name_column)) = False Then
                                        New_plat_name = Data_table_cu_valori.Rows(s).Item(New_plat_name_column)
                                    Else
                                        MsgBox("there is not a plat name")
                                        Freeze_operations = False
                                        Exit Sub
                                    End If




                                    Dim Fisierul_exista As Boolean = True
                                    Do Until Fisierul_exista = False
                                        If System.IO.File.Exists(TextBox_Output_Directory.Text & New_plat_name & ".dwg") = True Then
                                            Dim Fisierul_indexat_exista As Boolean = True
                                            Do Until Fisierul_indexat_exista = False
                                                If System.IO.File.Exists(TextBox_Output_Directory.Text & New_plat_name & "_" & Index_existent.ToString & ".dwg") = True Then
                                                    Index_existent = Index_existent + 1
                                                Else
                                                    New_plat_name = New_plat_name & "_" & Index_existent.ToString
                                                    Fisierul_indexat_exista = False
                                                End If
                                            Loop
                                        Else
                                            Fisierul_exista = False
                                        End If
                                    Loop




                                    If LockDatabase(SheetSet_database, True) = True Then

                                        Dim sheetSetFolder As String
                                        sheetSetFolder = TextBox_Output_Directory.Text
                                        Dim SS_name As String = System.IO.Path.GetFileNameWithoutExtension(TextBox_sheet_set_template.Text)
                                        Dim SS_descr As String = System.IO.Path.GetFileNameWithoutExtension(TextBox_sheet_set_template.Text)


                                        SetSheetSetDefaults(SheetSet_database, _
                                                SS_name, SS_descr, sheetSetFolder, TextBox_dwt_template.Text, New_plat_name)
                                        Dim Sheet1 As New AcSmSheet

                                        Try
                                            Try
                                                Sheet1 = AddSheet(SheetSet_database, New_plat_name, New_plat_name, New_plat_name, 1)
                                            Catch ex As System.IO.FileNotFoundException
                                                MsgBox("FILE exception: " & vbCrLf & ex.Message)
                                                LockDatabase(SheetSet_database, False)
                                                Freeze_operations = False
                                                Exit Sub
                                            End Try
                                        Catch ex As Runtime.InteropServices.COMException
                                            MsgBox("com exception: check the dwt settings inside dst " & vbCrLf & ex.Message)
                                            LockDatabase(SheetSet_database, False)
                                            Freeze_operations = False
                                            Exit Sub
                                        End Try

                                        With ComboBox_state
                                            If Not .Text = "" Then
                                                Dim Property_name As String
                                                If Not ComboBox_Sheet_Set_state.Text = "" Then
                                                    Property_name = ComboBox_Sheet_Set_state.Text
                                                Else
                                                    Property_name = .Text
                                                End If

                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(.Text)) = False Then
                                                    SetCustomProperty(Sheet1, Property_name, Data_table_cu_valori.Rows(s).Item(.Text), PropertyFlags.CUSTOM_SHEET_PROP)
                                                End If
                                            End If
                                        End With

                                        With ComboBox_county
                                            If Not .Text = "" Then
                                                Dim Property_name As String
                                                If Not ComboBox_Sheet_Set_county.Text = "" Then
                                                    Property_name = ComboBox_Sheet_Set_county.Text
                                                Else
                                                    Property_name = .Text
                                                End If

                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(.Text)) = False Then
                                                    SetCustomProperty(Sheet1, Property_name, Data_table_cu_valori.Rows(s).Item(.Text), PropertyFlags.CUSTOM_SHEET_PROP)
                                                End If
                                            End If
                                        End With



                                        With ComboBox_town
                                            If Not .Text = "" Then
                                                Dim Property_name As String
                                                If Not ComboBox_Sheet_Set_town.Text = "" Then
                                                    Property_name = ComboBox_Sheet_Set_town.Text
                                                Else
                                                    Property_name = .Text
                                                End If

                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(.Text)) = False Then
                                                    SetCustomProperty(Sheet1, Property_name, Data_table_cu_valori.Rows(s).Item(.Text), PropertyFlags.CUSTOM_SHEET_PROP)
                                                End If
                                            End If
                                        End With

                                        With ComboBox_linelist
                                            If Not .Text = "" Then
                                                Dim Property_name As String
                                                If Not ComboBox_Sheet_Set_linelist.Text = "" Then
                                                    Property_name = ComboBox_Sheet_Set_linelist.Text
                                                Else
                                                    Property_name = .Text
                                                End If
                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(.Text)) = False Then
                                                    SetCustomProperty(Sheet1, Property_name, Data_table_cu_valori.Rows(s).Item(.Text), PropertyFlags.CUSTOM_SHEET_PROP)
                                                End If
                                            End If
                                        End With

                                        With ComboBox_HMM_ID
                                            If Not .Text = "" Then
                                                Dim Property_name As String
                                                If Not ComboBox_Sheet_Set_hmm_ID.Text = "" Then
                                                    Property_name = ComboBox_Sheet_Set_hmm_ID.Text
                                                Else
                                                    Property_name = .Text
                                                End If
                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(.Text)) = False Then
                                                    SetCustomProperty(Sheet1, Property_name, Data_table_cu_valori.Rows(s).Item(.Text), PropertyFlags.CUSTOM_SHEET_PROP)
                                                End If
                                            End If
                                        End With

                                        With ComboBox_MBL
                                            If Not .Text = "" Then
                                                Dim Property_name As String
                                                If Not ComboBox_Sheet_Set_MBL.Text = "" Then
                                                    Property_name = ComboBox_Sheet_Set_MBL.Text
                                                Else
                                                    Property_name = .Text
                                                End If
                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(.Text)) = False Then
                                                    SetCustomProperty(Sheet1, Property_name, Data_table_cu_valori.Rows(s).Item(.Text), PropertyFlags.CUSTOM_SHEET_PROP)
                                                End If
                                            End If
                                        End With


                                        With ComboBox_deed_page
                                            If Not .Text = "" Then
                                                Dim Property_name As String
                                                If Not ComboBox_Sheet_Set_deed_page.Text = "" Then
                                                    Property_name = ComboBox_Sheet_Set_deed_page.Text
                                                Else
                                                    Property_name = .Text
                                                End If
                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(.Text)) = False Then
                                                    SetCustomProperty(Sheet1, Property_name, Data_table_cu_valori.Rows(s).Item(.Text), PropertyFlags.CUSTOM_SHEET_PROP)
                                                End If
                                            End If
                                        End With

                                        With ComboBox_APN
                                            If Not .Text = "" Then
                                                Dim Property_name As String
                                                If Not ComboBox_Sheet_Set_APN.Text = "" Then
                                                    Property_name = ComboBox_Sheet_Set_APN.Text
                                                Else
                                                    Property_name = .Text
                                                End If
                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(.Text)) = False Then
                                                    SetCustomProperty(Sheet1, Property_name, Data_table_cu_valori.Rows(s).Item(.Text), PropertyFlags.CUSTOM_SHEET_PROP)
                                                End If
                                            End If
                                        End With

                                        With ComboBox_access_road_length
                                            If Not .Text = "" Then
                                                Dim Property_name As String
                                                If Not ComboBox_Sheet_Set_access_road_length.Text = "" Then
                                                    Property_name = ComboBox_Sheet_Set_access_road_length.Text
                                                Else
                                                    Property_name = .Text
                                                End If
                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(.Text)) = False Then
                                                    Dim Round1 As Integer = 2
                                                    If IsNumeric(TextBox_round_access_road_length.Text) Then
                                                        Round1 = Abs(CInt(TextBox_round_access_road_length.Text))
                                                    End If
                                                    SetCustomProperty(Sheet1, Property_name, Get_String_Rounded(Data_table_cu_valori.Rows(s).Item(.Text), Round1), PropertyFlags.CUSTOM_SHEET_PROP)
                                                End If
                                            End If
                                        End With

                                        With ComboBox_Section_TWP_Range
                                            If Not .Text = "" Then
                                                Dim Property_name As String
                                                If Not ComboBox_Sheet_Set_sec_twp_range.Text = "" Then
                                                    Property_name = ComboBox_Sheet_Set_sec_twp_range.Text
                                                Else
                                                    Property_name = .Text
                                                End If
                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(.Text)) = False Then
                                                    SetCustomProperty(Sheet1, Property_name, Data_table_cu_valori.Rows(s).Item(.Text), PropertyFlags.CUSTOM_SHEET_PROP)
                                                End If
                                            End If
                                        End With

                                        With ComboBox_owner
                                            If Not .Text = "" Then
                                                Dim Property_name As String
                                                If Not ComboBox_Sheet_Set_owner.Text = "" Then
                                                    Property_name = ComboBox_Sheet_Set_owner.Text
                                                Else
                                                    Property_name = .Text
                                                End If
                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(.Text)) = False Then
                                                    SetCustomProperty(Sheet1, Property_name, Data_table_cu_valori.Rows(s).Item(.Text), PropertyFlags.CUSTOM_SHEET_PROP)
                                                End If
                                            End If
                                        End With

                                        With ComboBox_crossing_length

                                            If Not .Text = "" Then
                                                Dim Property_name1 As String
                                                If Not ComboBox_Sheet_Set_crossing_len_FT.Text = "" Then
                                                    Property_name1 = ComboBox_Sheet_Set_crossing_len_FT.Text
                                                Else
                                                    Property_name1 = .Text
                                                End If
                                                Dim Property_name2 As String = ""
                                                If Not ComboBox_Sheet_Set_crossing_len_rod.Text = "" Then
                                                    Property_name2 = ComboBox_Sheet_Set_crossing_len_rod.Text
                                                End If
                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(.Text)) = False Then
                                                    Dim Round1 As Integer = 2
                                                    If IsNumeric(TextBox_round_crossing_length_ft.Text) Then
                                                        Round1 = Abs(CInt(TextBox_round_crossing_length_ft.Text))
                                                    End If

                                                    SetCustomProperty(Sheet1, Property_name1, Get_String_Rounded(Data_table_cu_valori.Rows(s).Item(.Text), Round1), PropertyFlags.CUSTOM_SHEET_PROP)
                                                    If Not Property_name2 = "" Then
                                                        Dim Round2 As Integer = 2
                                                        If IsNumeric(TextBox_round_crossing_length_rod.Text) Then
                                                            Round2 = Abs(CInt(TextBox_round_crossing_length_rod.Text))
                                                        End If
                                                        SetCustomProperty(Sheet1, Property_name2, Get_String_Rounded(Data_table_cu_valori.Rows(s).Item(.Text) / 16.5, Round2), PropertyFlags.CUSTOM_SHEET_PROP)
                                                    End If
                                                End If
                                            End If
                                        End With

                                        If IsNothing(Data_table_FROM_CL) = False Then
                                            With ComboBox_pipe_diameter
                                                If Not .Text = "" And Data_table_FROM_CL.Columns.Contains(.Text) = True Then
                                                    Dim Property_name As String
                                                    If Not ComboBox_Sheet_Set_pipe_DIAMETER.Text = "" Then
                                                        Property_name = ComboBox_Sheet_Set_pipe_DIAMETER.Text
                                                    Else
                                                        Property_name = .Text
                                                    End If
                                                    If IsDBNull(Data_table_FROM_CL.Rows(0).Item(.Text)) = False Then
                                                        SetCustomProperty(Sheet1, Property_name, Data_table_FROM_CL.Rows(0).Item(.Text), PropertyFlags.CUSTOM_SHEET_PROP)
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_pipe_name
                                                If Not .Text = "" And Data_table_FROM_CL.Columns.Contains(.Text) = True Then
                                                    Dim Property_name As String
                                                    If Not ComboBox_Sheet_Set_pipe_name.Text = "" Then
                                                        Property_name = ComboBox_Sheet_Set_pipe_name.Text
                                                    Else
                                                        Property_name = .Text
                                                    End If
                                                    If Property_name = "NAME" Then
                                                        MsgBox("you can not have NAME as a sheet set custom field")
                                                        Freeze_operations = False
                                                        Exit Sub
                                                    End If

                                                    If IsDBNull(Data_table_FROM_CL.Rows(0).Item(.Text)) = False Then
                                                        SetCustomProperty(Sheet1, Property_name, Data_table_FROM_CL.Rows(0).Item(.Text), PropertyFlags.CUSTOM_SHEET_PROP)
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_pipe_segment
                                                If Not .Text = "" And Data_table_FROM_CL.Columns.Contains(.Text) = True Then
                                                    Dim Property_name As String
                                                    If Not ComboBox_Sheet_Set_pipe_segment.Text = "" Then
                                                        Property_name = ComboBox_Sheet_Set_pipe_segment.Text
                                                    Else
                                                        Property_name = .Text
                                                    End If
                                                    If IsDBNull(Data_table_FROM_CL.Rows(0).Item(.Text)) = False Then
                                                        SetCustomProperty(Sheet1, Property_name, Data_table_FROM_CL.Rows(0).Item(.Text), PropertyFlags.CUSTOM_SHEET_PROP)
                                                    End If
                                                End If
                                            End With
                                        End If






                                        ' into dst you can insert only text values!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

                                        With ComboBox_area_EX_E
                                            If Not .Text = "" Then
                                                Dim Property_name As String
                                                If Not ComboBox_Sheet_Set_ex_e.Text = "" Then
                                                    Property_name = ComboBox_Sheet_Set_ex_e.Text
                                                Else
                                                    Property_name = .Text
                                                End If
                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(.Text)) = False Then
                                                    Dim Round1 As Integer = 2
                                                    If IsNumeric(TextBox_round_EX_E.Text) Then
                                                        Round1 = Abs(CInt(TextBox_round_EX_E.Text))
                                                    End If
                                                    SetCustomProperty(Sheet1, Property_name, Get_String_Rounded(Data_table_cu_valori.Rows(s).Item(.Text), Round1), PropertyFlags.CUSTOM_SHEET_PROP)
                                                End If
                                            End If
                                        End With

                                        With ComboBox_area_P_E
                                            If Not .Text = "" Then
                                                Dim Property_name As String
                                                If Not ComboBox_Sheet_Set_P_E.Text = "" Then
                                                    Property_name = ComboBox_Sheet_Set_P_E.Text
                                                Else
                                                    Property_name = .Text
                                                End If
                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(.Text)) = False Then
                                                    Dim Round1 As Integer = 2
                                                    If IsNumeric(TextBox_round_P_E.Text) Then
                                                        Round1 = Abs(CInt(TextBox_round_P_E.Text))
                                                    End If
                                                    SetCustomProperty(Sheet1, Property_name, Get_String_Rounded(Data_table_cu_valori.Rows(s).Item(.Text), Round1), PropertyFlags.CUSTOM_SHEET_PROP)
                                                End If
                                            End If
                                        End With

                                        With ComboBox_area_TWS
                                            If Not .Text = "" Then
                                                Dim Property_name As String
                                                If Not ComboBox_Sheet_Set_TWS.Text = "" Then
                                                    Property_name = ComboBox_Sheet_Set_TWS.Text
                                                Else
                                                    Property_name = .Text
                                                End If
                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(.Text)) = False Then
                                                    Dim Round1 As Integer = 2
                                                    If IsNumeric(TextBox_round_TWS.Text) Then
                                                        Round1 = Abs(CInt(TextBox_round_TWS.Text))
                                                    End If
                                                    SetCustomProperty(Sheet1, Property_name, Get_String_Rounded(Data_table_cu_valori.Rows(s).Item(.Text), Round1), PropertyFlags.CUSTOM_SHEET_PROP)
                                                End If
                                            End If
                                        End With

                                        With ComboBox_area_ATWS
                                            If Not .Text = "" Then
                                                Dim Property_name As String
                                                If Not ComboBox_Sheet_Set_atws.Text = "" Then
                                                    Property_name = ComboBox_Sheet_Set_atws.Text
                                                Else
                                                    Property_name = .Text
                                                End If
                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(.Text)) = False Then
                                                    Dim Round1 As Integer = 2
                                                    If IsNumeric(TextBox_round_ATWS.Text) Then
                                                        Round1 = Abs(CInt(TextBox_round_ATWS.Text))
                                                    End If
                                                    SetCustomProperty(Sheet1, Property_name, Get_String_Rounded(Data_table_cu_valori.Rows(s).Item(.Text), Round1), PropertyFlags.CUSTOM_SHEET_PROP)
                                                End If
                                            End If
                                        End With







                                        With ComboBox_area_A_R
                                            If Not .Text = "" Then
                                                Dim Property_name As String
                                                If Not ComboBox_Sheet_Set_a_R.Text = "" Then
                                                    Property_name = ComboBox_Sheet_Set_a_R.Text
                                                Else
                                                    Property_name = .Text
                                                End If
                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(.Text)) = False Then
                                                    Dim Round1 As Integer = 2
                                                    If IsNumeric(TextBox_round_A_R.Text) Then
                                                        Round1 = Abs(CInt(TextBox_round_A_R.Text))
                                                    End If
                                                    SetCustomProperty(Sheet1, Property_name, Get_String_Rounded(Data_table_cu_valori.Rows(s).Item(.Text), Round1), PropertyFlags.CUSTOM_SHEET_PROP)
                                                End If
                                            End If
                                        End With

                                        With ComboBox_area_WARE
                                            If Not .Text = "" Then
                                                Dim Property_name As String
                                                If Not ComboBox_Sheet_Set_ware.Text = "" Then
                                                    Property_name = ComboBox_Sheet_Set_ware.Text
                                                Else
                                                    Property_name = .Text
                                                End If
                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(.Text)) = False Then
                                                    Dim Round1 As Integer = 2
                                                    If IsNumeric(TextBox_round_WARE.Text) Then
                                                        Round1 = Abs(CInt(TextBox_round_WARE.Text))
                                                    End If
                                                    SetCustomProperty(Sheet1, Property_name, Get_String_Rounded(Data_table_cu_valori.Rows(s).Item(.Text), Round1), PropertyFlags.CUSTOM_SHEET_PROP)
                                                End If
                                            End If
                                        End With


                                        With ComboBox_area_TWS_ABD
                                            If Not .Text = "" Then
                                                Dim Property_name As String
                                                If Not ComboBox_Sheet_Set_tws_abd.Text = "" Then
                                                    Property_name = ComboBox_Sheet_Set_tws_abd.Text
                                                Else
                                                    Property_name = .Text
                                                End If
                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(.Text)) = False Then
                                                    Dim Round1 As Integer = 2
                                                    If IsNumeric(TextBox_round_TWS_ABD.Text) Then
                                                        Round1 = Abs(CInt(TextBox_round_TWS_ABD.Text))
                                                    End If
                                                    SetCustomProperty(Sheet1, Property_name, Get_String_Rounded(Data_table_cu_valori.Rows(s).Item(.Text), Round1), PropertyFlags.CUSTOM_SHEET_PROP)
                                                End If
                                            End If
                                        End With

                                        With ComboBox_area_TWS_PD
                                            If Not .Text = "" Then
                                                Dim Property_name As String
                                                If Not ComboBox_Sheet_Set_tws_PD.Text = "" Then
                                                    Property_name = ComboBox_Sheet_Set_tws_PD.Text
                                                Else
                                                    Property_name = .Text
                                                End If
                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(.Text)) = False Then
                                                    Dim Round1 As Integer = 2
                                                    If IsNumeric(TextBox_round_TWS_PD.Text) Then
                                                        Round1 = Abs(CInt(TextBox_round_TWS_PD.Text))
                                                    End If
                                                    SetCustomProperty(Sheet1, Property_name, Get_String_Rounded(Data_table_cu_valori.Rows(s).Item(.Text), Round1), PropertyFlags.CUSTOM_SHEET_PROP)
                                                End If
                                            End If
                                        End With

                                        If Not ComboBox_SHEET_SET_scale_main.Text = "" Then
                                            If IsDBNull(Data_table_cu_valori.Rows(s).Item(VW_CUST_SCALE)) = False Then
                                                SetCustomProperty(Sheet1, ComboBox_SHEET_SET_scale_main.Text, "1" & Chr(34) & "=" & (1 / Data_table_cu_valori.Rows(s).Item(VW_CUST_SCALE)).ToString & "'", PropertyFlags.CUSTOM_SHEET_PROP)
                                                Vw_scale = Data_table_cu_valori.Rows(s).Item(VW_CUST_SCALE)
                                            End If
                                        End If


                                        If CheckBox_add_locus.Checked = True And Not Not ComboBox_SHEET_SET_scale_locus.Text = "" Then
                                            If Not ComboBox_SHEET_SET_scale_locus.Text = "" Then
                                                SetCustomProperty(Sheet1, ComboBox_SHEET_SET_scale_locus.Text, "1" & Chr(34) & "=2000'", PropertyFlags.CUSTOM_SHEET_PROP)
                                            End If
                                        End If

                                        LockDatabase(SheetSet_database, False)
                                    Else

                                        MsgBox(SheetSet_database.GetLockStatus.ToString)
                                        ' Display error message
                                        MsgBox("Sheet set could not be opened for write, it is blocked by Autocad most likely.")
                                        Freeze_operations = False
                                        Exit Sub

                                    End If

                                    ' Close the sheet set
                                    SheetSet_manager.Close(SheetSet_database)

                                    Dim Fisier_nou As String = TextBox_Output_Directory.Text & New_plat_name & ".dwg"

                                    Dim DocumentManager1 As DocumentCollection = Application.DocumentManager
                                    Dim New_doc As Document = DocumentCollectionExtension.Open(DocumentManager1, Fisier_nou, False)

                                    Using lock2 As DocumentLock = New_doc.LockDocument
                                        Using Trans_New_doc As Autodesk.AutoCAD.DatabaseServices.Transaction = New_doc.TransactionManager.StartTransaction
                                            HostApplicationServices.WorkingDatabase = New_doc.Database

                                            Creaza_layer(Layer_name_Main_Viewport, 40, Layer_name_Main_Viewport, False)
                                            Creaza_layer(Layer_name_locus_VP, 7, Layer_name_Main_Viewport, True)
                                            Creaza_layer(Layer_poly_parcela_new, 7, "Boundary", True)
                                            Creaza_layer(Layer_name_no_plot1, 40, Layer_name_no_plot1, False)
                                            If CheckBox_add_locus.Checked = True Then
                                                Creaza_layer(Layer_hatch_locus, 251, "hatch for locus", True)
                                            End If
                                            Trans_New_doc.Commit()
                                        End Using
                                    End Using



                                    Dim Anno_scale_name1 As String = ""
                                    Dim Anno_scale_name2 As String = "1" & Chr(34) & "=2000'"
                                    Dim DWG_units As Integer
                                    Dim OBJ_id_Viewport_MAIN As ObjectId
                                    Dim OBJ_id_Viewport_LOCUS As ObjectId

                                    Using lock2 As DocumentLock = New_doc.LockDocument

                                        Using Trans_New_doc As Autodesk.AutoCAD.DatabaseServices.Transaction = New_doc.TransactionManager.StartTransaction
                                            HostApplicationServices.WorkingDatabase = New_doc.Database

                                            Dim ocm As ObjectContextManager = New_doc.Database.ObjectContextManager
                                            Dim occ As ObjectContextCollection

                                            If IsNothing(ocm) = False Then
                                                occ = ocm.GetContextCollection("ACDB_ANNOTATIONSCALES")
                                            End If

                                            Anno_scale_name1 = "1" & Chr(34) & "=" & Round((1 / Data_table_cu_valori.Rows(s).Item(VW_CUST_SCALE)), 0).ToString & "'"
                                            DWG_units = 1 / Data_table_cu_valori.Rows(s).Item(VW_CUST_SCALE)

                                            If IsNothing(occ) = False Then


                                                Dim asc As New AnnotationScale
                                                asc.Name = Anno_scale_name1
                                                asc.PaperUnits = 1
                                                asc.DrawingUnits = DWG_units
                                                If occ.HasContext(asc.Name) = False Then
                                                    occ.AddContext(asc)
                                                End If

                                            End If

                                            Dim BlockTable_new_doc As BlockTable = Trans_New_doc.GetObject(New_doc.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                            Dim BTrecord_new_doc_PS As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                            BTrecord_new_doc_PS = Trans_New_doc.GetObject(BlockTable_new_doc(BlockTableRecord.PaperSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                                            Dim BTrecord_new_doc_MS As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                            BTrecord_new_doc_MS = Trans_New_doc.GetObject(BlockTable_new_doc(BlockTableRecord.ModelSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                                            'Dim Id_mapping1 As New IdMapping
                                            ' New_doc.Database.WblockCloneObjects(Colectie_blocksIDs, BTrecord_new_doc_MS.ObjectId, Id_mapping1, DuplicateRecordCloning.Ignore, False)

                                            Dim DrawOrderTable1 As Autodesk.AutoCAD.DatabaseServices.DrawOrderTable = Trans_New_doc.GetObject(BTrecord_new_doc_MS.DrawOrderTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                                            Dim oBJiD_COL_Mtext As New ObjectIdCollection

                                            Dim Viewport1 As New Viewport
                                            Viewport1.SetDatabaseDefaults()
                                            Viewport1.CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(Vw_CenX, Vw_CenY, 0) ' asta e pozitia viewport in paper space
                                            Viewport1.Height = Vw_height
                                            Viewport1.Width = Vw_width
                                            Viewport1.Layer = Layer_name_Main_Viewport
                                            Viewport1.ViewDirection = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis
                                            Viewport1.ViewTarget = New Point3d(Data_table_cu_valori.Rows(s).Item(VW_TARGET_X), Data_table_cu_valori.Rows(s).Item(VW_TARGET_Y), 0) ' asta e pozitia viewport in MODEL space
                                            Viewport1.ViewCenter = Autodesk.AutoCAD.Geometry.Point2d.Origin
                                            Viewport1.TwistAngle = Data_table_cu_valori.Rows(s).Item(VW_TWIST) ' asta e PT TWIST



                                            BTrecord_new_doc_PS.AppendEntity(Viewport1)
                                            Trans_New_doc.AddNewlyCreatedDBObject(Viewport1, True)

                                            Viewport1.On = True
                                            Viewport1.CustomScale = Data_table_cu_valori.Rows(s).Item(VW_CUST_SCALE) 'Vw_width / (1.05 * Width1)

                                            Viewport1.Locked = True
                                            OBJ_id_Viewport_MAIN = Viewport1.ObjectId

                                            If CheckBox_add_locus.Checked = True Then
                                                Dim Viewport2 As New Viewport
                                                Viewport2.SetDatabaseDefaults()
                                                Viewport2.CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(VwL_CenX, VwL_CenY, 0) ' asta e pozitia viewport in paper space
                                                Viewport2.Height = VwL_height
                                                Viewport2.Width = VwL_width
                                                Viewport2.Layer = Layer_name_locus_VP
                                                Viewport2.ViewDirection = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis
                                                Viewport2.ViewTarget = New Point3d(Data_table_cu_valori.Rows(s).Item(VW_TARGET_X), Data_table_cu_valori.Rows(s).Item(VW_TARGET_Y), 0) ' asta e pozitia viewport in MODEL space
                                                Viewport2.ViewCenter = Autodesk.AutoCAD.Geometry.Point2d.Origin
                                                Viewport2.TwistAngle = Data_table_cu_valori.Rows(s).Item(VW_TWIST) ' asta e PT TWIST



                                                BTrecord_new_doc_PS.AppendEntity(Viewport2)
                                                Trans_New_doc.AddNewlyCreatedDBObject(Viewport2, True)

                                                Viewport2.On = True
                                                Viewport2.CustomScale = Locus_scale

                                                Viewport2.Locked = True
                                                OBJ_id_Viewport_LOCUS = Viewport2.ObjectId
                                            End If


                                            If IsDBNull(Data_table_cu_valori.Rows(s).Item(Obj_type)) = True Then
                                                Dim Poly123 As New Polyline
                                                If IsNothing(Colectie_Parcele(s)) = False Then
                                                    Dim col_int As New Point3dCollection
                                                    Poly_CL.IntersectWith(Colectie_Parcele(s), Intersect.OnBothOperands, col_int, IntPtr.Zero, IntPtr.Zero)

                                                    Dim Colectie_pct_CL_parc As New Point3dCollection

                                                    If col_int.Count > 0 Then
                                                        If col_int.Count = 1 Then
                                                            Dim Param1 As Double = Poly_CL.GetParameterAtPoint(col_int(0))
                                                            Dim Dist1 As Double = Poly_CL.GetDistAtPoint(col_int(0))

                                                            If Dist1 < Poly_CL.Length / 2 Then
                                                                For k = 0 To Floor(Param1)
                                                                    Colectie_pct_CL_parc.Add(Poly_CL.GetPointAtParameter(k))
                                                                Next
                                                                Colectie_pct_CL_parc.Add(Poly_CL.GetPointAtParameter(Param1))
                                                            Else
                                                                Colectie_pct_CL_parc.Add(Poly_CL.GetPointAtParameter(Param1))
                                                                For k = Ceiling(Param1) To Poly_CL.NumberOfVertices - 1
                                                                    Colectie_pct_CL_parc.Add(Poly_CL.GetPointAtParameter(k))
                                                                Next
                                                            End If
                                                        End If

                                                        If col_int.Count > 1 Then
                                                            Dim Param1 As Double = Poly_CL.GetParameterAtPoint(col_int(0))
                                                            Dim Param2 As Double = Poly_CL.GetParameterAtPoint(col_int(col_int.Count - 1))

                                                            If Param2 < Param1 Then
                                                                Dim t As Double
                                                                t = Param1
                                                                Param1 = Param2
                                                                Param2 = t
                                                            End If


                                                            If Floor(Param1) = Floor(Param2) Then
                                                                Colectie_pct_CL_parc.Add(Poly_CL.GetPointAtParameter(Param1))
                                                                Colectie_pct_CL_parc.Add(Poly_CL.GetPointAtParameter(Param2))
                                                            Else
                                                                Colectie_pct_CL_parc.Add(Poly_CL.GetPointAtParameter(Param1))
                                                                For k = Ceiling(Param1) To Floor(Param2)
                                                                    Colectie_pct_CL_parc.Add(Poly_CL.GetPointAtParameter(k))
                                                                Next
                                                                Colectie_pct_CL_parc.Add(Poly_CL.GetPointAtParameter(Param2))
                                                            End If
                                                        End If


                                                        For k = 0 To Colectie_pct_CL_parc.Count - 1
                                                            Poly123.AddVertexAt(k, New Point2d(Colectie_pct_CL_parc(k).X, Colectie_pct_CL_parc(k).Y), 0, 0, 0)
                                                        Next

                                                        Creaza_layer(Layer_name_text, 7, "Text", True)
                                                        Poly123.Layer = Layer_name_no_plot1
                                                        BTrecord_new_doc_MS.AppendEntity(Poly123)
                                                        Trans_New_doc.AddNewlyCreatedDBObject(Poly123, True)
                                                    End If
                                                End If


                                                If CheckBox_display_stations.Checked = True Then
                                                    If IsNothing(Colectie_Parcele(s)) = False Then
                                                        'If IsNothing(Poly123) = False Then
                                                        If Not Poly123.NumberOfVertices = 0 Then


                                                            Dim Point0 As New Point3d
                                                            Dim Point1 As New Point3d
                                                            Point0 = Poly123.GetPointAtDist(0)
                                                            Point1 = Poly123.GetPointAtParameter(1)



                                                            Dim Angle123 As Double
                                                            Dim Linie123 As New Line(Point0, Point1)
                                                            Angle123 = Linie123.Angle

                                                            Dim mText123 As New MText

                                                            If Linie123.Length > 0 Then
                                                                Linie123.TransformBy(Matrix3d.Scaling(((Text_height / Vw_scale) / Linie123.Length), Point0))
                                                                Linie123.TransformBy(Matrix3d.Displacement(Linie123.GetPointAtDist(Linie123.Length / 2).GetVectorTo(Point0)))
                                                                Linie123.TransformBy(Matrix3d.Rotation(PI / 2, Vector3d.ZAxis, Point0))
                                                                'BTrecord_new_doc_MS.AppendEntity(Linie123)
                                                                'Trans_New_doc.AddNewlyCreatedDBObject(Linie123, True)


                                                                mText123.Contents = "{\Fromans|c0;" & Get_chainage_feet_from_double(0, 0) & " P.O.B.}"
                                                                mText123.Location = Linie123.EndPoint
                                                                mText123.Rotation = Angle123 + PI / 2
                                                                mText123.Attachment = AttachmentPoint.MiddleLeft
                                                                mText123.TextHeight = Text_height / Vw_scale
                                                                mText123.Layer = Layer_name_text
                                                                mText123.BackgroundFill = True
                                                                mText123.UseBackgroundColor = True
                                                                mText123.BackgroundScaleFactor = 1.2

                                                                BTrecord_new_doc_MS.AppendEntity(mText123)
                                                                Trans_New_doc.AddNewlyCreatedDBObject(mText123, True)

                                                                oBJiD_COL_Mtext.Add(mText123.ObjectId)
                                                            End If
                                                            Point1 = Poly123.GetPointAtParameter(Poly123.NumberOfVertices - 2)
                                                                Point0 = Poly123.GetPointAtDist(Poly123.Length)

                                                                Linie123 = New Line(Point0, Point1)
                                                            Angle123 = Linie123.Angle + PI
                                                            If Linie123.Length > 0 Then
                                                                Linie123.TransformBy(Matrix3d.Scaling(((Text_height / Vw_scale) / Linie123.Length), Point0))
                                                                Linie123.TransformBy(Matrix3d.Displacement(Linie123.GetPointAtDist(Linie123.Length / 2).GetVectorTo(Point0)))
                                                                Linie123.TransformBy(Matrix3d.Rotation(PI / 2, Vector3d.ZAxis, Point0))
                                                                'BTrecord_new_doc_MS.AppendEntity(Linie123)
                                                                'Trans_New_doc.AddNewlyCreatedDBObject(Linie123, True)

                                                                mText123 = New MText
                                                                mText123.Contents = "{\Fromans|c0;" & Get_chainage_feet_from_double(Poly123.Length, 0) & " P.O.T.}"
                                                                mText123.Location = Linie123.StartPoint
                                                                mText123.Rotation = Angle123 + PI / 2
                                                                mText123.Attachment = AttachmentPoint.MiddleLeft
                                                                mText123.TextHeight = Text_height / Vw_scale
                                                                mText123.Layer = Layer_name_text
                                                                mText123.BackgroundFill = True
                                                                mText123.UseBackgroundColor = True
                                                                mText123.BackgroundScaleFactor = 1.2
                                                                BTrecord_new_doc_MS.AppendEntity(mText123)
                                                                Trans_New_doc.AddNewlyCreatedDBObject(mText123, True)

                                                                oBJiD_COL_Mtext.Add(mText123.ObjectId)
                                                            End If
                                                            If Poly123.Length > 500 Then
                                                                    For k = 1 To Floor(Poly123.Length / 500)

                                                                        Point0 = Poly123.GetPointAtDist(k * 500)
                                                                        Dim Param0 As Double = Poly123.GetParameterAtDistance(k * 500)
                                                                        Point1 = Poly123.GetPointAtParameter(Floor(Param0) + 1)

                                                                        Linie123 = New Line(Point0, Point1)
                                                                        Angle123 = Linie123.Angle
                                                                        If Linie123.Length > 0 Then
                                                                            Linie123.TransformBy(Matrix3d.Scaling(((Text_height / Vw_scale) / Linie123.Length), Point0))
                                                                            Linie123.TransformBy(Matrix3d.Displacement(Linie123.GetPointAtDist(Linie123.Length / 2).GetVectorTo(Point0)))
                                                                            Linie123.TransformBy(Matrix3d.Rotation(PI / 2, Vector3d.ZAxis, Point0))
                                                                            'BTrecord_new_doc_MS.AppendEntity(Linie123)
                                                                            'Trans_New_doc.AddNewlyCreatedDBObject(Linie123, True)

                                                                            mText123 = New MText
                                                                            mText123.Contents = "{\Fromans|c0;" & Get_chainage_feet_from_double(k * 500, 0) & "}"
                                                                            mText123.Location = Linie123.EndPoint
                                                                            mText123.Rotation = Angle123
                                                                            mText123.Attachment = AttachmentPoint.BottomCenter
                                                                            mText123.TextHeight = Text_height / Vw_scale
                                                                            mText123.Layer = Layer_name_text
                                                                            mText123.BackgroundFill = True
                                                                            mText123.UseBackgroundColor = True
                                                                            mText123.BackgroundScaleFactor = 1.2
                                                                            BTrecord_new_doc_MS.AppendEntity(mText123)
                                                                            Trans_New_doc.AddNewlyCreatedDBObject(mText123, True)
                                                                            oBJiD_COL_Mtext.Add(mText123.ObjectId)
                                                                        End If

                                                                    Next
                                                                End If
                                                            End If
                                                        End If
                                                End If

                                                If CheckBox_add_bearing_and_distances.Checked = True Then
                                                    If IsNothing(Colectie_Parcele(s)) = False Then
                                                        If Not Poly123.NumberOfVertices = 0 Then
                                                            Dim n As Integer = 1

                                                            Do While n < Poly123.NumberOfVertices

                                                                Dim ParamPicked As Double = n - 0.5
                                                                Dim Start3D As New Point3d
                                                                Dim End3D As New Point3d


                                                                Dim Poly2 As Polyline

                                                                If CheckBox_ignore_deflections_less_0_5.Checked = True Then
                                                                    Poly2 = Bearing_dist_calc_with_min_deflection(0.5 * PI / 180, Poly123, ParamPicked)
                                                                Else
                                                                    Poly2 = Bearing_dist_calc_with_min_deflection(0, Poly123, ParamPicked)
                                                                End If

                                                                Start3D = Poly2.StartPoint
                                                                End3D = Poly2.EndPoint

                                                                Dim Bear_rad As Double = GET_Bearing_rad(Start3D.X, Start3D.Y, End3D.X, End3D.Y)

                                                                Dim Start2D As New Point3d(Start3D.X, Start3D.Y, 0)
                                                                Dim End2D As New Point3d(End3D.X, End3D.Y, 0)
                                                                Dim Dist2D As Double = Abs(Poly123.GetDistAtPoint(Start3D) - Poly123.GetDistAtPoint(End3D))


                                                                Dim Param_end As Integer = Poly123.GetParameterAtPoint(End3D)
                                                                n = Param_end + 1

                                                                Dim Continut As String = "{\Fromans|c0;\Q20;\L" & Quadrant_bearings(Bear_rad) & " - " & Round(Dist2D, 0).ToString & "'±}"

                                                                Dim Mtext_Bear_dist As New MText
                                                                Mtext_Bear_dist.Contents = Continut
                                                                Mtext_Bear_dist.Location = New Point3d((Start3D.X + End3D.X) / 2, (Start3D.Y + End3D.Y) / 2, 0)
                                                                Mtext_Bear_dist.Rotation = 2 * PI - Data_table_cu_valori.Rows(s).Item(VW_TWIST)
                                                                Mtext_Bear_dist.Attachment = AttachmentPoint.BottomLeft
                                                                Mtext_Bear_dist.TextHeight = Text_height / Vw_scale
                                                                Mtext_Bear_dist.Layer = Layer_name_text
                                                                Mtext_Bear_dist.BackgroundFill = True
                                                                Mtext_Bear_dist.UseBackgroundColor = True
                                                                Mtext_Bear_dist.BackgroundScaleFactor = 1.2
                                                                BTrecord_new_doc_MS.AppendEntity(Mtext_Bear_dist)
                                                                Trans_New_doc.AddNewlyCreatedDBObject(Mtext_Bear_dist, True)
                                                                oBJiD_COL_Mtext.Add(Mtext_Bear_dist.ObjectId)

                                                            Loop


                                                        End If
                                                    End If
                                                End If
                                            End If

                                            Dim Poly_parcela_from_points As New Polyline

                                            If IsDBNull(Data_table_cu_valori.Rows(s).Item(Obj_type)) = False Then

                                                If IsNothing(Colectie_Parcele_for_points(s)) = False Then

                                                    For k = 0 To Colectie_Parcele_for_points(s).NumberOfVertices - 1
                                                        Poly_parcela_from_points.AddVertexAt(k, Colectie_Parcele_for_points(s).GetPoint2dAt(k), 0, 0, 0)
                                                    Next

                                                    BTrecord_new_doc_MS.AppendEntity(Poly_parcela_from_points)
                                                    Trans_New_doc.AddNewlyCreatedDBObject(Poly_parcela_from_points, True)

                                                End If
                                            End If


                                            If Not ComboBox_county.Text = "" Then
                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(ComboBox_county.Text)) = False Then
                                                    Dim TextString As String = Data_table_cu_valori.Rows(s).Item(ComboBox_county.Text)
                                                    Dim Mtext1 As New MText
                                                    Mtext1.TextHeight = 0.08
                                                    Mtext1.Location = County_pt
                                                    Mtext1.Contents = "COUNTY = " & TextString
                                                    Mtext1.Layer = Layer_name_no_plot1
                                                    BTrecord_new_doc_PS.AppendEntity(Mtext1)
                                                    Trans_New_doc.AddNewlyCreatedDBObject(Mtext1, True)
                                                End If
                                            End If
                                            If Not ComboBox_state.Text = "" Then
                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(ComboBox_state.Text)) = False Then
                                                    Dim TextString As String = Data_table_cu_valori.Rows(s).Item(ComboBox_state.Text)
                                                    Dim Mtext1 As New MText
                                                    Mtext1.TextHeight = 0.08
                                                    Mtext1.Location = State_pt
                                                    Mtext1.Contents = "STATE = " & TextString
                                                    Mtext1.Layer = Layer_name_no_plot1
                                                    BTrecord_new_doc_PS.AppendEntity(Mtext1)
                                                    Trans_New_doc.AddNewlyCreatedDBObject(Mtext1, True)
                                                End If
                                            End If

                                            If Not ComboBox_town.Text = "" Then
                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(ComboBox_town.Text)) = False Then
                                                    Dim TextString As String = Data_table_cu_valori.Rows(s).Item(ComboBox_town.Text)
                                                    Dim Mtext1 As New MText
                                                    Mtext1.TextHeight = 0.08
                                                    Mtext1.Location = Town_pt
                                                    Mtext1.Contents = "TOWN = " & TextString
                                                    Mtext1.Layer = Layer_name_no_plot1
                                                    BTrecord_new_doc_PS.AppendEntity(Mtext1)
                                                    Trans_New_doc.AddNewlyCreatedDBObject(Mtext1, True)
                                                End If
                                            End If
                                            If Not ComboBox_owner.Text = "" Then
                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(ComboBox_owner.Text)) = False Then
                                                    Dim TextString As String = Data_table_cu_valori.Rows(s).Item(ComboBox_owner.Text)
                                                    Dim Mtext1 As New MText
                                                    Mtext1.TextHeight = 0.08
                                                    Mtext1.Location = Owner_pt
                                                    Mtext1.Contents = "OWNER = " & TextString
                                                    Mtext1.Layer = Layer_name_no_plot1
                                                    BTrecord_new_doc_PS.AppendEntity(Mtext1)
                                                    Trans_New_doc.AddNewlyCreatedDBObject(Mtext1, True)

                                                End If
                                            End If




                                            If Not ComboBox_linelist.Text = "" Then
                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(ComboBox_linelist.Text)) = False Then
                                                    Dim TextString As String = Data_table_cu_valori.Rows(s).Item(ComboBox_linelist.Text)
                                                    Dim Mtext1 As New MText
                                                    Mtext1.TextHeight = 0.08
                                                    Mtext1.Location = Line_list_pt
                                                    Mtext1.Contents = "LINELIST = " & TextString
                                                    Mtext1.Layer = Layer_name_no_plot1
                                                    BTrecord_new_doc_PS.AppendEntity(Mtext1)
                                                    Trans_New_doc.AddNewlyCreatedDBObject(Mtext1, True)
                                                End If
                                            End If

                                            If Not ComboBox_HMM_ID.Text = "" Then
                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(ComboBox_HMM_ID.Text)) = False Then
                                                    Dim TextString As String = Data_table_cu_valori.Rows(s).Item(ComboBox_HMM_ID.Text)
                                                    Dim Mtext1 As New MText
                                                    Mtext1.TextHeight = 0.08
                                                    Mtext1.Location = HMMID_pt
                                                    Mtext1.Contents = "HMM_ID = " & TextString
                                                    Mtext1.Layer = Layer_name_no_plot1
                                                    BTrecord_new_doc_PS.AppendEntity(Mtext1)
                                                    Trans_New_doc.AddNewlyCreatedDBObject(Mtext1, True)
                                                End If
                                            End If

                                            If Not ComboBox_pipe_diameter.Text = "" And Data_table_FROM_CL.Columns.Contains(ComboBox_pipe_diameter.Text) = True Then
                                                If IsDBNull(Data_table_FROM_CL.Rows(0).Item(ComboBox_pipe_diameter.Text)) = False Then
                                                    Dim TextString As String = Data_table_FROM_CL.Rows(0).Item(ComboBox_pipe_diameter.Text)
                                                    Dim Mtext1 As New MText
                                                    Mtext1.TextHeight = 0.08
                                                    Mtext1.Location = Diameter_pt
                                                    Mtext1.Contents = "DIAMETER = " & TextString
                                                    Mtext1.Layer = Layer_name_no_plot1
                                                    BTrecord_new_doc_PS.AppendEntity(Mtext1)
                                                    Trans_New_doc.AddNewlyCreatedDBObject(Mtext1, True)
                                                End If
                                            End If

                                            If Not ComboBox_pipe_name.Text = "" And Data_table_FROM_CL.Columns.Contains(ComboBox_pipe_name.Text) = True Then
                                                If IsDBNull(Data_table_FROM_CL.Rows(0).Item(ComboBox_pipe_name.Text)) = False Then
                                                    Dim TextString As String = Data_table_FROM_CL.Rows(0).Item(ComboBox_pipe_name.Text)
                                                    Dim Mtext1 As New MText
                                                    Mtext1.TextHeight = 0.08
                                                    Mtext1.Location = Name_pt
                                                    Mtext1.Layer = Layer_name_no_plot1
                                                    Mtext1.Contents = "NAME = " & TextString
                                                    BTrecord_new_doc_PS.AppendEntity(Mtext1)
                                                    Trans_New_doc.AddNewlyCreatedDBObject(Mtext1, True)
                                                End If
                                            End If

                                            If Not ComboBox_pipe_segment.Text = "" And Data_table_FROM_CL.Columns.Contains(ComboBox_pipe_segment.Text) = True Then
                                                If IsDBNull(Data_table_FROM_CL.Rows(0).Item(ComboBox_pipe_segment.Text)) = False Then
                                                    Dim TextString As String = Data_table_FROM_CL.Rows(0).Item(ComboBox_pipe_segment.Text)
                                                    Dim Mtext1 As New MText
                                                    Mtext1.TextHeight = 0.08
                                                    Mtext1.Location = Segment_pt
                                                    Mtext1.Layer = Layer_name_no_plot1
                                                    Mtext1.Contents = "SEGMENT = " & TextString
                                                    BTrecord_new_doc_PS.AppendEntity(Mtext1)
                                                    Trans_New_doc.AddNewlyCreatedDBObject(Mtext1, True)

                                                End If
                                            End If

                                            Dim Colectie_atr_name As New Specialized.StringCollection
                                            Dim Colectie_atr_value As New Specialized.StringCollection

                                            If Not TextBox_north_arrow.Text = "" Then
                                                Dim Block_north As Autodesk.AutoCAD.DatabaseServices.BlockReference
                                                If IsNumeric(TextBox_north_arrow_Big_X.Text) = True And IsNumeric(TextBox_north_arrow_Big_y.Text) = True Then
                                                    Punct_insertie_north_arrow = New Point3d(CDbl(TextBox_north_arrow_Big_X.Text), CDbl(TextBox_north_arrow_Big_y.Text), 0)
                                                End If

                                                Dim Nume_North_arrow As String = TextBox_north_arrow.Text


                                                Block_north = InsertBlock_with_multiple_atributes("", Nume_North_arrow, Punct_insertie_north_arrow, 1, BTrecord_new_doc_PS, Layer_name_Blocks, Colectie_atr_name, Colectie_atr_value)
                                                Block_north.Rotation = Data_table_cu_valori.Rows(s).Item(VW_TWIST)
                                            End If


                                            If Not TextBox_north_arrow_locus.Text = "" Then
                                                If CheckBox_add_locus.Checked = True Then
                                                    Dim Block_Lnorth As Autodesk.AutoCAD.DatabaseServices.BlockReference
                                                    If IsNumeric(TextBox_north_arrow_small_X.Text) = True And IsNumeric(TextBox_north_arrow_small_Y.Text) = True Then
                                                        Punct_insertie_Lnorth_arrow = New Point3d(CDbl(TextBox_north_arrow_small_X.Text), CDbl(TextBox_north_arrow_small_Y.Text), 0)
                                                    End If

                                                    Dim Nume_LNorth_arrow As String = TextBox_north_arrow_locus.Text


                                                    Block_Lnorth = InsertBlock_with_multiple_atributes("", Nume_LNorth_arrow, Punct_insertie_Lnorth_arrow, 1, BTrecord_new_doc_PS, Layer_name_Blocks, Colectie_atr_name, Colectie_atr_value)
                                                    Block_Lnorth.Rotation = Data_table_cu_valori.Rows(s).Item(VW_TWIST)
                                                End If
                                            End If



                                            If System.IO.File.Exists(TextBox_xref_model_space.Text) = True Then
                                                Dim xrefModelSpace As ObjectId = New_doc.Database.AttachXref(TextBox_xref_model_space.Text, "basemapXref")
                                                Dim br_ms As New BlockReference(New Point3d(0, 0, 0), xrefModelSpace)
                                                BTrecord_new_doc_MS.AppendEntity(br_ms)
                                                Trans_New_doc.AddNewlyCreatedDBObject(br_ms, True)

                                                Dim oBJiD_COL As New ObjectIdCollection
                                                oBJiD_COL.Add(br_ms.ObjectId)
                                                DrawOrderTable1.MoveToBottom(oBJiD_COL)


                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(Obj_type)) = True Then
                                                    Dim Poly_parcela_new As New Polyline
                                                    For h = 0 To Colectie_Parcele(s).NumberOfVertices - 1
                                                        Poly_parcela_new.AddVertexAt(h, Colectie_Parcele(s).GetPoint2dAt(h), Colectie_Parcele(s).GetBulgeAt(h), Colectie_Parcele(s).GetStartWidthAt(h), Colectie_Parcele(s).GetEndWidthAt(h))
                                                    Next
                                                    Poly_parcela_new.LinetypeScale = LinetypeScale_Poly_parcela_new
                                                    Poly_parcela_new.Layer = Layer_poly_parcela_new
                                                    Poly_parcela_new.Closed = True
                                                    BTrecord_new_doc_MS.AppendEntity(Poly_parcela_new)
                                                    Trans_New_doc.AddNewlyCreatedDBObject(Poly_parcela_new, True)

                                                    If CheckBox_add_locus.Checked = True Then
                                                        Dim Hatch1 As New Hatch
                                                        BTrecord_new_doc_MS.AppendEntity(Hatch1)
                                                        Trans_New_doc.AddNewlyCreatedDBObject(Hatch1, True)
                                                        Hatch1.SetHatchPattern(HatchPatternType.PreDefined, "SOLID")
                                                        Dim oBJiD_COL_H As New ObjectIdCollection
                                                        oBJiD_COL_H.Add(Poly_parcela_new.ObjectId)
                                                        Hatch1.AppendLoop(HatchLoopTypes.External, oBJiD_COL_H)
                                                        Hatch1.Layer = Layer_hatch_locus
                                                        Hatch1.EvaluateHatch(True)
                                                    End If
                                                End If

                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(Obj_type)) = True Then
                                                    Using filter As New Filters.SpatialFilter
                                                        Dim dictName As String = "ACAD_FILTER"
                                                        Dim spName As String = "SPATIAL"
                                                        Dim ptCol As New Point2dCollection
                                                        For h = 0 To Colectie_Parcele(s).NumberOfVertices - 1
                                                            ptCol.Add(Colectie_Parcele(s).GetPoint2dAt(h))
                                                        Next
                                                        Dim elev As Double = Colectie_Parcele(s).Elevation
                                                        Dim normal As Vector3d
                                                        If Application.DocumentManager.MdiActiveDocument.Database.TileMode = True Then
                                                            normal = Application.DocumentManager.MdiActiveDocument.Database.Ucsxdir.CrossProduct(Application.DocumentManager.MdiActiveDocument.Database.Ucsydir)
                                                        Else
                                                            normal = Application.DocumentManager.MdiActiveDocument.Database.Pucsxdir.CrossProduct(Application.DocumentManager.MdiActiveDocument.Database.Pucsydir)
                                                        End If

                                                        Dim filterDef As New Filters.SpatialFilterDefinition(ptCol, normal, elev, 0, 0, True)
                                                        filter.Definition = filterDef

                                                        If br_ms.ExtensionDictionary.IsNull Then
                                                            br_ms.UpgradeOpen()
                                                            br_ms.CreateExtensionDictionary()
                                                            br_ms.DowngradeOpen()
                                                        End If

                                                        Dim extDict As DBDictionary = Trans_New_doc.GetObject(br_ms.ExtensionDictionary, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                                                        ' Check to see if the dictionary for clipped boundaries exists, 
                                                        ' and add the spatial filter to the dictionary
                                                        If extDict.Contains(dictName) Then
                                                            Dim filterDict As DBDictionary = Trans_New_doc.GetObject(extDict.GetAt(dictName), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                                                            If (filterDict.Contains(spName)) Then filterDict.Remove(spName)

                                                            filterDict.SetAt(spName, filter)
                                                        Else
                                                            Using filterDict As New DBDictionary
                                                                extDict.SetAt(dictName, filterDict)

                                                                Trans_New_doc.AddNewlyCreatedDBObject(filterDict, True)
                                                                filterDict.SetAt(spName, filter)
                                                            End Using
                                                        End If

                                                        ' Append the spatial filter to the drawing
                                                        Trans_New_doc.AddNewlyCreatedDBObject(filter, True)
                                                    End Using
                                                End If

                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(Obj_type)) = False Then
                                                    Using filter As New Filters.SpatialFilter
                                                        Dim dictName As String = "ACAD_FILTER"
                                                        Dim spName As String = "SPATIAL"
                                                        Dim ptCol As New Point2dCollection
                                                        For h = 0 To Colectie_Parcele_for_points(s).NumberOfVertices - 1
                                                            ptCol.Add(Colectie_Parcele_for_points(s).GetPoint2dAt(h))
                                                        Next
                                                        Dim elev As Double = Colectie_Parcele_for_points(s).Elevation
                                                        Dim normal As Vector3d
                                                        If Application.DocumentManager.MdiActiveDocument.Database.TileMode = True Then
                                                            normal = Application.DocumentManager.MdiActiveDocument.Database.Ucsxdir.CrossProduct(Application.DocumentManager.MdiActiveDocument.Database.Ucsydir)
                                                        Else
                                                            normal = Application.DocumentManager.MdiActiveDocument.Database.Pucsxdir.CrossProduct(Application.DocumentManager.MdiActiveDocument.Database.Pucsydir)
                                                        End If

                                                        Dim filterDef As New Filters.SpatialFilterDefinition(ptCol, normal, elev, 0, 0, True)
                                                        filter.Definition = filterDef

                                                        If br_ms.ExtensionDictionary.IsNull Then
                                                            br_ms.UpgradeOpen()
                                                            br_ms.CreateExtensionDictionary()
                                                            br_ms.DowngradeOpen()
                                                        End If

                                                        Dim extDict As DBDictionary = Trans_New_doc.GetObject(br_ms.ExtensionDictionary, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                                                        ' Check to see if the dictionary for clipped boundaries exists, 
                                                        ' and add the spatial filter to the dictionary
                                                        If extDict.Contains(dictName) Then
                                                            Dim filterDict As DBDictionary = Trans_New_doc.GetObject(extDict.GetAt(dictName), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                                                            If (filterDict.Contains(spName)) Then filterDict.Remove(spName)

                                                            filterDict.SetAt(spName, filter)
                                                        Else
                                                            Using filterDict As New DBDictionary
                                                                extDict.SetAt(dictName, filterDict)

                                                                Trans_New_doc.AddNewlyCreatedDBObject(filterDict, True)
                                                                filterDict.SetAt(spName, filter)
                                                            End Using
                                                        End If

                                                        ' Append the spatial filter to the drawing
                                                        Trans_New_doc.AddNewlyCreatedDBObject(filter, True)
                                                    End Using
                                                End If

                                            End If
                                            'If System.IO.File.Exists(TextBox_xref_paper_space.Text) = True Then
                                            'Dim xrefPaperSpace As ObjectId = New_doc.Database.AttachXref(TextBox_xref_paper_space.Text, "TBLKXref")
                                            'Dim br_ps As New BlockReference(New Point3d(0, 0, 0), xrefPaperSpace)
                                            'BTrecord_new_doc_PS.AppendEntity(br_ps)
                                            'Trans_New_doc.AddNewlyCreatedDBObject(br_ps, True)
                                            'End If
                                            If oBJiD_COL_Mtext.Count > 0 Then
                                                DrawOrderTable1.MoveToTop(oBJiD_COL_Mtext)
                                            End If
                                            Trans_New_doc.Commit()
                                        End Using




                                        If Data_table_layers.Rows.Count > 0 Then
                                            Using Trans_New_doc As Autodesk.AutoCAD.DatabaseServices.Transaction = New_doc.TransactionManager.StartTransaction
                                                HostApplicationServices.WorkingDatabase = New_doc.Database
                                                Dim Layer_table As Autodesk.AutoCAD.DatabaseServices.LayerTable = Trans_New_doc.GetObject(New_doc.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                                                Dim Colectie_Layers_OFF_Main_Viewport As New ObjectIdCollection
                                                Dim Colectie_Layers_OFF_Locus_Viewport As New ObjectIdCollection
                                                Dim Colectie_Layers_ON_Main_Viewport As New ObjectIdCollection
                                                Dim Colectie_Layers_ON_Locus_Viewport As New ObjectIdCollection


                                                For Each ID1 As ObjectId In Layer_table
                                                    Dim LayerTRec1 As LayerTableRecord = Trans_New_doc.GetObject(ID1, OpenMode.ForRead)

                                                    For i = 0 To Data_table_layers.Rows.Count - 1

                                                        If IsDBNull(Data_table_layers.Rows(i).Item("VIEWPORT_TYPE")) = False And _
                                                            IsDBNull(Data_table_layers.Rows(i).Item("LAYER_NAME")) = False And _
                                                            IsDBNull(Data_table_layers.Rows(i).Item("THAW_FREEZE")) = False Then

                                                            Dim Viewport_type As String = Data_table_layers.Rows(i).Item("VIEWPORT_TYPE")
                                                            Dim Nume_layer As String = Data_table_layers.Rows(i).Item("LAYER_NAME")
                                                            Dim Thaw_freeze As String = Data_table_layers.Rows(i).Item("THAW_FREEZE")

                                                            If LayerTRec1.Name = Nume_layer Then
                                                                If Viewport_type = "MAIN" Then
                                                                    If Thaw_freeze = "THAW" Then
                                                                        Colectie_Layers_ON_Main_Viewport.Add(ID1)
                                                                    End If
                                                                    If Thaw_freeze = "FROZEN" Then
                                                                        Colectie_Layers_OFF_Main_Viewport.Add(ID1)
                                                                    End If

                                                                End If
                                                                If Viewport_type = "LOCUS" Then
                                                                    If Thaw_freeze = "THAW" Then
                                                                        Colectie_Layers_ON_Locus_Viewport.Add(ID1)
                                                                    End If
                                                                    If Thaw_freeze = "FROZEN" Then
                                                                        Colectie_Layers_OFF_Locus_Viewport.Add(ID1)
                                                                    End If
                                                                End If
                                                                If IsDBNull(Data_table_layers.Rows(i).Item("COLOR_INDEX")) = False Then
                                                                    LayerTRec1.UpgradeOpen()
                                                                    LayerTRec1.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, Data_table_layers.Rows(i).Item("COLOR_INDEX"))

                                                                End If
                                                                Exit For

                                                            End If
                                                        End If



                                                    Next

                                                Next

                                                Dim ocm As ObjectContextManager = New_doc.Database.ObjectContextManager
                                                Dim occ As ObjectContextCollection
                                                If IsNothing(ocm) = False Then
                                                    occ = ocm.GetContextCollection("ACDB_ANNOTATIONSCALES")
                                                End If
                                                Dim Viewport1 As Viewport = Trans_New_doc.GetObject(OBJ_id_Viewport_MAIN, OpenMode.ForWrite)
                                                If IsNothing(occ) = False And Not Anno_scale_name1 = "" Then
                                                    Dim Anno_scale As ObjectContext = occ.GetContext(Anno_scale_name1)
                                                    If IsNothing(Anno_scale) = False Then
                                                        Viewport1.AnnotationScale = Anno_scale
                                                    End If
                                                End If

                                                If IsNothing(Colectie_Layers_ON_Main_Viewport) = False Then
                                                    If Colectie_Layers_ON_Main_Viewport.Count > 0 Then
                                                        Viewport1.ThawLayersInViewport(Colectie_Layers_ON_Main_Viewport.GetEnumerator)
                                                    End If
                                                End If

                                                If IsNothing(Colectie_Layers_OFF_Main_Viewport) = False Then
                                                    If Colectie_Layers_OFF_Main_Viewport.Count > 0 Then
                                                        Viewport1.FreezeLayersInViewport(Colectie_Layers_OFF_Main_Viewport.GetEnumerator)
                                                    End If
                                                End If

                                                If CheckBox_add_locus.Checked = True Then
                                                    Dim Viewport2 As Viewport = Trans_New_doc.GetObject(OBJ_id_Viewport_LOCUS, OpenMode.ForWrite)
                                                    If IsNothing(occ) = False And Not Anno_scale_name2 = "" Then
                                                        Dim Anno_scale As ObjectContext = occ.GetContext(Anno_scale_name2)
                                                        If IsNothing(Anno_scale) = False Then
                                                            Viewport2.AnnotationScale = Anno_scale
                                                        End If
                                                    End If

                                                    If IsNothing(Colectie_Layers_ON_Locus_Viewport) = False Then
                                                        If Colectie_Layers_ON_Locus_Viewport.Count > 0 Then
                                                            Viewport2.ThawLayersInViewport(Colectie_Layers_ON_Locus_Viewport.GetEnumerator)
                                                        End If
                                                    End If

                                                    If IsNothing(Colectie_Layers_OFF_Locus_Viewport) = False Then
                                                        If Colectie_Layers_OFF_Locus_Viewport.Count > 0 Then
                                                            Viewport2.FreezeLayersInViewport(Colectie_Layers_OFF_Locus_Viewport.GetEnumerator)
                                                        End If
                                                    End If
                                                End If








                                                Trans_New_doc.Commit()
                                            End Using
                                        End If
                                        New_doc.Database.SaveAs(TextBox_Output_Directory.Text & New_plat_name & ".dwg", True, DwgVersion.Current, BaseMap_drawing.Database.SecurityParameters)
                                    End Using
                                    DocumentExtension.CloseAndDiscard(New_doc)
                                    HostApplicationServices.WorkingDatabase = BaseMap_drawing.Database
                                Next
                            End If 'If System.IO.Directory.Exists(TextBox_Output_Directory.Text) = True Then
                        End If
                    End If

                    MsgBox("You are done")
                Catch ex As Exception
                    Freeze_operations = False
                    MsgBox(ex.Message)
                End Try
            Else
                MsgBox("XREF " & TextBox_xref_model_space.Text & " DOES NOT EXIST")
            End If
            Freeze_operations = False
        End If
    End Sub

    Private Sub RadioButton10_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton10.CheckedChanged
        If RadioButton10.Checked = True Then
            Vw_scale = 1 / 10
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        End If
    End Sub

    Private Sub RadioButton20_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton20.CheckedChanged
        If RadioButton20.Checked = True Then
            Vw_scale = 1 / 20
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        End If
    End Sub

    Private Sub RadioButton30_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton30.CheckedChanged
        If RadioButton30.Checked = True Then
            Vw_scale = 1 / 30
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        End If
    End Sub

    Private Sub RadioButton40_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton40.CheckedChanged
        If RadioButton40.Checked = True Then
            Vw_scale = 1 / 40
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        End If
    End Sub

    Private Sub RadioButton50_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton50.CheckedChanged
        If RadioButton50.Checked = True Then
            Vw_scale = 1 / 50
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        End If
    End Sub

    Private Sub RadioButton60_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton60.CheckedChanged
        If RadioButton60.Checked = True Then
            Vw_scale = 1 / 60
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        End If
    End Sub

    Private Sub RadioButton100_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton100.CheckedChanged
        If RadioButton100.Checked = True Then
            Vw_scale = 1 / 100
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        End If
    End Sub

    Private Sub RadioButton200_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton200.CheckedChanged
        If RadioButton200.Checked = True Then
            Vw_scale = 1 / 200
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        End If
    End Sub

    Private Sub RadioButton300_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton300.CheckedChanged
        If RadioButton300.Checked = True Then
            Vw_scale = 1 / 300
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        End If
    End Sub

    Private Sub RadioButton400_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton400.CheckedChanged
        If RadioButton400.Checked = True Then
            Vw_scale = 1 / 400
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        End If
    End Sub

    Private Sub RadioButton500_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton500.CheckedChanged
        If RadioButton500.Checked = True Then
            Vw_scale = 1 / 500
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        End If
    End Sub

    Private Sub RadioButton600_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton600.CheckedChanged
        If RadioButton600.Checked = True Then
            Vw_scale = 1 / 600
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        End If
    End Sub

    Private Sub RadioButton1000_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1000.CheckedChanged
        If RadioButton1000.Checked = True Then
            Vw_scale = 1 / 1000
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        End If
    End Sub

    Private Sub HScrollBar_rotate_Scroll(sender As Object, e As Windows.Forms.ScrollEventArgs) Handles HScrollBar_rotate.Scroll
        Dim Valoare_rot As Double = HScrollBar_rotate.Value
        TextBox_rotation_ammount.Text = Valoare_rot
        Rotatie = Rotatie_originala - Valoare_rot * PI / 180

        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
    End Sub

    Private Sub Button_load_object_data_FROM_CL_Click(sender As Object, e As EventArgs) Handles Button_load_object_data_from_CL.Click
        Dim Empty_array() As ObjectId
        If Freeze_operations = False Then
            Freeze_operations = True

            Dim BaseMap_drawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = BaseMap_drawing.Editor
            Try
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()


                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptEntityResult

                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select entity containing object data:")

                Rezultat1 = Editor1.GetEntity(Object_Prompt)

                If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.WriteMessage(vbLf & "Command:")
                    Freeze_operations = False
                    Editor1.SetImpliedSelection(Empty_array)
                    Exit Sub
                End If

                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                    If IsNothing(Rezultat1) = False Then

                        Using Lock_dwg As DocumentLock = BaseMap_drawing.LockDocument


                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = BaseMap_drawing.TransactionManager.StartTransaction

                                Dim Ent1 As Entity = Trans1.GetObject(Rezultat1.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                                Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                                Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables
                                Dim Id1 As ObjectId = Ent1.ObjectId

                                Using Records1 As Autodesk.Gis.Map.ObjectData.Records = Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                    If IsNothing(Records1) = False Then
                                        If Records1.Count > 0 Then
                                            Dim Record1 As Autodesk.Gis.Map.ObjectData.Record



                                            For Each Record1 In Records1
                                                Dim Tabla1 As Autodesk.Gis.Map.ObjectData.Table
                                                Tabla1 = Tables1(Record1.TableName)


                                                Dim Field_defs1 As Autodesk.Gis.Map.ObjectData.FieldDefinitions
                                                Field_defs1 = Tabla1.FieldDefinitions
                                                For i = 0 To Record1.Count - 1
                                                    Dim Field_def1 As Autodesk.Gis.Map.ObjectData.FieldDefinition
                                                    Field_def1 = Field_defs1(i)

                                                    With ComboBox_pipe_diameter
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With
                                                    With ComboBox_pipe_name
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With
                                                    With ComboBox_pipe_segment
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With
                                                Next



                                            Next


                                            With ComboBox_pipe_diameter
                                                If .Items.Count > 0 Then
                                                    If ComboBox_pipe_diameter.Items.Contains("DIAMETER") = True Then
                                                        .SelectedIndex = .Items.IndexOf("DIAMETER")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_pipe_name
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("NAME") = True Then
                                                        .SelectedIndex = .Items.IndexOf("NAME")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_pipe_segment
                                                If .Items.Count > 0 Then
                                                    If ComboBox_pipe_segment.Items.Contains("SEGMENT") = True Then
                                                        .SelectedIndex = .Items.IndexOf("SEGMENT")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                        End If
                                    End If
                                End Using






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

    Private Sub Button_browse_Output_Directory_Click(sender As Object, e As EventArgs) Handles Button_browse_Output_Directory.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                Dim FolderBrowserDialog1 As New Windows.Forms.FolderBrowserDialog
                FolderBrowserDialog1.ShowNewFolderButton = False
                If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    TextBox_Output_Directory.Text = FolderBrowserDialog1.SelectedPath
                End If

            Catch ex As System.Exception
                Freeze_operations = False
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If


    End Sub

    Private Sub Button_xref_Modelspace_Click(sender As Object, e As EventArgs) Handles Button_xref_Modelspace.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                Dim FileBrowserDialog1 As New Windows.Forms.OpenFileDialog
                FileBrowserDialog1.Filter = "Drawing Files (*.dwg)|*.dwg|All Files (*.*)|*.*"


                If FileBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    TextBox_xref_model_space.Text = FileBrowserDialog1.FileName
                End If

            Catch ex As System.Exception
                Freeze_operations = False
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_dwt_template_Click(sender As Object, e As EventArgs) Handles Button_dwt_template.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                Dim FileBrowserDialog1 As New Windows.Forms.OpenFileDialog
                FileBrowserDialog1.Filter = "Template Files (*.dwt)|*.dwt|All Files (*.*)|*.*"
                If FileBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    TextBox_dwt_template.Text = FileBrowserDialog1.FileName
                End If
            Catch ex As System.Exception
                Freeze_operations = False
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_sheet_set_template_Click(sender As Object, e As EventArgs) Handles Button_sheet_set_template.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                Dim FileBrowserDialog1 As New Windows.Forms.OpenFileDialog
                FileBrowserDialog1.Filter = "Sheet Set Files (*.dst)|*.dst|All Files (*.*)|*.*"
                If FileBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    TextBox_sheet_set_template.Text = FileBrowserDialog1.FileName
                End If
            Catch ex As System.Exception
                Freeze_operations = False
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub

    Private Function AddSheet(ByVal component As IAcSmComponent,
                              ByVal name As String,
                              ByVal description As String,
                              ByVal title As String,
                              ByVal number As String) As AcSmSheet

        Dim sheet As AcSmSheet

        ' Check to see if the component is a sheet set or subset, 
        ' and create the new sheet based on the component's type
        If component.GetTypeName = "AcSmSubset" Then
            Dim subset As AcSmSubset = component
            sheet = subset.AddNewSheet(name, description)

            ' Add the sheet as the first one in the subset
            subset.InsertComponent(sheet, Nothing)
        Else
            sheet = component.GetDatabase().GetSheetSet().AddNewSheet(name,
                                                                      description)

            ' Add the sheet as the first one in the sheet set
            component.GetDatabase().GetSheetSet().InsertComponent(sheet, Nothing)
        End If

        ' Set the number and title of the sheet
        sheet.SetNumber(number)
        sheet.SetTitle(title)

        AddSheet = sheet
    End Function

    Private Function ImportASheet(ByVal component As IAcSmComponent, _
                                  ByVal title As String, _
                                  ByVal description As String, _
                                  ByVal number As String, _
                                  ByVal fileName As String, _
                                  ByVal layout As String) As AcSmSheet

        Dim sheet As AcSmSheet

        ' Create a reference to a Layout Reference object
        Dim layoutReference As New AcSmAcDbLayoutReference
        layoutReference.InitNew(component)

        ' Set the layout and drawing file to use for the sheet
        layoutReference.SetFileName(fileName)
        layoutReference.SetName(layout)

        ' Import the sheet into the sheet set
        ' Check to see if the Component is a Subset or Sheet Set
        If component.GetTypeName = "AcSmSubset" Then
            Dim subset As AcSmSubset = component

            sheet = subset.ImportSheet(layoutReference)
            subset.InsertComponent(sheet, Nothing)
        Else
            Dim sheetSetDatabase As AcSmDatabase = component

            sheet = sheetSetDatabase.GetSheetSet().ImportSheet(layoutReference)
            sheetSetDatabase.GetSheetSet().InsertComponent(sheet, Nothing)
        End If

        ' Set the properties of the sheet
        sheet.SetDesc(description)
        sheet.SetTitle(title)
        sheet.SetNumber(number)

        ImportASheet = sheet
    End Function

    Private Sub SetSheetSetDefaults(ByVal sheetSetDatabase As AcSmDatabase, _
                                    ByVal name As String, _
                                  ByVal description As String, _
                                   ByVal newSheetLocation As String, _
                                  Optional ByVal newSheetDWTLocation As String = "", _
                                  Optional ByVal newSheetDWTLayout As String = "", _
                                  Optional ByVal promptForDWT As Boolean = False)

        ' Set the Name and Description for the sheet set
        sheetSetDatabase.GetSheetSet().SetName(name)
        sheetSetDatabase.GetSheetSet().SetDesc(description)

        ' Check to see if a Storage Location was provided
        If newSheetLocation <> "" Then
            ' Get the folder the sheet set is stored in
            Dim sheetSetFolder As String
            sheetSetFolder = Mid(sheetSetDatabase.GetFileName(), 1, InStrRev(sheetSetDatabase.GetFileName(), "\"))

            ' Create a reference to a File Reference object
            Dim fileReference As IAcSmFileReference
            fileReference = sheetSetDatabase.GetSheetSet().GetNewSheetLocation()

            ' Set the default storage location based on the location of the sheet set
            fileReference.SetFileName(newSheetLocation)

            ' Set the new Sheet location for the sheet set
            sheetSetDatabase.GetSheetSet().SetNewSheetLocation(fileReference)
        End If

        ' Check to see if a Template was provided
        If newSheetDWTLocation <> "" Then
            ' Set the Default Template for the sheet set
            Dim layoutReference As AcSmAcDbLayoutReference
            layoutReference = sheetSetDatabase.GetSheetSet().GetDefDwtLayout()

            ' Set the template location and name of the layout 
            ' for the Layout Reference object
            layoutReference.SetFileName(newSheetDWTLocation)
            layoutReference.SetName(newSheetDWTLayout)

            ' Set the Layout Reference for the sheet set
            sheetSetDatabase.GetSheetSet().SetDefDwtLayout(layoutReference)
        End If

        ' Set the Prompt for Template option of the subset
        sheetSetDatabase.GetSheetSet().SetPromptForDwt(promptForDWT)
    End Sub

    Private Sub Button_read_from_excel_Click(sender As Object, e As EventArgs) Handles Button_read_from_excel.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                Button_read_from_excel.Visible = False
                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Colectie_ID = New Specialized.StringCollection
                Dim Start1 As Integer = 0
                Dim End1 As Integer = 0
                Dim Column_id As String = TextBox_Column_ID_Excel.Text.ToUpper
                If IsNumeric(TextBox_Start_row_excel.Text) = True Then
                    Start1 = CDbl(TextBox_Start_row_excel.Text)
                End If
                If IsNumeric(TextBox_End_row_excel.Text) = True Then
                    End1 = CDbl(TextBox_End_row_excel.Text)
                End If

                If Not Start1 = 0 Or Not End1 = 0 Then
                    If End1 > 0 And Start1 > 0 And End1 >= Start1 Then
                        For i = Start1 To End1
                            If Not Replace(W1.Range(Column_id & i).Value, " ", "") = "" Then
                                Colectie_ID.Add(W1.Range(Column_id & i).Value)
                            End If

                        Next
                    End If
                End If


            Catch ex As System.Exception
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try
            Button_read_from_excel.Visible = True
            Freeze_operations = False
        End If
    End Sub

    Private Sub SetCustomProperty(ByVal owner As IAcSmPersist, _
                                      ByVal propertyName As String, _
                                      ByVal propertyValue As Object, _
                                      ByVal sheetSetFlag As PropertyFlags)

        ' Create a reference to the Custom Property Bag

        Dim customPropertyBag As AcSmCustomPropertyBag


        If owner.GetTypeName() = "AcSmSheet" Then
            Dim sheet As AcSmSheet = owner
            customPropertyBag = sheet.GetCustomPropertyBag()
        Else
            Dim sheetSet As AcSmSheetSet = owner
            customPropertyBag = sheetSet.GetCustomPropertyBag()
        End If
        ' Create a reference to a Custom Property Value
        Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
        customPropertyValue.InitNew(owner)
        ' Set the flag for the property
        customPropertyValue.SetFlags(sheetSetFlag)
        ' Set the value for the property
        customPropertyValue.SetValue(propertyValue)
        ' Create the property
        customPropertyBag.SetProperty(propertyName, customPropertyValue)

    End Sub

    Private Sub Button_station_at_point_Click(sender As Object, e As EventArgs) Handles Button_station_at_point.Click
        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Empty_array() As ObjectId
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument



            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Dim Current_UCS_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem

            Dim Text_height As Double = 0.08
            Dim Dist_fixed As Double = 0.12
            Dim Tilemode1 As Integer = Application.GetSystemVariable("TILEMODE")
            Dim CVport1 As Integer = Application.GetSystemVariable("CVPORT")




            If Viewport_loaded = True And Tilemode1 = 1 Then
                Text_height = Text_height * Scale_factor
                Dist_fixed = Dist_fixed * Scale_factor
            Else


                If (Tilemode1 = 0 And Not CVport1 = 1) Or Tilemode1 = 1 Then
                    For Each ctrl2 As Windows.Forms.Control In Panel_SCALE_SELECTION.Controls
                        Dim Radiob2 As Windows.Forms.RadioButton
                        If TypeOf ctrl2 Is Windows.Forms.RadioButton Then
                            Radiob2 = ctrl2
                            If Radiob2.Checked = True Then
                                Dim Nume1 As String = Replace(Radiob2.Name, "RadioButton", "")
                                If IsNumeric(Nume1) = True Then
                                    Text_height = Text_height * CInt(Nume1)
                                    Dist_fixed = Dist_fixed * CInt(Nume1)
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                End If

            End If


            Try
                Using Lock1 As DocumentLock = ThisDrawing.LockDocument
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                    Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt.MessageForAdding = vbLf & "Select polyline:"

                    Object_Prompt.SingleOnly = True

                    Rezultat1 = Editor1.GetSelection(Object_Prompt)


                    If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Editor1.WriteMessage(vbLf & "Command:")
                        Freeze_operations = False
                        Exit Sub
                    End If
                    Dim Poly1 As Polyline
                    Dim Point_on_poly As New Point3d
                    If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        If IsNothing(Rezultat1) = False Then
                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Rezultat1.Value.Item(0)
                                Dim Ent1 As Entity
                                Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                    Poly1 = Ent1
                                Else
                                    Editor1.WriteMessage("No Polyline")
                                    Editor1.WriteMessage(vbLf & "Command:")
                                    Freeze_operations = False
                                    Exit Sub
                                End If
                            End Using
                        End If
                    End If

                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                        Creaza_layer(Layer_name_no_plot1, 40, "No plot", False)
                        Creaza_layer(Station_at_point_layer, 7, "Text", True)
                        Trans1.Commit()
                    End Using

                    Dim Planul_curent As Plane = New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)
                    Dim Xaxis1 As Vector3d = Current_UCS_matrix.CoordinateSystem3d.Xaxis
                    Dim Angle_Xaxis As Double = (Xaxis1.AngleOnPlane(Planul_curent))

1234:





                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                        Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Pick a position:")
                        Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                        PP1.AllowNone = True
                        Point1 = Editor1.GetPoint(PP1)
                        If Not Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Editor1.SetImpliedSelection(Empty_array)
                            Trans1.Commit()
                            Freeze_operations = False
                            Exit Sub
                        End If





                        Dim Distanta_pana_la_xing As Double
                        If IsNothing(Poly1) = False Then
                            Point_on_poly = Poly1.GetClosestPointTo(Point1.Value.TransformBy(Current_UCS_matrix), Vector3d.ZAxis, False)
                            Distanta_pana_la_xing = Poly1.GetDistAtPoint(Point_on_poly)
                        End If


                        Dim Station1 As Double = Distanta_pana_la_xing
                        Dim Param_mid As Double = Poly1.GetParameterAtDistance(Station1)
                        Dim Param1 As Integer = Floor(Param_mid)
                        If Param1 = Poly1.NumberOfVertices - 1 Then
                            Param1 = Poly1.NumberOfVertices - 2
                        End If
                        Dim Param2 As Integer = Param1 + 1

                        Dim Station_as_string As String = Get_chainage_feet_from_double(Station1, 0)
                        If Station_as_string = "-0+00" Then Station_as_string = "0+00"

                        Dim Line1 As New Line(Point_on_poly, Point1.Value.TransformBy(Current_UCS_matrix))
                        Line1.TransformBy(Matrix3d.Scaling((Line1.Length + Dist_fixed) / Line1.Length, Point_on_poly))

                        Dim Side As String = "L"

                        Dim Line_segment As New Line(Poly1.GetPointAtParameter(Param1), Poly1.GetPointAtParameter(Param2))

                        If Directie_offset(Line_segment, Point1.Value.TransformBy(Current_UCS_matrix)) = -1 Then
                            Side = "R"
                        End If

                        'Dim Point2 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                        'Dim Jig1 As jig_Mtext_class
                        'Jig1 = New jig_Mtext_class(New MText, Text_height, Bearing_segment + PI / 2, "{\Fromans|c0;" & Station_as_string & "}")
                        'Point2 = Jig1.BeginJig

                        'If IsNothing(Point2) = True Then
                        'Editor1.SetImpliedSelection(Empty_array)
                        'Trans1.Commit()
                        'Freeze_operations = False
                        'Exit Sub
                        'End If


                        'If Not Point2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        'Editor1.SetImpliedSelection(Empty_array)
                        'Trans1.Commit()
                        'Freeze_operations = False
                        'Exit Sub
                        'End If



                        Dim Insertion_point As New Point3d
                        Insertion_point = Line1.EndPoint

                        ''xxx()

                        Dim Mtext1 As New MText
                        Mtext1.Location = Insertion_point
                        Mtext1.TextHeight = Text_height
                        Mtext1.Contents = "{\Fromans|c0;" & Station_as_string & "}"



                        'GET_Bearing_rad(Line1.StartPoint.X, Line1.StartPoint.Y, Line1.EndPoint.X, Line1.EndPoint.Y)
                        If Side = "R" Then
                            Mtext1.Attachment = AttachmentPoint.MiddleRight
                            Mtext1.Rotation = Line1.Angle - Angle_Xaxis + PI
                        Else
                            Mtext1.Attachment = AttachmentPoint.MiddleLeft
                            Mtext1.Rotation = Line1.Angle - Angle_Xaxis
                        End If

                        Mtext1.Layer = Station_at_point_layer
                        Mtext1.BackgroundFill = True
                        Mtext1.UseBackgroundColor = True
                        Mtext1.BackgroundScaleFactor = 1.2
                        BTrecord.AppendEntity(Mtext1)
                        Trans1.AddNewlyCreatedDBObject(Mtext1, True)

                        Dim Linie3 As New Line(Point1.Value.TransformBy(Current_UCS_matrix), Insertion_point)
                        Linie3.Layer = Layer_name_no_plot1
                        BTrecord.AppendEntity(Linie3)
                        Trans1.AddNewlyCreatedDBObject(Linie3, True)
                        Trans1.TransactionManager.QueueForGraphicsFlush()

                        Trans1.Commit()

                        GoTo 1234

                    End Using


                    Editor1.SetImpliedSelection(Empty_array)
                    Editor1.WriteMessage(vbLf & "Command:")
                End Using
            Catch ex As Exception
                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_Dim_dist_Click(sender As Object, e As EventArgs) Handles Button_Dim_dist.Click
        Dim Empty_array() As ObjectId
        If Freeze_operations = False Then
            Freeze_operations = True
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Dim Current_UCS_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
            Dim rad_small As Double = 0.25
            Dim Inaltimea1 As Double = 1000 * 0.0625
            Dim Text_height As Double = 0.08
            Dim Text_rotation As Double = 0



            Dim Tilemode1 As Integer = Application.GetSystemVariable("TILEMODE")
            Dim CVport1 As Integer = Application.GetSystemVariable("CVPORT")

            Dim view1 As ViewTableRecord = ThisDrawing.Editor.GetCurrentView
            Dim Planul_curent As Plane = New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)
            Dim Xaxis1 As Vector3d = Current_UCS_matrix.CoordinateSystem3d.Xaxis
            Dim Angle_Xaxis As Double = (Xaxis1.AngleOnPlane(Planul_curent))

            Text_rotation = (2 * PI - view1.ViewTwist) - Angle_Xaxis


            If Viewport_loaded = True And Tilemode1 = 1 Then
                Inaltimea1 = Inaltimea1 * Scale_factor
                rad_small = rad_small * Scale_factor
                Text_height = Text_height * Scale_factor
                Text_rotation = View_rotation
            Else
                If (Tilemode1 = 0 And Not CVport1 = 1) Or Tilemode1 = 1 Then
                    For Each ctrl2 As Windows.Forms.Control In Panel_SCALE_SELECTION.Controls
                        Dim Radiob2 As Windows.Forms.RadioButton
                        If TypeOf ctrl2 Is Windows.Forms.RadioButton Then
                            Radiob2 = ctrl2
                            If Radiob2.Checked = True Then
                                Dim Nume1 As String = Replace(Radiob2.Name, "RadioButton", "")
                                If IsNumeric(Nume1) = True Then
                                    Inaltimea1 = Inaltimea1 * CInt(Nume1)
                                    rad_small = rad_small * CInt(Nume1)
                                    Text_height = Text_height * CInt(Nume1)
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                End If

            End If







            Dim Arr_len As Double = Inaltimea1 / 500
            Dim Arr_width As Double = Inaltimea1 / 1250

            Dim OLD_OSnap As Integer = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE")

            Try
                Using Lock1 As DocumentLock = ThisDrawing.LockDocument

                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                        Creaza_layer(Tie_distance_layer, 7, "", True)
                        Trans1.Commit()
                    End Using

                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()


                    Dim NEW_OSnap As Integer = Autodesk.AutoCAD.EditorInput.ObjectSnapMasks.End + Autodesk.AutoCAD.EditorInput.ObjectSnapMasks.Intersection
                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap)


                    Using Trans1 As Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                        Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Pick 1st point:")
                        Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                        PP1.AllowNone = True
                        Point1 = Editor1.GetPoint(PP1)
                        If Not Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
                            Freeze_operations = False
                            Exit Sub
                        End If

                        Dim PP2 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Pick 2nd point:")
                        Dim Point2 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                        PP2.AllowNone = True
                        PP2.BasePoint = Point1.Value
                        PP2.UseBasePoint = True
                        Point2 = Editor1.GetPoint(PP2)
                        If Not Point2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
                            Freeze_operations = False
                            Exit Sub
                        End If


                        If Point1.Value.GetVectorTo(Point2.Value).Length < 2 * Arr_len Then
                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
                            Freeze_operations = False
                            Exit Sub
                        End If

                        Dim JigAcolade As New JIG_ACOLADE
                        Dim Point3 As Autodesk.AutoCAD.EditorInput.PromptPointResult

                        Point3 = JigAcolade.StartJig(Point1.Value.TransformBy(Current_UCS_matrix), Point2.Value.TransformBy(Current_UCS_matrix), Arr_len, Arr_width, rad_small)
                        If IsNothing(Point3) = True Then
                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
                            Freeze_operations = False
                            Exit Sub
                        End If
                        If Not Point3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
                            Freeze_operations = False
                            Exit Sub
                        End If


                        Dim Arc3d As New CircularArc3d(Point1.Value.TransformBy(Current_UCS_matrix), Point2.Value.TransformBy(Current_UCS_matrix), Point3.Value)
                        Dim Circle1 As New Circle(Arc3d.Center, Vector3d.ZAxis, Point1.Value.TransformBy(Current_UCS_matrix).GetVectorTo(Arc3d.Center).Length)


                        Dim Line1 As New Line(Circle1.Center, Point3.Value)
                        Dim Scale_f As Double = (((Circle1.Radius + rad_small) ^ 2 - rad_small ^ 2) ^ 0.5) / Circle1.Radius
                        Line1.TransformBy(Matrix3d.Scaling(Scale_f, Circle1.Center))



                        Dim PointT As New Point3d
                        PointT = Line1.EndPoint
                        Dim PointA As New Point3d
                        PointA = Line1.GetPointAtDist(Line1.Length - rad_small)
                        Dim LinieR As New Line(PointT, PointA)

                        Dim LinieL As New Line
                        LinieL = LinieR.Clone
                        LinieL.TransformBy(Matrix3d.Rotation(PI / 2, Vector3d.ZAxis, PointT))
                        LinieR.TransformBy(Matrix3d.Rotation(-PI / 2, Vector3d.ZAxis, PointT))


                        Dim Circle2 As New Circle(LinieR.EndPoint, Vector3d.ZAxis, rad_small)

                        Dim Circle3 As New Circle(LinieL.EndPoint, Vector3d.ZAxis, rad_small)

                        Dim PtI1 As New Point3d
                        Dim PtI2 As New Point3d

                        Dim Colint1 As New Point3dCollection
                        Circle1.IntersectWith(Circle2, Intersect.OnBothOperands, Colint1, IntPtr.Zero, IntPtr.Zero)
                        If Colint1.Count > 0 Then
                            PtI1 = Colint1(0)
                        End If

                        Dim Colint2 As New Point3dCollection
                        Circle1.IntersectWith(Circle3, Intersect.OnBothOperands, Colint2, IntPtr.Zero, IntPtr.Zero)
                        If Colint2.Count > 0 Then
                            PtI2 = Colint2(0)
                        End If

                        If Colint1.Count > 0 And Colint2.Count > 0 Then
                            If IsNothing(PtI1) = False And IsNothing(PtI2) = False Then

                                Dim AngleStart As Double = GET_Bearing_rad(Circle2.Center.X, Circle2.Center.Y, PtI1.X, PtI1.Y)
                                Dim AngleEnd As Double = GET_Bearing_rad(Circle2.Center.X, Circle2.Center.Y, PointT.X, PointT.Y)
                                Dim Arc1 As New Arc(Circle2.Center, rad_small, AngleStart, AngleEnd)

                                AngleStart = GET_Bearing_rad(Circle3.Center.X, Circle3.Center.Y, PtI2.X, PtI2.Y)
                                AngleEnd = GET_Bearing_rad(Circle3.Center.X, Circle3.Center.Y, PointT.X, PointT.Y)
                                Dim Arc2 As New Arc(Circle3.Center, rad_small, AngleEnd, AngleStart)

                                Dim PointB3 As New Point3d
                                Dim PointB4 As New Point3d
                                PointB3 = Point1.Value.TransformBy(Current_UCS_matrix)
                                PointB4 = Point2.Value.TransformBy(Current_UCS_matrix)

                                If PointB3.GetVectorTo(Circle2.Center).Length < PointB3.GetVectorTo(Circle3.Center).Length Then
                                    Dim T As New Point3d
                                    T = PointB3
                                    PointB3 = PointB4
                                    PointB4 = T
                                End If

                                AngleStart = GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PtI2.X, PtI2.Y)
                                AngleEnd = GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PointB3.X, PointB3.Y)
                                Dim Arc3 As New Arc(Circle1.Center, Circle1.Radius, AngleEnd, AngleStart)


                                AngleStart = GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PtI1.X, PtI1.Y)
                                AngleEnd = GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PointB4.X, PointB4.Y)
                                Dim Arc4 As New Arc(Circle1.Center, Circle1.Radius, AngleStart, AngleEnd)

                                Dim Poly123 As New Polyline
                                Dim b0, b1, b2, b3 As Double

                                b0 = Tan(Arc3.TotalAngle / 4)
                                b1 = -Tan(Arc2.TotalAngle / 4)
                                b2 = -Tan(Arc1.TotalAngle / 4)
                                b3 = Tan(Arc4.TotalAngle / 4)

                                If Arc3.Length > Arr_len Then
                                    Dim PtArr3 As New Point3d
                                    PtArr3 = Arc3.GetPointAtDist(Arr_len)


                                    AngleStart = GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PtArr3.X, PtArr3.Y)
                                    AngleEnd = GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PointB3.X, PointB3.Y)
                                    Dim Arc31 As New Arc(Circle1.Center, Circle1.Radius, AngleEnd, AngleStart)
                                    Dim b01 As Double = Tan(Arc31.TotalAngle / 4)


                                    Poly123.AddVertexAt(0, New Point2d(PointB3.X, PointB3.Y), b01, 0, Arr_width)

                                    AngleStart = GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PtI2.X, PtI2.Y)
                                    AngleEnd = GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PtArr3.X, PtArr3.Y)
                                    Dim Arc32 As New Arc(Circle1.Center, Circle1.Radius, AngleEnd, AngleStart)
                                    Dim b02 As Double = Tan(Arc32.TotalAngle / 4)


                                    Poly123.AddVertexAt(1, New Point2d(PtArr3.X, PtArr3.Y), b02, 0, 0)
                                    Poly123.AddVertexAt(2, New Point2d(PtI2.X, PtI2.Y), b1, 0, 0)
                                    Poly123.AddVertexAt(3, New Point2d(PointT.X, PointT.Y), b2, 0, 0)

                                    If Arc4.Length > Arr_len Then
                                        Dim PtArr4 As New Point3d
                                        PtArr4 = Arc4.GetPointAtDist(Arc4.Length - Arr_len)

                                        AngleStart = GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PtArr4.X, PtArr4.Y)
                                        AngleEnd = GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PtI1.X, PtI1.Y)
                                        Dim Arc41 As New Arc(Circle1.Center, Circle1.Radius, AngleEnd, AngleStart)
                                        Dim b41 As Double = Tan(Arc41.TotalAngle / 4)
                                        Poly123.AddVertexAt(4, New Point2d(PtI1.X, PtI1.Y), b41, 0, 0)

                                        AngleStart = GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PointB4.X, PointB4.Y)
                                        AngleEnd = GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PtArr4.X, PtArr4.Y)
                                        Dim Arc42 As New Arc(Circle1.Center, Circle1.Radius, AngleEnd, AngleStart)
                                        Dim b42 As Double = Tan(Arc42.TotalAngle / 4)

                                        Poly123.AddVertexAt(5, New Point2d(PtArr4.X, PtArr4.Y), b42, Arr_width, 0)
                                        Poly123.AddVertexAt(6, New Point2d(PointB4.X, PointB4.Y), 0, 0, 0)
                                        Poly123.Layer = Tie_distance_layer

                                        BTrecord.AppendEntity(Poly123)
                                        Trans1.AddNewlyCreatedDBObject(Poly123, True)
                                        Trans1.TransactionManager.QueueForGraphicsFlush()

                                        Dim Distance_string As String = "{\Fromans|c0;" & Get_String_Rounded(Point1.Value.GetVectorTo(Point2.Value).Length, 0) & "'±}"

                                        Dim PointMtext As Autodesk.AutoCAD.EditorInput.PromptPointResult
                                        Dim Jig_mt As jig_Mtext_class
                                        Jig_mt = New jig_Mtext_class(New MText, Text_height, Text_rotation, Distance_string)
                                        PointMtext = Jig_mt.BeginJig

                                        'yyy()

                                        If IsNothing(PointMtext) = True Then
                                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
                                            Freeze_operations = False
                                            Exit Sub
                                        End If

                                        If PointMtext.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
                                            Freeze_operations = False
                                            Exit Sub
                                        End If

                                        Dim Mtext1 As New MText
                                        Mtext1.Location = PointMtext.Value
                                        Mtext1.TextHeight = Text_height
                                        Mtext1.Contents = Distance_string
                                        Mtext1.Attachment = AttachmentPoint.MiddleCenter
                                        Mtext1.BackgroundFill = True
                                        Mtext1.UseBackgroundColor = True
                                        Mtext1.BackgroundScaleFactor = 1.2
                                        Mtext1.Rotation = Text_rotation
                                        BTrecord.AppendEntity(Mtext1)
                                        Mtext1.Layer = Tie_distance_layer
                                        Trans1.AddNewlyCreatedDBObject(Mtext1, True)
                                        Trans1.TransactionManager.QueueForGraphicsFlush()


                                        Trans1.Commit()




                                    End If
                                End If
                            End If
                        End If




                    End Using






                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)

                    Editor1.SetImpliedSelection(Empty_array)
                    Editor1.WriteMessage(vbLf & "Command:")
                End Using
            Catch ex As Exception
                Freeze_operations = False
                Editor1.SetImpliedSelection(Empty_array)
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
                Editor1.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_bear_dist_Click(sender As Object, e As EventArgs) Handles Button_bear_dist.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Dim Empty_array() As ObjectId
            Dim Text_height As Double = 0.08
            Dim Text_rotation As Double = 0

            Dim OLD_OSnap As Integer = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE")

            Dim NEW_OSnap As Integer = Autodesk.AutoCAD.EditorInput.ObjectSnapMasks.Near
            Dim Current_UCS_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Dim Inaltimea1 As Double = 1000 * 0.0625

            Dim Tilemode1 As Integer = Application.GetSystemVariable("TILEMODE")
            Dim CVport1 As Integer = Application.GetSystemVariable("CVPORT")
            Dim view1 As ViewTableRecord = ThisDrawing.Editor.GetCurrentView
            Dim Planul_curent As Plane = New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)
            Dim Xaxis1 As Vector3d = Current_UCS_matrix.CoordinateSystem3d.Xaxis
            Dim Angle_Xaxis As Double = (Xaxis1.AngleOnPlane(Planul_curent))

            Text_rotation = (2 * PI - view1.ViewTwist) - Angle_Xaxis

            If Viewport_loaded = True And Tilemode1 = 1 Then
                Text_height = Text_height * Scale_factor
                Inaltimea1 = Inaltimea1 * Scale_factor
                Text_rotation = View_rotation
            Else
                If (Tilemode1 = 0 And Not CVport1 = 1) Or Tilemode1 = 1 Then
                    For Each ctrl2 As Windows.Forms.Control In Panel_SCALE_SELECTION.Controls
                        Dim Radiob2 As Windows.Forms.RadioButton
                        If TypeOf ctrl2 Is Windows.Forms.RadioButton Then
                            Radiob2 = ctrl2
                            If Radiob2.Checked = True Then
                                Dim Nume1 As String = Replace(Radiob2.Name, "RadioButton", "")
                                If IsNumeric(Nume1) = True Then
                                    Inaltimea1 = Inaltimea1 * CInt(Nume1)
                                    Text_height = Text_height * CInt(Nume1)
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                End If
            End If

            Try
                Using Lock1 As DocumentLock = ThisDrawing.LockDocument
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Creaza_layer(Bearing_and_distance_layer, 7, "b&D", True)
                        Creaza_layer(Arc_leader_polyline_layer, 7, "b&D", True)
                        Trans1.Commit()
                    End Using
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)

                        Dim Prompt_optionsPL As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select segment:")
                        Prompt_optionsPL.SetRejectMessage(vbLf & "You did not selected a polyline or line")
                        Prompt_optionsPL.AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Polyline), True)
                        Prompt_optionsPL.AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Line), True)
                        Dim Rezultat_PL As Autodesk.AutoCAD.EditorInput.PromptEntityResult = Editor1.GetEntity(Prompt_optionsPL)
                        If Not Rezultat_PL.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Editor1.SetImpliedSelection(Empty_array)
                            Editor1.WriteMessage(vbLf & "Command:")
                            Freeze_operations = False
                            Exit Sub
                        End If

                        Dim Curve1 As Curve = Trans1.GetObject(Rezultat_PL.ObjectId, OpenMode.ForRead)
                        Dim PickedPT As New Point3d

                        Dim Bear_rad As Double
                        Dim Dist2D As Double
                        Dim Min_defl As Double = 0.5 * PI / 180
                        Dim Start3D As New Point3d
                        Dim End3D As New Point3d

                        If TypeOf (Curve1) Is Line Then
                            Dim Line1 As Line = Curve1
                            Bear_rad = GET_Bearing_rad(Line1.StartPoint.X, Line1.StartPoint.Y, Line1.EndPoint.X, Line1.EndPoint.Y)
                            Dim Start2D As New Point3d(Line1.StartPoint.X, Line1.StartPoint.Y, 0)
                            Dim End2D As New Point3d(Line1.EndPoint.X, Line1.EndPoint.Y, 0)
                            Dist2D = Start2D.GetVectorTo(End2D).Length
                            PickedPT = Line1.GetClosestPointTo(Rezultat_PL.PickedPoint.TransformBy(Current_UCS_matrix), Vector3d.ZAxis, False)
                        End If

                        If TypeOf (Curve1) Is Polyline Then
                            Dim Poly1 As Polyline = Curve1

                            PickedPT = Poly1.GetClosestPointTo(Rezultat_PL.PickedPoint.TransformBy(Current_UCS_matrix), Vector3d.ZAxis, False)
                            Dim ParamPicked As Double = Poly1.GetParameterAtPoint(PickedPT)


                            If CheckBox_ignore_deflections_less_0_5.Checked = True Then
                                Dim Poly2 As Polyline = Bearing_dist_calc_with_min_deflection(Min_defl, Poly1, ParamPicked)
                                Start3D = Poly2.StartPoint
                                End3D = Poly2.EndPoint

                            End If

                            If CheckBox_ignore_deflections_less_0_5.Checked = False Then
                                Dim Poly2 As Polyline = Bearing_dist_calc_with_min_deflection(0, Poly1, ParamPicked)
                                Start3D = Poly2.StartPoint
                                End3D = Poly2.EndPoint
                            End If
                            Bear_rad = GET_Bearing_rad(Start3D.X, Start3D.Y, End3D.X, End3D.Y)

                            Dim Start2D As New Point3d(Start3D.X, Start3D.Y, 0)
                            Dim End2D As New Point3d(End3D.X, End3D.Y, 0)
                            Dist2D = Abs(Poly1.GetDistAtPoint(Start3D) - Poly1.GetDistAtPoint(End3D))
                        End If

                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap)

                        Dim Jig1 As New Draw_JIG1
                        Dim Point2 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                        Point2 = Jig1.StartJig(PickedPT, Inaltimea1)

                        If Point2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
                            Freeze_operations = False
                            Exit Sub
                        End If

                        Dim Jig2 As New Draw_JIG2
                        Dim Point3 As Autodesk.AutoCAD.EditorInput.PromptPointResult

                        Point3 = Jig2.StartJig(PickedPT, Point2.Value, Inaltimea1)
                        If Point3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
                            Freeze_operations = False
                            Exit Sub
                        End If

                        Dim x0 As Double = PickedPT.X
                        Dim y0 As Double = PickedPT.Y
                        Dim x2 As Double = Point2.Value.X
                        Dim y2 As Double = Point2.Value.Y
                        Dim x3 As Double = Point3.Value.X
                        Dim y3 As Double = Point3.Value.Y
                        Dim Wdth1 As Double = (Inaltimea1) / 1250

                        Dim x1, y1 As Double
                        x1 = x1_for_arc_leader(x0, y0, x2, y2, Inaltimea1)
                        y1 = y1_for_arc_leader(x0, y0, x2, y2, Inaltimea1)

                        Dim Bulge1 As Double
                        Bulge1 = Bulge_for_arc_leader(x0, y0, x2, y2, x3, y3, Inaltimea1)

                        Dim Poly_arrow As New Autodesk.AutoCAD.DatabaseServices.Polyline
                        Poly_arrow.AddVertexAt(0, New Point2d(x0, y0), 0, 0, Wdth1)
                        Poly_arrow.AddVertexAt(1, New Point2d(x1, y1), Bulge1, 0, 0)
                        Poly_arrow.AddVertexAt(2, New Point2d(x3, y3), 0, 0, 0)
                        Poly_arrow.Layer = Arc_leader_polyline_layer
                        BTrecord.AppendEntity(Poly_arrow)
                        Trans1.AddNewlyCreatedDBObject(Poly_arrow, True)
                        Trans1.TransactionManager.QueueForGraphicsFlush()

                        Dim Continut As String = "{\Fromans|c0;\Q20;\L" & Quadrant_bearings(Bear_rad) & " - " & Round(Dist2D, 0).ToString & "'±}"
                        Dim Point4 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                        Dim Jig3 As New jig_Mtext_class(New MText, Text_height, Text_rotation, Continut)
                        Point4 = Jig3.BeginJig

                        If Not Point4.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Editor1.SetImpliedSelection(Empty_array)
                            Editor1.WriteMessage(vbLf & "Command:")
                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
                            Freeze_operations = False
                            Exit Sub
                        End If

                        Dim mText123 As New MText
                        mText123.Contents = Continut
                        mText123.Location = Point4.Value
                        mText123.Rotation = Text_rotation
                        mText123.Attachment = AttachmentPoint.MiddleCenter
                        mText123.TextHeight = Text_height
                        mText123.Layer = Bearing_and_distance_layer
                        mText123.BackgroundFill = True
                        mText123.UseBackgroundColor = True
                        mText123.BackgroundScaleFactor = 1.2
                        BTrecord.AppendEntity(mText123)
                        Trans1.AddNewlyCreatedDBObject(mText123, True)
                        Trans1.Commit()
                    End Using
                End Using
            Catch ex As Exception
                MsgBox(ex.Message)
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
            End Try
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_arc_leader_Click(sender As Object, e As EventArgs) Handles Button_arc_leader.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Dim OLD_OSnap As Integer = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE")

            Dim NEW_OSnap As Integer = Autodesk.AutoCAD.EditorInput.ObjectSnapMasks.Near

            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap)

            Try
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Database1 As Autodesk.AutoCAD.DatabaseServices.Database

                ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                Using lock1 As DocumentLock = ThisDrawing.LockDocument
                    Database1 = ThisDrawing.Database
                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor

                    Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    Dim Inaltimea1 As Double = 1000 * 0.0625

                    Dim Tilemode1 As Integer = Application.GetSystemVariable("TILEMODE")
                    Dim CVport1 As Integer = Application.GetSystemVariable("CVPORT")


                    If Viewport_loaded = True And Tilemode1 = 1 Then
                        Inaltimea1 = Inaltimea1 * Scale_factor
                    Else
                        If (Tilemode1 = 0 And Not CVport1 = 1) Or Tilemode1 = 1 Then
                            For Each ctrl2 As Windows.Forms.Control In Panel_SCALE_SELECTION.Controls
                                Dim Radiob2 As Windows.Forms.RadioButton
                                If TypeOf ctrl2 Is Windows.Forms.RadioButton Then
                                    Radiob2 = ctrl2
                                    If Radiob2.Checked = True Then
                                        Dim Nume1 As String = Replace(Radiob2.Name, "RadioButton", "")
                                        If IsNumeric(Nume1) = True Then
                                            Inaltimea1 = Inaltimea1 * CInt(Nume1)
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If

                    Dim Jig1 As New Draw_JIG1
                    Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions("Pick first point : ")
                    Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult

                    PP1.AllowNone = False
                    Point1 = Editor1.GetPoint(PP1)

                    If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                        Freeze_operations = False
                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
                        Exit Sub
                    End If

                    Dim Point2 As Autodesk.AutoCAD.EditorInput.PromptPointResult

                    Point2 = Jig1.StartJig(Point1.Value, Inaltimea1)

                    If Point2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                        Freeze_operations = False
                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
                        Exit Sub
                    End If

                    Dim Jig2 As New Draw_JIG2
                    Dim Point3 As Autodesk.AutoCAD.EditorInput.PromptPointResult

                    Point3 = Jig2.StartJig(Point1.Value, Point2.Value, Inaltimea1)
                    If Point3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                        Freeze_operations = False
                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
                        Exit Sub
                    End If

                    Dim x0 As Double = Point1.Value.X
                    Dim y0 As Double = Point1.Value.Y
                    Dim x2 As Double = Point2.Value.X
                    Dim y2 As Double = Point2.Value.Y
                    Dim x3 As Double = Point3.Value.X
                    Dim y3 As Double = Point3.Value.Y
                    Dim Wdth1 As Double = (Inaltimea1) / 1250

                    Dim x1, y1 As Double
                    x1 = x1_for_arc_leader(x0, y0, x2, y2, Inaltimea1)
                    y1 = y1_for_arc_leader(x0, y0, x2, y2, Inaltimea1)

                    Dim Bulge1 As Double
                    Bulge1 = Bulge_for_arc_leader(x0, y0, x2, y2, x3, y3, Inaltimea1)
                    Using Trans2 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Creaza_layer(Arc_leader_polyline_layer, 7, "AL", True)
                    End Using


                    Using Trans2 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans2.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                        Dim Poly1 As New Autodesk.AutoCAD.DatabaseServices.Polyline
                        Poly1.AddVertexAt(0, New Point2d(x0, y0), 0, 0, Wdth1)
                        Poly1.AddVertexAt(1, New Point2d(x1, y1), Bulge1, 0, 0)
                        Poly1.AddVertexAt(2, New Point2d(x3, y3), 0, 0, 0)
                        Poly1.Layer = Arc_leader_polyline_layer
                        BTrecord.AppendEntity(Poly1)
                        Trans2.AddNewlyCreatedDBObject(Poly1, True)
                        Trans2.Commit()
                    End Using
                End Using
            Catch ex As Exception
                Freeze_operations = False
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
                MsgBox(ex.Message)
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_load_viewport_Click(sender As Object, e As EventArgs) Handles Button_load_viewport.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            If Label_viewport_loaded.Text = "Viewport Not Loaded" Then

                Dim Empty_array() As ObjectId
                Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
                Dim Current_UCS_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
                Dim Tilemode1 As Integer = Application.GetSystemVariable("TILEMODE")
                Dim CVport1 As Integer = Application.GetSystemVariable("CVPORT")
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                Try
                    Using Lock1 As DocumentLock = ThisDrawing.LockDocument
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                            Dim Prompt_options_Viewport As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select viewport (for clipped viewports select the clipping polyline) :")
                            Prompt_options_Viewport.SetRejectMessage(vbLf & "You did not selected a Viewport")
                            Prompt_options_Viewport.AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Viewport), True)
                            Prompt_options_Viewport.AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Polyline), True)

                            Dim Rezultat_Viewport As Autodesk.AutoCAD.EditorInput.PromptEntityResult = Editor1.GetEntity(Prompt_options_Viewport)
                            If Not Rezultat_Viewport.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Editor1.SetImpliedSelection(Empty_array)
                                Editor1.WriteMessage(vbLf & "Command:")
                                Freeze_operations = False
                                Exit Sub
                            End If

                            Dim vpId As ObjectId
                            If TypeOf Trans1.GetObject(Rezultat_Viewport.ObjectId, OpenMode.ForRead) Is Polyline Then
                                vpId = LayoutManager.Current.GetNonRectangularViewportIdFromClipId(Rezultat_Viewport.ObjectId)
                                If vpId = Nothing Then
                                    Freeze_operations = False
                                    Exit Sub
                                End If
                            Else
                                vpId = Rezultat_Viewport.ObjectId
                            End If

                            If IsNothing(vpId) = False Then
                                Dim Viewport1 As Viewport = Trans1.GetObject(vpId, OpenMode.ForRead)
                                If IsNothing(Viewport1) = False Then
                                    Scale_factor = 1 / Viewport1.CustomScale
                                    Dim Twist1 As Double = Viewport1.TwistAngle

                                    ThisDrawing.Editor.SwitchToModelSpace()
                                    Application.SetSystemVariable("CVPORT", Viewport1.Number)

                                    Dim view1 As ViewTableRecord = ThisDrawing.Editor.GetCurrentView
                                    Dim Planul_curent As Plane = New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)
                                    Dim Xaxis1 As Vector3d = Current_UCS_matrix.CoordinateSystem3d.Xaxis
                                    Dim Angle_Xaxis As Double = (Xaxis1.AngleOnPlane(Planul_curent))

                                    View_rotation = (2 * PI - view1.ViewTwist) - Angle_Xaxis
                                    ThisDrawing.Editor.SwitchToPaperSpace()

                                End If

                                Trans1.Commit()
                                Editor1.SetImpliedSelection(Empty_array)
                                Editor1.WriteMessage(vbLf & "Command:")
                                Label_viewport_loaded.Text = "Viewport Loaded"
                                Button_load_viewport.Text = "Unload Viewport"
                                Viewport_loaded = True

                            End If
                        End Using
                    End Using
                Catch ex As Exception
                    Freeze_operations = False
                    Editor1.SetImpliedSelection(Empty_array)
                    Editor1.WriteMessage(vbLf & "Command:")
                    MsgBox(ex.Message)
                End Try
            Else
                Label_viewport_loaded.Text = "Viewport Not Loaded"
                Button_load_viewport.Text = "Load Viewport"
                Scale_factor = 0
                View_rotation = 0
                Viewport_loaded = False
            End If
            Freeze_operations = False
        End If

    End Sub

    Private Sub CheckBox_add_locus_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_add_locus.CheckedChanged
        If Freeze_operations = False Then
            Freeze_operations = True
            If CheckBox_add_locus.Checked = True Then
                Panel_locus.Visible = True
                If ComboBox_SHEET_SET_scale_locus.Items.Count > 0 Then
                    ComboBox_SHEET_SET_scale_locus.Items.Clear()
                End If
                ComboBox_SHEET_SET_scale_locus.Items.Add("SCALE_LOCUS")
                ComboBox_SHEET_SET_scale_locus.SelectedIndex = 0
            Else
                Panel_locus.Visible = False
            End If
            Freeze_operations = False
        End If
    End Sub

    Private Sub RadioButtonL_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButtonL3000.CheckedChanged, _
                                                                                        RadioButtonL6000.CheckedChanged, _
                                                                                        RadioButtonL600.CheckedChanged, _
                                                                                        RadioButtonL5000.CheckedChanged, _
                                                                                        RadioButtonL500.CheckedChanged, _
                                                                                        RadioButtonL4000.CheckedChanged, _
                                                                                        RadioButtonL400.CheckedChanged, _
                                                                                        RadioButtonL300.CheckedChanged, _
                                                                                        RadioButtonL2000.CheckedChanged, _
                                                                                        RadioButtonL200.CheckedChanged, _
                                                                                        RadioButtonL10000.CheckedChanged, _
                                                                                        RadioButtonL1000.CheckedChanged, _
                                                                                        RadioButtonL100.CheckedChanged



        Dim Name1 As String = sender.name

        If sender.checked = True Then
            If Freeze_operations = False Then
                Freeze_operations = True
                If IsNumeric(Replace(Name1, "RadioButtonL", "")) = True Then
                    Dim Scale1 As Double = CDbl(Replace(Name1, "RadioButtonL", ""))
                    Dim Empty_array() As ObjectId
                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    ThisDrawing.Editor.SetImpliedSelection(Empty_array)

                    Try
                        Using Lock1 As DocumentLock = ThisDrawing.LockDocument
                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                                Dim Prompt_options_vW As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select Viewport:")
                                Prompt_options_vW.SetRejectMessage(vbLf & "You did not selected a Viewport")
                                Prompt_options_vW.AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Viewport), True)

                                Dim Rezultat_Viewport As Autodesk.AutoCAD.EditorInput.PromptEntityResult = ThisDrawing.Editor.GetEntity(Prompt_options_vW)
                                If Not Rezultat_Viewport.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                    ThisDrawing.Editor.SetImpliedSelection(Empty_array)
                                    ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                                    Freeze_operations = False
                                    Exit Sub
                                End If

                                Dim Viewport1 As Viewport = Trans1.GetObject(Rezultat_Viewport.ObjectId, OpenMode.ForWrite)
                                Viewport1.Locked = False
                                Viewport1.CustomScale = 1 / Scale1
                                Viewport1.Locked = True
                                Trans1.Commit()
                            End Using
                        End Using
                    Catch ex As Exception
                        MsgBox(ex.Message)

                    End Try

                    ThisDrawing.Editor.SetImpliedSelection(Empty_array)

                    Dim SheetSet_manager As New AcSmSheetSetMgr
                    Dim SheetSet_database As AcSmDatabase = SheetSet_manager.OpenDatabase(TextBox_sheet_set_template.Text, False)
                    Dim sheetSet As AcSmSheetSet
                    sheetSet = SheetSet_database.GetSheetSet()
                    If LockDatabase(SheetSet_database, True) = True Then
                        Dim EnumSheets As IAcSmEnumComponent = sheetSet.GetSheetEnumerator()
                        Dim smComponent As IAcSmComponent
                        Dim sheet1 As IAcSmSheet
                        smComponent = EnumSheets.Next()

                        Dim Sheet_name As String = System.IO.Path.GetFileNameWithoutExtension(ThisDrawing.Name)

                        If IsNothing(Sheet_name) = False Then
                            While True
                                If smComponent Is Nothing Then
                                    Exit While
                                End If

                                sheet1 = TryCast(smComponent, IAcSmSheet)

                                If sheet1.GetTitle = Sheet_name Then
                                    Dim customPropertyBag As AcSmCustomPropertyBag = sheet1.GetCustomPropertyBag()
                                    Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()

                                    customPropertyValue.InitNew(sheet1)
                                    ' Set the flag for the property
                                    customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                    ' Set the value for the property
                                    customPropertyValue.SetValue("1" & Chr(34) & "=" & Round(Scale1, 0).ToString & "'")

                                    customPropertyBag.SetProperty(ComboBox_SHEET_SET_scale_locus.Text, customPropertyValue)
                                    Exit While
                                End If

                                smComponent = EnumSheets.Next()
                            End While
                        End If


                        LockDatabase(SheetSet_database, False)



                        ThisDrawing.Editor.Regen()

                    End If


                End If
                Freeze_operations = False
            End If
        End If


    End Sub

    Private Sub Panel_Owner_LineLIST_Click(sender As Object, e As EventArgs)
        If Freeze_operations = False Then
            Freeze_operations = True

            Incarca_existing_Blocks_with_attributes_to_combobox(ComboBox_blocks)
            If ComboBox_blocks.Items.Count > 0 Then
                If ComboBox_blocks.Items.Contains("Owner_Linelist") = True Then
                    ComboBox_blocks.SelectedIndex = ComboBox_blocks.Items.IndexOf("Owner_Linelist")
                Else
                    ComboBox_blocks.SelectedIndex = 0
                End If
            End If

            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_add_owner_linelist_Click(sender As Object, e As EventArgs) Handles Button_add_owner_linelist.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Dim Empty_array() As ObjectId
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Try
                Using Lock1 As DocumentLock = ThisDrawing.LockDocument
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Creaza_layer(Layer_name_no_plot1, 40, "No Plot", False)
                        Trans1.Commit()
                    End Using
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                        Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                        Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables

                        For Each Id1 As ObjectId In BTrecord
                            Dim Ent1 As Entity
                            Ent1 = TryCast(Trans1.GetObject(Id1, OpenMode.ForRead), Entity)
                            If IsNothing(Ent1) = False Then
                                If TypeOf Ent1 Is Polyline Then
                                    Dim Poly1 As Polyline = Ent1
                                    If Poly1.Closed = True Then
                                        Dim DBobj_col1 As New DBObjectCollection()
                                        DBobj_col1.Add(Poly1)
                                        Dim Region1 As Autodesk.AutoCAD.DatabaseServices.Region


                                        Try
                                            Region1 = TryCast(Autodesk.AutoCAD.DatabaseServices.Region.CreateFromCurves(DBobj_col1)(0), Autodesk.AutoCAD.DatabaseServices.Region)
                                        Catch ex As Exception

                                        End Try


                                        Dim InsPoint As New Point3d

                                        Try
                                            Dim Centru As New Point2d
                                            Centru = Region1.AreaProperties(Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis).Centroid
                                            InsPoint = New Point3d(Centru.X, Centru.Y, 0)
                                        Catch ex As Exception
                                            InsPoint = Poly1.GetPointAtParameter(1)
                                        End Try



                                        Using Records1 As Autodesk.Gis.Map.ObjectData.Records = Tables1.GetObjectRecords(Convert.ToInt32(0), Poly1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                            If IsNothing(Records1) = False Then
                                                If Records1.Count > 0 Then

                                                    Dim Record1 As Autodesk.Gis.Map.ObjectData.Record

                                                    Dim Value1 As String = ""
                                                    Dim Value2 As String = ""
                                                    Dim Value3 As String = ""
                                                    Dim Value4 As String = ""
                                                    Dim Value5 As String = ""
                                                    Dim INSEREAZA As Boolean = False

                                                    For Each Record1 In Records1
                                                        Dim Tabla1 As Autodesk.Gis.Map.ObjectData.Table
                                                        Tabla1 = Tables1(Record1.TableName)
                                                        If Tabla1.Name.ToUpper = ComboBox_OD_Table_name.Text.ToUpper Then

                                                            INSEREAZA = True


                                                            Dim Field_defs1 As Autodesk.Gis.Map.ObjectData.FieldDefinitions
                                                            Field_defs1 = Tabla1.FieldDefinitions

                                                            For i = 0 To Record1.Count - 1
                                                                Dim Field_def1 As Autodesk.Gis.Map.ObjectData.FieldDefinition
                                                                Field_def1 = Field_defs1(i)

                                                                If Not ComboBox_OD1.Text.ToUpper = "" Then
                                                                    If Field_def1.Name.ToUpper.Contains(ComboBox_OD1.Text.ToUpper) = True Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(i)
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Value1 = Valoare_record1.StrValue
                                                                            If CheckBox_UPPER_CASE.Checked = True Then
                                                                                Value1 = Value1.ToUpper
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If


                                                                If Not ComboBox_OD2.Text.ToUpper = "" Then
                                                                    If Field_def1.Name.ToUpper.Contains(ComboBox_OD2.Text.ToUpper) = True Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(i)
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Value2 = Valoare_record1.StrValue
                                                                            If CheckBox_UPPER_CASE.Checked = True Then
                                                                                Value2 = Value2.ToUpper
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If

                                                                If Not ComboBox_OD3.Text.ToUpper = "" Then
                                                                    If Field_def1.Name.ToUpper.Contains(ComboBox_OD3.Text.ToUpper) = True Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(i)
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Value3 = Valoare_record1.StrValue
                                                                            If CheckBox_UPPER_CASE.Checked = True Then
                                                                                Value3 = Value3.ToUpper
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If

                                                                If Not ComboBox_OD4.Text.ToUpper = "" Then
                                                                    If Field_def1.Name.ToUpper.Contains(ComboBox_OD4.Text.ToUpper) = True Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(i)
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Value4 = Valoare_record1.StrValue
                                                                            If CheckBox_UPPER_CASE.Checked = True Then
                                                                                Value4 = Value4.ToUpper
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If

                                                                If Not ComboBox_OD5.Text.ToUpper = "" Then
                                                                    If Field_def1.Name.ToUpper.Contains(ComboBox_OD5.Text.ToUpper) = True Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(i)
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Value5 = Valoare_record1.StrValue
                                                                            If CheckBox_UPPER_CASE.Checked = True Then
                                                                                Value5 = Value5.ToUpper
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                            Next

                                                        End If

                                                    Next

                                                    Dim Colectie_nume As New Specialized.StringCollection
                                                    Dim Colectie_valori As New Specialized.StringCollection

                                                    If Not Value1 = "" Then
                                                        If Not ComboBox_bl_atr1.Text = "" Then
                                                            Colectie_nume.Add(ComboBox_bl_atr1.Text)
                                                            Colectie_valori.Add(Value1)
                                                        End If
                                                    End If

                                                    Dim Owner1 As String = ""
                                                    If Not Value2 = "" Then
                                                        Owner1 = Value2
                                                    End If

                                                    If Not Value3 = "" Then
                                                        Dim Extra1 As String = TextBox_CONCAT1.Text
                                                        If Extra1.ToUpper.Contains("COMMA") = True Then
                                                            Extra1 = Replace(Extra1.ToUpper, "COMMA", ",")
                                                        End If
                                                        If Extra1.ToUpper.Contains("SPACE") = True Then
                                                            Extra1 = Replace(Extra1.ToUpper, "SPACE", " ")
                                                        End If
                                                        Owner1 = Owner1 & Extra1 & Value3
                                                    End If

                                                    If Not Value4 = "" Then
                                                        Dim Extra1 As String = TextBox_CONCAT2.Text
                                                        If Extra1.ToUpper.Contains("COMMA") = True Then
                                                            Extra1 = Replace(Extra1.ToUpper, "COMMA", ",")
                                                        End If
                                                        If Extra1.ToUpper.Contains("SPACE") = True Then
                                                            Extra1 = Replace(Extra1.ToUpper, "SPACE", " ")
                                                        End If
                                                        Owner1 = Owner1 & Extra1 & Value4
                                                    End If

                                                    If Not Value5 = "" Then
                                                        Dim Extra1 As String = TextBox_CONCAT3.Text
                                                        If Extra1.ToUpper.Contains("COMMA") = True Then
                                                            Extra1 = Replace(Extra1.ToUpper, "COMMA", ",")
                                                        End If
                                                        If Extra1.ToUpper.Contains("SPACE") = True Then
                                                            Extra1 = Replace(Extra1.ToUpper, "SPACE", " ")
                                                        End If
                                                        Owner1 = Owner1 & Extra1 & Value5
                                                    End If

                                                    If Not Owner1 = "" Then
                                                        If Not ComboBox_bl_atr2.Text = "" Then
                                                            Colectie_nume.Add(ComboBox_bl_atr2.Text)
                                                            Colectie_valori.Add(Owner1)
                                                        End If
                                                    End If



                                                    If INSEREAZA = True Then
                                                        InsertBlock_with_multiple_atributes("", ComboBox_blocks.Text, InsPoint, 1, BTrecord, Layer_name_no_plot1, Colectie_nume, Colectie_valori)
                                                    End If


                                                End If
                                            End If
                                        End Using
                                    End If
                                End If
                            End If
                        Next








                        Trans1.Commit()
                    End Using
                End Using
            Catch ex As Exception
                MsgBox(ex.Message)

            End Try
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_rotate_block_to_viewport_Click(sender As Object, e As EventArgs) Handles Button_rotate_block_to_viewport.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Dim Empty_array() As ObjectId
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            If Viewport_loaded = True Then
                Try
                    Dim Nr_blocks As Double = 0
                    Using Lock1 As DocumentLock = ThisDrawing.LockDocument

                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(BlockTable1(BlockTableRecord.ModelSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                            If Not ComboBox_blocks.Text = "" Then

                                For Each Id1 As ObjectId In BTrecord
                                    Dim Ent1 As Entity
                                    Ent1 = TryCast(Trans1.GetObject(Id1, OpenMode.ForRead), Entity)
                                    If IsNothing(Ent1) = False Then
                                        If TypeOf Ent1 Is BlockReference Then
                                            Dim block1 As BlockReference = Ent1
                                            If block1.Name = ComboBox_blocks.Text Then
                                                block1.UpgradeOpen()
                                                block1.Rotation = View_rotation
                                                Nr_blocks = Nr_blocks + 1
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                            Trans1.Commit()
                        End Using
                    End Using
                    MsgBox(Nr_blocks.ToString & " Blocks are rotated to viewport" & vbCrLf & _
                           "Please run ATTSYNC on the block")

                Catch ex As Exception
                    MsgBox(ex.Message)
                    Freeze_operations = False
                End Try
                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
            End If

            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_load_object_data_from_parcels_update_Click(sender As Object, e As EventArgs) Handles Button_load_object_data_from_parcels_update.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Dim BaseMap_drawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = BaseMap_drawing.Editor
            Try
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()



                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptEntityResult

                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select entity containing object data:")

                Rezultat1 = Editor1.GetEntity(Object_Prompt)

                If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.WriteMessage(vbLf & "Command:")
                    Freeze_operations = False
                    Exit Sub
                End If

                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                    If IsNothing(Rezultat1) = False Then

                        Using Lock_dwg As DocumentLock = BaseMap_drawing.LockDocument
                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = BaseMap_drawing.TransactionManager.StartTransaction
                                Dim Ent1 As Entity = Trans1.GetObject(Rezultat1.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                                Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables
                                Dim Id1 As ObjectId = Ent1.ObjectId

                                Using Records1 As Autodesk.Gis.Map.ObjectData.Records = Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                    If IsNothing(Records1) = False Then
                                        If Records1.Count > 0 Then
                                            With ComboBox_state_update
                                                .Items.Clear()
                                                .Text = ""
                                            End With
                                            With ComboBox_county_update
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_town_update
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_linelist_update
                                                .Items.Clear()
                                                .Text = ""
                                            End With
                                            With ComboBox_HMM_ID_update
                                                .Items.Clear()
                                                .Text = ""
                                            End With
                                            With ComboBox_MBL_update
                                                .Items.Clear()
                                                .Text = ""
                                            End With


                                            With ComboBox_deed_page_update
                                                .Items.Clear()
                                                .Text = ""
                                            End With


                                            With ComboBox_APN_update
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_access_road_length_update
                                                .Items.Clear()
                                                .Text = ""
                                            End With



                                            With ComboBox_Section_TWP_Range_update
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_owner_update
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_crossing_length_update
                                                .Items.Clear()
                                                .Text = ""
                                            End With


                                            With ComboBox_area_EX_E_update
                                                .Items.Clear()
                                                .Text = ""
                                            End With
                                            With ComboBox_area_P_E_update
                                                .Items.Clear()
                                                .Text = ""
                                            End With
                                            With ComboBox_area_TWS_update
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_area_ATWS_update
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_area_A_R_update
                                                .Items.Clear()
                                                .Text = ""
                                            End With
                                            With ComboBox_area_WARE_update
                                                .Items.Clear()
                                                .Text = ""
                                            End With
                                            With ComboBox_area_TWS_ABD_update
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_area_TWS_PD_update
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_shape_user1
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_shape_user2
                                                .Items.Clear()
                                                .Text = ""
                                            End With


                                            With ComboBox_shape_user3
                                                .Items.Clear()
                                                .Text = ""
                                            End With


                                            With ComboBox_shape_user4
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_shape_user5
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_shape_user6
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_shape_user7
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_shape_user8
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_shape_user9
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_shape_user10
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_shape_user11
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_shape_user12
                                                .Items.Clear()
                                                .Text = ""
                                            End With

                                            With ComboBox_shape_user13
                                                .Items.Clear()
                                                .Text = ""
                                            End With


                                            Dim Record1 As Autodesk.Gis.Map.ObjectData.Record


                                            For Each Record1 In Records1
                                                Dim Tabla1 As Autodesk.Gis.Map.ObjectData.Table
                                                Tabla1 = Tables1(Record1.TableName)


                                                Dim Field_defs1 As Autodesk.Gis.Map.ObjectData.FieldDefinitions
                                                Field_defs1 = Tabla1.FieldDefinitions



                                                For i = 0 To Record1.Count - 1
                                                    Dim Field_def1 As Autodesk.Gis.Map.ObjectData.FieldDefinition
                                                    Field_def1 = Field_defs1(i)
                                                    With ComboBox_state_update
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With
                                                    With ComboBox_county_update
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With
                                                    With ComboBox_town_update
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With


                                                    With ComboBox_linelist_update
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With
                                                    With ComboBox_HMM_ID_update
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With
                                                    With ComboBox_MBL_update
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With

                                                    With ComboBox_deed_page_update
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With
                                                    With ComboBox_APN_update
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With

                                                    With ComboBox_access_road_length_update
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With

                                                    With ComboBox_Section_TWP_Range_update
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With

                                                    With ComboBox_owner_update
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With
                                                    With ComboBox_crossing_length_update
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With



                                                    With ComboBox_area_EX_E_update
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With
                                                    With ComboBox_area_P_E_update
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With
                                                    With ComboBox_area_TWS_update
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With
                                                    With ComboBox_area_ATWS_update
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With

                                                    With ComboBox_area_A_R_update
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With
                                                    With ComboBox_area_WARE_update
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With
                                                    With ComboBox_area_TWS_ABD_update
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With
                                                    With ComboBox_area_TWS_PD_update
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With

                                                    With ComboBox_shape_user1
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With

                                                    With ComboBox_shape_user2
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With


                                                    With ComboBox_shape_user3
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With


                                                    With ComboBox_shape_user4
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With

                                                    With ComboBox_shape_user5
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With

                                                    With ComboBox_shape_user6
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With

                                                    With ComboBox_shape_user7
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With

                                                    With ComboBox_shape_user8
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With

                                                    With ComboBox_shape_user9
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With

                                                    With ComboBox_shape_user10
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With

                                                    With ComboBox_shape_user11
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With

                                                    With ComboBox_shape_user12
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With

                                                    With ComboBox_shape_user13
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With


                                                Next



                                            Next


                                            With ComboBox_state_update
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("PHYSSTATE") = True Then
                                                        .SelectedIndex = .Items.IndexOf("PHYSSTATE")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_county_update
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("PHYSCOUNTY") = True Then
                                                        .SelectedIndex = .Items.IndexOf("PHYSCOUNTY")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_town_update
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("SGC_TOWN") = True Then
                                                        .SelectedIndex = .Items.IndexOf("SGC_TOWN")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_owner_update
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("SGC_OWNER") = True Then
                                                        .SelectedIndex = .Items.IndexOf("SGC_OWNER")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With


                                            With ComboBox_linelist_update
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("LINELIST") = True Then
                                                        .SelectedIndex = .Items.IndexOf("LINELIST")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_MBL_update
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("SGC_MBL") = True Then
                                                        .SelectedIndex = .Items.IndexOf("SGC_MBL")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_area_ATWS_update
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("ATWSSQFT") = True Then
                                                        .SelectedIndex = .Items.IndexOf("ATWSSQFT")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_area_EX_E_update
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("E_EASE_SF") = True Then
                                                        .SelectedIndex = .Items.IndexOf("E_EASE_SF")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_area_P_E_update
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("PERMSQFT") = True Then
                                                        .SelectedIndex = .Items.IndexOf("PERMSQFT")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_area_TWS_update
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("TWSSQFT") = True Then
                                                        .SelectedIndex = .Items.IndexOf("TWSSQFT")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_area_A_R_update
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("ACSSQFT") = True Then
                                                        .SelectedIndex = .Items.IndexOf("ACSSQFT")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With


                                            With ComboBox_area_WARE_update
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("WARESQFT") = True Then
                                                        .SelectedIndex = .Items.IndexOf("WARESQFT")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_HMM_ID_update
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("PARCEL_HMM") = True Then
                                                        .SelectedIndex = .Items.IndexOf("PARCEL_HMM")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With









                                            With ComboBox_Section_TWP_Range
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("PHYS_S_T_R") = True Then
                                                        .SelectedIndex = .Items.IndexOf("PHYS_S_T_R")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_area_TWS_ABD
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("TWS_AB_SF") = True Then
                                                        .SelectedIndex = .Items.IndexOf("TWS_AB_SF")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_area_TWS_PD
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("TWS_PD_SF") = True Then
                                                        .SelectedIndex = .Items.IndexOf("TWS_PD_SF")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With



                                            With ComboBox_deed_page_update
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("DEED_BK_PK") = True Then
                                                        .SelectedIndex = .Items.IndexOf("DEED_BK_PK")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                            With ComboBox_APN_update
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("APN") = True Then
                                                        .SelectedIndex = .Items.IndexOf("APN")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With
                                            With ComboBox_crossing_length_update
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("CL_CROSSFT") = True Then
                                                        .SelectedIndex = .Items.IndexOf("CL_CROSSFT")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With
                                            With ComboBox_Section_TWP_Range_update
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("PHYS_S_T_R") = True Then
                                                        .SelectedIndex = .Items.IndexOf("PHYS_S_T_R")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With
                                            With ComboBox_area_TWS_ABD_update
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("TWS_AB_SF") = True Then
                                                        .SelectedIndex = .Items.IndexOf("TWS_AB_SF")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With
                                            With ComboBox_area_TWS_PD_update
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("TWS_PD_SF") = True Then
                                                        .SelectedIndex = .Items.IndexOf("TWS_PD_SF")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With
                                            With ComboBox_access_road_length_update
                                                If .Items.Count > 0 Then
                                                    If .Items.Contains("AR_LENGTH") = True Then
                                                        .SelectedIndex = .Items.IndexOf("AR_LENGTH")
                                                    Else
                                                        .SelectedIndex = .Items.IndexOf("")
                                                    End If
                                                End If
                                            End With

                                        End If
                                    End If
                                End Using






                            End Using
                        End Using

                    End If
                End If

                Dim Empty_array() As ObjectId
                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")

            Catch ex As Exception
                Editor1.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_read_from_excel_update_Click(sender As Object, e As EventArgs) Handles Button_read_from_excel_update.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                Button_read_from_excel_update.Visible = False
                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Colectie_ID = New Specialized.StringCollection
                Dim Start1 As Integer = 0
                Dim End1 As Integer = 0
                Dim Column_id As String = TextBox_Column_ID_Excel_update.Text.ToUpper
                If IsNumeric(TextBox_Start_row_excel_update.Text) = True Then
                    Start1 = CDbl(TextBox_Start_row_excel_update.Text)
                End If
                If IsNumeric(TextBox_End_row_excel_update.Text) = True Then
                    End1 = CDbl(TextBox_End_row_excel_update.Text)
                End If

                If Not Start1 = 0 Or Not End1 = 0 Then
                    If End1 > 0 And Start1 > 0 And End1 >= Start1 Then
                        For i = Start1 To End1
                            If Not Replace(W1.Range(Column_id & i).Value, " ", "") = "" Then
                                Colectie_ID.Add(W1.Range(Column_id & i).Value)
                            End If

                        Next
                    End If
                End If


            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Button_read_from_excel_update.Visible = True
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_update_Plats_Click(sender As Object, e As EventArgs) Handles Button_update_Plats.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            If IsNothing(Colectie_ID) = True Then
                MsgBox("Please read HMM ID excel file")
                Freeze_operations = False
                Exit Sub
            End If

            If Colectie_ID.Count = 0 Then
                MsgBox("Please read HMM ID excel file")
                Freeze_operations = False
                Exit Sub
            End If

            If TextBox_sheet_set_template.Text = "" Then
                MsgBox("Please specify the SHEET SET file")
                Freeze_operations = False
                Exit Sub
            End If

            If Not Strings.Right(TextBox_sheet_set_template.Text, 3).ToUpper = "DST" Then
                MsgBox("Please specify the SHEET SET file")
                Freeze_operations = False
                Exit Sub
            End If

            Dim Empty_array() As ObjectId
            Dim BaseMap_drawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = BaseMap_drawing.Editor
            Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Try


                Dim Data_table_with_values As New System.Data.DataTable
                Dim Index_data_table_valori As Double = 0

                If Not ComboBox_state_update.Text = "" Then
                    If Data_table_with_values.Columns.Contains(ComboBox_state_update.Text) = False Then
                        Data_table_with_values.Columns.Add(ComboBox_state_update.Text, GetType(String))
                    End If
                End If

                If Not ComboBox_county_update.Text = "" Then
                    If Data_table_with_values.Columns.Contains(ComboBox_county_update.Text) = False Then
                        Data_table_with_values.Columns.Add(ComboBox_county_update.Text, GetType(String))
                    End If
                End If

                If Not ComboBox_town_update.Text = "" Then
                    If Data_table_with_values.Columns.Contains(ComboBox_town_update.Text) = False Then
                        Data_table_with_values.Columns.Add(ComboBox_town_update.Text, GetType(String))
                    End If
                End If

                If Not ComboBox_owner_update.Text = "" Then
                    If Data_table_with_values.Columns.Contains(ComboBox_owner_update.Text) = False Then
                        Data_table_with_values.Columns.Add(ComboBox_owner_update.Text, GetType(String))
                    End If
                End If



                If Not ComboBox_linelist_update.Text = "" Then
                    If Data_table_with_values.Columns.Contains(ComboBox_linelist_update.Text) = False Then
                        Data_table_with_values.Columns.Add(ComboBox_linelist_update.Text, GetType(String))
                    End If
                End If

                If Not ComboBox_MBL_update.Text = "" Then
                    If Data_table_with_values.Columns.Contains(ComboBox_MBL_update.Text) = False Then
                        Data_table_with_values.Columns.Add(ComboBox_MBL_update.Text, GetType(String))
                    End If
                End If

                If Not ComboBox_area_ATWS_update.Text = "" Then
                    If Data_table_with_values.Columns.Contains(ComboBox_area_ATWS_update.Text) = False Then
                        Data_table_with_values.Columns.Add(ComboBox_area_ATWS_update.Text, GetType(Double))
                    End If
                End If

                If Not ComboBox_area_EX_E_update.Text = "" Then
                    If Data_table_with_values.Columns.Contains(ComboBox_area_EX_E_update.Text) = False Then
                        Data_table_with_values.Columns.Add(ComboBox_area_EX_E_update.Text, GetType(Double))
                    End If
                End If

                If Not ComboBox_area_P_E_update.Text = "" Then
                    If Data_table_with_values.Columns.Contains(ComboBox_area_P_E_update.Text) = False Then
                        Data_table_with_values.Columns.Add(ComboBox_area_P_E_update.Text, GetType(Double))
                    End If
                End If

                If Not ComboBox_area_TWS_update.Text = "" Then
                    If Data_table_with_values.Columns.Contains(ComboBox_area_TWS_update.Text) = False Then
                        Data_table_with_values.Columns.Add(ComboBox_area_TWS_update.Text, GetType(Double))
                    End If
                End If

                If Not ComboBox_area_A_R_update.Text = "" Then
                    If Data_table_with_values.Columns.Contains(ComboBox_area_A_R_update.Text) = False Then
                        Data_table_with_values.Columns.Add(ComboBox_area_A_R_update.Text, GetType(Double))
                    End If
                End If

                If Not ComboBox_area_WARE_update.Text = "" Then
                    If Data_table_with_values.Columns.Contains(ComboBox_area_WARE_update.Text) = False Then
                        Data_table_with_values.Columns.Add(ComboBox_area_WARE_update.Text, GetType(Double))
                    End If
                End If



                With ComboBox_deed_page_update
                    If Not .Text = "" Then
                        If Data_table_with_values.Columns.Contains(.Text) = False Then
                            Data_table_with_values.Columns.Add(.Text, GetType(String))
                        End If
                    End If
                End With
                With ComboBox_APN_update
                    If Not .Text = "" Then
                        If Data_table_with_values.Columns.Contains(.Text) = False Then
                            Data_table_with_values.Columns.Add(.Text, GetType(String))
                        End If
                    End If
                End With
                With ComboBox_crossing_length_update
                    If Not .Text = "" Then
                        If Data_table_with_values.Columns.Contains(.Text) = False Then
                            Data_table_with_values.Columns.Add(.Text, GetType(Double))
                        End If
                    End If
                End With

                With ComboBox_Section_TWP_Range_update
                    If Not .Text = "" Then
                        If Data_table_with_values.Columns.Contains(.Text) = False Then
                            Data_table_with_values.Columns.Add(.Text, GetType(String))
                        End If
                    End If
                End With

                With ComboBox_area_TWS_ABD_update
                    If Not .Text = "" Then
                        If Data_table_with_values.Columns.Contains(.Text) = False Then
                            Data_table_with_values.Columns.Add(.Text, GetType(Double))
                        End If
                    End If
                End With

                With ComboBox_area_TWS_PD_update
                    If Not .Text = "" Then
                        If Data_table_with_values.Columns.Contains(.Text) = False Then
                            Data_table_with_values.Columns.Add(.Text, GetType(Double))
                        End If
                    End If
                End With

                With ComboBox_access_road_length_update
                    If Not .Text = "" Then
                        If Data_table_with_values.Columns.Contains(.Text) = False Then
                            Data_table_with_values.Columns.Add(.Text, GetType(Double))
                        End If
                    End If
                End With

                With ComboBox_shape_user1
                    If Not .Text = "" Then
                        If Data_table_with_values.Columns.Contains(.Text) = False Then
                            Data_table_with_values.Columns.Add(.Text, GetType(String))
                        End If
                    End If
                End With

                With ComboBox_shape_user2
                    If Not .Text = "" Then
                        If Data_table_with_values.Columns.Contains(.Text) = False Then
                            Data_table_with_values.Columns.Add(.Text, GetType(String))
                        End If
                    End If
                End With

                With ComboBox_shape_user3
                    If Not .Text = "" Then
                        If Data_table_with_values.Columns.Contains(.Text) = False Then
                            Data_table_with_values.Columns.Add(.Text, GetType(String))
                        End If
                    End If
                End With
                With ComboBox_shape_user4
                    If Not .Text = "" Then
                        If Data_table_with_values.Columns.Contains(.Text) = False Then
                            Data_table_with_values.Columns.Add(.Text, GetType(String))
                        End If
                    End If
                End With
                With ComboBox_shape_user5
                    If Not .Text = "" Then
                        If Data_table_with_values.Columns.Contains(.Text) = False Then
                            Data_table_with_values.Columns.Add(.Text, GetType(String))
                        End If
                    End If
                End With
                With ComboBox_shape_user6
                    If Not .Text = "" Then
                        If Data_table_with_values.Columns.Contains(.Text) = False Then
                            Data_table_with_values.Columns.Add(.Text, GetType(String))
                        End If
                    End If
                End With
                With ComboBox_shape_user7
                    If Not .Text = "" Then
                        If Data_table_with_values.Columns.Contains(.Text) = False Then
                            Data_table_with_values.Columns.Add(.Text, GetType(String))
                        End If
                    End If
                End With
                With ComboBox_shape_user8
                    If Not .Text = "" Then
                        If Data_table_with_values.Columns.Contains(.Text) = False Then
                            Data_table_with_values.Columns.Add(.Text, GetType(String))
                        End If
                    End If
                End With
                With ComboBox_shape_user9
                    If Not .Text = "" Then
                        If Data_table_with_values.Columns.Contains(.Text) = False Then
                            Data_table_with_values.Columns.Add(.Text, GetType(String))
                        End If
                    End If
                End With
                With ComboBox_shape_user10
                    If Not .Text = "" Then
                        If Data_table_with_values.Columns.Contains(.Text) = False Then
                            Data_table_with_values.Columns.Add(.Text, GetType(String))
                        End If
                    End If
                End With
                With ComboBox_shape_user11
                    If Not .Text = "" Then
                        If Data_table_with_values.Columns.Contains(.Text) = False Then
                            Data_table_with_values.Columns.Add(.Text, GetType(String))
                        End If
                    End If
                End With
                With ComboBox_shape_user12
                    If Not .Text = "" Then
                        If Data_table_with_values.Columns.Contains(.Text) = False Then
                            Data_table_with_values.Columns.Add(.Text, GetType(String))
                        End If
                    End If
                End With
                With ComboBox_shape_user13
                    If Not .Text = "" Then
                        If Data_table_with_values.Columns.Contains(.Text) = False Then
                            Data_table_with_values.Columns.Add(.Text, GetType(String))
                        End If
                    End If
                End With

                If Not ComboBox_HMM_ID_update.Text = "" Then
                    If Data_table_with_values.Columns.Contains(ComboBox_HMM_ID_update.Text) = False Then
                        Data_table_with_values.Columns.Add(ComboBox_HMM_ID_update.Text, GetType(Double))
                    End If
                Else
                    MsgBox("No HMM_ID field specified")
                    Freeze_operations = False
                    Exit Sub
                End If


                Dim Index_parcele As Integer = 0
                Dim Conversie_SQFT_ACREES As Double = 1 * 2.2956841 / 100000
                If CheckBox_convert_sqft_to_acres_UPDATE.Checked = False Then
                    Conversie_SQFT_ACREES = 1
                End If


                Using lock1 As DocumentLock = BaseMap_drawing.LockDocument
                    Using Trans_basemap As Autodesk.AutoCAD.DatabaseServices.Transaction = BaseMap_drawing.TransactionManager.StartTransaction
                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        BlockTable1 = BaseMap_drawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        Dim BTrecord_MS_basemap As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans_basemap.GetObject(BlockTable1(BlockTableRecord.ModelSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                        Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                        Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables
                        Dim parcela_gasita As Boolean = False


                        For Each ObjID1 As ObjectId In BTrecord_MS_basemap
                            Dim Ent1 As Entity
                            Ent1 = TryCast(Trans_basemap.GetObject(ObjID1, OpenMode.ForRead), Entity)
                            If IsNothing(Ent1) = False Then
                                If TypeOf Ent1 Is Polyline Or TypeOf Ent1 Is DBPoint Then

                                    Dim Id1 As ObjectId = Ent1.ObjectId


                                    Using Records1 As Autodesk.Gis.Map.ObjectData.Records = Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                        If IsNothing(Records1) = False Then
                                            If Records1.Count > 0 Then

                                                Dim HMM_ID As String

                                                For Each Record1 As Autodesk.Gis.Map.ObjectData.Record In Records1
                                                    Dim Tabla1 As Autodesk.Gis.Map.ObjectData.Table
                                                    Tabla1 = Tables1(Record1.TableName)

                                                    Dim Field_defs1 As Autodesk.Gis.Map.ObjectData.FieldDefinitions
                                                    Field_defs1 = Tabla1.FieldDefinitions

                                                    For j = 0 To Record1.Count - 1
                                                        Dim Field_def1 As Autodesk.Gis.Map.ObjectData.FieldDefinition
                                                        Field_def1 = Field_defs1(j)

                                                        If Not ComboBox_HMM_ID_update.Text = "" Then
                                                            If Field_def1.Name.ToUpper = ComboBox_HMM_ID_update.Text Then
                                                                Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                Valoare_record1 = Record1(j)
                                                                HMM_ID = Valoare_record1.StrValue
                                                                If Colectie_ID.Contains(HMM_ID) = True Then
                                                                    Data_table_with_values.Rows.Add()
                                                                    Data_table_with_values.Rows(Index_data_table_valori).Item(ComboBox_HMM_ID_update.Text) = Valoare_record1.StrValue
                                                                    parcela_gasita = True
                                                                    Exit For
                                                                End If

                                                                If IsNumeric(HMM_ID) = True Then
                                                                    Dim hmmid_int As Integer = CInt(HMM_ID)
                                                                    Dim hmmid_dbl As Double = CDbl(HMM_ID)

                                                                    If CDbl(hmmid_int) - hmmid_dbl = 0 Then
                                                                        If Colectie_ID.Contains(hmmid_int.ToString) = True Then
                                                                            Data_table_with_values.Rows.Add()
                                                                            Data_table_with_values.Rows(Index_data_table_valori).Item(ComboBox_HMM_ID_update.Text) = Valoare_record1.StrValue
                                                                            parcela_gasita = True
                                                                            Exit For
                                                                        End If
                                                                    End If
                                                                End If

                                                            End If
                                                        End If
                                                    Next

                                                    If parcela_gasita = True Then


                                                        For j = 0 To Record1.Count - 1
                                                            Dim Field_def1 As Autodesk.Gis.Map.ObjectData.FieldDefinition
                                                            Field_def1 = Field_defs1(j)



                                                            With ComboBox_state_update
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        Dim Valoare1 As String = ""
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Valoare1 = Valoare_record1.StrValue
                                                                        End If
                                                                        Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = Valoare1

                                                                    End If
                                                                End If
                                                            End With

                                                            With ComboBox_county_update
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        Dim Valoare1 As String = " "
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Valoare1 = Valoare_record1.StrValue
                                                                        End If
                                                                        Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = Valoare1
                                                                    End If
                                                                End If
                                                            End With

                                                            With ComboBox_town_update
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        Dim Valoare1 As String = " "
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Valoare1 = Valoare_record1.StrValue
                                                                        End If
                                                                        Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = Valoare1
                                                                    End If
                                                                End If
                                                            End With


                                                            With ComboBox_linelist_update
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text.ToUpper Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        Dim Valoare1 As String = " "
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Valoare1 = Valoare_record1.StrValue
                                                                        End If
                                                                        Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = Valoare1
                                                                    End If
                                                                End If
                                                            End With

                                                            With ComboBox_HMM_ID_update
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        Dim Valoare1 As String = " "
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Valoare1 = Valoare_record1.StrValue
                                                                        End If
                                                                        Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = Valoare1
                                                                    End If
                                                                End If
                                                            End With

                                                            With ComboBox_MBL_update
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        Dim Valoare1 As String = " "
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Valoare1 = Valoare_record1.StrValue
                                                                        End If
                                                                        Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = Valoare1
                                                                    End If
                                                                End If
                                                            End With


                                                            With ComboBox_deed_page_update
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        Dim Valoare1 As String = " "
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Valoare1 = Valoare_record1.StrValue
                                                                        End If
                                                                        Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = Valoare1
                                                                    End If
                                                                End If
                                                            End With
                                                            With ComboBox_APN_update
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        Dim Valoare1 As String = " "
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Valoare1 = Valoare_record1.StrValue
                                                                        End If
                                                                        Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = Valoare1
                                                                    End If
                                                                End If
                                                            End With

                                                            With ComboBox_access_road_length_update
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                            Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue)
                                                                        End If
                                                                    End If
                                                                End If
                                                            End With


                                                            With ComboBox_Section_TWP_Range_update
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        Dim Valoare1 As String = " "
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Valoare1 = Valoare_record1.StrValue
                                                                        End If
                                                                        Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = Valoare1
                                                                    End If
                                                                End If
                                                            End With

                                                            With ComboBox_owner_update
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = ComboBox_owner_update.Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        Dim Valoare1 As String = " "
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Valoare1 = Valoare_record1.StrValue
                                                                        End If
                                                                        Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = Valoare1
                                                                    End If
                                                                End If
                                                            End With

                                                            With ComboBox_crossing_length_update
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                            Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue)
                                                                        End If
                                                                    End If
                                                                End If
                                                            End With

                                                            With ComboBox_area_EX_E_update
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                            Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES
                                                                        End If
                                                                    End If
                                                                End If
                                                            End With

                                                            With ComboBox_area_P_E_update
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                            Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES
                                                                        End If
                                                                    End If
                                                                End If
                                                            End With





                                                            With ComboBox_area_TWS_update
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                            Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES
                                                                        End If
                                                                    End If
                                                                End If
                                                            End With

                                                            With ComboBox_area_ATWS_update
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                            Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES
                                                                        End If
                                                                    End If
                                                                End If
                                                            End With

                                                            With ComboBox_area_A_R_update
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                            Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES
                                                                        End If
                                                                    End If
                                                                End If
                                                            End With

                                                            With ComboBox_area_WARE_update
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                            Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES
                                                                        End If
                                                                    End If
                                                                End If
                                                            End With






                                                            With ComboBox_area_TWS_ABD_update
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                            Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES
                                                                        End If
                                                                    End If
                                                                End If
                                                            End With

                                                            With ComboBox_area_TWS_PD_update
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                            Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES
                                                                        End If
                                                                    End If
                                                                End If
                                                            End With


                                                            With ComboBox_shape_user1
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        Dim Valoare1 As String = " "
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Valoare1 = Valoare_record1.StrValue
                                                                        End If
                                                                        Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = Valoare1
                                                                    End If
                                                                End If
                                                            End With

                                                            With ComboBox_shape_user2
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        Dim Valoare1 As String = " "
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Valoare1 = Valoare_record1.StrValue
                                                                        End If
                                                                        Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = Valoare1
                                                                    End If
                                                                End If
                                                            End With

                                                            With ComboBox_shape_user3
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        Dim Valoare1 As String = " "
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Valoare1 = Valoare_record1.StrValue
                                                                        End If
                                                                        Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = Valoare1
                                                                    End If
                                                                End If
                                                            End With
                                                            With ComboBox_shape_user4
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        Dim Valoare1 As String = " "
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Valoare1 = Valoare_record1.StrValue
                                                                        End If
                                                                        Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = Valoare1
                                                                    End If
                                                                End If
                                                            End With

                                                            With ComboBox_shape_user5
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        Dim Valoare1 As String = " "
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Valoare1 = Valoare_record1.StrValue
                                                                        End If
                                                                        Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = Valoare1
                                                                    End If
                                                                End If
                                                            End With
                                                            With ComboBox_shape_user6
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        Dim Valoare1 As String = " "
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Valoare1 = Valoare_record1.StrValue
                                                                        End If
                                                                        Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = Valoare1
                                                                    End If
                                                                End If
                                                            End With
                                                            With ComboBox_shape_user7
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        Dim Valoare1 As String = " "
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Valoare1 = Valoare_record1.StrValue
                                                                        End If
                                                                        Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = Valoare1
                                                                    End If
                                                                End If
                                                            End With
                                                            With ComboBox_shape_user8
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        Dim Valoare1 As String = " "
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Valoare1 = Valoare_record1.StrValue
                                                                        End If
                                                                        Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = Valoare1
                                                                    End If
                                                                End If
                                                            End With
                                                            With ComboBox_shape_user9
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        Dim Valoare1 As String = " "
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Valoare1 = Valoare_record1.StrValue
                                                                        End If
                                                                        Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = Valoare1
                                                                    End If
                                                                End If
                                                            End With
                                                            With ComboBox_shape_user10
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        Dim Valoare1 As String = " "
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Valoare1 = Valoare_record1.StrValue
                                                                        End If
                                                                        Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = Valoare1
                                                                    End If
                                                                End If
                                                            End With
                                                            With ComboBox_shape_user11
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        Dim Valoare1 As String = " "
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Valoare1 = Valoare_record1.StrValue
                                                                        End If
                                                                        Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = Valoare1
                                                                    End If
                                                                End If
                                                            End With
                                                            With ComboBox_shape_user12
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        Dim Valoare1 As String = " "
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Valoare1 = Valoare_record1.StrValue
                                                                        End If
                                                                        Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = Valoare1
                                                                    End If
                                                                End If
                                                            End With
                                                            With ComboBox_shape_user13
                                                                If Not .Text = "" Then
                                                                    If Field_def1.Name.ToUpper = .Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        Dim Valoare1 As String = " "
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Valoare1 = Valoare_record1.StrValue
                                                                        End If
                                                                        Data_table_with_values.Rows(Index_data_table_valori).Item(.Text) = Valoare1
                                                                    End If
                                                                End If
                                                            End With
                                                        Next

                                                    End If
                                                Next

                                                If parcela_gasita = True Then
                                                    Index_data_table_valori = Index_data_table_valori + 1
                                                    parcela_gasita = False
                                                End If


                                            End If
                                        End If
                                    End Using
                                End If
                            End If


                        Next
                        Trans_basemap.Commit()

                    End Using
                End Using

                If IsNothing(Data_table_with_values) = False Then
                    If Data_table_with_values.Rows.Count > 0 Then


                        Dim Index_existent As Integer = 1

                        Dim SheetSet_manager As New AcSmSheetSetMgr
                        Dim SheetSet_database As AcSmDatabase = SheetSet_manager.OpenDatabase(TextBox_sheet_set_template.Text, False)
                        Dim sheetSet As AcSmSheetSet = SheetSet_database.GetSheetSet()


                        For s = 0 To Data_table_with_values.Rows.Count - 1

                            If LockDatabase(SheetSet_database, True) = True Then

                                Dim EnumSheets As IAcSmEnumComponent = sheetSet.GetSheetEnumerator()
                                Dim smComponent As IAcSmComponent
                                Dim sheet1 As IAcSmSheet
                                smComponent = EnumSheets.Next()

                                Dim ID1 As String = ""

                                If Not ComboBox_HMM_ID_update.Text = "" Then
                                    If IsDBNull(Data_table_with_values.Rows(s).Item(ComboBox_HMM_ID_update.Text)) = False Then
                                        ID1 = Data_table_with_values.Rows(s).Item(ComboBox_HMM_ID_update.Text)
                                    End If
                                End If

                                While True
                                    If smComponent Is Nothing Then
                                        Exit While
                                    End If

                                    sheet1 = TryCast(smComponent, IAcSmSheet)
                                    If IsNothing(sheet1) = False Then
                                        Dim customPropertyBag As AcSmCustomPropertyBag = sheet1.GetCustomPropertyBag()
                                        Dim EnumProp As IAcSmEnumProperty = customPropertyBag.GetPropertyEnumerator()
                                        Do
                                            Dim Prop_name As String = ""
                                            Dim Prop_value As AcSmCustomPropertyValue = Nothing
                                            EnumProp.Next(Prop_name, Prop_value)
                                            If Prop_name = "" Then Exit Do
                                            If Prop_name = ComboBox_sheet_set_HMMID_update.Text Then
                                                If Prop_value.GetValue = ID1 Then



                                                    With ComboBox_state
                                                        If Not .Text = "" Then
                                                            Dim Valoare1 As String = ""
                                                            Dim Prefix1 As String = ""
                                                            Dim Suffix1 As String = ""
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    Valoare1 = Data_table_with_values.Rows(s).Item(.Text)
                                                                    If Not TextBox_prefix_state_update.Text = "" Then Prefix1 = TextBox_prefix_state_update.Text
                                                                    If Not TextBox_suffix_state_update.Text = "" Then Suffix1 = TextBox_suffix_state_update.Text
                                                                End If

                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)




                                                                Dim State1 As String = Valoare1
                                                                If CheckBox_state_abreviation_UPDATE.Checked = True Then
                                                                    Select Case State1.ToUpper
                                                                        Case "PA"
                                                                            State1 = Prefix1 & "Pennsylvania".ToUpper & Suffix1
                                                                        Case "NY"
                                                                            State1 = Prefix1 & "New York".ToUpper & Suffix1
                                                                        Case "MA"
                                                                            State1 = Prefix1 & "Massachusetts".ToUpper & Suffix1
                                                                        Case "NH"
                                                                            State1 = Prefix1 & "New Hampshire".ToUpper & Suffix1
                                                                        Case "CT"
                                                                            State1 = Prefix1 & "Connecticut".ToUpper & Suffix1
                                                                        Case "OH"
                                                                            State1 = Prefix1 & "ohio".ToUpper & Suffix1
                                                                    End Select
                                                                End If

                                                                Dim Property_name As String = .Text
                                                                If Not ComboBox_sheet_set_state_update.Text = "" Then
                                                                    Property_name = ComboBox_sheet_set_state_update.Text
                                                                End If


                                                                customPropertyValue.SetValue(State1)
                                                                customPropertyBag.SetProperty(Property_name, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With

                                                    With ComboBox_county_update
                                                        If Not .Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_county_update.Text = "" Then Prefix1 = TextBox_prefix_county_update.Text
                                                                If Not TextBox_suffix_county_update.Text = "" Then Suffix1 = TextBox_suffix_county_update.Text

                                                                Dim Valoare1 As String = ""
                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    Valoare1 = Prefix1 & Data_table_with_values.Rows(s).Item(.Text) & Suffix1
                                                                End If


                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                customPropertyValue.SetValue(Valoare1)
                                                                Dim Property_name As String = .Text
                                                                If Not ComboBox_sheet_set_county_update.Text = "" Then
                                                                    Property_name = ComboBox_sheet_set_county_update.Text
                                                                End If

                                                                customPropertyBag.SetProperty(Property_name, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With

                                                    With ComboBox_town_update
                                                        If Not .Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_town_update.Text = "" Then Prefix1 = TextBox_prefix_town_update.Text
                                                                If Not TextBox_suffix_town_update.Text = "" Then Suffix1 = TextBox_suffix_town_update.Text

                                                                Dim Valoare1 As String = ""
                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    Valoare1 = Prefix1 & Data_table_with_values.Rows(s).Item(.Text) & Suffix1
                                                                End If

                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                customPropertyValue.SetValue(Valoare1)
                                                                Dim Property_name As String = .Text
                                                                If Not ComboBox_sheet_set_town_update.Text = "" Then
                                                                    Property_name = ComboBox_sheet_set_town_update.Text
                                                                End If
                                                                customPropertyBag.SetProperty(Property_name, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With

                                                    With ComboBox_linelist_update
                                                        If Not .Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_linelist_update.Text = "" Then Prefix1 = TextBox_prefix_linelist_update.Text
                                                                If Not TextBox_suffix_linelist_update.Text = "" Then Suffix1 = TextBox_suffix_linelist_update.Text

                                                                Dim Valoare1 As String = ""
                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    Valoare1 = Prefix1 & Data_table_with_values.Rows(s).Item(.Text) & Suffix1
                                                                End If

                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                customPropertyValue.SetValue(Valoare1)
                                                                Dim Property_name As String = .Text
                                                                If Not ComboBox_sheet_set_linelist_update.Text = "" Then
                                                                    Property_name = ComboBox_sheet_set_linelist_update.Text
                                                                End If
                                                                customPropertyBag.SetProperty(Property_name, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With

                                                    With ComboBox_MBL_update
                                                        If Not .Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_mbl_update.Text = "" Then Prefix1 = TextBox_prefix_mbl_update.Text
                                                                If Not TextBox_suffix_mbl_update.Text = "" Then Suffix1 = TextBox_suffix_mbl_update.Text

                                                                Dim Valoare1 As String = ""
                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    Valoare1 = Prefix1 & Data_table_with_values.Rows(s).Item(.Text) & Suffix1
                                                                End If

                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                customPropertyValue.SetValue(Valoare1)
                                                                Dim Property_name As String = .Text
                                                                If Not ComboBox_sheet_set_MBL_update.Text = "" Then
                                                                    Property_name = ComboBox_sheet_set_MBL_update.Text
                                                                End If
                                                                customPropertyBag.SetProperty(Property_name, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With


                                                    With ComboBox_deed_page_update
                                                        If Not .Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_deed_page_update.Text = "" Then Prefix1 = TextBox_prefix_deed_page_update.Text
                                                                If Not TextBox_suffix_deed_page_update.Text = "" Then Suffix1 = TextBox_suffix_deed_page_update.Text

                                                                Dim Valoare1 As String = ""
                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    Valoare1 = Prefix1 & Data_table_with_values.Rows(s).Item(.Text) & Suffix1
                                                                End If

                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                customPropertyValue.SetValue(Valoare1)
                                                                Dim Property_name As String = .Text
                                                                If Not ComboBox_sheet_set_deed_page_update.Text = "" Then
                                                                    Property_name = ComboBox_sheet_set_deed_page_update.Text
                                                                End If
                                                                customPropertyBag.SetProperty(Property_name, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With

                                                    With ComboBox_APN_update
                                                        If Not .Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_apn_update.Text = "" Then Prefix1 = TextBox_prefix_apn_update.Text
                                                                If Not TextBox_suffix_apn_update.Text = "" Then Suffix1 = TextBox_suffix_apn_update.Text

                                                                Dim Valoare1 As String = ""
                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    Valoare1 = Prefix1 & Data_table_with_values.Rows(s).Item(.Text) & Suffix1
                                                                End If

                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                customPropertyValue.SetValue(Valoare1)
                                                                Dim Property_name As String = .Text
                                                                If Not ComboBox_sheet_set_apn_update.Text = "" Then
                                                                    Property_name = ComboBox_sheet_set_apn_update.Text
                                                                End If
                                                                customPropertyBag.SetProperty(Property_name, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With


                                                    With ComboBox_access_road_length_update
                                                        If Not .Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim VALOARE1 As Double = 0
                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    VALOARE1 = Data_table_with_values.Rows(s).Item(.Text)
                                                                End If
                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                Dim Round1 As Integer = 2
                                                                If IsNumeric(TextBox_round_access_road_length_update.Text) = True Then
                                                                    Round1 = Abs(CInt(TextBox_round_access_road_length_update.Text))
                                                                End If

                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_a_r_length_update.Text = "" Then Prefix1 = TextBox_prefix_a_r_length_update.Text
                                                                If Not TextBox_suffix_a_r_length_update.Text = "" Then Suffix1 = TextBox_suffix_a_r_length_update.Text

                                                                customPropertyValue.SetValue(Prefix1 & Get_String_Rounded(VALOARE1, Round1) & Suffix1)
                                                                Dim Property_name As String = .Text
                                                                If Not ComboBox_sheet_set_access_road_length_update.Text = "" Then
                                                                    Property_name = ComboBox_sheet_set_access_road_length_update.Text
                                                                End If
                                                                customPropertyBag.SetProperty(Property_name, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With

                                                    With ComboBox_Section_TWP_Range_update
                                                        If Not .Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_sec_twp_rge_update.Text = "" Then Prefix1 = TextBox_prefix_sec_twp_rge_update.Text
                                                                If Not TextBox_suffix_sec_twp_rge_update.Text = "" Then Suffix1 = TextBox_suffix_sec_twp_rge_update.Text

                                                                Dim Valoare1 As String = ""
                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    Valoare1 = Prefix1 & Data_table_with_values.Rows(s).Item(.Text) & Suffix1
                                                                End If

                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                customPropertyValue.SetValue(Valoare1)
                                                                Dim Property_name As String = .Text
                                                                If Not ComboBox_sheet_set_sec_twp_range_update.Text = "" Then
                                                                    Property_name = ComboBox_sheet_set_sec_twp_range_update.Text
                                                                End If
                                                                customPropertyBag.SetProperty(Property_name, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With


                                                    With ComboBox_owner_update
                                                        If Not .Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_owner_update.Text = "" Then Prefix1 = TextBox_prefix_owner_update.Text
                                                                If Not TextBox_suffix_owner_update.Text = "" Then Suffix1 = TextBox_suffix_owner_update.Text

                                                                Dim Valoare1 As String = ""
                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    Valoare1 = Prefix1 & Data_table_with_values.Rows(s).Item(.Text) & Suffix1
                                                                End If

                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                customPropertyValue.SetValue(Valoare1)
                                                                Dim Property_name As String = .Text
                                                                If Not ComboBox_sheet_set_owner_update.Text = "" Then
                                                                    Property_name = ComboBox_sheet_set_owner_update.Text
                                                                End If
                                                                customPropertyBag.SetProperty(Property_name, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With


                                                    With ComboBox_crossing_length_update
                                                        If Not .Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim VALOARE1 As Double = 0
                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    VALOARE1 = Data_table_with_values.Rows(s).Item(.Text)
                                                                End If

                                                                Dim customPropertyValue1 As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue1.InitNew(sheet1)
                                                                customPropertyValue1.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                Dim Round1 As Integer = 2
                                                                If IsNumeric(TextBox_round_crossing_length_ft_update.Text) = True Then
                                                                    Round1 = Abs(CInt(TextBox_round_crossing_length_ft_update.Text))
                                                                End If

                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_crossing_length_ft_update.Text = "" Then Prefix1 = TextBox_prefix_crossing_length_ft_update.Text
                                                                If Not TextBox_suffix_crossing_length_ft_update.Text = "" Then Suffix1 = TextBox_suffix_crossing_length_ft_update.Text
                                                                customPropertyValue1.SetValue(Prefix1 & Get_String_Rounded(VALOARE1, Round1) & Suffix1)


                                                                Dim Prefix2 As String = ""
                                                                Dim Suffix2 As String = ""
                                                                If Not TextBox_prefix_crossing_length_rod_update.Text = "" Then Prefix2 = TextBox_prefix_crossing_length_rod_update.Text
                                                                If Not TextBox_suffix_crossing_length_rod_update.Text = "" Then Suffix2 = TextBox_suffix_crossing_length_rod_update.Text

                                                                Dim Property_name1 As String = .Text
                                                                If Not ComboBox_sheet_set_crosing_length_ft_update.Text = "" Then
                                                                    Property_name1 = ComboBox_sheet_set_crosing_length_ft_update.Text
                                                                End If
                                                                customPropertyBag.SetProperty(Property_name1, customPropertyValue1)

                                                                Dim customPropertyValue2 As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue2.InitNew(sheet1)
                                                                customPropertyValue2.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                Dim Round2 As Integer = 2
                                                                If IsNumeric(TextBox_round_crossing_length_rod_update.Text) = True Then
                                                                    Round2 = Abs(CInt(TextBox_round_crossing_length_rod_update.Text))
                                                                End If

                                                                customPropertyValue2.SetValue(Prefix2 & Get_String_Rounded(VALOARE1 / 16.5, Round2) & Suffix2)
                                                                Dim Property_name2 As String = ""
                                                                If Not ComboBox_sheet_set_crosing_length_rod_update.Text = "" Then
                                                                    Property_name2 = ComboBox_sheet_set_crosing_length_rod_update.Text
                                                                    customPropertyBag.SetProperty(Property_name2, customPropertyValue2)
                                                                End If
                                                            End If
                                                        End If
                                                    End With



                                                    With ComboBox_area_EX_E_update
                                                        If Not .Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim VALOARE1 As Double = 0
                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    VALOARE1 = Data_table_with_values.Rows(s).Item(.Text)
                                                                End If
                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                Dim Round1 As Integer = 2
                                                                If IsNumeric(TextBox_round_EX_E_update.Text) = True Then
                                                                    Round1 = Abs(CInt(TextBox_round_EX_E_update.Text))
                                                                End If

                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_AREA_EX_E_update.Text = "" Then Prefix1 = TextBox_prefix_AREA_EX_E_update.Text
                                                                If Not TextBox_suffix_AREA_EX_E_update.Text = "" Then Suffix1 = TextBox_suffix_AREA_EX_E_update.Text
                                                                customPropertyValue.SetValue(Prefix1 & Get_String_Rounded(VALOARE1, Round1) & Suffix1)

                                                                Dim Property_name As String = .Text
                                                                If Not ComboBox_sheet_set_ex_e_update.Text = "" Then
                                                                    Property_name = ComboBox_sheet_set_ex_e_update.Text
                                                                End If
                                                                customPropertyBag.SetProperty(Property_name, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With


                                                    With ComboBox_area_P_E_update
                                                        If Not .Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim VALOARE1 As Double = 0
                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    VALOARE1 = Data_table_with_values.Rows(s).Item(.Text)
                                                                End If
                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                Dim Round1 As Integer = 2
                                                                If IsNumeric(TextBox_round_P_E_update.Text) = True Then
                                                                    Round1 = Abs(CInt(TextBox_round_P_E_update.Text))
                                                                End If

                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_AREA_P_E_update.Text = "" Then Prefix1 = TextBox_prefix_AREA_P_E_update.Text
                                                                If Not TextBox_suffix_AREA_P_E_update.Text = "" Then Suffix1 = TextBox_suffix_AREA_P_E_update.Text
                                                                customPropertyValue.SetValue(Prefix1 & Get_String_Rounded(VALOARE1, Round1) & Suffix1)

                                                                Dim Property_name As String = .Text
                                                                If Not ComboBox_sheet_set_p_e_update.Text = "" Then
                                                                    Property_name = ComboBox_sheet_set_p_e_update.Text
                                                                End If
                                                                customPropertyBag.SetProperty(Property_name, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With

                                                    With ComboBox_area_TWS_update
                                                        If Not .Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim VALOARE1 As Double = 0
                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    VALOARE1 = Data_table_with_values.Rows(s).Item(.Text)
                                                                End If
                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                Dim Round1 As Integer = 2
                                                                If IsNumeric(TextBox_round_TWS_update.Text) = True Then
                                                                    Round1 = Abs(CInt(TextBox_round_TWS_update.Text))
                                                                End If

                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_AREA_TWS_update.Text = "" Then Prefix1 = TextBox_prefix_AREA_TWS_update.Text
                                                                If Not TextBox_suffix_AREA_TWS_update.Text = "" Then Suffix1 = TextBox_suffix_AREA_TWS_update.Text
                                                                customPropertyValue.SetValue(Prefix1 & Get_String_Rounded(VALOARE1, Round1) & Suffix1)

                                                                Dim Property_name As String = .Text
                                                                If Not ComboBox_sheet_set_tws_update.Text = "" Then
                                                                    Property_name = ComboBox_sheet_set_tws_update.Text
                                                                End If
                                                                customPropertyBag.SetProperty(Property_name, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With


                                                    With ComboBox_area_ATWS_update
                                                        If Not .Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim VALOARE1 As Double = 0
                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    VALOARE1 = Data_table_with_values.Rows(s).Item(.Text)
                                                                End If
                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                Dim Round1 As Integer = 2
                                                                If IsNumeric(TextBox_round_ATWS_update.Text) = True Then
                                                                    Round1 = Abs(CInt(TextBox_round_ATWS_update.Text))
                                                                End If

                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_area_ATWS_update.Text = "" Then Prefix1 = TextBox_prefix_area_ATWS_update.Text
                                                                If Not TextBox_suffix_area_ATWS_update.Text = "" Then Suffix1 = TextBox_suffix_area_ATWS_update.Text
                                                                customPropertyValue.SetValue(Prefix1 & Get_String_Rounded(VALOARE1, Round1) & Suffix1)

                                                                Dim Property_name As String = .Text
                                                                If Not ComboBox_sheet_set_atws_update.Text = "" Then
                                                                    Property_name = ComboBox_sheet_set_atws_update.Text
                                                                End If
                                                                customPropertyBag.SetProperty(Property_name, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With







                                                    With ComboBox_area_A_R_update
                                                        If Not .Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim VALOARE1 As Double = 0
                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    VALOARE1 = Data_table_with_values.Rows(s).Item(.Text)
                                                                End If
                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                Dim Round1 As Integer = 2
                                                                If IsNumeric(TextBox_round_A_R_update.Text) = True Then
                                                                    Round1 = Abs(CInt(TextBox_round_A_R_update.Text))
                                                                End If

                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_AREA_A_R_update.Text = "" Then Prefix1 = TextBox_prefix_AREA_A_R_update.Text
                                                                If Not TextBox_suffix_AREA_A_R_update.Text = "" Then Suffix1 = TextBox_suffix_AREA_A_R_update.Text
                                                                customPropertyValue.SetValue(Prefix1 & Get_String_Rounded(VALOARE1, Round1) & Suffix1)

                                                                Dim Property_name As String = .Text
                                                                If Not ComboBox_sheet_set_a_r_update.Text = "" Then
                                                                    Property_name = ComboBox_sheet_set_a_r_update.Text
                                                                End If
                                                                customPropertyBag.SetProperty(Property_name, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With

                                                    With ComboBox_area_WARE_update
                                                        If Not .Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim VALOARE1 As Double = 0
                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    VALOARE1 = Data_table_with_values.Rows(s).Item(.Text)
                                                                End If
                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                Dim Round1 As Integer = 2
                                                                If IsNumeric(TextBox_round_WARE_update.Text) = True Then
                                                                    Round1 = Abs(CInt(TextBox_round_WARE_update.Text))
                                                                End If

                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_AREA_WARE_update.Text = "" Then Prefix1 = TextBox_prefix_AREA_WARE_update.Text
                                                                If Not TextBox_suffix_AREA_WARE_update.Text = "" Then Suffix1 = TextBox_suffix_AREA_WARE_update.Text
                                                                customPropertyValue.SetValue(Prefix1 & Get_String_Rounded(VALOARE1, Round1) & Suffix1)

                                                                Dim Property_name As String = .Text
                                                                If Not ComboBox_sheet_set_ware_update.Text = "" Then
                                                                    Property_name = ComboBox_sheet_set_ware_update.Text
                                                                End If
                                                                customPropertyBag.SetProperty(Property_name, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With



                                                    With ComboBox_area_TWS_ABD_update
                                                        If Not .Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim VALOARE1 As Double = 0
                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    VALOARE1 = Data_table_with_values.Rows(s).Item(.Text)
                                                                End If
                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                Dim Round1 As Integer = 2
                                                                If IsNumeric(TextBox_round_TWS_ABD_update.Text) = True Then
                                                                    Round1 = Abs(CInt(TextBox_round_TWS_ABD_update.Text))
                                                                End If

                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_AREA_TWS_ABD_update.Text = "" Then Prefix1 = TextBox_prefix_AREA_TWS_ABD_update.Text
                                                                If Not TextBox_suffix_AREA_TWS_ABD_update.Text = "" Then Suffix1 = TextBox_suffix_AREA_TWS_ABD_update.Text
                                                                customPropertyValue.SetValue(Prefix1 & Get_String_Rounded(VALOARE1, Round1) & Suffix1)

                                                                Dim Property_name As String = .Text
                                                                If Not ComboBox_sheet_set_tws_ABD_update.Text = "" Then
                                                                    Property_name = ComboBox_sheet_set_tws_ABD_update.Text
                                                                End If
                                                                customPropertyBag.SetProperty(Property_name, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With

                                                    With ComboBox_area_TWS_PD_update
                                                        If Not .Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim VALOARE1 As Double = 0
                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    VALOARE1 = Data_table_with_values.Rows(s).Item(.Text)
                                                                End If
                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                Dim Round1 As Integer = 2
                                                                If IsNumeric(TextBox_round_TWS_PD_update.Text) = True Then
                                                                    Round1 = Abs(CInt(TextBox_round_TWS_PD_update.Text))
                                                                End If

                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_AREA_TWS_PD_update.Text = "" Then Prefix1 = TextBox_prefix_AREA_TWS_PD_update.Text
                                                                If Not TextBox_suffix_AREA_TWS_PD_update.Text = "" Then Suffix1 = TextBox_suffix_AREA_TWS_PD_update.Text
                                                                customPropertyValue.SetValue(Prefix1 & Get_String_Rounded(VALOARE1, Round1) & Suffix1)

                                                                Dim Property_name As String = .Text
                                                                If Not ComboBox_sheet_set_TWS_PD_update.Text = "" Then
                                                                    Property_name = ComboBox_sheet_set_TWS_PD_update.Text
                                                                End If
                                                                customPropertyBag.SetProperty(Property_name, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With


                                                    With ComboBox_shape_user1
                                                        If Not .Text = "" And Not ComboBox_sheet_user1.Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim Valoare1 As String = ""
                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_user1.Text = "" Then Prefix1 = TextBox_prefix_user1.Text
                                                                If Not TextBox_suffix_user1.Text = "" Then Suffix1 = TextBox_suffix_user1.Text

                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    Valoare1 = Prefix1 & Data_table_with_values.Rows(s).Item(.Text) & Suffix1
                                                                End If

                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                customPropertyValue.SetValue(Valoare1)
                                                                customPropertyBag.SetProperty(ComboBox_sheet_user1.Text, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With


                                                    With ComboBox_shape_user2
                                                        If Not .Text = "" And Not ComboBox_sheet_user2.Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim Valoare1 As String = ""
                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_user2.Text = "" Then Prefix1 = TextBox_prefix_user2.Text
                                                                If Not TextBox_suffix_user2.Text = "" Then Suffix1 = TextBox_suffix_user2.Text

                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    Valoare1 = Prefix1 & Data_table_with_values.Rows(s).Item(.Text) & Suffix1
                                                                End If

                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                customPropertyValue.SetValue(Valoare1)
                                                                customPropertyBag.SetProperty(ComboBox_sheet_user2.Text, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With

                                                    With ComboBox_shape_user3
                                                        If Not .Text = "" And Not ComboBox_sheet_user3.Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim Valoare1 As String = ""
                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_user3.Text = "" Then Prefix1 = TextBox_prefix_user3.Text
                                                                If Not TextBox_suffix_user3.Text = "" Then Suffix1 = TextBox_suffix_user3.Text

                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    Valoare1 = Prefix1 & Data_table_with_values.Rows(s).Item(.Text) & Suffix1
                                                                End If

                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                customPropertyValue.SetValue(Valoare1)
                                                                customPropertyBag.SetProperty(ComboBox_sheet_user3.Text, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With

                                                    With ComboBox_shape_user4
                                                        If Not .Text = "" And Not ComboBox_sheet_user4.Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim Valoare1 As String = ""
                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_user4.Text = "" Then Prefix1 = TextBox_prefix_user4.Text
                                                                If Not TextBox_suffix_user4.Text = "" Then Suffix1 = TextBox_suffix_user4.Text

                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    Valoare1 = Prefix1 & Data_table_with_values.Rows(s).Item(.Text) & Suffix1
                                                                End If

                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                customPropertyValue.SetValue(Valoare1)
                                                                customPropertyBag.SetProperty(ComboBox_sheet_user4.Text, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With


                                                    With ComboBox_shape_user5
                                                        If Not .Text = "" And Not ComboBox_sheet_user5.Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim Valoare1 As String = ""
                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_user5.Text = "" Then Prefix1 = TextBox_prefix_user5.Text
                                                                If Not TextBox_suffix_user5.Text = "" Then Suffix1 = TextBox_suffix_user5.Text

                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    Valoare1 = Prefix1 & Data_table_with_values.Rows(s).Item(.Text) & Suffix1
                                                                End If

                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                customPropertyValue.SetValue(Valoare1)
                                                                customPropertyBag.SetProperty(ComboBox_sheet_user5.Text, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With

                                                    With ComboBox_shape_user6
                                                        If Not .Text = "" And Not ComboBox_sheet_user6.Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim Valoare1 As String = ""
                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_user6.Text = "" Then Prefix1 = TextBox_prefix_user6.Text
                                                                If Not TextBox_suffix_user6.Text = "" Then Suffix1 = TextBox_suffix_user6.Text

                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    Valoare1 = Prefix1 & Data_table_with_values.Rows(s).Item(.Text) & Suffix1
                                                                End If

                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                customPropertyValue.SetValue(Valoare1)
                                                                customPropertyBag.SetProperty(ComboBox_sheet_user6.Text, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With
                                                    With ComboBox_shape_user7
                                                        If Not .Text = "" And Not ComboBox_sheet_user7.Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim Valoare1 As String = ""
                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_user7.Text = "" Then Prefix1 = TextBox_prefix_user7.Text
                                                                If Not TextBox_suffix_user7.Text = "" Then Suffix1 = TextBox_suffix_user7.Text

                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    Valoare1 = Prefix1 & Data_table_with_values.Rows(s).Item(.Text) & Suffix1
                                                                End If

                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                customPropertyValue.SetValue(Valoare1)
                                                                customPropertyBag.SetProperty(ComboBox_sheet_user7.Text, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With
                                                    With ComboBox_shape_user8
                                                        If Not .Text = "" And Not ComboBox_sheet_user8.Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim Valoare1 As String = ""
                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_user8.Text = "" Then Prefix1 = TextBox_prefix_user8.Text
                                                                If Not TextBox_suffix_user8.Text = "" Then Suffix1 = TextBox_suffix_user8.Text

                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    Valoare1 = Prefix1 & Data_table_with_values.Rows(s).Item(.Text) & Suffix1
                                                                End If

                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                customPropertyValue.SetValue(Valoare1)
                                                                customPropertyBag.SetProperty(ComboBox_sheet_user8.Text, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With

                                                    With ComboBox_shape_user9
                                                        If Not .Text = "" And Not ComboBox_sheet_user9.Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim Valoare1 As String = ""
                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_user9.Text = "" Then Prefix1 = TextBox_prefix_user9.Text
                                                                If Not TextBox_suffix_user9.Text = "" Then Suffix1 = TextBox_suffix_user9.Text

                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    Valoare1 = Prefix1 & Data_table_with_values.Rows(s).Item(.Text) & Suffix1
                                                                End If

                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                customPropertyValue.SetValue(Valoare1)
                                                                customPropertyBag.SetProperty(ComboBox_sheet_user9.Text, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With

                                                    With ComboBox_shape_user10
                                                        If Not .Text = "" And Not ComboBox_sheet_user10.Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim Valoare1 As String = ""
                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_user10.Text = "" Then Prefix1 = TextBox_prefix_user10.Text
                                                                If Not TextBox_suffix_user10.Text = "" Then Suffix1 = TextBox_suffix_user10.Text

                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    Valoare1 = Prefix1 & Data_table_with_values.Rows(s).Item(.Text) & Suffix1
                                                                End If

                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                customPropertyValue.SetValue(Valoare1)
                                                                customPropertyBag.SetProperty(ComboBox_sheet_user10.Text, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With

                                                    With ComboBox_shape_user11
                                                        If Not .Text = "" And Not ComboBox_sheet_user11.Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim Valoare1 As String = ""
                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_user11.Text = "" Then Prefix1 = TextBox_prefix_user11.Text
                                                                If Not TextBox_suffix_user11.Text = "" Then Suffix1 = TextBox_suffix_user11.Text

                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    Valoare1 = Prefix1 & Data_table_with_values.Rows(s).Item(.Text) & Suffix1
                                                                End If

                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                customPropertyValue.SetValue(Valoare1)
                                                                customPropertyBag.SetProperty(ComboBox_sheet_user11.Text, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With

                                                    With ComboBox_shape_user12
                                                        If Not .Text = "" And Not ComboBox_sheet_user12.Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim Valoare1 As String = ""
                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_user12.Text = "" Then Prefix1 = TextBox_prefix_user12.Text
                                                                If Not TextBox_suffix_user12.Text = "" Then Suffix1 = TextBox_suffix_user12.Text

                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    Valoare1 = Prefix1 & Data_table_with_values.Rows(s).Item(.Text) & Suffix1
                                                                End If

                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                customPropertyValue.SetValue(Valoare1)
                                                                customPropertyBag.SetProperty(ComboBox_sheet_user12.Text, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With

                                                    With ComboBox_shape_user13
                                                        If Not .Text = "" And Not ComboBox_sheet_user13.Text = "" Then
                                                            If Data_table_with_values.Columns.Contains(.Text) = True Then
                                                                Dim Valoare1 As String = ""
                                                                Dim Prefix1 As String = ""
                                                                Dim Suffix1 As String = ""
                                                                If Not TextBox_prefix_user13.Text = "" Then Prefix1 = TextBox_prefix_user13.Text
                                                                If Not TextBox_suffix_user13.Text = "" Then Suffix1 = TextBox_suffix_user13.Text

                                                                If IsDBNull(Data_table_with_values.Rows(s).Item(.Text)) = False Then
                                                                    Valoare1 = Prefix1 & Data_table_with_values.Rows(s).Item(.Text) & Suffix1
                                                                End If

                                                                Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
                                                                customPropertyValue.InitNew(sheet1)
                                                                customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP)
                                                                customPropertyValue.SetValue(Valoare1)
                                                                customPropertyBag.SetProperty(ComboBox_sheet_user13.Text, customPropertyValue)
                                                            End If
                                                        End If
                                                    End With

                                                    Exit Do
                                                End If

                                            End If
                                        Loop
                                    End If

                                    smComponent = EnumSheets.Next()
                                End While


                                LockDatabase(SheetSet_database, False)
                            Else

                                MsgBox(SheetSet_database.GetLockStatus.ToString)
                                ' Display error message
                                MsgBox("Sheet set could not be opened for write.")
                            End If
                        Next

                        ' Close the sheet set
                        SheetSet_manager.Close(SheetSet_database)

                    End If
                End If

                MsgBox("You are done")
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_DST_to_Clipboard_Click(sender As Object, e As EventArgs) Handles Button_DST_to_Clipboard.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            If TextBox_sheet_set_template.Text = "" Then
                MsgBox("Please specify the SHEET SET file")
                Freeze_operations = False
                Exit Sub
            End If
            If Not Strings.Right(TextBox_sheet_set_template.Text, 3).ToUpper = "DST" Then
                MsgBox("Please specify the SHEET SET file")
                Freeze_operations = False
                Exit Sub
            End If
            Try
                Dim Data_table_sheetSet As New System.Data.DataTable
                Data_table_sheetSet.Columns.Add("SHEET_NAME", GetType(String))
                Dim Index1 As Integer = 0
                Button_DST_to_Clipboard.Visible = False

                Dim SheetSet_manager As New AcSmSheetSetMgr
                Dim SheetSet_database As AcSmDatabase = SheetSet_manager.OpenDatabase(TextBox_sheet_set_template.Text, False)
                Dim sheetSet As AcSmSheetSet = SheetSet_database.GetSheetSet()

                If LockDatabase(SheetSet_database, True) = True Then
                    Dim EnumSheets As IAcSmEnumComponent = sheetSet.GetSheetEnumerator()
                    Dim smComponent As IAcSmComponent
                    Dim sheet1 As IAcSmSheet
                    smComponent = EnumSheets.Next()
                    While True
                        If smComponent Is Nothing Then
                            Exit While
                        End If
                        sheet1 = TryCast(smComponent, IAcSmSheet)
                        If IsNothing(sheet1) = False Then
                            Dim customPropertyBag As AcSmCustomPropertyBag = sheet1.GetCustomPropertyBag()
                            Dim EnumProp As IAcSmEnumProperty = customPropertyBag.GetPropertyEnumerator()
                            Dim Add_row As Boolean = True
                            Do
                                Dim Prop_name As String = ""
                                Dim Prop_value As AcSmCustomPropertyValue = Nothing
                                EnumProp.Next(Prop_name, Prop_value)
                                If Prop_name = "" Then Exit Do
                                If IsNothing(Prop_value) = False Then
                                    If Not Prop_value.GetValue = "" Then
                                        If Add_row = True Then
                                            Data_table_sheetSet.Rows.Add()
                                            Data_table_sheetSet.Rows(Index1).Item("SHEET_NAME") = sheet1.GetTitle
                                            Add_row = False
                                        End If
                                        If Data_table_sheetSet.Columns.Contains(Prop_name) = False Then
                                            Data_table_sheetSet.Columns.Add(Prop_name, GetType(String))
                                        End If
                                        Data_table_sheetSet.Rows(Index1).Item(Prop_name) = Prop_value.GetValue
                                    End If
                                End If
                            Loop
                            Index1 = Index1 + 1
                        End If
                        smComponent = EnumSheets.Next()
                    End While
                    LockDatabase(SheetSet_database, False)

                    If Data_table_sheetSet.Rows.Count > 0 Then
                        Dim Values_string As String = ""
                        Values_string = Data_table_sheetSet.Columns(0).ColumnName
                        For r = 1 To Data_table_sheetSet.Columns.Count - 1
                            Values_string = Values_string & Chr(9) & Data_table_sheetSet.Columns(r).ColumnName
                        Next
                        For s = 0 To Data_table_sheetSet.Rows.Count - 1
                            Dim Temp1 As String = Data_table_sheetSet.Rows(s).Item(0)
                            For r = 1 To Data_table_sheetSet.Columns.Count - 1
                                Temp1 = Temp1 & Chr(9) & Data_table_sheetSet.Rows(s).Item(r)
                            Next
                            Values_string = Values_string & vbCrLf & Temp1
                        Next

                        My.Computer.Clipboard.SetText(Values_string)
                    End If
                    MsgBox("Data has been copied to the clipboard" & vbCrLf & "(" & (Data_table_sheetSet.Rows.Count - 1).ToString & " rows)")
                Else
                    MsgBox(SheetSet_database.GetLockStatus.ToString)
                    ' Display error message
                    MsgBox("Sheet set could not be opened for write.")
                End If
                Button_DST_to_Clipboard.Visible = True
            Catch ex As Exception
                Button_DST_to_Clipboard.Visible = True
                MsgBox(ex.Message)
            End Try
            Button_DST_to_Clipboard.Visible = True
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_compare_DST_Click(sender As Object, e As EventArgs) Handles Button_compare_DST.Click
        Dim Conversie_SQFT_ACREES As Double = 1 * 2.2956841 / 100000
        If CheckBox_convert_sqft_to_acres_UPDATE.Checked = False Then
            Conversie_SQFT_ACREES = 1
        End If
        If TextBox_sheet_set_template.Text = "" Then
            MsgBox("Please specify the SHEET SET file")
            Exit Sub
        End If
        If Not Strings.Right(TextBox_sheet_set_template.Text, 3).ToUpper = "DST" Then
            MsgBox("Please specify the SHEET SET file")
            Exit Sub
        End If

        If Freeze_operations = False Then
            Freeze_operations = True

            Colectie_ID = New Specialized.StringCollection

            Try
                Dim Data_table_sheetSet As New System.Data.DataTable
                Data_table_sheetSet.Columns.Add("SHEET_NAME", GetType(String))
                Dim Index1 As Integer = 0
                Button_compare_DST.Visible = False

                Dim SheetSet_manager As New AcSmSheetSetMgr
                Dim SheetSet_database As AcSmDatabase = SheetSet_manager.OpenDatabase(TextBox_sheet_set_template.Text, False)
                Dim sheetSet As AcSmSheetSet = SheetSet_database.GetSheetSet()

                If LockDatabase(SheetSet_database, True) = True Then
                    Dim EnumSheets As IAcSmEnumComponent = sheetSet.GetSheetEnumerator()
                    Dim smComponent As IAcSmComponent
                    Dim sheet1 As IAcSmSheet
                    smComponent = EnumSheets.Next()
                    While True
                        If smComponent Is Nothing Then
                            Exit While
                        End If
                        sheet1 = TryCast(smComponent, IAcSmSheet)
                        If IsNothing(sheet1) = False Then
                            Dim customPropertyBag As AcSmCustomPropertyBag = sheet1.GetCustomPropertyBag()
                            Dim EnumProp As IAcSmEnumProperty = customPropertyBag.GetPropertyEnumerator()
                            Dim Add_row As Boolean = True
                            Do
                                Dim Prop_name As String = ""
                                Dim Prop_value As AcSmCustomPropertyValue = Nothing
                                EnumProp.Next(Prop_name, Prop_value)
                                If Prop_name = "" Then Exit Do
                                If IsNothing(Prop_value) = False Then
                                    If Not Prop_value.GetValue = "" Then
                                        If Add_row = True Then
                                            Data_table_sheetSet.Rows.Add()
                                            Data_table_sheetSet.Rows(Index1).Item("SHEET_NAME") = sheet1.GetTitle
                                            Add_row = False
                                        End If
                                        If Data_table_sheetSet.Columns.Contains(Prop_name) = False Then
                                            Data_table_sheetSet.Columns.Add(Prop_name, GetType(String))
                                        End If

                                        Data_table_sheetSet.Rows(Index1).Item(Prop_name) = Prop_value.GetValue
                                        If Prop_name = ComboBox_HMM_ID_update.Text Then
                                            Colectie_ID.Add(Prop_value.GetValue)
                                        End If
                                    End If
                                End If
                            Loop
                            Index1 = Index1 + 1
                        End If
                        smComponent = EnumSheets.Next()
                    End While
                    LockDatabase(SheetSet_database, False)

                    Dim Data_table_cu_valori As New System.Data.DataTable
                    Dim Index_data_table_valori As Double = 0

                    If Not ComboBox_state_update.Text = "" Then
                        If Data_table_cu_valori.Columns.Contains(ComboBox_state_update.Text) = False Then
                            Data_table_cu_valori.Columns.Add(ComboBox_state_update.Text, GetType(String))
                        End If
                    End If

                    If Not ComboBox_county_update.Text = "" Then
                        If Data_table_cu_valori.Columns.Contains(ComboBox_county_update.Text) = False Then
                            Data_table_cu_valori.Columns.Add(ComboBox_county_update.Text, GetType(String))
                        End If
                    End If

                    If Not ComboBox_town_update.Text = "" Then
                        If Data_table_cu_valori.Columns.Contains(ComboBox_town_update.Text) = False Then
                            Data_table_cu_valori.Columns.Add(ComboBox_town_update.Text, GetType(String))
                        End If
                    End If
                    If Not ComboBox_owner_update.Text = "" Then
                        If Data_table_cu_valori.Columns.Contains(ComboBox_owner_update.Text) = False Then
                            Data_table_cu_valori.Columns.Add(ComboBox_owner_update.Text, GetType(String))
                        End If
                    End If





                    If Not ComboBox_linelist_update.Text = "" Then
                        If Data_table_cu_valori.Columns.Contains(ComboBox_linelist_update.Text) = False Then
                            Data_table_cu_valori.Columns.Add(ComboBox_linelist_update.Text, GetType(String))
                        End If
                    End If

                    If Not ComboBox_MBL_update.Text = "" Then
                        If Data_table_cu_valori.Columns.Contains(ComboBox_MBL_update.Text) = False Then
                            Data_table_cu_valori.Columns.Add(ComboBox_MBL_update.Text, GetType(String))
                        End If
                    End If

                    If Not ComboBox_area_ATWS_update.Text = "" Then
                        If Data_table_cu_valori.Columns.Contains(ComboBox_area_ATWS_update.Text) = False Then
                            Data_table_cu_valori.Columns.Add(ComboBox_area_ATWS_update.Text, GetType(Double))
                        End If
                    End If

                    If Not ComboBox_area_EX_E_update.Text = "" Then
                        If Data_table_cu_valori.Columns.Contains(ComboBox_area_EX_E_update.Text) = False Then
                            Data_table_cu_valori.Columns.Add(ComboBox_area_EX_E_update.Text, GetType(Double))
                        End If
                    End If

                    If Not ComboBox_area_P_E_update.Text = "" Then
                        If Data_table_cu_valori.Columns.Contains(ComboBox_area_P_E_update.Text) = False Then
                            Data_table_cu_valori.Columns.Add(ComboBox_area_P_E_update.Text, GetType(Double))
                        End If
                    End If

                    If Not ComboBox_area_TWS_update.Text = "" Then
                        If Data_table_cu_valori.Columns.Contains(ComboBox_area_TWS_update.Text) = False Then
                            Data_table_cu_valori.Columns.Add(ComboBox_area_TWS_update.Text, GetType(Double))
                        End If
                    End If

                    If Not ComboBox_area_A_R_update.Text = "" Then
                        If Data_table_cu_valori.Columns.Contains(ComboBox_area_A_R_update.Text) = False Then
                            Data_table_cu_valori.Columns.Add(ComboBox_area_A_R_update.Text, GetType(Double))
                        End If
                    End If

                    If Not ComboBox_area_WARE_update.Text = "" Then
                        If Data_table_cu_valori.Columns.Contains(ComboBox_area_WARE_update.Text) = False Then
                            Data_table_cu_valori.Columns.Add(ComboBox_area_WARE_update.Text, GetType(Double))
                        End If
                    End If

                    With ComboBox_area_TWS_ABD_update
                        If Not .Text = "" Then
                            If Data_table_cu_valori.Columns.Contains(.Text) = False Then
                                Data_table_cu_valori.Columns.Add(.Text, GetType(Double))
                            End If
                        End If
                    End With
                    With ComboBox_area_TWS_PD_update
                        If Not .Text = "" Then
                            If Data_table_cu_valori.Columns.Contains(.Text) = False Then
                                Data_table_cu_valori.Columns.Add(.Text, GetType(Double))
                            End If
                        End If
                    End With

                    With ComboBox_MBL_update
                        If Not .Text = "" Then
                            If Data_table_cu_valori.Columns.Contains(.Text) = False Then
                                Data_table_cu_valori.Columns.Add(.Text, GetType(String))
                            End If
                        End If
                    End With

                    With ComboBox_deed_page_update
                        If Not .Text = "" Then
                            If Data_table_cu_valori.Columns.Contains(.Text) = False Then
                                Data_table_cu_valori.Columns.Add(.Text, GetType(String))
                            End If
                        End If
                    End With
                    With ComboBox_APN_update
                        If Not .Text = "" Then
                            If Data_table_cu_valori.Columns.Contains(.Text) = False Then
                                Data_table_cu_valori.Columns.Add(.Text, GetType(String))
                            End If
                        End If
                    End With
                    With ComboBox_crossing_length_update
                        If Not .Text = "" Then
                            If Data_table_cu_valori.Columns.Contains(.Text) = False Then
                                Data_table_cu_valori.Columns.Add(.Text, GetType(Double))
                            End If
                        End If
                    End With

                    With ComboBox_access_road_length_update
                        If Not .Text = "" Then
                            If Data_table_cu_valori.Columns.Contains(.Text) = False Then
                                Data_table_cu_valori.Columns.Add(.Text, GetType(Double))
                            End If
                        End If
                    End With

                    With ComboBox_Section_TWP_Range_update
                        If Not .Text = "" Then
                            If Data_table_cu_valori.Columns.Contains(.Text) = False Then
                                Data_table_cu_valori.Columns.Add(.Text, GetType(String))
                            End If
                        End If
                    End With

                    If Not ComboBox_HMM_ID_update.Text = "" Then
                        If Data_table_cu_valori.Columns.Contains(ComboBox_HMM_ID_update.Text) = False Then
                            Data_table_cu_valori.Columns.Add(ComboBox_HMM_ID_update.Text, GetType(Double))
                        End If
                    Else
                        MsgBox("No HMM_ID field specified")
                        Button_compare_DST.Visible = True
                        Freeze_operations = False
                        Exit Sub
                    End If



                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    Using lock1 As DocumentLock = ThisDrawing.LockDocument
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            Dim BTrecord_MS_basemap As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(BlockTable1(BlockTableRecord.ModelSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                            Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                            Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables
                            Dim parcela_gasita As Boolean = False

                            For Each ObjID1 As ObjectId In BTrecord_MS_basemap
                                Dim Ent1 As Entity
                                Ent1 = TryCast(Trans1.GetObject(ObjID1, OpenMode.ForRead), Entity)
                                If IsNothing(Ent1) = False Then
                                    If TypeOf Ent1 Is Polyline Or TypeOf Ent1 Is DBPoint Then

                                        Dim Id1 As ObjectId = Ent1.ObjectId

                                        Using Records1 As Autodesk.Gis.Map.ObjectData.Records = Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                            If IsNothing(Records1) = False Then
                                                If Records1.Count > 0 Then

                                                    Dim HMM_ID As String

                                                    For Each Record1 As Autodesk.Gis.Map.ObjectData.Record In Records1
                                                        Dim Tabla1 As Autodesk.Gis.Map.ObjectData.Table
                                                        Tabla1 = Tables1(Record1.TableName)

                                                        Dim Field_defs1 As Autodesk.Gis.Map.ObjectData.FieldDefinitions
                                                        Field_defs1 = Tabla1.FieldDefinitions

                                                        For j = 0 To Record1.Count - 1
                                                            Dim Field_def1 As Autodesk.Gis.Map.ObjectData.FieldDefinition
                                                            Field_def1 = Field_defs1(j)

                                                            If Not ComboBox_HMM_ID_update.Text = "" Then
                                                                If Field_def1.Name.ToUpper = ComboBox_HMM_ID_update.Text Then
                                                                    Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                    Valoare_record1 = Record1(j)
                                                                    HMM_ID = Valoare_record1.StrValue
                                                                    If Colectie_ID.Contains(HMM_ID) = True Then
                                                                        Data_table_cu_valori.Rows.Add()
                                                                        Data_table_cu_valori.Rows(Index_data_table_valori).Item(ComboBox_HMM_ID_update.Text) = Valoare_record1.StrValue
                                                                        parcela_gasita = True
                                                                        Exit For
                                                                    End If
                                                                    If IsNumeric(HMM_ID) = True Then
                                                                        Dim hmmid_int As Integer = CInt(HMM_ID)
                                                                        Dim hmmid_dbl As Double = CDbl(HMM_ID)

                                                                        If CDbl(hmmid_int) - hmmid_dbl = 0 Then
                                                                            If Colectie_ID.Contains(hmmid_int.ToString) = True Then
                                                                                Data_table_cu_valori.Rows.Add()
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(ComboBox_HMM_ID_update.Text) = Valoare_record1.StrValue
                                                                                parcela_gasita = True
                                                                                Exit For
                                                                            End If
                                                                        End If
                                                                    End If

                                                                End If
                                                            End If
                                                        Next

                                                        If parcela_gasita = True Then

                                                            For j = 0 To Record1.Count - 1
                                                                Dim Field_def1 As Autodesk.Gis.Map.ObjectData.FieldDefinition
                                                                Field_def1 = Field_defs1(j)




                                                                If Not ComboBox_county_update.Text = "" Then
                                                                    If Field_def1.Name.ToUpper = ComboBox_county_update.Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Data_table_cu_valori.Rows(Index_data_table_valori).Item(ComboBox_county_update.Text) = Valoare_record1.StrValue
                                                                        End If
                                                                    End If
                                                                End If

                                                                If Not ComboBox_state_update.Text = "" Then
                                                                    If Field_def1.Name.ToUpper = ComboBox_state_update.Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Dim State1 As String = Valoare_record1.StrValue
                                                                            If CheckBox_state_abreviation_UPDATE.Checked = True Then
                                                                                Select Case State1.ToUpper
                                                                                    Case "PA"
                                                                                        State1 = "Pennsylvania".ToUpper
                                                                                    Case "NY"
                                                                                        State1 = "New York".ToUpper
                                                                                    Case "MA"
                                                                                        State1 = "Massachusetts".ToUpper
                                                                                    Case "NH"
                                                                                        State1 = "New Hampshire".ToUpper
                                                                                    Case "CT"
                                                                                        State1 = "Connecticut".ToUpper
                                                                                    Case "OH"
                                                                                        State1 = "OHIO".ToUpper
                                                                                End Select
                                                                            End If


                                                                            Data_table_cu_valori.Rows(Index_data_table_valori).Item(ComboBox_state_update.Text) = State1
                                                                        End If
                                                                    End If
                                                                End If

                                                                If Not ComboBox_town_update.Text = "" Then
                                                                    If Field_def1.Name.ToUpper = ComboBox_town_update.Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Data_table_cu_valori.Rows(Index_data_table_valori).Item(ComboBox_town_update.Text) = Valoare_record1.StrValue
                                                                        End If
                                                                    End If
                                                                End If
                                                                If Not ComboBox_owner_update.Text = "" Then
                                                                    If Field_def1.Name.ToUpper = ComboBox_owner_update.Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Data_table_cu_valori.Rows(Index_data_table_valori).Item(ComboBox_owner_update.Text) = Valoare_record1.StrValue
                                                                        End If
                                                                    End If
                                                                End If




                                                                If Not ComboBox_linelist_update.Text = "" Then
                                                                    If Field_def1.Name.ToUpper = ComboBox_linelist_update.Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Data_table_cu_valori.Rows(Index_data_table_valori).Item(ComboBox_linelist_update.Text) = Valoare_record1.StrValue
                                                                        End If
                                                                    End If
                                                                End If

                                                                If Not ComboBox_MBL_update.Text = "" Then
                                                                    If Field_def1.Name.ToUpper = ComboBox_MBL_update.Text Then
                                                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                        Valoare_record1 = Record1(j)
                                                                        If Not Valoare_record1.StrValue = "" Then
                                                                            Data_table_cu_valori.Rows(Index_data_table_valori).Item(ComboBox_MBL_update.Text) = Valoare_record1.StrValue
                                                                        End If
                                                                    End If
                                                                End If



                                                                With ComboBox_area_ATWS_update
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Round(CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES, 2)
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_area_EX_E_update
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Round(CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES, 2)
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_area_P_E_update
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Round(CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES, 2)
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_area_TWS_update
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Round(CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES, 2)
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_area_A_R_update
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Round(CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES, 2)
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_area_WARE_update
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Round(CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES, 2)
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With
                                                                With ComboBox_area_TWS_ABD_update
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Round(CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES, 2)
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With
                                                                With ComboBox_area_TWS_PD_update
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Round(CDbl(Valoare_record1.StrValue) * Conversie_SQFT_ACREES, 2)
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_deed_page_update
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If Not Valoare_record1.StrValue = "" Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Valoare_record1.StrValue
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_APN_update
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If Not Valoare_record1.StrValue = "" Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Valoare_record1.StrValue
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_Section_TWP_Range_update
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If Not Valoare_record1.StrValue = "" Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Valoare_record1.StrValue
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                                With ComboBox_crossing_length_update
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Round(CDbl(Valoare_record1.StrValue), 2)
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With
                                                                With ComboBox_access_road_length_update
                                                                    If Not .Text = "" Then
                                                                        If Field_def1.Name.ToUpper = .Text Then
                                                                            Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                            Valoare_record1 = Record1(j)
                                                                            If IsNumeric(Valoare_record1.StrValue) = True Then
                                                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(.Text) = Round(CDbl(Valoare_record1.StrValue), 2)
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With

                                                            Next
                                                        End If
                                                    Next

                                                    If parcela_gasita = True Then
                                                        Index_data_table_valori = Index_data_table_valori + 1
                                                        parcela_gasita = False
                                                    End If

                                                End If
                                            End If
                                        End Using
                                    End If
                                End If

                            Next
                            Trans1.Abort()
                        End Using
                    End Using

                    Dim Values_string As String = ""

                    With Data_table_sheetSet
                        If .Rows.Count > 0 Then
                            Values_string = .Columns(0).ColumnName
                            For r = 1 To .Columns.Count - 1
                                Values_string = Values_string & Chr(9) & .Columns(r).ColumnName
                            Next
                            For s = 0 To .Rows.Count - 1
                                Dim Temp1 As String = .Rows(s).Item(0)
                                For r = 1 To .Columns.Count - 1
                                    Temp1 = Temp1 & Chr(9) & .Rows(s).Item(r)
                                Next
                                Values_string = Values_string & vbCrLf & Temp1
                            Next
                        End If
                    End With

                    With Data_table_cu_valori
                        If .Rows.Count > 0 Then
                            Values_string = Values_string & vbCrLf & .Columns(0).ColumnName
                            For r = 1 To .Columns.Count - 1
                                Values_string = Values_string & Chr(9) & .Columns(r).ColumnName
                            Next
                            For s = 0 To .Rows.Count - 1
                                Dim Temp1 As String = .Rows(s).Item(0)
                                For r = 1 To .Columns.Count - 1
                                    Temp1 = Temp1 & Chr(9) & .Rows(s).Item(r)
                                Next
                                Values_string = Values_string & vbCrLf & Temp1
                            Next
                        End If
                    End With


                    Dim Comparation As String = ""
                    Dim Not_equal As Integer = 0
                    For i = 0 To Data_table_sheetSet.Rows.Count - 1
                        Dim SheetName As String = Data_table_sheetSet.Rows(i).Item("SHEET_NAME")

                        If i = 0 Then
                            Comparation = SheetName
                        Else

                            Comparation = Comparation & vbCrLf & " " & vbCrLf & SheetName
                        End If


                        Dim HmmID As String = Data_table_sheetSet.Rows(i).Item("HMMID")
                        Dim j As Integer = 0
                        Dim Is_ID_matched As Boolean = False
                        For j = 0 To Data_table_cu_valori.Rows.Count - 1
                            If HmmID = Data_table_cu_valori.Rows(j).Item(ComboBox_HMM_ID_update.Text) Then
                                Is_ID_matched = True
                                Exit For
                            End If
                        Next



                        For s = 0 To Data_table_sheetSet.Columns.Count - 1
                            Dim Nume_coloana_Sset As String = Data_table_sheetSet.Columns(s).ColumnName
                            Dim Nume_coloana_Dtable As String = ""

                            Select Case Nume_coloana_Sset.ToUpper
                                Case ComboBox_HMM_ID_update.Text
                                    Nume_coloana_Dtable = ComboBox_HMM_ID_update.Text
                                Case ComboBox_MBL_update.Text
                                    Nume_coloana_Dtable = ComboBox_MBL_update.Text
                                Case "SGC_OWNER"

                                    Nume_coloana_Dtable = ComboBox_owner_update.Text

                                Case ComboBox_owner_update.Text
                                    Nume_coloana_Dtable = ComboBox_owner_update.Text
                                Case ComboBox_state_update.Text
                                    Nume_coloana_Dtable = ComboBox_state_update.Text
                                Case ComboBox_county_update.Text
                                    Nume_coloana_Dtable = ComboBox_county_update.Text
                                Case ComboBox_town_update.Text
                                    Nume_coloana_Dtable = ComboBox_town_update.Text
                                Case ComboBox_linelist_update.Text
                                    Nume_coloana_Dtable = ComboBox_linelist_update.Text
                                Case ComboBox_area_EX_E_update.Text
                                    Nume_coloana_Dtable = ComboBox_area_EX_E_update.Text
                                Case ComboBox_area_P_E_update.Text
                                    Nume_coloana_Dtable = ComboBox_area_P_E_update.Text
                                Case ComboBox_area_TWS_update.Text
                                    Nume_coloana_Dtable = ComboBox_area_TWS_update.Text
                                Case ComboBox_area_A_R_update.Text
                                    Nume_coloana_Dtable = ComboBox_area_A_R_update.Text
                                Case ComboBox_area_WARE_update.Text
                                    Nume_coloana_Dtable = ComboBox_area_WARE_update.Text
                                Case ComboBox_deed_page_update.Text
                                    Nume_coloana_Dtable = ComboBox_deed_page_update.Text
                                Case ComboBox_APN_update.Text
                                    Nume_coloana_Dtable = ComboBox_APN_update.Text
                                Case ComboBox_crossing_length_update.Text
                                    Nume_coloana_Dtable = ComboBox_crossing_length_update.Text
                                Case ComboBox_Section_TWP_Range_update.Text
                                    Nume_coloana_Dtable = ComboBox_Section_TWP_Range_update.Text
                                Case ComboBox_area_TWS_ABD_update.Text
                                    Nume_coloana_Dtable = ComboBox_area_TWS_ABD_update.Text
                                Case ComboBox_area_TWS_PD_update.Text
                                    Nume_coloana_Dtable = ComboBox_area_TWS_PD_update.Text
                                Case ComboBox_access_road_length_update.Text
                                    Nume_coloana_Dtable = ComboBox_access_road_length_update.Text
                                Case ComboBox_MBL_update.Text
                                    Nume_coloana_Dtable = ComboBox_MBL_update.Text
                                Case ComboBox_area_ATWS_update.Text
                                    Nume_coloana_Dtable = ComboBox_area_ATWS_update.Text
                            End Select




                            If Not Nume_coloana_Dtable = "" Then
                                Dim Equal1 As String = "YES"
                                Dim Val2 As String = ""

                                Comparation = Comparation & vbCrLf & Nume_coloana_Sset

                                If IsDBNull(Data_table_sheetSet.Rows(i).Item(s)) = False Then
                                    Comparation = Comparation & Chr(9) & Data_table_sheetSet.Rows(i).Item(s)
                                    Val2 = Data_table_sheetSet.Rows(i).Item(s)
                                Else
                                    Comparation = Comparation & Chr(9) & " "
                                End If



                                If Is_ID_matched = True Then
                                    If Data_table_cu_valori.Columns.Contains(Nume_coloana_Dtable) = True Then
                                        If Not Nume_coloana_Dtable = ComboBox_owner_update.Text Then
                                            If IsDBNull(Data_table_cu_valori.Rows(j).Item(Nume_coloana_Dtable)) = False Then
                                                Dim Val1 As String = Data_table_cu_valori.Rows(j).Item(Nume_coloana_Dtable)

                                                If IsNumeric(Val1) = True And IsNumeric(Val2) = True Then
                                                    If Not CDbl(Val1) = CDbl(Val2) Then
                                                        Equal1 = "NO"
                                                        Not_equal = Not_equal + 1
                                                    End If
                                                Else

                                                    If Not Val1.ToUpper = Val2.ToUpper Then
                                                        Equal1 = "NO"
                                                        Not_equal = Not_equal + 1
                                                    End If
                                                End If


                                                Comparation = Comparation & Chr(9) & Val1 & Chr(9) & Equal1
                                            End If
                                        Else
                                            Dim Owner1 As String = ""
                                            If IsDBNull(Data_table_cu_valori.Rows(j).Item(Nume_coloana_Dtable)) = False Then
                                                Owner1 = Data_table_cu_valori.Rows(j).Item(Nume_coloana_Dtable)
                                            End If


                                            If Not Owner1.ToUpper = Val2.ToUpper Then
                                                Equal1 = "NO"
                                                Not_equal = Not_equal + 1
                                            End If
                                            Comparation = Comparation & Chr(9) & Owner1 & Chr(9) & Equal1
                                        End If

                                    End If
                                Else
                                    Comparation = Comparation & Chr(9) & "NO"
                                    Not_equal = Not_equal + 1
                                End If

                            End If

                        Next


                    Next

                    My.Computer.Clipboard.SetText(Comparation)
                    Dim DISCR1 As String = " DISCREPANCES"
                    If Not_equal = 1 Then
                        DISCR1 = " DISCREPANCE"
                    End If

                    MsgBox("Data has been copied to the clipboard" & vbCrLf & "(" & (Data_table_sheetSet.Rows.Count).ToString & " rows)" & vbCrLf & Not_equal.ToString & DISCR1)

                Else
                    MsgBox(SheetSet_database.GetLockStatus.ToString)
                    ' Display error message
                    MsgBox("Sheet set could not be opened for write.")
                End If
                Button_compare_DST.Visible = True
            Catch ex As Exception
                Button_compare_DST.Visible = True
                MsgBox(ex.Message)
            End Try
            Button_compare_DST.Visible = True
            Freeze_operations = False
        End If
    End Sub

    Private Sub TabPage_neighbors_Click(sender As Object, e As EventArgs) Handles TabPage_neighbors.Click
        Incarca_existing_Blocks_with_attributes_to_combobox(ComboBox_blocks)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks.Text, ComboBox_bl_atr1)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks.Text, ComboBox_bl_atr2)
        If ComboBox_bl_atr1.Items.Count > 1 Then
            ComboBox_bl_atr1.SelectedIndex = 1
        End If
        If ComboBox_bl_atr2.Items.Count > 2 Then
            ComboBox_bl_atr2.SelectedIndex = 2
        End If
    End Sub

    Private Sub ComboBox_blocks_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_blocks.SelectedIndexChanged
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks.Text, ComboBox_bl_atr1)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks.Text, ComboBox_bl_atr2)
        If ComboBox_bl_atr1.Items.Count > 1 Then
            ComboBox_bl_atr1.SelectedIndex = 1
        End If
        If ComboBox_bl_atr2.Items.Count > 2 Then
            ComboBox_bl_atr2.SelectedIndex = 2
        End If
    End Sub

    Private Sub Button_load_OD_vicinity_labels_Click(sender As Object, e As EventArgs) Handles Button_load_OD_vicinity_labels.Click

        Dim Empty_array() As ObjectId
        If Freeze_operations = False Then
            Freeze_operations = True


            Dim BaseMap_drawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = BaseMap_drawing.Editor
            Editor1.SetImpliedSelection(Empty_array)
            Try
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()



                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptEntityResult

                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select an entity containing object data:")

                Rezultat1 = Editor1.GetEntity(Object_Prompt)

                If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.WriteMessage(vbLf & "Command:")
                    Editor1.SetImpliedSelection(Empty_array)
                    Freeze_operations = False
                    Exit Sub
                End If

                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                    If IsNothing(Rezultat1) = False Then



                        Using Lock_dwg As DocumentLock = BaseMap_drawing.LockDocument
                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = BaseMap_drawing.TransactionManager.StartTransaction
                                Dim Ent1 As Entity = Trans1.GetObject(Rezultat1.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                                Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables
                                Dim Id1 As ObjectId = Ent1.ObjectId

                                Using Records1 As Autodesk.Gis.Map.ObjectData.Records = Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                    If IsNothing(Records1) = False Then
                                        If Records1.Count > 0 Then
                                            Dim Record1 As Autodesk.Gis.Map.ObjectData.Record
                                            With ComboBox_OD1
                                                If .Items.Count = 0 Then
                                                    .Items.Add("")
                                                End If
                                            End With
                                            With ComboBox_OD2
                                                If .Items.Count = 0 Then
                                                    .Items.Add("")
                                                End If
                                            End With
                                            With ComboBox_OD3
                                                If .Items.Count = 0 Then
                                                    .Items.Add("")
                                                End If
                                            End With
                                            With ComboBox_OD4
                                                If .Items.Count = 0 Then
                                                    .Items.Add("")
                                                End If
                                            End With
                                            With ComboBox_OD5
                                                If .Items.Count = 0 Then
                                                    .Items.Add("")
                                                End If
                                            End With
                                            For Each Record1 In Records1
                                                Dim Tabla1 As Autodesk.Gis.Map.ObjectData.Table
                                                Tabla1 = Tables1(Record1.TableName)
                                                With ComboBox_OD_Table_name
                                                    If .Items.Contains(Tabla1.Name.ToUpper) = False Then
                                                        .Items.Add(Tabla1.Name.ToUpper)
                                                    End If
                                                End With

                                                Dim Field_defs1 As Autodesk.Gis.Map.ObjectData.FieldDefinitions
                                                Field_defs1 = Tabla1.FieldDefinitions



                                                For i = 0 To Record1.Count - 1
                                                    Dim Field_def1 As Autodesk.Gis.Map.ObjectData.FieldDefinition
                                                    Field_def1 = Field_defs1(i)
                                                    With ComboBox_OD1
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With
                                                    With ComboBox_OD2
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With
                                                    With ComboBox_OD3
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With
                                                    With ComboBox_OD4
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With
                                                    With ComboBox_OD5
                                                        If .Items.Contains(Field_def1.Name.ToUpper) = False Then
                                                            .Items.Add(Field_def1.Name.ToUpper)
                                                        End If
                                                    End With


                                                Next



                                            Next



                                        End If
                                    End If
                                End Using






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

    Private Sub Button_layers_to_excel_Click(sender As Object, e As EventArgs) Handles Button_layers_to_excel.Click
        Dim Empty_array() As ObjectId
        If Freeze_operations = False Then
            Freeze_operations = True

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Editor1.SetImpliedSelection(Empty_array)
            Try

                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                Using Lock_dwg As DocumentLock = ThisDrawing.LockDocument
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(BlockTable1(BlockTableRecord.PaperSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                        Dim PromptEnt1 As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions("Select the viewport (or polyline for clipped viewports): ")

                        PromptEnt1.SetRejectMessage("Sorry, Not a viewport")

                        PromptEnt1.AddAllowedClass(GetType(Viewport), True)
                        PromptEnt1.AddAllowedClass(GetType(Polyline), True)

                        Dim Rezultat_viewport As Autodesk.AutoCAD.EditorInput.PromptEntityResult = Editor1.GetEntity(PromptEnt1)
                        Dim Colectie_frozen As New Specialized.StringCollection

                        Dim Este_viewport As Boolean = False

                        If Rezultat_viewport.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                            Dim vpId As ObjectId
                            If TypeOf Trans1.GetObject(Rezultat_viewport.ObjectId, OpenMode.ForRead) Is Polyline Then
                                vpId = LayoutManager.Current.GetNonRectangularViewportIdFromClipId(Rezultat_viewport.ObjectId)
                            Else
                                vpId = Rezultat_viewport.ObjectId
                            End If

                            If Not vpId = Nothing Then
                                Dim vp As Viewport = Trans1.GetObject(vpId, OpenMode.ForWrite)
                                Este_viewport = True
                                Dim resBuf As ResultBuffer
                                resBuf = vp.XData()

                                If resBuf Is Nothing Then
                                Else
                                    For Each tv As TypedValue In resBuf
                                        Dim typeCode As Short
                                        typeCode = tv.TypeCode

                                        If typeCode = 1003 Then
                                            Colectie_frozen.Add(tv.Value.ToString)

                                        End If
                                    Next
                                End If
                            End If

                        End If



                        Dim Layer_table As Autodesk.AutoCAD.DatabaseServices.LayerTable = Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                        W1 = Get_NEW_worksheet_from_Excel()
                        W1.Range("C" & 1).Value = "NAME"
                        W1.Range("D" & 1).Value = "COLOR INDEX"
                        W1.Range("E" & 1).Value = "SELECTED VIEWPORT THAW-FREEZE"

                        W1.Range("F" & 1).Value = "ON-OFF"
                        W1.Range("G" & 1).Value = "THAW-FREEZE"


                        W1.Range("H" & 1).Value = "PLOTABLE"
                        W1.Range("I" & 1).Value = "LINETYPE"

                        Dim Row1 As Integer = 2
                        For Each ID1 As ObjectId In Layer_table
                            Dim LayerTRec1 As LayerTableRecord = Trans1.GetObject(ID1, OpenMode.ForRead)
                            W1.Range("C" & Row1).Value = LayerTRec1.Name
                            W1.Range("D" & Row1).Value = LayerTRec1.Color.ColorIndex

                            If Colectie_frozen.Count > 0 Then
                                If Colectie_frozen.Contains(LayerTRec1.Name) = True Then
                                    W1.Range("E" & Row1).Value2 = "FROZEN"
                                Else
                                    W1.Range("E" & Row1).Value2 = "THAW"
                                End If
                            Else
                                If Este_viewport = True Then
                                    W1.Range("E" & Row1).Value2 = "THAW"
                                End If
                            End If

                            If LayerTRec1.IsOff = True Then
                                W1.Range("F" & Row1).Value = "OFF"
                            Else
                                W1.Range("F" & Row1).Value = "ON"
                            End If
                            If LayerTRec1.IsFrozen = True Then
                                W1.Range("G" & Row1).Value = "FROZEN"
                            Else
                                W1.Range("G" & Row1).Value = "THAW"
                            End If
                            If LayerTRec1.IsPlottable = True Then
                                W1.Range("H" & Row1).Value = "PLOTABLE"
                            Else
                                W1.Range("H" & Row1).Value = "NON PLOTABLE"
                            End If
                            Dim Ltype As LinetypeTableRecord = Trans1.GetObject(LayerTRec1.LinetypeObjectId, OpenMode.ForRead)
                            W1.Range("I" & Row1).Value = Ltype.Name

                            Row1 = Row1 + 1
                        Next
                        Trans1.Commit()
                    End Using
                End Using



                '
                '
                'W1.Cells(s + 2, Data_table1.Columns.Count + 1).Formula = "=A" & (s + 2) & "&" & Chr(34) & ":" & Chr(34) & "&" & "B" & (s + 2) & "&" & Chr(34) & "," & Chr(34) & "&" & "C" & (s + 2)
                '

                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")

            Catch ex As Exception

                Editor1.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If


    End Sub

    Private Sub Button_read_settings_old(sender As Object, e As EventArgs)
        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                If System.IO.File.Exists(Settings_file) = True Then
                    Using Reader1 As New System.IO.StreamReader(Settings_file)
                        Dim Line1 As String


                        While Reader1.Peek > 0
                            Line1 = Reader1.ReadLine
                            If Line1.Contains(":") = True And Line1.Contains("|") = True Then
                                Dim Line_no As String = Line1.Split(":")(0)
                                Dim Value1 As String = Line1.Split("|")(1)
                                Select Case Line_no
                                    Case 1
                                        TextBox_xref_model_space.Text = Value1
                                    Case 2
                                        TextBox_dwt_template.Text = Value1
                                    Case 3
                                        TextBox_sheet_set_template.Text = Value1
                                    Case 4
                                        TextBox_Output_Directory.Text = Value1
                                    Case 5
                                        TextBox_north_arrow_Big_X.Text = Value1
                                    Case 6
                                        TextBox_north_arrow_Big_y.Text = Value1
                                    Case 7
                                        TextBox_main_viewport_height.Text = Value1
                                    Case 8
                                        TextBox_main_viewport_width.Text = Value1
                                    Case 9
                                        TextBox_main_viewport_center_X.Text = Value1
                                    Case 10
                                        TextBox_main_viewport_center_Y.Text = Value1
                                    Case 11
                                        TextBox_north_arrow.Text = Value1

                                    Case 12
                                        TextBox_north_arrow_small_X.Text = Value1
                                    Case 13
                                        TextBox_north_arrow_small_Y.Text = Value1
                                    Case 14
                                        TextBox_locus_viewport_height.Text = Value1
                                    Case 15
                                        TextBox_locus_viewport_width.Text = Value1
                                    Case 16
                                        TextBox_locus_viewport_center_X.Text = Value1
                                    Case 17
                                        TextBox_locus_viewport_center_Y.Text = Value1
                                    Case 18
                                        TextBox_north_arrow.Text = Value1
                                    Case 19
                                        With ComboBox_state
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 20
                                        With ComboBox_county
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 21
                                        With ComboBox_town
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 22
                                        With ComboBox_linelist
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 23
                                        With ComboBox_HMM_ID
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 24
                                        With ComboBox_MBL
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 25
                                        With ComboBox_owner
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With

                                    Case 26

                                    Case 27

                                    Case 28
                                        With ComboBox_area_EX_E
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 29
                                        With ComboBox_area_P_E
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 30
                                        With ComboBox_area_TWS
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 31
                                        With ComboBox_area_ATWS
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 32
                                        With ComboBox_area_A_R
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 33
                                        With ComboBox_area_WARE
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 34
                                        With ComboBox_pipe_diameter
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 35
                                        With ComboBox_pipe_name
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 36
                                        With ComboBox_pipe_segment
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 37
                                        With ComboBox_state_update
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 38
                                        With ComboBox_county_update
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 39
                                        With ComboBox_town_update
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 40
                                        With ComboBox_linelist_update
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 41
                                        With ComboBox_HMM_ID_update
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 42
                                        With ComboBox_MBL_update
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 43
                                        With ComboBox_owner_update
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With


                                    Case 44

                                    Case 45


                                    Case 46
                                        With ComboBox_area_EX_E_update
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 47
                                        With ComboBox_area_P_E_update
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 48
                                        With ComboBox_area_TWS_update
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 49
                                        With ComboBox_area_ATWS_update
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 50
                                        With ComboBox_area_A_R_update
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 51
                                        With ComboBox_area_WARE_update
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 52
                                        With ComboBox_plat_name
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With


                                    Case 53
                                        If Value1 = 1 Then
                                            CheckBox_convert_sqft_to_acres.Checked = True
                                        Else
                                            CheckBox_convert_sqft_to_acres.Checked = False
                                        End If

                                    Case 54
                                        If Value1 = 1 Then
                                            CheckBox_ignore_deflections_less_0_5.Checked = True
                                        Else
                                            CheckBox_ignore_deflections_less_0_5.Checked = False
                                        End If

                                    Case 55
                                        If Value1 = 1 Then
                                            CheckBox_add_bearing_and_distances.Checked = True
                                        Else
                                            CheckBox_add_bearing_and_distances.Checked = False
                                        End If

                                    Case 56
                                        If Value1 = 1 Then
                                            CheckBox_display_stations.Checked = True
                                        Else
                                            CheckBox_display_stations.Checked = False
                                        End If

                                    Case 57
                                        If Value1 = 1 Then
                                            CheckBox_rotate_to_north.Checked = True
                                        Else
                                            CheckBox_rotate_to_north.Checked = False
                                        End If
                                    Case 58
                                        Station_at_point_layer = Value1
                                    Case 59
                                        Centerline_stationing_layer = Value1
                                    Case 60
                                        Bearing_and_distance_layer = Value1
                                    Case 61
                                        Tie_distance_layer = Value1
                                    Case 62
                                        Arc_leader_polyline_layer = Value1
                                    Case 63
                                        Layer_name_Main_Viewport = Value1
                                    Case 64
                                        Layer_name_locus_VP = Value1
                                    Case 65
                                        Layer_name_Blocks = Value1
                                    Case 66
                                        Layer_name_text = Value1
                                    Case 67
                                        Layer_poly_parcela_new = Value1
                                    Case 68
                                        Layer_hatch_locus = Value1
                                    Case 69
                                        With ComboBox_deed_page
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 70
                                        With ComboBox_APN
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 71
                                        With ComboBox_crossing_length
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 72
                                        With ComboBox_Section_TWP_Range
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 73
                                        With ComboBox_area_TWS_PD
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 74
                                        With ComboBox_area_TWS_ABD
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 75
                                        With ComboBox_access_road_length
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 76

                                    Case 77
                                        With ComboBox_deed_page_update
                                            .Items.Add(Value1)
                                            .SelectedIndex = 1
                                        End With
                                    Case 78
                                        With ComboBox_APN_update
                                            .Items.Add(Value1)
                                            .SelectedIndex = 1
                                        End With
                                    Case 79
                                        With ComboBox_crossing_length_update
                                            .Items.Add(Value1)
                                            .SelectedIndex = 1
                                        End With
                                    Case 80
                                        With ComboBox_Section_TWP_Range_update
                                            .Items.Add(Value1)
                                            .SelectedIndex = 1
                                        End With
                                    Case 81
                                        With ComboBox_area_TWS_PD_update
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 82
                                        With ComboBox_area_TWS_ABD_update
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                    Case 83
                                        With ComboBox_access_road_length_update
                                            If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                            .SelectedIndex = .Items.IndexOf(Value1)
                                        End With
                                End Select
                            End If
                        End While

                    End Using

                Else
                    MsgBox("There is no file on the desktop - no settings have been loaded")

                End If



            Catch ex As System.Exception
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_create_settings_Click_old(sender As Object, e As EventArgs)
        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                If System.IO.File.Exists(Settings_file) = True Then
                    IO.File.Delete(Settings_file)
                End If
                Using Fs As IO.FileStream = IO.File.Create(Settings_file)

                End Using


                Using sw As IO.StreamWriter = New IO.StreamWriter(Settings_file)
                    sw.Write("1: XREF (Basemap)|" & TextBox_xref_model_space.Text)
                    sw.Write(vbCrLf & "2: DWT template|" & TextBox_dwt_template.Text)
                    sw.Write(vbCrLf & "3: DST sheet set manager template|" & TextBox_sheet_set_template.Text)
                    sw.Write(vbCrLf & "4: Output Folder|" & TextBox_Output_Directory.Text)
                    sw.Write(vbCrLf & "5: North Arrow X Coordinate (Main Viewport)|" & TextBox_north_arrow_Big_X.Text)
                    sw.Write(vbCrLf & "6: North Arrow Y Coordinate (Main Viewport)|" & TextBox_north_arrow_Big_y.Text)
                    sw.Write(vbCrLf & "7: Main Viewport Height|" & TextBox_main_viewport_height.Text)
                    sw.Write(vbCrLf & "8: Main Viewport Width|" & TextBox_main_viewport_width.Text)
                    sw.Write(vbCrLf & "9: Main Viewport Center X Coordinate|" & TextBox_main_viewport_center_X.Text)
                    sw.Write(vbCrLf & "10: Main Viewport Center Y Coordinate|" & TextBox_main_viewport_center_Y.Text)
                    sw.Write(vbCrLf & "11: North Arrow Autocad Block name (Main Viewport)|" & TextBox_north_arrow.Text)
                    If CheckBox_add_locus.Checked = True Then
                        sw.Write(vbCrLf & "12: North Arrow X Coordinate (Locus Viewport)|" & TextBox_north_arrow_small_X.Text)
                        sw.Write(vbCrLf & "13: North Arrow Y Coordinate (Locus Viewport)|" & TextBox_north_arrow_small_Y.Text)
                        sw.Write(vbCrLf & "14: Locus Viewport Height|" & TextBox_locus_viewport_height.Text)
                        sw.Write(vbCrLf & "15: Locus Viewport Width|" & TextBox_locus_viewport_width.Text)
                        sw.Write(vbCrLf & "16: Locus Viewport Center X Coordinate|" & TextBox_locus_viewport_center_X.Text)
                        sw.Write(vbCrLf & "17: Locus Viewport Center Y Coordinate|" & TextBox_locus_viewport_center_Y.Text)
                        sw.Write(vbCrLf & "18: North Arrow Autocad Block name (Locus Viewport)|" & TextBox_north_arrow.Text)
                    End If
                    If Not ComboBox_state.Text = "" Then sw.Write(vbCrLf & "19: State (Object Data)|" & ComboBox_state.Text)
                    If Not ComboBox_county.Text = "" Then sw.Write(vbCrLf & "20: County (Object Data)|" & ComboBox_county.Text)
                    If Not ComboBox_town.Text = "" Then sw.Write(vbCrLf & "21: Town (Object Data)|" & ComboBox_town.Text)
                    If Not ComboBox_linelist.Text = "" Then sw.Write(vbCrLf & "22: Linelist (Object Data)|" & ComboBox_linelist.Text)
                    If Not ComboBox_HMM_ID.Text = "" Then sw.Write(vbCrLf & "23: Parcel ID (Object Data)|" & ComboBox_HMM_ID.Text)
                    If Not ComboBox_MBL.Text = "" Then sw.Write(vbCrLf & "24: MBL (Object Data)|" & ComboBox_MBL.Text)
                    If Not ComboBox_owner.Text = "" Then sw.Write(vbCrLf & "25: Owner on 1 Line (Object Data)|" & ComboBox_owner.Text)

                    'If Not ComboBox_owner2_1.Text = "" Then sw.Write(vbCrLf & "26: Owner on 2 Lines - first line (Object Data)|" & ComboBox_owner2_1.Text)
                    'If Not ComboBox_owner2_2.Text = "" Then sw.Write(vbCrLf & "27: Owner on 2 Lines - second line (Object Data)|" & ComboBox_owner2_2.Text)



                    If Not ComboBox_area_EX_E.Text = "" Then sw.Write(vbCrLf & "28: AREA Existing Easement (Object Data)|" & ComboBox_area_EX_E.Text)
                    If Not ComboBox_area_P_E.Text = "" Then sw.Write(vbCrLf & "29: AREA Permanent Easement (Object Data)|" & ComboBox_area_P_E.Text)
                    If Not ComboBox_area_TWS.Text = "" Then sw.Write(vbCrLf & "30: AREA Temporary Work Space (Object Data)|" & ComboBox_area_TWS.Text)
                    If Not ComboBox_area_ATWS.Text = "" Then sw.Write(vbCrLf & "31: AREA Additional Temporary Work Space (Object Data)|" & ComboBox_area_ATWS.Text)
                    If Not ComboBox_area_A_R.Text = "" Then sw.Write(vbCrLf & "32: AREA Access Road (Object Data)|" & ComboBox_area_A_R.Text)
                    If Not ComboBox_area_WARE.Text = "" Then sw.Write(vbCrLf & "33: AREA Wareyard (Object Data)|" & ComboBox_area_WARE.Text)
                    If Not ComboBox_pipe_diameter.Text = "" Then sw.Write(vbCrLf & "34: PIPE DIAMETER (Object Data)|" & ComboBox_pipe_diameter.Text)
                    If Not ComboBox_pipe_name.Text = "" Then sw.Write(vbCrLf & "35: NAME OF THE PROJECT (Object Data)|" & ComboBox_pipe_name.Text)
                    If Not ComboBox_pipe_segment.Text = "" Then sw.Write(vbCrLf & "36: SEGMENT (Object Data)|" & ComboBox_pipe_segment.Text)

                    If Not ComboBox_state_update.Text = "" Then sw.Write(vbCrLf & "37: State <UPDATE> (Object Data)|" & ComboBox_state_update.Text)
                    If Not ComboBox_county_update.Text = "" Then sw.Write(vbCrLf & "38: County <UPDATE> (Object Data)|" & ComboBox_county_update.Text)
                    If Not ComboBox_town_update.Text = "" Then sw.Write(vbCrLf & "39: Town <UPDATE> (Object Data)|" & ComboBox_town_update.Text)
                    If Not ComboBox_linelist_update.Text = "" Then sw.Write(vbCrLf & "40: Linelist <UPDATE> (Object Data)|" & ComboBox_linelist_update.Text)
                    If Not ComboBox_HMM_ID_update.Text = "" Then sw.Write(vbCrLf & "41: Parcel ID <UPDATE> (Object Data)|" & ComboBox_HMM_ID_update.Text)
                    If Not ComboBox_MBL_update.Text = "" Then sw.Write(vbCrLf & "42: MBL <UPDATE> (Object Data)|" & ComboBox_MBL_update.Text)
                    If Not ComboBox_owner_update.Text = "" Then sw.Write(vbCrLf & "43: Owner on 1 Line <UPDATE> (Object Data)|" & ComboBox_owner_update.Text)

                    'If Not ComboBox_owner2_1_update.Text = "" Then sw.Write(vbCrLf & "44: Owner on 2 Lines - first line <UPDATE> (Object Data)|" & ComboBox_owner2_1_update.Text)
                    'If Not ComboBox_owner2_2_update.Text = "" Then sw.Write(vbCrLf & "45: Owner on 2 Lines - second line <UPDATE> (Object Data)|" & ComboBox_owner2_2_update.Text)



                    If Not ComboBox_area_EX_E_update.Text = "" Then sw.Write(vbCrLf & "46: AREA Existing Easement <UPDATE> (Object Data)|" & ComboBox_area_EX_E_update.Text)
                    If Not ComboBox_area_P_E_update.Text = "" Then sw.Write(vbCrLf & "47: AREA Permanent Easement <UPDATE> (Object Data)|" & ComboBox_area_P_E_update.Text)
                    If Not ComboBox_area_TWS_update.Text = "" Then sw.Write(vbCrLf & "48: AREA Temporary Work Space <UPDATE> (Object Data)|" & ComboBox_area_TWS_update.Text)
                    If Not ComboBox_area_ATWS_update.Text = "" Then sw.Write(vbCrLf & "49: AREA Additional Temporary Work Space <UPDATE> (Object Data)|" & ComboBox_area_ATWS_update.Text)
                    If Not ComboBox_area_A_R_update.Text = "" Then sw.Write(vbCrLf & "50: AREA Access Road <UPDATE> (Object Data)|" & ComboBox_area_A_R_update.Text)
                    If Not ComboBox_area_WARE_update.Text = "" Then sw.Write(vbCrLf & "51: AREA Wareyard <UPDATE> (Object Data)|" & ComboBox_area_WARE_update.Text)
                    If Not ComboBox_plat_name.Text = "" Then sw.Write(vbCrLf & "52: PLAT NAME (Object Data)|" & ComboBox_plat_name.Text)

                    If CheckBox_convert_sqft_to_acres.Checked = False Then
                        sw.Write(vbCrLf & "53: Convert SQFT to Acres|0")
                    Else
                        sw.Write(vbCrLf & "53: Convert SQFT to Acres|1")
                    End If
                    If CheckBox_ignore_deflections_less_0_5.Checked = False Then
                        sw.Write(vbCrLf & "54: Ignore points with deflection less than 0.5|0")
                    Else
                        sw.Write(vbCrLf & "54: Ignore points with deflection less than 0.5|1")
                    End If

                    If CheckBox_add_bearing_and_distances.Checked = False Then
                        sw.Write(vbCrLf & "55: Add Bearing and distances along CL|0")
                    Else
                        sw.Write(vbCrLf & "55: Add Bearing and distances along CL|1")
                    End If

                    If CheckBox_display_stations.Checked = False Then
                        sw.Write(vbCrLf & "56: Display stations along CL|0")
                    Else
                        sw.Write(vbCrLf & "56: Display stations along CL|1")
                    End If

                    If CheckBox_rotate_to_north.Checked = False Then
                        sw.Write(vbCrLf & "57: Set View to North|0")
                    Else
                        sw.Write(vbCrLf & "57: Set View to North|1")
                    End If

                    If Not Station_at_point_layer = "" Then
                        sw.Write(vbCrLf & "58: Station at Point Layer|" & Station_at_point_layer)
                    End If

                    If Not Centerline_stationing_layer = "" Then
                        sw.Write(vbCrLf & "59: Centerline Stationing Layer|" & Centerline_stationing_layer)
                    End If

                    If Not Bearing_and_distance_layer = "" Then
                        sw.Write(vbCrLf & "60: Bearing and Distance Layer|" & Bearing_and_distance_layer)
                    End If
                    If Not Tie_distance_layer = "" Then
                        sw.Write(vbCrLf & "61: Tie Distance Layer|" & Tie_distance_layer)
                    End If
                    If Not Arc_leader_polyline_layer = "" Then
                        sw.Write(vbCrLf & "62: Arc Leader Polyline Layer|" & Arc_leader_polyline_layer)
                    End If
                    If Not Layer_name_Main_Viewport = "" Then
                        sw.Write(vbCrLf & "63: Layer name Main Viewport|" & Layer_name_Main_Viewport)
                    End If

                    If Not Layer_name_locus_VP = "" Then
                        sw.Write(vbCrLf & "64: Layer name Locus Viewport|" & Layer_name_locus_VP)
                    End If

                    If Not Layer_name_Blocks = "" Then
                        sw.Write(vbCrLf & "65: Layer name for Blocks|" & Layer_name_Blocks)
                    End If

                    If Not Layer_name_text = "" Then
                        sw.Write(vbCrLf & "66: Layer name for Text|" & Layer_name_text)
                    End If

                    If Not Layer_poly_parcela_new = "" Then
                        sw.Write(vbCrLf & "67: Layer name for new parcel|" & Layer_poly_parcela_new)
                    End If

                    If Not Layer_hatch_locus = "" Then
                        sw.Write(vbCrLf & "68: Layer name for Locus hatch|" & Layer_hatch_locus)
                    End If

                    If Not ComboBox_deed_page.Text = "" Then sw.Write(vbCrLf & "69: DEED BOOK PAGE (Object Data)|" & ComboBox_deed_page.Text)
                    If Not ComboBox_APN.Text = "" Then sw.Write(vbCrLf & "70: APN (Object Data)|" & ComboBox_APN.Text)
                    If Not ComboBox_crossing_length.Text = "" Then sw.Write(vbCrLf & "71: CL CROSSING LENGTH (Object Data)|" & ComboBox_crossing_length.Text)
                    If Not ComboBox_Section_TWP_Range.Text = "" Then sw.Write(vbCrLf & "72: SECTION - TWP - RANGE (Object Data)|" & ComboBox_Section_TWP_Range.Text)
                    If Not ComboBox_area_TWS_PD.Text = "" Then sw.Write(vbCrLf & "73: Temporary Workspace Previously Disturbed Square Footage (Object Data)|" & ComboBox_area_TWS_PD.Text)
                    If Not ComboBox_area_TWS_ABD.Text = "" Then sw.Write(vbCrLf & "74:  Temporary Workspace Abandoned Square Footage (Object Data)|" & ComboBox_area_TWS_ABD.Text)
                    If Not ComboBox_access_road_length.Text = "" Then sw.Write(vbCrLf & "75:  Access road length (Object Data)|" & ComboBox_access_road_length.Text)
                    'If Not ComboBox_plat_name_update.Text = "" Then sw.Write(vbCrLf & "76: PLAT NAME - update (Object Data)|" & ComboBox_plat_name_update.Text)
                    With ComboBox_deed_page_update
                        If Not .Text = "" Then sw.Write(vbCrLf & "77: DEED BOOK PAGE - update (Object Data)|" & .Text)
                    End With
                    With ComboBox_APN_update
                        If Not .Text = "" Then sw.Write(vbCrLf & "78: APN - update (Object Data)|" & .Text)
                    End With
                    With ComboBox_crossing_length_update
                        If Not .Text = "" Then sw.Write(vbCrLf & "79:  CL CROSSING LENGTH - update (Object Data)|" & .Text)
                    End With
                    With ComboBox_Section_TWP_Range_update
                        If Not .Text = "" Then sw.Write(vbCrLf & "80: SECTION - TWP - RANGE - update (Object Data)|" & .Text)
                    End With
                    With ComboBox_area_TWS_PD_update
                        If Not .Text = "" Then sw.Write(vbCrLf & "81: Temporary Workspace Previously Disturbed Square Footage - update (Object Data)|" & .Text)
                    End With
                    With ComboBox_area_TWS_ABD_update
                        If Not .Text = "" Then sw.Write(vbCrLf & "82: Temporary Workspace Abandoned Square Footage - update (Object Data)|" & .Text)
                    End With
                    With ComboBox_access_road_length_update
                        If Not .Text = "" Then sw.Write(vbCrLf & "83: Access road length - update (Object Data)|" & .Text)
                    End With

                End Using

            Catch ex As System.Exception
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub

    Private Sub Platt_Generator_form_Load_old(sender As Object, e As EventArgs)
        HScrollBar_rotate.Minimum = -180
        HScrollBar_rotate.Maximum = 180
        HScrollBar_rotate.Value = 0
        Settings_file = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\plat_settings.csv"

        ComboBox_plat_name.Items.Add("")
        ComboBox_state.Items.Add("")
        ComboBox_county.Items.Add("")
        ComboBox_town.Items.Add("")
        ComboBox_linelist.Items.Add("")
        ComboBox_HMM_ID.Items.Add("")
        ComboBox_MBL.Items.Add("")
        ComboBox_deed_page.Items.Add("")
        ComboBox_APN.Items.Add("")
        ComboBox_access_road_length.Items.Add("")
        ComboBox_Section_TWP_Range.Items.Add("")
        ComboBox_owner.Items.Add("")
        ComboBox_crossing_length.Items.Add("")




        ComboBox_area_EX_E.Items.Add("")
        ComboBox_area_P_E.Items.Add("")
        ComboBox_area_TWS.Items.Add("")
        ComboBox_area_ATWS.Items.Add("")
        ComboBox_area_A_R.Items.Add("")
        ComboBox_area_WARE.Items.Add("")
        ComboBox_area_TWS_ABD.Items.Add("")
        ComboBox_area_TWS_PD.Items.Add("")




        ComboBox_state_update.Items.Add("")
        ComboBox_county_update.Items.Add("")
        ComboBox_town_update.Items.Add("")
        ComboBox_linelist_update.Items.Add("")
        ComboBox_HMM_ID_update.Items.Add("")
        ComboBox_MBL_update.Items.Add("")
        ComboBox_deed_page_update.Items.Add("")
        ComboBox_APN_update.Items.Add("")
        ComboBox_crossing_length_update.Items.Add("")
        ComboBox_access_road_length_update.Items.Add("")
        ComboBox_Section_TWP_Range_update.Items.Add("")
        ComboBox_owner_update.Items.Add("")
        ComboBox_area_EX_E_update.Items.Add("")
        ComboBox_area_P_E_update.Items.Add("")
        ComboBox_area_TWS_update.Items.Add("")
        ComboBox_area_ATWS_update.Items.Add("")
        ComboBox_area_A_R_update.Items.Add("")
        ComboBox_area_WARE_update.Items.Add("")
        ComboBox_area_TWS_ABD_update.Items.Add("")
        ComboBox_area_TWS_PD_update.Items.Add("")
        ComboBox_pipe_diameter.Items.Add("")
        ComboBox_pipe_name.Items.Add("")
        ComboBox_pipe_segment.Items.Add("")


        Try

            If System.IO.File.Exists(Settings_file) = True Then
                Using Reader1 As New System.IO.StreamReader(Settings_file)
                    Dim Line1 As String
                    While Reader1.Peek > 0
                        Line1 = Reader1.ReadLine
                        If Line1.Contains(":") = True And Line1.Contains("|") = True Then
                            Dim Line_no As String = Line1.Split(":")(0)
                            Dim Value1 As String = Line1.Split("|")(1)
                            Select Case Line_no
                                Case 1
                                    TextBox_xref_model_space.Text = Value1
                                Case 2
                                    TextBox_dwt_template.Text = Value1
                                Case 3
                                    TextBox_sheet_set_template.Text = Value1
                                Case 4
                                    TextBox_Output_Directory.Text = Value1
                                Case 5
                                    TextBox_north_arrow_Big_X.Text = Value1
                                Case 6
                                    TextBox_north_arrow_Big_y.Text = Value1
                                Case 7
                                    TextBox_main_viewport_height.Text = Value1
                                Case 8
                                    TextBox_main_viewport_width.Text = Value1
                                Case 9
                                    TextBox_main_viewport_center_X.Text = Value1
                                Case 10
                                    TextBox_main_viewport_center_Y.Text = Value1
                                Case 11
                                    TextBox_north_arrow.Text = Value1

                                Case 12
                                    TextBox_north_arrow_small_X.Text = Value1
                                Case 13
                                    TextBox_north_arrow_small_Y.Text = Value1
                                Case 14
                                    TextBox_locus_viewport_height.Text = Value1
                                Case 15
                                    TextBox_locus_viewport_width.Text = Value1
                                Case 16
                                    TextBox_locus_viewport_center_X.Text = Value1
                                Case 17
                                    TextBox_locus_viewport_center_Y.Text = Value1
                                Case 18
                                    TextBox_north_arrow_locus.Text = Value1
                                Case 19
                                    With ComboBox_state
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 20
                                    With ComboBox_county
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 21
                                    With ComboBox_town
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 22
                                    With ComboBox_linelist
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 23
                                    With ComboBox_HMM_ID
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 24
                                    With ComboBox_MBL
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 25
                                    With ComboBox_owner
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With

                                Case 26


                                Case 27

                                Case 28
                                    With ComboBox_area_EX_E
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 29
                                    With ComboBox_area_P_E
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 30
                                    With ComboBox_area_TWS
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 31
                                    With ComboBox_area_ATWS
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 32
                                    With ComboBox_area_A_R
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 33
                                    With ComboBox_area_WARE
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 34
                                    With ComboBox_pipe_diameter
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 35
                                    With ComboBox_pipe_name
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 36
                                    With ComboBox_pipe_segment
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 37
                                    With ComboBox_state_update
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 38
                                    With ComboBox_county_update
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 39
                                    With ComboBox_town_update
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 40
                                    With ComboBox_linelist_update
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 41
                                    With ComboBox_HMM_ID_update
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 42
                                    With ComboBox_MBL_update
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 43
                                    With ComboBox_owner_update
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With


                                Case 44

                                Case 45

                                Case 46
                                    With ComboBox_area_EX_E_update
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 47
                                    With ComboBox_area_P_E_update
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 48
                                    With ComboBox_area_TWS_update
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 49
                                    With ComboBox_area_ATWS_update
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 50
                                    With ComboBox_area_A_R_update
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 51
                                    With ComboBox_area_WARE_update
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 52
                                    With ComboBox_plat_name
                                        If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                        .SelectedIndex = .Items.IndexOf(Value1)
                                    End With
                                Case 53
                                    If Value1 = 1 Then
                                        CheckBox_convert_sqft_to_acres.Checked = True
                                    Else
                                        CheckBox_convert_sqft_to_acres.Checked = False
                                    End If

                                Case 54
                                    If Value1 = 1 Then
                                        CheckBox_ignore_deflections_less_0_5.Checked = True
                                    Else
                                        CheckBox_ignore_deflections_less_0_5.Checked = False
                                    End If

                                Case 55
                                    If Value1 = 1 Then
                                        CheckBox_add_bearing_and_distances.Checked = True
                                    Else
                                        CheckBox_add_bearing_and_distances.Checked = False
                                    End If

                                Case 56
                                    If Value1 = 1 Then
                                        CheckBox_display_stations.Checked = True
                                    Else
                                        CheckBox_display_stations.Checked = False
                                    End If

                                Case 57
                                    If Value1 = 1 Then
                                        CheckBox_rotate_to_north.Checked = True
                                    Else
                                        CheckBox_rotate_to_north.Checked = False
                                    End If

                                Case 58
                                    Station_at_point_layer = Value1
                                Case 59
                                    Centerline_stationing_layer = Value1
                                Case 60
                                    Bearing_and_distance_layer = Value1
                                Case 61
                                    Tie_distance_layer = Value1
                                Case 62
                                    Arc_leader_polyline_layer = Value1
                                Case 63
                                    Layer_name_Main_Viewport = Value1
                                Case 64
                                    Layer_name_locus_VP = Value1
                                Case 65
                                    Layer_name_Blocks = Value1
                                Case 66
                                    Layer_name_text = Value1
                                Case 67
                                    Layer_poly_parcela_new = Value1
                                Case 68
                                    Layer_hatch_locus = Value1
                                Case 69
                                    With ComboBox_deed_page
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 70
                                    With ComboBox_APN
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 71
                                    With ComboBox_crossing_length
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 72
                                    With ComboBox_Section_TWP_Range
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 73
                                    With ComboBox_area_TWS_PD
                                        If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                        .SelectedIndex = .Items.IndexOf(Value1)
                                    End With
                                Case 74
                                    With ComboBox_area_TWS_ABD
                                        If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                        .SelectedIndex = .Items.IndexOf(Value1)
                                    End With
                                Case 75
                                    With ComboBox_access_road_length
                                        If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                        .SelectedIndex = .Items.IndexOf(Value1)
                                    End With

                                Case 76

                                Case 77
                                    With ComboBox_deed_page_update
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 78
                                    With ComboBox_APN_update
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 79
                                    With ComboBox_crossing_length_update
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 80
                                    With ComboBox_Section_TWP_Range_update
                                        .Items.Add(Value1)
                                        .SelectedIndex = 1
                                    End With
                                Case 81
                                    With ComboBox_area_TWS_PD_update
                                        If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                        .SelectedIndex = .Items.IndexOf(Value1)
                                    End With
                                Case 82
                                    With ComboBox_area_TWS_ABD_update
                                        If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                        .SelectedIndex = .Items.IndexOf(Value1)
                                    End With
                                Case 83
                                    With ComboBox_access_road_length_update
                                        If .Items.Contains(Value1) = False Then .Items.Add(Value1)
                                        .SelectedIndex = .Items.IndexOf(Value1)
                                    End With

                            End Select
                        End If
                    End While

                End Using
            Else
                'TextBox_xref_model_space.Text = "G:\KinderMorgan\339501_NEDProject\DataProd\DWG\Segment_A\A_Host_Basefile_Plats.dwg"
                'TextBox_dwt_template.Text = "C:\Users\pop70694\Documents\Work Files\2015-10-19 plat generator error\danF.dwt" '  "G:\KinderMorgan\339501_NEDProject\DataProd\Work\Drafting\Land_Plats\1_Platt_Generator\Platt_Generator_Temp_19F.dwt"
                'TextBox_sheet_set_template.Text = "C:\Users\pop70694\Documents\Work Files\2015-10-19 plat generator error\Hancock_plats.dst" ' "G:\KinderMorgan\339501_NEDProject\DataProd\Work\Drafting\Land_Plats\1_Platt_Generator\Platt_Sheet_Set_19F.dst"
                'TextBox_Output_Directory.Text = "C:\Users\pop70694\Documents\Work Files\" ' "G:\KinderMorgan\339501_NEDProject\DataProd\Work\Drafting\Land_Plats\Segment_K\"
            End If

        Catch ex As System.Exception
            MsgBox(ex.Message)
        End Try


        Incarca_existing_Blocks_with_attributes_to_combobox(ComboBox_blocks)

        If ComboBox_blocks.Items.Contains("Owner_Linelist") = True Then
            ComboBox_blocks.SelectedIndex = ComboBox_blocks.Items.IndexOf("Owner_Linelist")
        Else
            ComboBox_blocks.SelectedIndex = 0
        End If

        Incarca_existing_Atributes_to_combobox(ComboBox_blocks.Text, ComboBox_bl_atr1)
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks.Text, ComboBox_bl_atr2)
        If ComboBox_bl_atr1.Items.Count > 1 Then
            ComboBox_bl_atr1.SelectedIndex = 1
        End If
        If ComboBox_bl_atr2.Items.Count > 2 Then
            ComboBox_bl_atr2.SelectedIndex = 2
        End If
    End Sub

    Private Sub Button_load_fields_from_dst_Click(sender As Object, e As EventArgs) Handles Button_load_fields_from_dst.Click



        If TextBox_sheet_set_template.Text = "" Then
            MsgBox("Please specify the SHEET SET file")
            Exit Sub
        End If
        If Not Strings.Right(TextBox_sheet_set_template.Text, 3).ToUpper = "DST" Then
            MsgBox("Please specify the SHEET SET file")
            Exit Sub
        End If

        If Freeze_operations = False Then
            Freeze_operations = True


            Try


                Dim SheetSet_manager As New AcSmSheetSetMgr
                Dim SheetSet_database As AcSmDatabase = SheetSet_manager.OpenDatabase(TextBox_sheet_set_template.Text, False)
                Dim sheetSet As AcSmSheetSet = SheetSet_database.GetSheetSet()

                If LockDatabase(SheetSet_database, True) = True Then
                    Dim EnumSheets As IAcSmEnumComponent = sheetSet.GetSheetEnumerator()
                    Dim smComponent As IAcSmComponent
                    Dim sheet1 As IAcSmSheet
                    smComponent = EnumSheets.Next()
                    While True
                        If smComponent Is Nothing Then
                            Exit While
                        End If
                        sheet1 = TryCast(smComponent, IAcSmSheet)
                        If IsNothing(sheet1) = False Then
                            With ComboBox_Sheet_Set_state
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_Sheet_Set_county
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_Sheet_Set_town
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_Sheet_Set_linelist
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_Sheet_Set_hmm_ID
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_Sheet_Set_MBL
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_Sheet_Set_deed_page
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_Sheet_Set_APN
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_Sheet_Set_access_road_length
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_Sheet_Set_sec_twp_range
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_Sheet_Set_owner
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_Sheet_Set_crossing_len_FT
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_Sheet_Set_crossing_len_rod
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_Sheet_Set_ex_e
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_Sheet_Set_P_E
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_Sheet_Set_TWS
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_Sheet_Set_atws
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_Sheet_Set_a_R
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_Sheet_Set_ware
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_Sheet_Set_tws_abd
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_Sheet_Set_tws_PD
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_Sheet_Set_pipe_DIAMETER
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_Sheet_Set_pipe_name
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_Sheet_Set_pipe_segment
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_SHEET_SET_scale_locus
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_SHEET_SET_scale_main
                                .Items.Clear()
                                .Text = ""
                            End With

                            Dim customPropertyBag As AcSmCustomPropertyBag = sheet1.GetCustomPropertyBag()
                            Dim EnumProp As IAcSmEnumProperty = customPropertyBag.GetPropertyEnumerator()
                            Do
                                Dim Prop_name As String = ""
                                Dim Prop_value As AcSmCustomPropertyValue = Nothing
                                EnumProp.Next(Prop_name, Prop_value)
                                If Prop_name = "" Then Exit Do

                                With ComboBox_Sheet_Set_state
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_Sheet_Set_county
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_Sheet_Set_town
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_Sheet_Set_linelist
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_Sheet_Set_hmm_ID
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_Sheet_Set_MBL
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_Sheet_Set_deed_page
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_Sheet_Set_APN
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_Sheet_Set_access_road_length
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_Sheet_Set_sec_twp_range
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_Sheet_Set_owner
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_Sheet_Set_crossing_len_FT
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_Sheet_Set_crossing_len_rod
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_Sheet_Set_ex_e
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_Sheet_Set_P_E
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_Sheet_Set_TWS
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_Sheet_Set_atws
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_Sheet_Set_a_R
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_Sheet_Set_ware
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_Sheet_Set_tws_abd
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_Sheet_Set_tws_PD
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_Sheet_Set_pipe_name
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_Sheet_Set_pipe_DIAMETER
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_Sheet_Set_pipe_segment
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_SHEET_SET_scale_locus
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_SHEET_SET_scale_main
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                            Loop

                        End If
                        smComponent = EnumSheets.Next()
                    End While
                    LockDatabase(SheetSet_database, False)

                    MsgBox("DONE")

                Else
                    MsgBox(SheetSet_database.GetLockStatus.ToString)
                    ' Display error message
                    MsgBox("Sheet set could not be opened for write.")
                End If

            Catch ex As Exception

                MsgBox(ex.Message)
            End Try

            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_load_fields_from_dst_UPDATE_Click(sender As Object, e As EventArgs) Handles Button_load_fields_from_dst_UPDATE.Click


        If TextBox_sheet_set_template.Text = "" Then
            MsgBox("Please specify the SHEET SET file")
            Exit Sub
        End If
        If Not Strings.Right(TextBox_sheet_set_template.Text, 3).ToUpper = "DST" Then
            MsgBox("Please specify the SHEET SET file")
            Exit Sub
        End If

        If Freeze_operations = False Then
            Freeze_operations = True


            Try


                Dim SheetSet_manager As New AcSmSheetSetMgr
                Dim SheetSet_database As AcSmDatabase = SheetSet_manager.OpenDatabase(TextBox_sheet_set_template.Text, False)
                Dim sheetSet As AcSmSheetSet = SheetSet_database.GetSheetSet()

                If LockDatabase(SheetSet_database, True) = True Then
                    Dim EnumSheets As IAcSmEnumComponent = sheetSet.GetSheetEnumerator()
                    Dim smComponent As IAcSmComponent
                    Dim sheet1 As IAcSmSheet
                    smComponent = EnumSheets.Next()
                    While True
                        If smComponent Is Nothing Then
                            Exit While
                        End If
                        sheet1 = TryCast(smComponent, IAcSmSheet)
                        If IsNothing(sheet1) = False Then
                            With ComboBox_sheet_set_state_update
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_sheet_set_county_update
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_sheet_set_town_update
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_sheet_set_linelist_update
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_sheet_set_HMMID_update
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_sheet_set_MBL_update
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_sheet_set_deed_page_update
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_sheet_set_apn_update
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_sheet_set_access_road_length_update
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_sheet_set_sec_twp_range_update
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_sheet_set_owner_update
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_sheet_set_crosing_length_ft_update
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_sheet_set_crosing_length_rod_update
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_sheet_set_ex_e_update
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_sheet_set_p_e_update
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_sheet_set_tws_update
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_sheet_set_atws_update
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_sheet_set_a_r_update
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_sheet_set_ware_update
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_sheet_set_tws_ABD_update
                                .Items.Clear()
                                .Text = ""
                            End With
                            With ComboBox_sheet_set_TWS_PD_update
                                .Items.Clear()
                                .Text = ""
                            End With

                            With ComboBox_sheet_user1
                                .Items.Clear()
                                .Text = ""
                            End With

                            With ComboBox_sheet_user2
                                .Items.Clear()
                                .Text = ""
                            End With

                            With ComboBox_sheet_user3
                                .Items.Clear()
                                .Text = ""
                            End With

                            With ComboBox_sheet_user4
                                .Items.Clear()
                                .Text = ""
                            End With

                            With ComboBox_sheet_user5
                                .Items.Clear()
                                .Text = ""
                            End With

                            With ComboBox_sheet_user6
                                .Items.Clear()
                                .Text = ""
                            End With

                            With ComboBox_sheet_user7
                                .Items.Clear()
                                .Text = ""
                            End With

                            With ComboBox_sheet_user8
                                .Items.Clear()
                                .Text = ""
                            End With

                            With ComboBox_sheet_user9
                                .Items.Clear()
                                .Text = ""
                            End With

                            With ComboBox_sheet_user10
                                .Items.Clear()
                                .Text = ""
                            End With

                            With ComboBox_sheet_user11
                                .Items.Clear()
                                .Text = ""
                            End With

                            With ComboBox_sheet_user12
                                .Items.Clear()
                                .Text = ""
                            End With

                            With ComboBox_sheet_user13
                                .Items.Clear()
                                .Text = ""
                            End With




                            Dim customPropertyBag As AcSmCustomPropertyBag = sheet1.GetCustomPropertyBag()
                            Dim EnumProp As IAcSmEnumProperty = customPropertyBag.GetPropertyEnumerator()
                            Do
                                Dim Prop_name As String = ""
                                Dim Prop_value As AcSmCustomPropertyValue = Nothing
                                EnumProp.Next(Prop_name, Prop_value)
                                If Prop_name = "" Then Exit Do

                                With ComboBox_sheet_set_state_update
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_sheet_set_county_update
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_sheet_set_town_update
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_sheet_set_linelist_update
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_sheet_set_HMMID_update
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_sheet_set_MBL_update
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_sheet_set_deed_page_update
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_sheet_set_apn_update
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_sheet_set_access_road_length_update
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_sheet_set_sec_twp_range_update
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_sheet_set_owner_update
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_sheet_set_crosing_length_ft_update
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_sheet_set_crosing_length_rod_update
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_sheet_set_ex_e_update
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_sheet_set_p_e_update
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_sheet_set_tws_update
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_sheet_set_atws_update
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_sheet_set_a_r_update
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_sheet_set_ware_update
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_sheet_set_tws_ABD_update
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With
                                With ComboBox_sheet_set_TWS_PD_update
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With

                                With ComboBox_sheet_user1
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With

                                With ComboBox_sheet_user2
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With

                                With ComboBox_sheet_user3
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With

                                With ComboBox_sheet_user4
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With

                                With ComboBox_sheet_user5
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With

                                With ComboBox_sheet_user6
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With

                                With ComboBox_sheet_user7
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With

                                With ComboBox_sheet_user8
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With

                                With ComboBox_sheet_user9
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With

                                With ComboBox_sheet_user10
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With

                                With ComboBox_sheet_user11
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With

                                With ComboBox_sheet_user12
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With

                                With ComboBox_sheet_user13
                                    If .Items.Contains(Prop_name) = False Then
                                        .Items.Add(Prop_name)
                                    End If
                                End With


                            Loop

                        End If
                        smComponent = EnumSheets.Next()
                    End While
                    LockDatabase(SheetSet_database, False)

                    MsgBox("DONE")

                Else
                    MsgBox(SheetSet_database.GetLockStatus.ToString)
                    ' Display error message
                    MsgBox("Sheet set could not be opened for write.")
                End If

            Catch ex As Exception

                MsgBox(ex.Message)
            End Try

            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_pick_and_label_point_Click(sender As Object, e As EventArgs) Handles Button_pick_and_label_point.Click



        If Freeze_operations = False Then
            Freeze_operations = True




            Dim OLD_OSnap As Integer = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE")

            Dim NEW_OSnap As Integer = Autodesk.AutoCAD.EditorInput.ObjectSnapMasks.End + Autodesk.AutoCAD.EditorInput.ObjectSnapMasks.Intersection

            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap)

            Try
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Database1 As Autodesk.AutoCAD.DatabaseServices.Database

                ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                Using lock1 As DocumentLock = ThisDrawing.LockDocument
                    Database1 = ThisDrawing.Database
                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
                    Dim Arrowid As ObjectId = Get_Arrow_dimension_ID("DIMBLK2", "_DotSmall")

                    Dim Tilemode1 As Integer = Application.GetSystemVariable("TILEMODE")
                    Dim CVport1 As Integer = Application.GetSystemVariable("CVPORT")
                    Dim Landinggap1 As Double
                    Dim Doglength1 As Double
                    Dim Texth As Double
                    Dim Arrowsize1 As Double

                    Landinggap1 = 0.05
                    Doglength1 = 0.1
                    Texth = 0.06
                    Arrowsize1 = 0.2




123:
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction


                        If Viewport_loaded = True And Tilemode1 = 1 Then
                            Landinggap1 = Scale_factor / 20
                            Doglength1 = Scale_factor / 10
                            Texth = Scale_factor / 16.66667
                            Arrowsize1 = Scale_factor / 5
                        Else
                            If (Tilemode1 = 0 And Not CVport1 = 1) Or Tilemode1 = 1 Then
                                For Each ctrl2 As Windows.Forms.Control In Panel_SCALE_SELECTION.Controls
                                    Dim Radiob2 As Windows.Forms.RadioButton
                                    If TypeOf ctrl2 Is Windows.Forms.RadioButton Then
                                        Radiob2 = ctrl2
                                        If Radiob2.Checked = True Then
                                            Dim Nume1 As String = Replace(Radiob2.Name, "RadioButton", "")
                                            If IsNumeric(Nume1) = True Then

                                                Landinggap1 = CInt(Nume1) / 20
                                                Doglength1 = CInt(Nume1) / 10
                                                Texth = CInt(Nume1) / 16.66667
                                                Arrowsize1 = CInt(Nume1) / 5
                                                Exit For
                                            End If
                                        End If
                                    End If
                                Next
                            End If

                        End If


                        Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Point 1: ")
                        Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult

                        PP1.AllowNone = False
                        Point1 = Editor1.GetPoint(PP1)

                        If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                            Freeze_operations = False
                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
                            Exit Sub
                        End If

                        Dim PP2 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions("Point 2:")
                        Dim Point2 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                        PP2.AllowNone = False
                        PP2.UseBasePoint = True
                        PP2.BasePoint = Point1.Value
                        Point2 = Editor1.GetPoint(PP2)


                        If Point2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                            Freeze_operations = False
                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
                            Exit Sub
                        End If



                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)



                        Dim Mleader1 As New MLeader
                        Dim Nr1 As Integer = Mleader1.AddLeader()
                        Dim Nr2 As Integer = Mleader1.AddLeaderLine(Nr1)
                        Mleader1.AddFirstVertex(Nr2, New Point3d(Point1.Value.X, Point1.Value.Y, 0))
                        Mleader1.AddLastVertex(Nr2, New Point3d(Point2.Value.X, Point2.Value.Y, 0))
                        Mleader1.LeaderLineType = LeaderType.StraightLeader

                        Mleader1.ContentType = ContentType.MTextContent




                        Mleader1.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.LeftLeader)
                        Mleader1.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.RightLeader)
                        Mleader1.Annotative = AnnotativeStates.False


                        Dim Mtext1 As New MText

                        Mtext1.Contents = "N:" & Get_String_Rounded(Point1.Value.Y, 1) & "'" & vbCrLf & "E:" & Get_String_Rounded(Point1.Value.X, 1) & "'"
                        Mtext1.ColorIndex = 0

                        Mleader1.MText = Mtext1


                        Mleader1.TextHeight = Texth
                        Mleader1.ArrowSymbolId = Arrowid
                        Mleader1.LandingGap = Landinggap1
                        Mleader1.ArrowSize = Arrowsize1
                        Mleader1.DoglegLength = Doglength1



                        BTrecord.AppendEntity(Mleader1)
                        Trans1.AddNewlyCreatedDBObject(Mleader1, True)
                        Trans1.Commit()
                        GoTo 123

                    End Using

                End Using
            Catch ex As Exception
                Freeze_operations = False
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
                MsgBox(ex.Message)
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
            Freeze_operations = False
        End If
    End Sub
End Class


