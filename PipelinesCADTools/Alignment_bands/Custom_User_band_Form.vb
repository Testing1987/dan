Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class Custom_User_band_Form
    Dim Colectie1 As New Specialized.StringCollection

    Dim Data_table_Parcels As System.Data.DataTable
    Dim Data_table_Matchlines As System.Data.DataTable
    Dim PolyCL As Polyline
    Dim PolyCL3D As Polyline3d
    Dim Poly_length As Double
    Dim Empty_array() As ObjectId
    Dim Data_table_crossing_band_Excel_data As System.Data.DataTable
    Dim Data_table_station_equation As System.Data.DataTable
    Dim Data_table_class_band_Excel_data As System.Data.DataTable
    Dim Data_table_materials As System.Data.DataTable
    Dim Data_table_water_band_Excel_data As System.Data.DataTable

    Dim Freeze_operations As Boolean = False
    Dim Is_canada As Boolean = False

    Dim No_plot As String = "NO PLOT"


    Private Sub Line_list_band_Form_Load(sender As Object, e As EventArgs) Handles Me.Load
        Incarca_existing_layers_to_combobox(ComboBox_layers)
        Incarca_existing_layers_to_combobox(ComboBox_layers_for_Mtext)


        If ComboBox_layers.Items.Count > 0 Then
            If ComboBox_layers.Items.Contains("Text_ROW") = True Then
                ComboBox_layers.SelectedIndex = ComboBox_layers.Items.IndexOf("Text_ROW")
            Else
                ComboBox_layers.SelectedIndex = 0
            End If
        End If

        If ComboBox_layers_for_Mtext.Items.Count > 0 Then
            If ComboBox_layers_for_Mtext.Items.Contains("Text_ROW") = True Then
                ComboBox_layers_for_Mtext.SelectedIndex = ComboBox_layers_for_Mtext.Items.IndexOf("Text_Material")
            Else
                ComboBox_layers_for_Mtext.SelectedIndex = 0
            End If
        End If

        Incarca_existing_Blocks_with_attributes_to_combobox(ComboBox_blocks)



        Incarca_existing_textstyles_to_combobox(ComboBox_text_style)
        Incarca_existing_textstyles_to_combobox(ComboBox_text_style_mtext)


        If ComboBox_text_style_mtext.Items.Count > 0 Then
            If ComboBox_text_style_mtext.Items.Contains("Arial") = True Then
                ComboBox_text_style_mtext.SelectedIndex = ComboBox_text_style_mtext.Items.IndexOf("Arial")
            Else
                ComboBox_text_style_mtext.SelectedIndex = 0
            End If
        End If

        Incarca_existing_layers_to_combobox(ComboBox_layers_deflections)
        Incarca_existing_layers_to_combobox(ComboBox_layers_crossings)

        If ComboBox_layers_deflections.Items.Count > 0 Then
            If ComboBox_layers_deflections.Items.Contains("TEXT_PI") = True Then
                ComboBox_layers_deflections.SelectedIndex = ComboBox_layers_deflections.Items.IndexOf("TEXT_PI")
            Else
                ComboBox_layers_deflections.SelectedIndex = 0
            End If
        End If
        If ComboBox_layers_crossings.Items.Count > 0 Then
            If ComboBox_layers_crossings.Items.Contains("Text") = True Then
                ComboBox_layers_crossings.SelectedIndex = ComboBox_layers_crossings.Items.IndexOf("Text")
            Else
                ComboBox_layers_crossings.SelectedIndex = 0
            End If
        End If

        Incarca_existing_textstyles_to_combobox(ComboBox_text_styles_excel_data)

        If ComboBox_text_styles_excel_data.Items.Count > 0 Then
            If ComboBox_text_styles_excel_data.Items.Contains("ALIGNDB") = True Then
                ComboBox_text_styles_excel_data.SelectedIndex = ComboBox_text_styles_excel_data.Items.IndexOf("ALIGNDB")
            Else
                ComboBox_text_styles_excel_data.SelectedIndex = 0
            End If
        End If

        Incarca_existing_layers_to_combobox(ComboBox_class_layer)

        If ComboBox_class_layer.Items.Count > 0 Then
            If ComboBox_class_layer.Items.Contains("CLASS") = True Then
                ComboBox_class_layer.SelectedIndex = ComboBox_class_layer.Items.IndexOf("CLASS")
            Else
                ComboBox_class_layer.SelectedIndex = 0
            End If
        End If


        Incarca_existing_textstyles_to_combobox(ComboBox_class_text_style)
        If ComboBox_class_text_style.Items.Count > 0 Then
            If ComboBox_class_text_style.Items.Contains("CLASS") = True Then
                ComboBox_class_text_style.SelectedIndex = ComboBox_class_text_style.Items.IndexOf("CLASS")
            Else
                ComboBox_class_text_style.SelectedIndex = 0
            End If
        End If

        Incarca_existing_layers_to_combobox(ComboBox_layer_water)

        If ComboBox_layer_water.Items.Contains("TEXT") = True Then
            ComboBox_layer_water.SelectedIndex = ComboBox_layer_water.Items.IndexOf("TEXT")
        ElseIf ComboBox_layer_water.Items.Contains("CLASS") = True Then
            ComboBox_layer_water.SelectedIndex = ComboBox_layer_water.Items.IndexOf("CLASS")
        Else
            ComboBox_layer_water.SelectedIndex = 0
        End If

        Incarca_existing_textstyles_to_combobox(ComboBox_text_style_water)
        If ComboBox_text_style_water.Items.Count > 0 Then

            If ComboBox_text_style_water.Items.Contains("CLASS") = True Then
                ComboBox_text_style_water.SelectedIndex = ComboBox_text_style_water.Items.IndexOf("CLASS")
            ElseIf ComboBox_text_style_water.Items.Contains("Arial") = True Then
                ComboBox_text_style_water.SelectedIndex = ComboBox_text_style_water.Items.IndexOf("Arial")
            Else
                ComboBox_text_style_water.SelectedIndex = 0
            End If
        End If

        TextBox_viewport_Width.Text = "6068.3117"
        TextBox_viewport_Height.Text = "260"
        TextBox_X_MS.Text = "0"
        TextBox_Y_MS.Text = "-1000"
        TextBox_X_PS.Text = "1056.6883"
        TextBox_Y_PS.Text = "4405.0000"
        TextBox_BAND_SPACING.Text = "600"
        TextBox_shift_viewport_y.Text = "0"
        TextBox_CSF.Text = "1"
        Label_scale_1_to.Text = "Scale 1" & Chr(34) & " to"
        Label_inches.Text = "'"
        TextBox_prefix_excel_data.Text = ""

        TextBox_textOffset_X.Text = "5"
        TextBox_textOffset_Y.Text = "75"
        TextBox_text_height.Text = "16"
        With ComboBox_text_style
            If .Items.Contains("Romans") = True Then
                .SelectedIndex = .Items.IndexOf("Romans")
            End If
        End With
        TextBox_textwidth.Text = "0.8"
        TextBox_minimum_distance.Text = "400"






        Panel_blocks_insertion.Visible = False

    End Sub
    Private Sub Panel_design_param_Click(sender As Object, e As EventArgs) Handles Panel_design_param.Click
        Incarca_existing_layers_to_combobox(ComboBox_layers)

        If ComboBox_layers.Items.Count > 0 Then
            If ComboBox_layers.Items.Contains("Text_ROW") = True Then
                ComboBox_layers.SelectedIndex = ComboBox_layers.Items.IndexOf("Text_ROW")
            Else
                ComboBox_layers.SelectedIndex = 0
            End If
        End If


        Incarca_existing_textstyles_to_combobox(ComboBox_text_style)

        If ComboBox_text_style.Items.Count > 0 Then
            If ComboBox_text_style.Items.Contains("PROPERTY") = True Then
                ComboBox_text_style.SelectedIndex = ComboBox_text_style.Items.IndexOf("PROPERTY")
            Else
                ComboBox_text_style.SelectedIndex = 0
            End If
        End If

    End Sub
    Private Sub Panel_Blocks_Click(sender As Object, e As EventArgs) Handles Panel_Blocks.Click
        Incarca_existing_Blocks_with_attributes_to_combobox(ComboBox_blocks)
        ComboBox_line_list_attribute.Items.Clear()
        ComboBox_TRACT_attribute.Items.Clear()
        ComboBox_LENGTH_attribute.Items.Clear()
    End Sub


    Private Sub Panel_xcel_formating_Click(sender As Object, e As EventArgs) Handles Panel_xcel_formating.Click
        Incarca_existing_textstyles_to_combobox(ComboBox_text_styles_excel_data)
        Incarca_existing_layers_to_combobox(ComboBox_layers_deflections)
        Incarca_existing_layers_to_combobox(ComboBox_layers_crossings)
    End Sub

    Private Sub Button_load_OD_Click(sender As Object, e As EventArgs) Handles Button_load_OD.Click

        If Freeze_operations = False Then
            Freeze_operations = True
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor

            Try
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                Colectie1 = New Specialized.StringCollection

                Editor1.SetImpliedSelection(Empty_array)

                Dim Rezultat_Parc As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select a sample parcel containing object data:"
                Object_Prompt.SingleOnly = True
                Rezultat_Parc = Editor1.GetSelection(Object_Prompt)
                If Not Rezultat_Parc.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Freeze_operations = False
                    Editor1.SetImpliedSelection(Empty_array)
                    Editor1.WriteMessage(vbLf & "Command:")
                    Freeze_operations = False
                    Exit Sub
                End If

                If Rezultat_Parc.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    If IsNothing(Rezultat_Parc) = False Then
                        Using Lock_dwg As DocumentLock = ThisDrawing.LockDocument
                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                                If ComboBox_object_data_linelist.Items.Count = 0 Then ComboBox_object_data_linelist.Items.Add("")
                                If ComboBox_object_data_tract.Items.Count = 0 Then ComboBox_object_data_tract.Items.Add("")
                                For j = 0 To Rezultat_Parc.Value.Count - 1
                                    Dim Ent1 As Entity = Trans1.GetObject(Rezultat_Parc.Value.Item(j).ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

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
                                                        With ComboBox_object_data_linelist
                                                            If .Items.Contains(Field_def1.Name) = False Then
                                                                .Items.Add(Field_def1.Name)
                                                            End If
                                                        End With
                                                        With ComboBox_object_data_tract
                                                            If .Items.Contains(Field_def1.Name) = False Then
                                                                .Items.Add(Field_def1.Name)
                                                            End If
                                                        End With

                                                    Next
                                                Next
                                            End If
                                        End If
                                    End Using
                                Next


                                With ComboBox_object_data_linelist
                                    If .Items.Count > 0 Then
                                        For i = 0 To .Items.Count - 1

                                            If .Items(i).ToString.ToUpper = "FIELD1" Then
                                                .SelectedIndex = i
                                                Exit For
                                            Else
                                                .SelectedIndex = 0
                                            End If
                                        Next

                                    End If
                                End With
                                With ComboBox_object_data_tract
                                    For i = 0 To .Items.Count - 1

                                        If .Items(i).ToString.ToUpper = "OWNER" Then
                                            .SelectedIndex = i
                                            Exit For
                                        Else
                                            .SelectedIndex = 0
                                        End If
                                    Next
                                End With

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

    Private Sub Button_SELECT_PARCELS_Click(sender As Object, e As EventArgs) Handles Button_SELECT_PARCELS.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Dim Error_index1 As Integer
            Dim Error_index2 As Integer

            Dim Round1 As Integer = 0
            If IsNumeric(TextBox_dec1.Text) = True Then
                Round1 = CInt(TextBox_dec1.Text)
            End If

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor

            Dim RezultatCL As Autodesk.AutoCAD.EditorInput.PromptEntityResult

            Dim Object_PromptCL As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select centerline:")

            Object_PromptCL.SetRejectMessage(vbLf & "Please select a lightweight polyline or a 3D polyline")
            Object_PromptCL.AddAllowedClass(GetType(Polyline), True)
            Object_PromptCL.AddAllowedClass(GetType(Polyline3d), True)
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            RezultatCL = Editor1.GetEntity(Object_PromptCL)


            If Not RezultatCL.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                MsgBox("NO centerline")
                Editor1.WriteMessage(vbLf & "Command:")
                Freeze_operations = False
                Editor1.SetImpliedSelection(Empty_array)
                Exit Sub
            End If



            Try

                Colectie1 = New Specialized.StringCollection

                Editor1.SetImpliedSelection(Empty_array)
                Label_od_LOADED.Text = "Identified:"

                Dim idarrayEmpty() As ObjectId
                ThisDrawing.Editor.SetImpliedSelection(idarrayEmpty)


                Data_table_Parcels = New System.Data.DataTable
                Data_table_Parcels.Columns.Add("FIELD1", GetType(String))
                Data_table_Parcels.Columns.Add("FIELD2", GetType(String))
                Data_table_Parcels.Columns.Add("X1", GetType(Double))
                Data_table_Parcels.Columns.Add("Y1", GetType(Double))
                Data_table_Parcels.Columns.Add("X2", GetType(Double))
                Data_table_Parcels.Columns.Add("Y2", GetType(Double))
                Data_table_Parcels.Columns.Add("BEGSTA", GetType(Double))
                Data_table_Parcels.Columns.Add("ENDSTA", GetType(Double))
                Data_table_Parcels.Columns.Add("FIELD3", GetType(Double))





                If RezultatCL.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                    If IsNothing(RezultatCL) = False Then
                        Using Lock_dwg As DocumentLock = ThisDrawing.LockDocument



                            Dim Index_data_parcels As Double = 0

                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                                PolyCL = TryCast(Trans1.GetObject(RezultatCL.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Polyline)

                                PolyCL3D = TryCast(Trans1.GetObject(RezultatCL.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Polyline3d)

                                If IsNothing(PolyCL3D) = False Then
                                    PolyCL = Creaza_polyline_din_polyline3d(Trans1, PolyCL3D)
                                End If



                                If PolyCL.NumberOfVertices < 2 Then
                                    MsgBox("NO centerline")
                                    Editor1.WriteMessage(vbLf & "Command:")
                                    Freeze_operations = False
                                    Editor1.SetImpliedSelection(Empty_array)
                                    Exit Sub
                                End If

                                Dim Poly_CL As New Polyline
                                Poly_CL = PolyCL.Clone
                                Poly_CL.Elevation = 0

                                Dim BTrecord As BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                                For Each OD1 As ObjectId In BTrecord
                                    Dim Ent1 As Entity = TryCast(Trans1.GetObject(OD1, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Entity)



                                    If IsNothing(Ent1) = False Then

                                        Dim ruleaza As Boolean = False
                                        If IsNothing(PolyCL3D) = False Then
                                            If Not OD1 = PolyCL3D.ObjectId Then
                                                ruleaza = True
                                            End If
                                        Else
                                            If Not OD1 = PolyCL.ObjectId Then
                                                ruleaza = True
                                            End If
                                        End If

                                        If ruleaza = True Then
                                            If TypeOf Ent1 Is Polyline Then
                                                Dim Poly_with_object_data As Polyline = Ent1
                                                Dim Poly_pt_intersectie As New Polyline
                                                Poly_pt_intersectie = Poly_with_object_data.Clone
                                                Poly_pt_intersectie.Elevation = 0

                                                Dim COL_INT As New Point3dCollection
                                                COL_INT = Intersect_on_both_operands(Poly_CL, Poly_pt_intersectie)

                                                Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                                                Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables
                                                Dim Id1 As ObjectId = Ent1.ObjectId


                                                If COL_INT.Count > 0 Then

                                                    Dim Data_table_unique As New System.Data.DataTable
                                                    Data_table_unique.Columns.Add("FIELD1", GetType(String))
                                                    Data_table_unique.Columns.Add("FIELD2", GetType(String))
                                                    Data_table_unique.Columns.Add("X", GetType(Double))
                                                    Data_table_unique.Columns.Add("Y", GetType(Double))
                                                    Data_table_unique.Columns.Add("STATION", GetType(Double))
                                                    Data_table_unique.Columns.Add("FIELD3", GetType(Double))

                                                    Dim Valoare_field1 As String = "XXX"
                                                    Dim Valoare_field2 As String = "XXX"



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

                                                                        If ComboBox_object_data_linelist.Text = Field_def1.Name Then
                                                                            If Not Record1(i).StrValue = "" Then
                                                                                Valoare_field1 = Record1(i).StrValue
                                                                            End If
                                                                        End If
                                                                        If ComboBox_object_data_tract.Text = Field_def1.Name Then
                                                                            If Not Record1(i).StrValue = "" Then
                                                                                Valoare_field2 = Record1(i).StrValue
                                                                            End If
                                                                        End If
                                                                    Next
                                                                Next
                                                            End If

                                                        End If

                                                    End Using


                                                    Dim Start_end As Integer = 0

                                                    For m = 0 To COL_INT.Count - 1
                                                        Data_table_unique.Rows.Add()
                                                        Data_table_unique.Rows(m).Item("X") = COL_INT(m).X
                                                        Data_table_unique.Rows(m).Item("Y") = COL_INT(m).Y


                                                        Dim Station2D As Double = Round(Poly_CL.GetDistAtPoint(COL_INT(m)), Round1)

                                                        If IsNothing(PolyCL3D) = False Then
                                                            Dim Param2d As Double = PolyCL.GetParameterAtPoint(COL_INT(m))

                                                            Dim Station3D As Double = Round(PolyCL3D.GetDistanceAtParameter(Param2d), Round1)
                                                            Data_table_unique.Rows(m).Item("STATION") = Station3D


                                                        Else

                                                            Data_table_unique.Rows(m).Item("STATION") = Station2D
                                                        End If


                                                        Data_table_unique.Rows(m).Item("FIELD1") = Valoare_field1
                                                        Data_table_unique.Rows(m).Item("FIELD2") = Valoare_field2
                                                        If Ent1.Layer = "FIELD3" Then
                                                            Data_table_unique.Rows(m).Item("FIELD3") = PI / 2
                                                        Else
                                                            Data_table_unique.Rows(m).Item("FIELD3") = 0
                                                        End If
                                                    Next

                                                    Data_table_unique = Sort_data_table(Data_table_unique, "STATION")

                                                    For m = 0 To COL_INT.Count - 1

                                                        If COL_INT.Count = 1 Then
                                                            Dim Station1 As Double = Poly_CL.GetDistAtPoint(COL_INT(0))
                                                            Dim Station0 As Double = 0
                                                            Dim Station2 As Double = Poly_CL.Length




                                                            If Station1 - Station0 <= Station2 - Station1 Then
                                                                Data_table_unique.Rows.Add()
                                                                Data_table_unique.Rows(1).Item("X") = Data_table_unique.Rows(0).Item("X")
                                                                Data_table_unique.Rows(1).Item("Y") = Data_table_unique.Rows(0).Item("Y")
                                                                Data_table_unique.Rows(1).Item("STATION") = Data_table_unique.Rows(0).Item("STATION")
                                                                Data_table_unique.Rows(1).Item("FIELD1") = Data_table_unique.Rows(0).Item("FIELD1")
                                                                Data_table_unique.Rows(1).Item("FIELD2") = Data_table_unique.Rows(0).Item("FIELD2")
                                                                If Ent1.Layer = "FIELD3" Then
                                                                    Data_table_unique.Rows(1).Item("FIELD3") = PI / 2
                                                                Else
                                                                    Data_table_unique.Rows(1).Item("FIELD3") = 0
                                                                End If

                                                                Data_table_unique.Rows(0).Item("X") = Poly_CL.GetPointAtDist(0).X
                                                                Data_table_unique.Rows(0).Item("Y") = Poly_CL.GetPointAtDist(0).Y
                                                                Data_table_unique.Rows(0).Item("STATION") = 0
                                                                Data_table_unique.Rows(0).Item("FIELD1") = Valoare_field1
                                                                Data_table_unique.Rows(0).Item("FIELD2") = Valoare_field2
                                                                If Ent1.Layer = "FIELD3" Then
                                                                    Data_table_unique.Rows(0).Item("FIELD3") = PI / 2
                                                                Else
                                                                    Data_table_unique.Rows(0).Item("FIELD3") = 0
                                                                End If
                                                            Else
                                                                Data_table_unique.Rows.Add()
                                                                Data_table_unique.Rows(1).Item("X") = Poly_CL.GetPointAtDist(Station2).X
                                                                Data_table_unique.Rows(1).Item("Y") = Poly_CL.GetPointAtDist(Station2).Y
                                                                Data_table_unique.Rows(1).Item("STATION") = Round(Station2, Round1)
                                                                Data_table_unique.Rows(1).Item("FIELD1") = Valoare_field1
                                                                Data_table_unique.Rows(1).Item("FIELD2") = Valoare_field2
                                                                If Ent1.Layer = "FIELD3" Then
                                                                    Data_table_unique.Rows(1).Item("FIELD3") = PI / 2
                                                                Else
                                                                    Data_table_unique.Rows(1).Item("FIELD3") = 0
                                                                End If
                                                            End If
                                                        End If



                                                    Next

                                                    Data_table_unique = Sort_data_table(Data_table_unique, "STATION")

                                                    For s = 0 To Data_table_unique.Rows.Count - 2 Step 2

                                                        Data_table_Parcels.Rows.Add()
                                                        Data_table_Parcels.Rows(Index_data_parcels).Item("X1") = Data_table_unique.Rows(s).Item("X")
                                                        Data_table_Parcels.Rows(Index_data_parcels).Item("Y1") = Data_table_unique.Rows(s).Item("Y")
                                                        Data_table_Parcels.Rows(Index_data_parcels).Item("X2") = Data_table_unique.Rows(s + 1).Item("X")
                                                        Data_table_Parcels.Rows(Index_data_parcels).Item("Y2") = Data_table_unique.Rows(s + 1).Item("Y")
                                                        Data_table_Parcels.Rows(Index_data_parcels).Item("BEGSTA") = Data_table_unique.Rows(s).Item("STATION")
                                                        Data_table_Parcels.Rows(Index_data_parcels).Item("ENDSTA") = Data_table_unique.Rows(s + 1).Item("STATION")
                                                        Data_table_Parcels.Rows(Index_data_parcels).Item("FIELD1") = Data_table_unique.Rows(s).Item("FIELD1")
                                                        Data_table_Parcels.Rows(Index_data_parcels).Item("FIELD2") = Data_table_unique.Rows(s).Item("FIELD2")
                                                        Data_table_Parcels.Rows(Index_data_parcels).Item("FIELD3") = Data_table_unique.Rows(s).Item("FIELD3")
                                                        Index_data_parcels = Index_data_parcels + 1
                                                    Next

                                                End If
                                            End If
                                        End If
                                    Else


                                    End If
                                Next

                                Trans1.Abort()
                            End Using
                        End Using
                    End If
                End If






                If Data_table_Matchlines.Rows.Count > 0 Then

                    Dim Index_data_table_add As Integer = Data_table_Parcels.Rows.Count
                    Dim Nr_rand As Integer = Data_table_Parcels.Rows.Count


                    For i = 0 To Nr_rand - 1
                        If IsDBNull(Data_table_Parcels.Rows(i).Item("BEGSTA")) = False And IsDBNull(Data_table_Parcels.Rows(i).Item("ENDSTA")) = False Then
                            Error_index1 = i

                            Dim Station1 As Double = Data_table_Parcels.Rows(i).Item("BEGSTA")
                            Dim Station2 As Double = Data_table_Parcels.Rows(i).Item("ENDSTA")



                            Dim LL1 As String = ""
                            Dim TRACT1 As String = ""
                            Dim ROAD1 As Double = 0
                            Dim X1 As Double = 0
                            Dim X2 As Double = 0
                            Dim Y1 As Double = 0
                            Dim Y2 As Double = 0

                            If IsDBNull(Data_table_Parcels.Rows(i).Item("FIELD1")) = False Then
                                LL1 = Data_table_Parcels.Rows(i).Item("FIELD1")
                            End If
                            If IsDBNull(Data_table_Parcels.Rows(i).Item("FIELD2")) = False Then
                                TRACT1 = Data_table_Parcels.Rows(i).Item("FIELD2")
                            End If
                            If IsDBNull(Data_table_Parcels.Rows(i).Item("FIELD3")) = False Then
                                ROAD1 = Data_table_Parcels.Rows(i).Item("FIELD3")
                            End If

                            If IsDBNull(Data_table_Parcels.Rows(i).Item("X1")) = False Then
                                X1 = Data_table_Parcels.Rows(i).Item("X1")
                            End If
                            If IsDBNull(Data_table_Parcels.Rows(i).Item("X2")) = False Then
                                X2 = Data_table_Parcels.Rows(i).Item("X2")
                            End If
                            If IsDBNull(Data_table_Parcels.Rows(i).Item("Y1")) = False Then
                                Y1 = Data_table_Parcels.Rows(i).Item("Y1")
                            End If
                            If IsDBNull(Data_table_Parcels.Rows(i).Item("Y2")) = False Then
                                Y2 = Data_table_Parcels.Rows(i).Item("Y2")
                            End If

                            Dim I_start As Integer = 0
                            Dim go_to_add_S1_S2 As Boolean = False




123:





                            For j = I_start To Data_table_Matchlines.Rows.Count - 1
                                If IsDBNull(Data_table_Matchlines.Rows(j).Item("BEGSTA")) = False And IsDBNull(Data_table_Matchlines.Rows(j).Item("ENDSTA")) = False Then


                                    Error_index2 = j
                                    If i = 119 And j = 50 Then
                                        ' MsgBox("investigate")
                                    End If

                                    Dim M1 As Double = Data_table_Matchlines.Rows(j).Item("BEGSTA")
                                    Dim M2 As Double = Data_table_Matchlines.Rows(j).Item("ENDSTA")



                                    If IsNothing(PolyCL3D) = False Then
                                        If M2 > PolyCL3D.Length Then
                                            M2 = PolyCL3D.Length
                                        End If

                                        If Station2 > PolyCL3D.Length Then
                                            Station2 = PolyCL3D.Length
                                        End If
                                    Else
                                        If M2 > PolyCL.Length Then
                                            M2 = PolyCL.Length
                                        End If
                                        If Station2 > PolyCL.Length Then
                                            Station2 = PolyCL.Length
                                        End If
                                    End If



                                    If go_to_add_S1_S2 = True Then
                                        GoTo label_add_S1_S2
                                    End If




                                    'case 1
                                    If M1 <= Station1 And M2 <= Station2 And M1 <= Station2 And M2 >= Station1 Then
                                        Data_table_Parcels.Rows(i).Item("ENDSTA") = M2
                                        Station1 = M2
                                        I_start = j + 1
                                        go_to_add_S1_S2 = True
                                        GoTo 123
                                    End If

                                    ' case 5
                                    If Station1 >= M1 And Station2 <= M2 Then
                                        Exit For
                                    End If




label_add_S1_S2:

                                    ' add S1, S2
                                    If Station1 >= M1 And Station2 <= M2 Then
                                        Data_table_Parcels.Rows.Add()
                                        Data_table_Parcels.Rows(Index_data_table_add).Item("BEGSTA") = Station1
                                        Data_table_Parcels.Rows(Index_data_table_add).Item("ENDSTA") = Station2

                                        If IsNothing(PolyCL3D) = False Then
                                            Data_table_Parcels.Rows(Index_data_table_add).Item("X2") = PolyCL3D.GetPointAtDist(Station2).X
                                            Data_table_Parcels.Rows(Index_data_table_add).Item("Y2") = PolyCL3D.GetPointAtDist(Station2).Y
                                            Data_table_Parcels.Rows(Index_data_table_add).Item("X1") = PolyCL3D.GetPointAtDist(Station1).X
                                            Data_table_Parcels.Rows(Index_data_table_add).Item("Y1") = PolyCL3D.GetPointAtDist(Station1).Y
                                        Else

                                            Data_table_Parcels.Rows(Index_data_table_add).Item("X2") = PolyCL.GetPointAtDist(Station2).X
                                            Data_table_Parcels.Rows(Index_data_table_add).Item("Y2") = PolyCL.GetPointAtDist(Station2).Y
                                            Data_table_Parcels.Rows(Index_data_table_add).Item("X1") = PolyCL.GetPointAtDist(Station1).X
                                            Data_table_Parcels.Rows(Index_data_table_add).Item("Y1") = PolyCL.GetPointAtDist(Station1).Y
                                        End If

                                        Data_table_Parcels.Rows(Index_data_table_add).Item("FIELD1") = LL1
                                        Data_table_Parcels.Rows(Index_data_table_add).Item("FIELD2") = TRACT1
                                        Data_table_Parcels.Rows(Index_data_table_add).Item("FIELD3") = ROAD1
                                        Index_data_table_add = Index_data_table_add + 1
                                        Exit For

                                    ElseIf Station1 <= M2 And Station1 >= M1 Then
                                        Data_table_Parcels.Rows.Add()
                                        Data_table_Parcels.Rows(Index_data_table_add).Item("BEGSTA") = Station1
                                        Data_table_Parcels.Rows(Index_data_table_add).Item("ENDSTA") = M2

                                        If IsNothing(PolyCL3D) = False Then
                                            Data_table_Parcels.Rows(Index_data_table_add).Item("X2") = PolyCL3D.GetPointAtDist(M2).X
                                            Data_table_Parcels.Rows(Index_data_table_add).Item("Y2") = PolyCL3D.GetPointAtDist(M2).Y
                                            Data_table_Parcels.Rows(Index_data_table_add).Item("X1") = PolyCL3D.GetPointAtDist(Station1).X
                                            Data_table_Parcels.Rows(Index_data_table_add).Item("Y1") = PolyCL3D.GetPointAtDist(Station1).Y
                                        Else

                                            Data_table_Parcels.Rows(Index_data_table_add).Item("X2") = PolyCL.GetPointAtDist(M2).X
                                            Data_table_Parcels.Rows(Index_data_table_add).Item("Y2") = PolyCL.GetPointAtDist(M2).Y
                                            Data_table_Parcels.Rows(Index_data_table_add).Item("X1") = PolyCL.GetPointAtDist(Station1).X
                                            Data_table_Parcels.Rows(Index_data_table_add).Item("Y1") = PolyCL.GetPointAtDist(Station1).Y
                                        End If

                                        Data_table_Parcels.Rows(Index_data_table_add).Item("FIELD1") = LL1
                                        Data_table_Parcels.Rows(Index_data_table_add).Item("FIELD2") = TRACT1
                                        Data_table_Parcels.Rows(Index_data_table_add).Item("FIELD3") = ROAD1
                                        Index_data_table_add = Index_data_table_add + 1


                                        Station1 = M2
                                        I_start = j + 1
                                        go_to_add_S1_S2 = True
                                        GoTo 123

                                    End If

                                End If




                            Next

                        End If
                    Next



                End If



                Data_table_Parcels = Sort_data_table(Data_table_Parcels, "BEGSTA")

                Transfer_datatable_to_new_excel_spreadsheet(Data_table_Parcels)


                Label_od_LOADED.Text = Label_od_LOADED.Text & "  " & Data_table_Parcels.Rows.Count & " items"

                TextBox_debug_row2.Text = (Data_table_Parcels.Rows.Count + 1).ToString

                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")

            Catch ex As Exception
                Editor1.WriteMessage(vbLf & "Command:")
                Freeze_operations = False
                MsgBox(ex.Message)

                MsgBox(Error_index1 & vbCrLf & Error_index2)

            End Try


            Freeze_operations = False
        End If
    End Sub


    Public Function Remove_spaces_start_end(ByVal String1 As String) As String
        Dim Left1 As Boolean = False
        Do While Left1 = False
            If Strings.Left(String1, 1) = " " Then
                String1 = Mid(String1, 2)
            Else
                Left1 = True
            End If
        Loop
        Dim Right1 As Boolean = False
        Do While Right1 = False
            If Strings.Right(String1, 1) = " " Then
                String1 = Strings.Left(String1, Len(String1) - 1)
            Else
                Right1 = True
            End If
        Loop
        Return String1
    End Function


    Private Sub ComboBox_blocks_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_blocks.SelectedIndexChanged
        Incarca_existing_Atributes_to_combobox(ComboBox_blocks.Text, ComboBox_line_list_attribute)
        If ComboBox_line_list_attribute.Items.Count > 1 Then
            ComboBox_line_list_attribute.SelectedIndex = 1
        Else
            ComboBox_line_list_attribute.SelectedIndex = 0
        End If

        Incarca_existing_Atributes_to_combobox(ComboBox_blocks.Text, ComboBox_TRACT_attribute)
        If ComboBox_TRACT_attribute.Items.Count > 2 Then
            ComboBox_TRACT_attribute.SelectedIndex = 2
        Else
            ComboBox_TRACT_attribute.SelectedIndex = 0
        End If

        Incarca_existing_Atributes_to_combobox(ComboBox_blocks.Text, ComboBox_LENGTH_attribute)
        If ComboBox_LENGTH_attribute.Items.Count > 3 Then
            ComboBox_LENGTH_attribute.SelectedIndex = 3
        Else
            ComboBox_LENGTH_attribute.SelectedIndex = 0
        End If
    End Sub


    Private Sub RadioButton_use_mtext_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton_use_mtext.CheckedChanged, RadioButton_use_block.CheckedChanged
        If RadioButton_use_mtext.Checked = True Then
            Panel_Blocks.Visible = False
        Else
            Panel_Blocks.Visible = True
        End If
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




    Private Sub Button_draw_all_property_bands_Click(sender As Object, e As EventArgs) Handles Button_draw_all_bands.Click
        If IsNumeric(TextBox_viewport_Height.Text) = False Then
            MsgBox("Not numeric viewport height specified")
            Exit Sub
        End If
        If IsNumeric(TextBox_viewport_SCALE.Text) = False Then
            MsgBox("Not numeric viewport scale specified")
            Exit Sub
        End If
        If IsNumeric(TextBox_viewport_Width.Text) = False Then
            MsgBox("Not numeric viewport width specified")
            Exit Sub
        End If
        If IsNumeric(TextBox_X_MS.Text) = False Then
            MsgBox("Not numeric 0+000 X position in modelspace specified")
            Exit Sub
        End If
        If IsNumeric(TextBox_Y_MS.Text) = False Then
            MsgBox("Not numeric 0+000 Y position in modelspace specified")
            Exit Sub
        End If
        If Freeze_operations = False Then
            Freeze_operations = True
            Dim Round1 As Integer = 0
            If IsNumeric(TextBox_dec1.Text) = True Then
                Round1 = CInt(TextBox_dec1.Text)
            End If

            Try
                If IsNothing(Data_table_Parcels) = False And IsNothing(Data_table_Matchlines) = False Then
                    If Data_table_Parcels.Rows.Count > 0 Then



                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor

                        Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)


                                Dim Text_style_table As Autodesk.AutoCAD.DatabaseServices.TextStyleTable = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                Dim Layer_table As Autodesk.AutoCAD.DatabaseServices.LayerTable = Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                                Dim Start_point_for_bands As New Point3d(CDbl(TextBox_X_MS.Text), CDbl(TextBox_Y_MS.Text), 0)



                                Dim Bands_Y_spacing As Double = 500




                                Dim Line_len As Double = Abs(CDbl(TextBox_viewport_Height.Text) * CDbl(TextBox_viewport_SCALE.Text))

                                Bands_Y_spacing = Ceiling(2.25 * Line_len / 50) * 50
                                TextBox_BAND_SPACING.Text = Bands_Y_spacing





                                If Line_len = 0 Then
                                    MsgBox("y1=y2")
                                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                    Freeze_operations = False
                                    Exit Sub
                                End If





                                Dim Min_dist As Double = 225


                                If IsNumeric(TextBox_minimum_distance.Text) = True Then
                                    Min_dist = CDbl(TextBox_minimum_distance.Text)
                                End If

                                If RadioButton_use_mtext.Checked = True Then
                                    If IsNumeric(TextBox_minimum_distance_mtext.Text) = True Then
                                        Min_dist = CDbl(TextBox_minimum_distance_mtext.Text)
                                    End If
                                End If




                                Dim Max_dist_rotation As Double = 220


                                Dim Text_offset_ptr_blocks_X As Double = 5
                                If IsNumeric(TextBox_textOffset_X.Text) = True Then
                                    Text_offset_ptr_blocks_X = CDbl(TextBox_textOffset_X.Text)
                                End If

                                Dim Text_offset_ptr_blocks_Y As Double = 5
                                If IsNumeric(TextBox_textOffset_Y.Text) = True Then
                                    Text_offset_ptr_blocks_Y = CDbl(TextBox_textOffset_Y.Text)
                                End If


                                Dim TextHeight As Double = 16
                                If IsNumeric(TextBox_text_height.Text) = True Then
                                    TextHeight = CDbl(TextBox_text_height.Text)
                                End If

                                Dim Mtext_height As Double = TextHeight
                                Dim Mtext_rotation As Double = PI / 2

                                Dim Block_scale As Double = 1









                                Creaza_layer(No_plot, 40, No_plot, False)


                                Dim Station2_prev As Double = 0

                                Dim X2prev As Double = Start_point_for_bands.X

                                For i = 0 To Data_table_Parcels.Rows.Count - 1
                                    If IsDBNull(Data_table_Parcels.Rows(i).Item("BEGSTA")) = False And IsDBNull(Data_table_Parcels.Rows(i).Item("ENDSTA")) = False Then
                                        Dim Station1 As Double = Data_table_Parcels.Rows(i).Item("BEGSTA")
                                        Dim Station2 As Double = Data_table_Parcels.Rows(i).Item("ENDSTA")
                                        Dim Length1 As Double = Station2 - Station1
                                        Dim LineList As String = "XXX"
                                        Dim Tract As String = "XXX"
                                        If IsDBNull(Data_table_Parcels.Rows(i).Item("FIELD1")) = False Then
                                            LineList = Data_table_Parcels.Rows(i).Item("FIELD1")
                                        End If
                                        If IsDBNull(Data_table_Parcels.Rows(i).Item("FIELD2")) = False Then
                                            Tract = Data_table_Parcels.Rows(i).Item("FIELD2")
                                        End If

                                        If RadioButton_use_block.Checked = True Then
                                            If IsDBNull(Data_table_Parcels.Rows(i).Item("FIELD3")) = False Then
                                                Dim rot_road As Double = Data_table_Parcels.Rows(i).Item("FIELD3")
                                                If Not rot_road = 0 Then
                                                    Min_dist = 110
                                                Else
                                                    If IsNumeric(TextBox_minimum_distance.Text) = True Then
                                                        Min_dist = CDbl(TextBox_minimum_distance.Text)
                                                    End If
                                                End If


                                            End If
                                        End If





                                        Dim Viewport_line As New Line(New Point3d(0, 0, 0), New Point3d(0, 0, 0))

                                        Dim Point_B1 As New Point3d
                                        Dim Point_B2 As New Point3d

                                        Dim Dist_from_start1 As Double
                                        Dim Dist_from_start2 As Double

                                        Dim M1 As Double = 0
                                        Dim M2 As Double = 0

                                        Dim Band_number As Integer = -1

                                        For j = 0 To Data_table_Matchlines.Rows.Count - 1
                                            If IsDBNull(Data_table_Matchlines.Rows(j).Item("BEGSTA")) = False And IsDBNull(Data_table_Matchlines.Rows(j).Item("ENDSTA")) = False Then

                                                M1 = Data_table_Matchlines.Rows(j).Item("BEGSTA")
                                                M2 = Data_table_Matchlines.Rows(j).Item("ENDSTA")
                                                Band_number = Band_number + 1
                                                If Round(Station1, Round1) <= Round(M2, Round1) And Round(Station1, Round1) >= Round(M1, Round1) And Round(Station2, Round1) <= Round(M2, Round1) And Round(Station2, Round1) >= Round(M1, Round1) Then
                                                    If IsDBNull(Data_table_Matchlines.Rows(j).Item("X1")) = False And IsDBNull(Data_table_Matchlines.Rows(j).Item("Y1")) = False Then
                                                        If IsDBNull(Data_table_Matchlines.Rows(j).Item("X2")) = False And IsDBNull(Data_table_Matchlines.Rows(j).Item("Y2")) = False Then
                                                            Dim Start1 As New Point3d
                                                            Dim End1 As New Point3d

                                                            Start1 = New Point3d(Data_table_Matchlines.Rows(j).Item("X1"), Data_table_Matchlines.Rows(j).Item("Y1"), 0)
                                                            End1 = New Point3d(Data_table_Matchlines.Rows(j).Item("X2"), Data_table_Matchlines.Rows(j).Item("Y2"), 0)

                                                            If IsNothing(PolyCL3D) = False Then

                                                                Start1 = PolyCL.GetClosestPointTo(Start1, Vector3d.ZAxis, False)
                                                                End1 = PolyCL.GetClosestPointTo(End1, Vector3d.ZAxis, False)
                                                                Dim Param1 As Double = PolyCL.GetParameterAtPoint(Start1)
                                                                Dim Param2 As Double = PolyCL.GetParameterAtPoint(End1)

                                                                Dim Point1 As New Point3d
                                                                Point1 = PolyCL3D.GetPointAtParameter(Param1)
                                                                Dim Point2 As New Point3d
                                                                Point2 = PolyCL3D.GetPointAtParameter(Param2)

                                                                If PolyCL3D.GetPointAtDist(M1).GetVectorTo(Point2).Length < PolyCL3D.GetPointAtDist(M1).GetVectorTo(Point1).Length Then
                                                                    Dim ttt As New Point3d
                                                                    ttt = Start1
                                                                    Start1 = End1
                                                                    End1 = ttt
                                                                End If

                                                            Else
                                                                If PolyCL.GetPointAtDist(M1).GetVectorTo(End1).Length < PolyCL.GetPointAtDist(M1).GetVectorTo(Start1).Length Then
                                                                    Dim ttt As New Point3d
                                                                    ttt = Start1
                                                                    Start1 = End1
                                                                    End1 = ttt
                                                                End If
                                                            End If
                                                            Viewport_line = New Line(Start1, End1)

                                                            If RadioButton_Left_right.Checked = True Then
                                                                Start_point_for_bands = New Point3d(CDbl(TextBox_X_MS.Text) + (CDbl(TextBox_viewport_Width.Text) / 2) * CDbl(TextBox_viewport_SCALE.Text) - Viewport_line.Length / 2, Start_point_for_bands.Y, 0)
                                                            Else
                                                                Start_point_for_bands = New Point3d(CDbl(TextBox_X_MS.Text) - (CDbl(TextBox_viewport_Width.Text) / 2) * CDbl(TextBox_viewport_SCALE.Text) + Viewport_line.Length / 2, Start_point_for_bands.Y, 0)
                                                            End If

                                                            Exit For
                                                        End If
                                                    End If





                                                End If



                                            End If
                                        Next

                                        If Viewport_line.Length > 0 Then

                                            If IsNothing(PolyCL3D) = False Then
                                                Dim Point1 As New Point3d
                                                Point1 = PolyCL3D.GetPointAtDist(Station1)

                                                Dim Point2 As New Point3d
                                                Point2 = PolyCL3D.GetPointAtDist(Station2)

                                                Point1 = New Point3d(Point1.X, Point1.Y, 0)
                                                Point2 = New Point3d(Point2.X, Point2.Y, 0)

                                                Dim PointV1 As New Point3d
                                                PointV1 = Viewport_line.GetClosestPointTo(Point1, Vector3d.ZAxis, False)
                                                Dim PointV2 As New Point3d
                                                PointV2 = Viewport_line.GetClosestPointTo(Point2, Vector3d.ZAxis, False)
                                                Dist_from_start1 = PointV1.GetVectorTo(Viewport_line.StartPoint).Length
                                                Dist_from_start2 = PointV2.GetVectorTo(Viewport_line.StartPoint).Length


                                            Else
                                                If Station1 > PolyCL.Length Then
                                                    Station1 = PolyCL.Length
                                                End If
                                                Dim Point1 As New Point3d
                                                Point1 = PolyCL.GetPointAtDist(Station1)

                                                Dim Point2 As New Point3d
                                                Point2 = PolyCL.GetPointAtDist(Station2)


                                                Dim PointV1 As New Point3d
                                                PointV1 = Viewport_line.GetClosestPointTo(Point1, Vector3d.ZAxis, False)
                                                Dim PointV2 As New Point3d
                                                PointV2 = Viewport_line.GetClosestPointTo(Point2, Vector3d.ZAxis, False)
                                                Dist_from_start1 = PointV1.GetVectorTo(Viewport_line.StartPoint).Length
                                                Dist_from_start2 = PointV2.GetVectorTo(Viewport_line.StartPoint).Length
                                            End If


                                            Dim Width1 As Double = Dist_from_start2 - Dist_from_start1
                                            If Width1 < Min_dist Then
                                                Width1 = Min_dist
                                            End If

                                            If RadioButton_Left_right.Checked = True Then
                                                Point_B1 = New Point3d(X2prev, Start_point_for_bands.Y - (Band_number * Bands_Y_spacing), 0)
                                                Point_B2 = New Point3d(Point_B1.X + Width1, Point_B1.Y, 0)
                                            Else
                                                Point_B1 = New Point3d(X2prev, Start_point_for_bands.Y - (Band_number * Bands_Y_spacing), 0)
                                                Point_B2 = New Point3d(Point_B1.X - Width1, Point_B1.Y, 0)
                                            End If



                                            If Not Round(Station1, Round1) = Round(Station2_prev, Round1) Then

                                                Length1 = Station2 - Station2_prev

                                                Dim Mtext1 As New MText
                                                Mtext1.Location = New Point3d(Point_B1.X + Text_offset_ptr_blocks_X, Start_point_for_bands.Y - (Band_number * Bands_Y_spacing) + Text_offset_ptr_blocks_Y, 0)
                                                Mtext1.TextHeight = Mtext_height * 2
                                                Mtext1.Rotation = Mtext_rotation
                                                Mtext1.Layer = No_plot
                                                Mtext1.ColorIndex = 1
                                                Dim ContinutMtext1 As String

                                                If CheckBox_use_equation.Checked = True Then

                                                    If IsNumeric(TextBox_textwidth.Text) = True Then
                                                        ContinutMtext1 = "{\W" & TextBox_textwidth.Text & ";" & Get_chainage_feet_from_double(Station1 + Get_equation_value(Station1), Round1) & "}"
                                                    Else
                                                        ContinutMtext1 = Get_chainage_feet_from_double(Station1 + Get_equation_value(Station1), Round1)
                                                    End If


                                                Else
                                                    If IsNumeric(TextBox_textwidth.Text) = True Then
                                                        ContinutMtext1 = "{\W" & TextBox_textwidth.Text & ";" & Get_chainage_feet_from_double(Station1, Round1) & "}"
                                                    Else
                                                        ContinutMtext1 = Get_chainage_feet_from_double(Station1, Round1)
                                                    End If

                                                End If



                                                Mtext1.Contents = ContinutMtext1



                                                Mtext1.Attachment = AttachmentPoint.TopLeft
                                                Dim ObjId1 As TextStyleTableRecord
                                                If Text_style_table.Has(ComboBox_text_style.Text) = True Then
                                                    ObjId1 = Text_style_table(ComboBox_text_style.Text).GetObject(OpenMode.ForRead)
                                                    Mtext1.TextStyleId = ObjId1.ObjectId
                                                End If


                                                BTrecord.AppendEntity(Mtext1)
                                                Trans1.AddNewlyCreatedDBObject(Mtext1, True)
                                            End If


                                            If Round(Station1, Round1) = Round(M1, Round1) Then

                                                Dim Band_text As New DBText
                                                Band_text.Layer = No_plot
                                                If RadioButton_Left_right.Checked = True Then

                                                    Band_text.Justify = AttachmentPoint.MiddleRight
                                                    Band_text.AlignmentPoint = New Point3d(Start_point_for_bands.X - 5 * Mtext_height, Start_point_for_bands.Y + Line_len / 2 - (Band_number * Bands_Y_spacing), 0)
                                                Else

                                                    Band_text.Justify = AttachmentPoint.MiddleLeft
                                                    Band_text.AlignmentPoint = New Point3d(Start_point_for_bands.X + 5 * Mtext_height, Start_point_for_bands.Y + Line_len / 2 - (Band_number * Bands_Y_spacing), 0)
                                                End If

                                                Band_text.TextString = CStr(Band_number + 1)
                                                Band_text.Height = 7.5 * Mtext_height

                                                BTrecord.AppendEntity(Band_text)
                                                Trans1.AddNewlyCreatedDBObject(Band_text, True)

                                                If RadioButton_Left_right.Checked = True Then
                                                    Point_B1 = New Point3d(Start_point_for_bands.X, Start_point_for_bands.Y - (Band_number * Bands_Y_spacing), 0)
                                                    Point_B2 = New Point3d(Point_B1.X + Width1, Point_B1.Y, 0)
                                                Else
                                                    Point_B1 = New Point3d(Start_point_for_bands.X, Start_point_for_bands.Y - (Band_number * Bands_Y_spacing), 0)
                                                    Point_B2 = New Point3d(Point_B1.X - Width1, Point_B1.Y, 0)
                                                End If


                                                Dim Linie1 As New Line(New Point3d(Point_B1.X, Start_point_for_bands.Y - (Band_number * Bands_Y_spacing), 0), New Point3d(Point_B1.X, Start_point_for_bands.Y - (Band_number * Bands_Y_spacing) + Line_len, 0))
                                                If Not ComboBox_layers.Text = "" Then
                                                    Linie1.Layer = ComboBox_layers.Text
                                                End If
                                                BTrecord.AppendEntity(Linie1)
                                                Trans1.AddNewlyCreatedDBObject(Linie1, True)

                                                Dim Mtext1 As New MText
                                                Mtext1.Location = New Point3d(Point_B1.X - Text_offset_ptr_blocks_X, Start_point_for_bands.Y - (Band_number * Bands_Y_spacing) + Text_offset_ptr_blocks_Y, 0)
                                                Mtext1.TextHeight = Mtext_height
                                                Mtext1.Rotation = Mtext_rotation
                                                If Layer_table.Has(ComboBox_layers.Text) = True Then
                                                    Mtext1.Layer = ComboBox_layers.Text
                                                End If


                                                Dim ContinutMtext1 As String

                                                If CheckBox_use_equation.Checked = True Then

                                                    If IsNumeric(TextBox_textwidth.Text) = True Then
                                                        ContinutMtext1 = "{\W" & TextBox_textwidth.Text & ";" & Get_chainage_feet_from_double(Station1 + Get_equation_value(Station1), Round1) & "}"
                                                    Else
                                                        ContinutMtext1 = Get_chainage_feet_from_double(Station1 + Get_equation_value(Station1), Round1)
                                                    End If

                                                Else
                                                    If IsNumeric(TextBox_textwidth.Text) = True Then
                                                        ContinutMtext1 = "{\W" & TextBox_textwidth.Text & ";" & Get_chainage_feet_from_double(Station1, Round1) & "}"
                                                    Else
                                                        ContinutMtext1 = Get_chainage_feet_from_double(Station1, Round1)
                                                    End If

                                                End If



                                                Mtext1.Contents = ContinutMtext1



                                                Mtext1.Attachment = AttachmentPoint.BottomLeft
                                                Dim ObjId1 As TextStyleTableRecord
                                                If Text_style_table.Has(ComboBox_text_style.Text) = True Then
                                                    ObjId1 = Text_style_table(ComboBox_text_style.Text).GetObject(OpenMode.ForRead)
                                                    Mtext1.TextStyleId = ObjId1.ObjectId
                                                End If


                                                BTrecord.AppendEntity(Mtext1)
                                                Trans1.AddNewlyCreatedDBObject(Mtext1, True)



                                            End If 'este de la If Station1 = M1 

                                            Dim Linie2 As New Line(New Point3d(Point_B2.X, Start_point_for_bands.Y - (Band_number * Bands_Y_spacing), 0), New Point3d(Point_B2.X, Start_point_for_bands.Y - (Band_number * Bands_Y_spacing) + Line_len, 0))
                                            If Not ComboBox_layers.Text = "" Then
                                                Linie2.Layer = ComboBox_layers.Text
                                            End If
                                            BTrecord.AppendEntity(Linie2)
                                            Trans1.AddNewlyCreatedDBObject(Linie2, True)

                                            Dim Mtext2 As New MText
                                            Mtext2.Location = New Point3d(Point_B2.X - Text_offset_ptr_blocks_X, Start_point_for_bands.Y - (Band_number * Bands_Y_spacing) + Text_offset_ptr_blocks_Y, 0)
                                            Mtext2.TextHeight = Mtext_height
                                            Mtext2.Rotation = Mtext_rotation
                                            If Layer_table.Has(ComboBox_layers.Text) = True Then
                                                Mtext2.Layer = ComboBox_layers.Text
                                            End If


                                            Dim ContinutMtext2 As String

                                            If CheckBox_use_equation.Checked = True Then

                                                If IsNumeric(TextBox_textwidth.Text) = True Then
                                                    ContinutMtext2 = "{\W" & TextBox_textwidth.Text & ";" & Get_chainage_feet_from_double(Station2 + Get_equation_value(Station2), Round1) & "}"
                                                Else
                                                    ContinutMtext2 = Get_chainage_feet_from_double(Station2 + Get_equation_value(Station2), Round1)
                                                End If
                                            Else
                                                If IsNumeric(TextBox_textwidth.Text) = True Then
                                                    ContinutMtext2 = "{\W" & TextBox_textwidth.Text & ";" & Get_chainage_feet_from_double(Station2, Round1) & "}"
                                                Else
                                                    ContinutMtext2 = Get_chainage_feet_from_double(Station2, Round1)
                                                End If

                                            End If



                                            Mtext2.Contents = ContinutMtext2



                                            Mtext2.Attachment = AttachmentPoint.BottomLeft
                                            Dim ObjId2 As TextStyleTableRecord
                                            If Text_style_table.Has(ComboBox_text_style.Text) = True Then
                                                ObjId2 = Text_style_table(ComboBox_text_style.Text).GetObject(OpenMode.ForRead)
                                                Mtext2.TextStyleId = ObjId2.ObjectId
                                            End If


                                            BTrecord.AppendEntity(Mtext2)
                                            Trans1.AddNewlyCreatedDBObject(Mtext2, True)

                                            Dim Insertion_point As New Point3d
                                            Insertion_point = New Point3d((Point_B1.X + Point_B2.X) / 2, Point_B1.Y + Line_len / 2, 0)


                                            If RadioButton_use_block.Checked = True Then
                                                Dim Colectie_atrib_value As New Specialized.StringCollection
                                                Dim Colectie_atrib_name As New Specialized.StringCollection
                                                Colectie_atrib_value.Add(LineList)
                                                Colectie_atrib_name.Add(ComboBox_line_list_attribute.Text)
                                                Colectie_atrib_value.Add(Tract)
                                                Colectie_atrib_name.Add(ComboBox_TRACT_attribute.Text)
                                                Colectie_atrib_value.Add(Get_String_Rounded(Length1, Round1) & "'")
                                                Colectie_atrib_name.Add(ComboBox_LENGTH_attribute.Text)


                                                If Not ComboBox_blocks.Text = "" Then


                                                    Dim Block1 As BlockReference
                                                    Block1 = InsertBlock_with_multiple_atributes("", ComboBox_blocks.Text, Insertion_point, Block_scale, BTrecord, ComboBox_layers.Text, Colectie_atrib_name, Colectie_atrib_value)
                                                    Dim rot_road As Double = Data_table_Parcels.Rows(i).Item("FIELD3")
                                                    Block1.Rotation = rot_road

                                                End If
                                            Else

                                                Dim Mtext_ll As New MText
                                                Mtext_ll.Contents = LineList
                                                Mtext_ll.TextHeight = CDbl(TextBox_mtext_height.Text)
                                                Mtext_ll.Location = Insertion_point
                                                Mtext_ll.Rotation = 0
                                                If Layer_table.Has(ComboBox_layers_for_Mtext.Text) = True Then
                                                    Mtext_ll.Layer = ComboBox_layers_for_Mtext.Text
                                                End If
                                                Mtext_ll.Attachment = AttachmentPoint.MiddleCenter
                                                Dim ObjId1 As TextStyleTableRecord
                                                If Text_style_table.Has(ComboBox_text_style_mtext.Text) = True Then
                                                    ObjId1 = Text_style_table(ComboBox_text_style_mtext.Text).GetObject(OpenMode.ForRead)
                                                    Mtext_ll.TextStyleId = ObjId1.ObjectId
                                                End If
                                                BTrecord.AppendEntity(Mtext_ll)
                                                Trans1.AddNewlyCreatedDBObject(Mtext_ll, True)

                                            End If


                                            X2prev = Point_B2.X



                                            Station2_prev = Station2

                                        End If 'asta e de la viewport length>0

                                    End If ' asta e de la If IsDBNull(Data_table_Parcels.Rows(i).Item("BEGSTA")) = False And IsDBNull(Data_table_Parcels.Rows(i).Item("ENDSTA")) = False
                                Next





                                Trans1.Commit()
                            End Using
                        End Using


                        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                    End If
                Else

                    MsgBox("You did not have data loaded for matchlines", MsgBoxStyle.Critical, "Dan says...")
                End If


            Catch ex As Exception
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub



    Private Sub Button_read_centerline_and_viewports_Click(sender As Object, e As EventArgs) Handles Button_read_centerline_and_viewports.Click, Button_read_cl_and_VW.Click, Button_load_cl_viewports_mtext.Click, Button_read_CL_VIEWPORTS_EXCEL.Click, Button_read_CL_Water.Click

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
                Colectie1 = New Specialized.StringCollection


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

                                Data_table_Matchlines = New System.Data.DataTable
                                Data_table_Matchlines.Columns.Add("BEGSTA", GetType(Double))
                                Data_table_Matchlines.Columns.Add("ENDSTA", GetType(Double))
                                Data_table_Matchlines.Columns.Add("X1", GetType(Double))
                                Data_table_Matchlines.Columns.Add("Y1", GetType(Double))
                                Data_table_Matchlines.Columns.Add("X2", GetType(Double))
                                Data_table_Matchlines.Columns.Add("Y2", GetType(Double))

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


                                                Data_table_Matchlines.Rows.Add()
                                                Data_table_Matchlines.Rows(Index_dataT).Item("BEGSTA") = Round(Station1, Round1)
                                                Data_table_Matchlines.Rows(Index_dataT).Item("ENDSTA") = Round(Station2, Round1)
                                                Data_table_Matchlines.Rows(Index_dataT).Item("X1") = X11 'Viewport_poly.GetPointAtParameter(3).X
                                                Data_table_Matchlines.Rows(Index_dataT).Item("Y1") = Y11 'Viewport_poly.GetPointAtParameter(3).Y
                                                Data_table_Matchlines.Rows(Index_dataT).Item("X2") = X21 'Viewport_poly.GetPointAtParameter(2).X
                                                Data_table_Matchlines.Rows(Index_dataT).Item("Y2") = Y22 'Viewport_poly.GetPointAtParameter(2).Y
                                                Index_dataT = Index_dataT + 1


                                            End If


                                        End If

                                    End If


                                Next

                                Data_table_Matchlines = Sort_data_table(Data_table_Matchlines, "BEGSTA")

                                If Data_table_Matchlines.Rows.Count > 0 Then
                                    Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                                    W1 = Get_NEW_worksheet_from_Excel()
                                    Dim Column_match As String = "A"
                                    Dim Idx_col As Integer = W1.Range(Column_match & "1").Column

                                    For i = 0 To Data_table_Matchlines.Rows.Count - 1
                                        W1.Range(Column_match & (i + 2).ToString).Value2 = Data_table_Matchlines.Rows(i).Item("BEGSTA") & " - " & Data_table_Matchlines.Rows(i).Item("ENDSTA")

                                        W1.Cells(i + 2, Idx_col + 2).Value2 = Data_table_Matchlines.Rows(i).Item("BEGSTA")
                                        W1.Cells(i + 2, Idx_col + 3).Value2 = Data_table_Matchlines.Rows(i).Item("ENDSTA")
                                        W1.Cells(i + 2, Idx_col + 4).FormulaR1C1 = "=RC[-2]-R[-1]C[-1]"
                                    Next
                                    W1.Cells(Data_table_Matchlines.Rows.Count + 2, Idx_col + 4).FormulaR1C1 = "=SUM(R[-" & (Data_table_Matchlines.Rows.Count).ToString & "]C:R[-1]C)"


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


    Private Sub Button_crossing_band_load_From_excel_Click(sender As Object, e As EventArgs) Handles Button_crossing_band_load_From_excel.Click
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

                If End1 = 0 Or Start1 = 0 Then
                    Freeze_operations = False
                    Exit Sub
                End If
                If End1 < Start1 Then
                    Freeze_operations = False
                    Exit Sub
                End If
                Dim Column_sta As String = ""
                Column_sta = TextBox_column_Station_excel_data.Text.ToUpper
                Dim Column_descr As String = ""
                Column_descr = TextBox_column_description_excel_data.Text.ToUpper

                Data_table_crossing_band_Excel_data = New System.Data.DataTable
                Data_table_crossing_band_Excel_data.Columns.Add("STATION", GetType(Double))
                Data_table_crossing_band_Excel_data.Columns.Add("DESCRIPTION", GetType(String))
                Data_table_crossing_band_Excel_data.Columns.Add("TYPE", GetType(String))

                Dim Index_data_table As Double



                For i = Start1 To End1
                    Dim Station_string As String = W1.Range(Column_sta & i).Value2
                    Dim Descriptie As String = W1.Range(Column_descr & i).Value2
                    If IsNumeric(Station_string) = True And Not Descriptie = "" Then
                        If CDbl(Station_string) >= 0 Then
                            Data_table_crossing_band_Excel_data.Rows.Add()
                            Data_table_crossing_band_Excel_data.Rows(Index_data_table).Item("STATION") = CDbl(Station_string)
                            Data_table_crossing_band_Excel_data.Rows(Index_data_table).Item("DESCRIPTION") = Descriptie
                            If Descriptie.Contains("'") = True And (Descriptie.Contains("RT") = True Or Descriptie.Contains("LT") = True) And (Strings.Right(Descriptie, 2) = "RT" Or Strings.Right(Descriptie, 2) = "LT") Then
                                Data_table_crossing_band_Excel_data.Rows(Index_data_table).Item("TYPE") = "DEFLECTION"
                            End If

                            Index_data_table = Index_data_table + 1
                        End If
                    End If
                Next


                Data_table_crossing_band_Excel_data = Sort_data_table(Data_table_crossing_band_Excel_data, "STATION")

                'MsgBox(Data_table_Centerline.Rows.Count)



            Catch ex As Exception
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try
        End If
        Freeze_operations = False
    End Sub

    Private Sub TabPage5_Click(sender As Object, e As EventArgs) Handles TabPage_excel_data.Click
        Incarca_existing_layers_to_combobox(ComboBox_layers_deflections)
        Incarca_existing_layers_to_combobox(ComboBox_layers_crossings)

        If ComboBox_layers_deflections.Items.Count > 0 Then
            If ComboBox_layers_deflections.Items.Contains("TEXT_PI") = True Then
                ComboBox_layers_deflections.SelectedIndex = ComboBox_layers_deflections.Items.IndexOf("TEXT_PI")
            Else
                ComboBox_layers_deflections.SelectedIndex = 0
            End If
        End If
        If ComboBox_layers_crossings.Items.Count > 0 Then
            If ComboBox_layers_crossings.Items.Contains("Text") = True Then
                ComboBox_layers_crossings.SelectedIndex = ComboBox_layers_crossings.Items.IndexOf("Text")
            Else
                ComboBox_layers_crossings.SelectedIndex = 0
            End If
        End If

        Incarca_existing_textstyles_to_combobox(ComboBox_text_styles_excel_data)

        If ComboBox_text_styles_excel_data.Items.Count > 0 Then
            If ComboBox_text_styles_excel_data.Items.Contains("ALIGNDB") = True Then
                ComboBox_text_styles_excel_data.SelectedIndex = ComboBox_text_styles_excel_data.Items.IndexOf("ALIGNDB")
            Else
                ComboBox_text_styles_excel_data.SelectedIndex = 0
            End If
        End If
    End Sub




    Private Sub Button_DRAW_EXCEL_DATA_Click(sender As Object, e As EventArgs) Handles Button_DRAW_EXCEL_DATA.Click

        If IsNumeric(TextBox_viewport_Height.Text) = False Then
            MsgBox("Not numeric viewport height specified")
            Exit Sub
        End If
        If IsNumeric(TextBox_viewport_SCALE.Text) = False Then
            MsgBox("Not numeric viewport scale specified")
            Exit Sub
        End If
        If IsNumeric(TextBox_X_MS.Text) = False Then
            MsgBox("Not numeric 0+000 X position in modelspace specified")
            Exit Sub
        End If
        If IsNumeric(TextBox_Y_MS.Text) = False Then
            MsgBox("Not numeric 0+000 Y position in modelspace specified")
            Exit Sub
        End If
        If IsNumeric(TextBox_Text_height_excel_data.Text) = False Then
            MsgBox("Not numeric text height specified")
            Exit Sub
        End If

        Dim CSF As Double = 1
        If IsNumeric(TextBox_CSF.Text) = True Then
            CSF = CDbl(TextBox_CSF.Text)
        End If

        If Freeze_operations = False Then
            Freeze_operations = True
            Dim Round1 As Integer = 0
            If IsNumeric(TextBox_dec1.Text) = True Then
                Round1 = CInt(TextBox_dec1.Text)
            End If
            Try
                If IsNothing(Data_table_crossing_band_Excel_data) = False Then
                    If Data_table_crossing_band_Excel_data.Rows.Count > 0 Then
                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()


                        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor

                        Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument

                            ' Dim k As Double = 1
                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                                Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord


                                BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                                'BTrecord = Trans1.GetObject(BlockTable_data1(BlockTableRecord.PaperSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                                BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)




                                Dim Text_style_table As Autodesk.AutoCAD.DatabaseServices.TextStyleTable = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                Dim ObjId_text_table_record As TextStyleTableRecord = Nothing

                                If Text_style_table.Has(ComboBox_text_styles_excel_data.Text) = True Then
                                    ObjId_text_table_record = Text_style_table(ComboBox_text_styles_excel_data.Text).GetObject(OpenMode.ForRead)
                                End If


                                Dim Layer_table As Autodesk.AutoCAD.DatabaseServices.LayerTable = Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                Dim Start_point_for_bands As New Point3d(CDbl(TextBox_X_MS.Text), CDbl(TextBox_Y_MS.Text), 0)


                                Dim Prefix As String = TextBox_prefix_excel_data.Text



                                Dim Line_len As Double = Abs(CDbl(TextBox_viewport_Height.Text) * CDbl(TextBox_viewport_SCALE.Text))

                                Dim Bands_Y_spacing As Double = 1000

                                If IsNumeric(TextBox_BAND_SPACING.Text) = True Then
                                    Bands_Y_spacing = CDbl(TextBox_BAND_SPACING.Text)
                                End If




                                Dim Point_MS As New Point3d

                                Dim Chainage_previous As Double = -1

                                Dim x_BLOCK As Double = 0


                                Dim Min_dist As Double = 45

                                If IsNumeric(TextBox_min_dist_excel_data.Text) = True Then
                                    Min_dist = CDbl(TextBox_min_dist_excel_data.Text)
                                End If


                                Dim TextHeight As Double = 16
                                If IsNumeric(TextBox_Text_height_excel_data.Text) = True Then
                                    TextHeight = CDbl(TextBox_Text_height_excel_data.Text)
                                End If

                                Dim Mtext_rotation As Double = PI / 2

                                If IsNumeric(TextBox_rotation_excel_data.Text) = True Then
                                    Mtext_rotation = PI * CDbl(TextBox_rotation_excel_data.Text) / 180
                                End If

                                If IsNothing(Data_table_Matchlines) = False Then
                                    If IsNothing(Data_table_crossing_band_Excel_data) = False Then

                                        Dim Nr_previous As Integer = -1
                                        Dim Station_previous As Double = -1
                                        Dim Match2_previous As Double = -1
                                        Dim Xprev As Double = 0
                                        Dim Line_viewport_previous As New Line
                                        Creaza_layer(No_plot, 40, No_plot, False)
                                        Dim Viewport_scale As Double = 1
                                        If IsNumeric(TextBox_viewport_SCALE.Text) = True Then
                                            Viewport_scale = CDbl(TextBox_viewport_SCALE.Text)
                                        End If
                                        Dim PolyTemp_2d As New Polyline

                                        If IsNothing(PolyCL) = True Then
                                            PolyTemp_2d = Creaza_polyline_din_polyline3d(Trans1, PolyCL3D)
                                        End If

                                        For i = 0 To Data_table_crossing_band_Excel_data.Rows.Count - 1
                                            If IsDBNull(Data_table_crossing_band_Excel_data.Rows(i).Item("STATION")) = False Then
                                                Dim Station As Double = Data_table_crossing_band_Excel_data.Rows(i).Item("STATION")
                                                If Station > Poly_length Then
                                                    Station = Poly_length
                                                End If
                                                Dim Type1 As String = ""

                                                Dim Descriptie As String = ""
                                                Dim Layer_xing As String = ""
                                                Dim TextWidth As Double = 0.8
                                                If IsNumeric(TextBox_textwidth.Text) = True Then
                                                    TextWidth = CDbl(TextBox_textwidth.Text)
                                                End If



                                                If IsDBNull(Data_table_crossing_band_Excel_data.Rows(i).Item("DESCRIPTION")) = False Then
                                                    Dim Format1 As String = "{\W" & TextWidth & ";"
                                                    Dim String1 As String = Data_table_crossing_band_Excel_data.Rows(i).Item("DESCRIPTION")
                                                    Layer_xing = ComboBox_layers_crossings.Text

                                                    If Is_canada = False Then
                                                        If String1.Contains("P.I.") = True Then
                                                            Format1 = "{\W" & TextWidth & ";\L"
                                                            Layer_xing = ComboBox_layers_deflections.Text
                                                        End If
                                                    Else

                                                        If IsDBNull(Data_table_crossing_band_Excel_data.Rows(i).Item("TYPE")) = False Then
                                                            Type1 = Data_table_crossing_band_Excel_data.Rows(i).Item("TYPE")
                                                        End If
                                                        If Type1 = "DEFLECTION" Then
                                                            Layer_xing = ComboBox_layers_deflections.Text
                                                        End If

                                                    End If


                                                    If CheckBox_use_equation.Checked = False Then
                                                        If Is_canada = False Then
                                                            If CheckBox_no_station.Checked = False Then
                                                                Descriptie = Format1 & Prefix & Get_chainage_feet_from_double(Station, Round1) & " " & String1 & "}"
                                                            Else
                                                                Descriptie = Format1 & Prefix & " " & String1 & "}"
                                                            End If

                                                        Else
                                                            Descriptie = Format1 & String1 & "}"
                                                        End If

                                                    Else
                                                        If Is_canada = False Then
                                                            If CheckBox_no_station.Checked = False Then
                                                                Descriptie = Format1 & Prefix & Get_chainage_feet_from_double(Station + Get_equation_value(Station), Round1) & " " & String1 & "}"
                                                            Else
                                                                Descriptie = Format1 & Prefix & " " & String1 & "}"
                                                            End If

                                                        Else
                                                            Descriptie = Format1 & String1 & "}"
                                                        End If
                                                    End If

                                                End If


                                                If Not Descriptie = "" Then
                                                    If IsNothing(PolyCL) = False Then
                                                        Point_MS = PolyCL.GetPointAtDist(Station * CSF)
                                                    Else
                                                        Point_MS = PolyCL3D.GetPointAtDist(Station * CSF)
                                                    End If



                                                    If Data_table_Matchlines.Rows.Count > 0 Then

                                                        Dim Nr_rand As Integer = -1

                                                        Dim Point_PS As New Point3d(0, 0, 0)
                                                        Dim Start1 As New Point3d
                                                        Dim End1 As New Point3d
                                                        Dim M1 As Double = 0
                                                        Dim M2 As Double = 0

                                                        For j = 0 To Data_table_Matchlines.Rows.Count - 1
                                                            If IsDBNull(Data_table_Matchlines.Rows(j).Item("BEGSTA")) = False And IsDBNull(Data_table_Matchlines.Rows(j).Item("ENDSTA")) = False Then

                                                                M1 = Data_table_Matchlines.Rows(j).Item("BEGSTA")
                                                                M2 = Data_table_Matchlines.Rows(j).Item("ENDSTA")
                                                                Nr_rand = Nr_rand + 1

                                                                If Station * CSF <= M2 And Station * CSF >= M1 Then

                                                                    If IsDBNull(Data_table_Matchlines.Rows(j).Item("X1")) = False And IsDBNull(Data_table_Matchlines.Rows(j).Item("Y1")) = False Then
                                                                        If IsDBNull(Data_table_Matchlines.Rows(j).Item("X2")) = False And IsDBNull(Data_table_Matchlines.Rows(j).Item("Y2")) = False Then
                                                                            Start1 = New Point3d(Data_table_Matchlines.Rows(j).Item("X1"), Data_table_Matchlines.Rows(j).Item("Y1"), 0)
                                                                            End1 = New Point3d(Data_table_Matchlines.Rows(j).Item("X2"), Data_table_Matchlines.Rows(j).Item("Y2"), 0)



                                                                            If IsNothing(PolyCL) = False Then
                                                                                If PolyCL.GetPointAtDist(M1).GetVectorTo(End1).Length < PolyCL.GetPointAtDist(M1).GetVectorTo(Start1).Length Then
                                                                                    Dim ttt As New Point3d
                                                                                    ttt = Start1
                                                                                    Start1 = End1
                                                                                    End1 = ttt
                                                                                End If
                                                                            Else
                                                                                If PolyCL3D.GetPointAtDist(M1).GetVectorTo(End1).Length < PolyCL3D.GetPointAtDist(M1).GetVectorTo(Start1).Length Then
                                                                                    Dim ttt As New Point3d
                                                                                    ttt = Start1
                                                                                    Start1 = End1
                                                                                    End1 = ttt
                                                                                End If
                                                                            End If

                                                                            If j - 1 >= 0 Then

                                                                                Line_viewport_previous = New Line(New Point3d(Data_table_Matchlines.Rows(j - 1).Item("X1"), Data_table_Matchlines.Rows(j - 1).Item("Y1"), 0),
                                                                                                                  New Point3d(Data_table_Matchlines.Rows(j - 1).Item("X2"), Data_table_Matchlines.Rows(j - 1).Item("Y2"), 0))
                                                                            End If


                                                                            If RadioButton_xcel_Left_right.Checked = True Then
                                                                                Start_point_for_bands = New Point3d(CDbl(TextBox_X_MS.Text) + (CDbl(TextBox_viewport_Width.Text) / 2) * CDbl(TextBox_viewport_SCALE.Text) - New Line(Start1, End1).Length / 2,
                                                                                                                    Start_point_for_bands.Y, 0)
                                                                            Else
                                                                                Start_point_for_bands = New Point3d(CDbl(TextBox_X_MS.Text) - (CDbl(TextBox_viewport_Width.Text) / 2) * CDbl(TextBox_viewport_SCALE.Text) + New Line(Start1, End1).Length / 2,
                                                                                                                    Start_point_for_bands.Y, 0)

                                                                            End If



                                                                            Exit For
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        Next


                                                        If Nr_previous <> Nr_rand Then



                                                            Nr_previous = Nr_rand
                                                            Xprev = Start_point_for_bands.X

                                                            Dim Band_text As New DBText
                                                            Band_text.Layer = No_plot

                                                            Dim Amount_label_Y As Double = Line_len / 2
                                                            If Is_canada = True Then
                                                                Amount_label_Y = 20 * Viewport_scale
                                                            End If

                                                            If RadioButton_xcel_Left_right.Checked = True Then

                                                                Band_text.Justify = AttachmentPoint.MiddleRight
                                                                Band_text.AlignmentPoint = New Point3d(Start_point_for_bands.X - 5 * CDbl(TextBox_Text_height_excel_data.Text), Start_point_for_bands.Y - (Nr_rand * Bands_Y_spacing) + Amount_label_Y, 0)
                                                            Else

                                                                Band_text.Justify = AttachmentPoint.MiddleLeft
                                                                Band_text.AlignmentPoint = New Point3d(Start_point_for_bands.X + 5 * CDbl(TextBox_Text_height_excel_data.Text), Start_point_for_bands.Y - (Nr_rand * Bands_Y_spacing) + Amount_label_Y, 0)
                                                            End If

                                                            Band_text.TextString = CStr(Nr_rand + 1)
                                                            Band_text.Height = 7.5 * TextHeight

                                                            BTrecord.AppendEntity(Band_text)
                                                            Trans1.AddNewlyCreatedDBObject(Band_text, True)



                                                            Dim Viewport_matchline As New Polyline
                                                            Viewport_matchline.Layer = "NO PLOT"
                                                            Viewport_matchline.Closed = True
                                                            Viewport_matchline.ColorIndex = 1
                                                            BTrecord.AppendEntity(Viewport_matchline)
                                                            Trans1.AddNewlyCreatedDBObject(Viewport_matchline, True)

                                                            If RadioButton_Left_right.Checked = True Then
                                                                Dim Pt1 As New Point2d(CDbl(TextBox_X_MS.Text) + CDbl(TextBox_viewport_Width.Text) / 2 - Start1.DistanceTo(End1) / 2, Start_point_for_bands.Y - Nr_rand * Bands_Y_spacing + CDbl(TextBox_viewport_Height.Text))
                                                                Dim Pt2 As New Point2d(CDbl(TextBox_X_MS.Text) + CDbl(TextBox_viewport_Width.Text) / 2 + Start1.DistanceTo(End1) / 2, Start_point_for_bands.Y - Nr_rand * Bands_Y_spacing + CDbl(TextBox_viewport_Height.Text))
                                                                Dim Pt3 As New Point2d(CDbl(TextBox_X_MS.Text) + CDbl(TextBox_viewport_Width.Text) / 2 + Start1.DistanceTo(End1) / 2, Start_point_for_bands.Y - Nr_rand * Bands_Y_spacing)
                                                                Dim Pt4 As New Point2d(CDbl(TextBox_X_MS.Text) + CDbl(TextBox_viewport_Width.Text) / 2 - Start1.DistanceTo(End1) / 2, Start_point_for_bands.Y - Nr_rand * Bands_Y_spacing)
                                                                Viewport_matchline.AddVertexAt(0, Pt1, 0, 0, 0)
                                                                Viewport_matchline.AddVertexAt(1, Pt2, 0, 0, 0)
                                                                Viewport_matchline.AddVertexAt(2, Pt3, 0, 0, 0)
                                                                Viewport_matchline.AddVertexAt(3, Pt4, 0, 0, 0)
                                                            Else
                                                                Dim Pt1 As New Point2d(CDbl(TextBox_X_MS.Text) - CDbl(TextBox_viewport_Width.Text) / 2 + Start1.DistanceTo(End1) / 2, Start_point_for_bands.Y - Nr_rand * Bands_Y_spacing + CDbl(TextBox_viewport_Height.Text))
                                                                Dim Pt2 As New Point2d(CDbl(TextBox_X_MS.Text) - CDbl(TextBox_viewport_Width.Text) / 2 - Start1.DistanceTo(End1) / 2, Start_point_for_bands.Y - Nr_rand * Bands_Y_spacing + CDbl(TextBox_viewport_Height.Text))
                                                                Dim Pt3 As New Point2d(CDbl(TextBox_X_MS.Text) - CDbl(TextBox_viewport_Width.Text) / 2 - Start1.DistanceTo(End1) / 2, Start_point_for_bands.Y - Nr_rand * Bands_Y_spacing)
                                                                Dim Pt4 As New Point2d(CDbl(TextBox_X_MS.Text) - CDbl(TextBox_viewport_Width.Text) / 2 + Start1.DistanceTo(End1) / 2, Start_point_for_bands.Y - Nr_rand * Bands_Y_spacing)
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

                                                            If RadioButton_Left_right.Checked = True Then
                                                                Dim Pt1 As New Point2d(CDbl(TextBox_X_MS.Text), Start_point_for_bands.Y - Nr_rand * Bands_Y_spacing + CDbl(TextBox_viewport_Height.Text))
                                                                Dim Pt2 As New Point2d(CDbl(TextBox_X_MS.Text) + CDbl(TextBox_viewport_Width.Text), Start_point_for_bands.Y - Nr_rand * Bands_Y_spacing + CDbl(TextBox_viewport_Height.Text))
                                                                Dim Pt3 As New Point2d(CDbl(TextBox_X_MS.Text) + CDbl(TextBox_viewport_Width.Text), Start_point_for_bands.Y - Nr_rand * Bands_Y_spacing)
                                                                Dim Pt4 As New Point2d(CDbl(TextBox_X_MS.Text), Start_point_for_bands.Y - Nr_rand * Bands_Y_spacing)
                                                                Viewport_big.AddVertexAt(0, Pt1, 0, 0, 0)
                                                                Viewport_big.AddVertexAt(1, Pt2, 0, 0, 0)
                                                                Viewport_big.AddVertexAt(2, Pt3, 0, 0, 0)
                                                                Viewport_big.AddVertexAt(3, Pt4, 0, 0, 0)
                                                            Else
                                                                Dim Pt1 As New Point2d(CDbl(TextBox_X_MS.Text), Start_point_for_bands.Y - Nr_rand * Bands_Y_spacing + CDbl(TextBox_viewport_Height.Text))
                                                                Dim Pt2 As New Point2d(CDbl(TextBox_X_MS.Text) - CDbl(TextBox_viewport_Width.Text), Start_point_for_bands.Y - Nr_rand * Bands_Y_spacing + CDbl(TextBox_viewport_Height.Text))
                                                                Dim Pt3 As New Point2d(CDbl(TextBox_X_MS.Text) - CDbl(TextBox_viewport_Width.Text), Start_point_for_bands.Y - Nr_rand * Bands_Y_spacing)
                                                                Dim Pt4 As New Point2d(CDbl(TextBox_X_MS.Text), Start_point_for_bands.Y - Nr_rand * Bands_Y_spacing)
                                                                Viewport_big.AddVertexAt(0, Pt1, 0, 0, 0)
                                                                Viewport_big.AddVertexAt(1, Pt2, 0, 0, 0)
                                                                Viewport_big.AddVertexAt(2, Pt3, 0, 0, 0)
                                                                Viewport_big.AddVertexAt(3, Pt4, 0, 0, 0)
                                                            End If





                                                        End If

                                                        Station_previous = Station

                                                        Dim X_aligned As Double = 0


                                                        Dim Viewport_line As New Line(Start1, End1)
                                                        Dim Dist_from_match As Double = -1.1111
                                                        Dim Point_on_line As New Point3d
                                                        Point_on_line = Viewport_line.GetClosestPointTo(Point_MS, Vector3d.ZAxis, True)
                                                        Dim Point_on_Match As New Point3d
                                                        If IsNothing(PolyCL) = False Then
                                                            Point_on_Match = Viewport_line.GetClosestPointTo(PolyCL.GetPointAtDist(M1), Vector3d.ZAxis, True)
                                                        Else
                                                            Dim param3d As Double = PolyCL3D.GetParameterAtDistance(M1)

                                                            Point_on_Match = Viewport_line.GetClosestPointTo(PolyTemp_2d.GetPointAtParameter(param3d), Vector3d.ZAxis, True)
                                                        End If

                                                        If IsNothing(Point_on_line) = False Then
                                                            Dist_from_match = Point_on_Match.GetVectorTo(Point_on_line).Length
                                                        End If
                                                        Dim Val1 As Double = Station - M1
                                                        If Not Dist_from_match = -1.1111 Then
                                                            Val1 = Dist_from_match
                                                        End If




                                                        If RadioButton_xcel_Left_right.Checked = True Then
                                                            Point_PS = New Point3d(Start_point_for_bands.X + Val1, Start_point_for_bands.Y - (Nr_rand * Bands_Y_spacing), 0)
                                                            If Point_PS.X - Xprev < Min_dist Then
                                                                Point_PS = New Point3d(Xprev + Min_dist, Start_point_for_bands.Y - (Nr_rand * Bands_Y_spacing), 0)
                                                            End If
                                                        Else
                                                            Point_PS = New Point3d(Start_point_for_bands.X - Val1, Start_point_for_bands.Y - (Nr_rand * Bands_Y_spacing), 0)

                                                            If Xprev - Point_PS.X < Min_dist Then
                                                                Point_PS = New Point3d(Xprev - Min_dist, Start_point_for_bands.Y - (Nr_rand * Bands_Y_spacing), 0)
                                                            End If
                                                        End If


                                                        If Is_canada = True Then
                                                            If Type1 = "DEFLECTION" Then
                                                                Dim Poly_l As New Polyline
                                                                Poly_l.AddVertexAt(0, New Point2d(Point_PS.X + (2.435 - 2.435) * Viewport_scale + TextHeight / 2, Point_PS.Y + (25.09 + 4.16) * Viewport_scale), 0, 0, 0)
                                                                Poly_l.AddVertexAt(1, New Point2d(Point_PS.X + (2.435 - 2.435) * Viewport_scale + TextHeight / 2, Point_PS.Y + 25.09 * Viewport_scale), 0, 0, 0)
                                                                Poly_l.AddVertexAt(2, New Point2d(Point_PS.X + (2.435 - 2.8 - 2.435) * Viewport_scale + TextHeight / 2, Point_PS.Y + (25.09 + 4.16) * Viewport_scale), 0, 0, 0)
                                                                Poly_l.Layer = Layer_xing
                                                                BTrecord.AppendEntity(Poly_l)
                                                                Trans1.AddNewlyCreatedDBObject(Poly_l, True)

                                                                Dim Poly_arc As New Polyline
                                                                Poly_arc.AddVertexAt(0, New Point2d(Point_PS.X + (2.435 - 3.377 - 2.435) * Viewport_scale + TextHeight / 2, Point_PS.Y + (25.09 + 2.446) * Viewport_scale), -Tan((117 * PI / 180) / 4), 0, 0)
                                                                Poly_arc.AddVertexAt(1, New Point2d(Point_PS.X + (2.435 + 0.997 - 2.435) * Viewport_scale + TextHeight / 2, Point_PS.Y + (25.09 + 3.092) * Viewport_scale), 0, 0, 0)
                                                                Poly_arc.Layer = Layer_xing
                                                                BTrecord.AppendEntity(Poly_arc)
                                                                Trans1.AddNewlyCreatedDBObject(Poly_arc, True)

                                                                Dim Mtext1 As New MText
                                                                Mtext1.Contents = Descriptie
                                                                Mtext1.TextHeight = TextHeight
                                                                Mtext1.Rotation = Mtext_rotation
                                                                Mtext1.Attachment = AttachmentPoint.MiddleLeft
                                                                Mtext1.Location = New Point3d(Point_PS.X, Point_PS.Y + 31 * Viewport_scale, 0)
                                                                Mtext1.Layer = Layer_xing
                                                                If IsNothing(ObjId_text_table_record) = False Then
                                                                    Mtext1.TextStyleId = ObjId_text_table_record.ObjectId
                                                                End If
                                                                BTrecord.AppendEntity(Mtext1)
                                                                Trans1.AddNewlyCreatedDBObject(Mtext1, True)

                                                                Dim Mtext2 As New MText
                                                                Mtext2.Contents = Get_chainage_from_double(Station, Round1)
                                                                Mtext2.TextHeight = TextHeight
                                                                Mtext2.Rotation = Mtext_rotation
                                                                Mtext2.Location = New Point3d(Point_PS.X, Point_PS.Y, 0)
                                                                Mtext2.Attachment = AttachmentPoint.MiddleLeft
                                                                Mtext2.Layer = Layer_xing
                                                                If IsNothing(ObjId_text_table_record) = False Then
                                                                    Mtext2.TextStyleId = ObjId_text_table_record.ObjectId
                                                                End If
                                                                BTrecord.AppendEntity(Mtext2)
                                                                Trans1.AddNewlyCreatedDBObject(Mtext2, True)


                                                            Else

                                                                Dim Mtext1 As New MText
                                                                Mtext1.Contents = Descriptie
                                                                Mtext1.TextHeight = TextHeight
                                                                Mtext1.Rotation = Mtext_rotation
                                                                Mtext1.Attachment = AttachmentPoint.MiddleLeft
                                                                Mtext1.Location = New Point3d(Point_PS.X, Point_PS.Y + 25.09 * Viewport_scale, 0)
                                                                Mtext1.Layer = Layer_xing
                                                                If IsNothing(ObjId_text_table_record) = False Then
                                                                    Mtext1.TextStyleId = ObjId_text_table_record.ObjectId
                                                                End If
                                                                BTrecord.AppendEntity(Mtext1)
                                                                Trans1.AddNewlyCreatedDBObject(Mtext1, True)

                                                                Dim Mtext2 As New MText
                                                                Mtext2.Contents = Get_chainage_from_double(Station, Round1)
                                                                Mtext2.TextHeight = TextHeight
                                                                Mtext2.Rotation = Mtext_rotation
                                                                Mtext2.Location = New Point3d(Point_PS.X, Point_PS.Y, 0)
                                                                Mtext2.Attachment = AttachmentPoint.MiddleLeft
                                                                Mtext2.Layer = Layer_xing
                                                                If IsNothing(ObjId_text_table_record) = False Then
                                                                    Mtext2.TextStyleId = ObjId_text_table_record.ObjectId
                                                                End If
                                                                BTrecord.AppendEntity(Mtext2)
                                                                Trans1.AddNewlyCreatedDBObject(Mtext2, True)

                                                            End If
                                                        End If



                                                        If Is_canada = False Then

                                                            Dim Mtext_crossing As New MText
                                                            Mtext_crossing.Location = Point_PS

                                                            Mtext_crossing.TextHeight = TextHeight
                                                            Mtext_crossing.Rotation = Mtext_rotation
                                                            Mtext_crossing.Layer = Layer_xing
                                                            Mtext_crossing.Contents = Descriptie

                                                            Mtext_crossing.Attachment = AttachmentPoint.BottomLeft


                                                            If IsNothing(ObjId_text_table_record) = False Then
                                                                Mtext_crossing.TextStyleId = ObjId_text_table_record.ObjectId
                                                            End If
                                                            BTrecord.AppendEntity(Mtext_crossing)
                                                            Trans1.AddNewlyCreatedDBObject(Mtext_crossing, True)
                                                        End If

                                                        Xprev = Point_PS.X
                                                    End If


                                                End If

                                            End If
                                        Next
                                    End If
                                End If


                                Trans1.Commit()
                            End Using
                            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                            ' asta e de la lock
                        End Using

                        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                    End If
                End If


            Catch ex As Exception
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try
        End If
        Freeze_operations = False
    End Sub
    Private Sub TabPage3_Click(sender As Object, e As EventArgs) Handles TabPage3.Click
        Incarca_existing_layers_to_combobox(ComboBox_layers_for_Mtext)
        Incarca_existing_textstyles_to_combobox(ComboBox_text_style_mtext)
    End Sub
    Private Sub RadioButton_linelist_ned_CheckedChanged(sender As Object, e As EventArgs) Handles _
        RadioButton_prop_band_ETC.CheckedChanged,
        RadioButton_STATION_band_etc.CheckedChanged,
        RadioButton_reference_band.CheckedChanged,
        RadioButton_class_location_NED.CheckedChanged,
        RadioButton_land_use_peneast.CheckedChanged,
        RadioButton_soils_band.CheckedChanged,
        RadioButton_prop_band_Ozark.CheckedChanged,
        RadioButton_propBand_spectra.CheckedChanged,
        RadioButton_env_band_spectra.CheckedChanged,
        RadioButton_class_location_ETC.CheckedChanged,
        RadioButton_crossing_band_spectra.CheckedChanged,
        RadioButton_watersheed_peneast.CheckedChanged,
        RadioButton_crossing_band_thornbury.CheckedChanged,
        RadioButton_HOP_peneast.CheckedChanged, RadioButton_class_band_peneast.CheckedChanged, RadioButton_prop_band_peneast.CheckedChanged, RadioButton_station_band_peneast.CheckedChanged

        If RadioButton_prop_band_ETC.Checked = True Then
            TextBox_viewport_Width.Text = "6068.3117"
            TextBox_viewport_Height.Text = "260"
            TextBox_X_MS.Text = "0"
            TextBox_Y_MS.Text = "-1000"
            TextBox_X_PS.Text = "1056.6883"
            TextBox_Y_PS.Text = "4405.0000"
            TextBox_BAND_SPACING.Text = "600"
            TextBox_shift_viewport_y.Text = "0"
            TextBox_CSF.Text = "1"
            Label_scale_1_to.Text = "Scale 1" & Chr(34) & " to"
            Label_inches.Text = "'"
            TextBox_prefix_excel_data.Text = ""

            TextBox_textOffset_X.Text = "5"
            TextBox_textOffset_Y.Text = "75"
            TextBox_text_height.Text = "16"
            With ComboBox_text_style
                If .Items.Contains("Romans") = True Then
                    .SelectedIndex = .Items.IndexOf("Romans")
                End If
            End With
            TextBox_textwidth.Text = "0.8"
            TextBox_minimum_distance.Text = "400"

            Is_canada = False
        ElseIf RadioButton_STATION_band_etc.Checked = True Then
            TextBox_viewport_Width.Text = "6068.3117"
            TextBox_viewport_Height.Text = "703.2"
            TextBox_X_MS.Text = "20000"
            TextBox_Y_MS.Text = "-1000"
            TextBox_X_PS.Text = "1062.6252"
            TextBox_Y_PS.Text = "3641.7232"
            TextBox_BAND_SPACING.Text = "1600"
            TextBox_shift_viewport_y.Text = "-120"
            TextBox_CSF.Text = "1"
            Label_scale_1_to.Text = "Scale 1" & Chr(34) & " to"
            Label_inches.Text = "'"
            TextBox_prefix_excel_data.Text = ""

            TextBox_Text_height_excel_data.Text = "16"

            With ComboBox_text_styles_excel_data
                If .Items.Contains("Romans") = True Then
                    .SelectedIndex = .Items.IndexOf("Romans")
                End If
            End With

            With ComboBox_layers_deflections
                If .Items.Contains("DEFL_TEXT") = True Then
                    .SelectedIndex = .Items.IndexOf("DEFL_TEXT")
                End If
            End With

            With ComboBox_layers_crossings
                If .Items.Contains("TEXT") = True Then
                    .SelectedIndex = .Items.IndexOf("TEXT")
                End If
            End With

            TextBox_width_factor_Excel_data.Text = "0.8"
            TextBox_min_dist_excel_data.Text = "30"
            TextBox_rotation_excel_data.Text = "90"

            Is_canada = False
        ElseIf RadioButton_reference_band.Checked = True Then
            TextBox_viewport_Width.Text = "787.98"
            TextBox_viewport_Height.Text = "481"
            TextBox_X_MS.Text = "20050"
            TextBox_Y_MS.Text = "-1000"
            TextBox_X_PS.Text = "3449.0369"
            TextBox_Y_PS.Text = "279.2064"
            TextBox_BAND_SPACING.Text = "500"
            TextBox_shift_viewport_y.Text = "-351.37"
            TextBox_CSF.Text = "1"
            Label_scale_1_to.Text = "Scale 1" & Chr(34) & " to"
            Label_inches.Text = "'"
            TextBox_prefix_excel_data.Text = ""
            Is_canada = False
        ElseIf RadioButton_class_location_NED.Checked = True Then
            TextBox_viewport_Width.Text = "6596"
            TextBox_viewport_Height.Text = "50"
            TextBox_X_MS.Text = "30000"
            TextBox_Y_MS.Text = "-1000"
            TextBox_X_PS.Text = "504"
            TextBox_Y_PS.Text = "1965"
            TextBox_BAND_SPACING.Text = "200"
            TextBox_shift_viewport_y.Text = "-85.87"
            TextBox_CSF.Text = "1"
            Label_scale_1_to.Text = "Scale 1" & Chr(34) & " to"
            Label_inches.Text = "'"
            TextBox_prefix_excel_data.Text = ""
            Is_canada = False
        ElseIf RadioButton_soils_band.Checked = True Then
            TextBox_viewport_Width.Text = "3050"
            TextBox_viewport_Height.Text = "25"
            TextBox_X_MS.Text = "40000"
            TextBox_Y_MS.Text = "-1000"
            TextBox_X_PS.Text = "294.998"
            TextBox_Y_PS.Text = "960.001"
            TextBox_BAND_SPACING.Text = "100"
            TextBox_shift_viewport_X.Text = "0"
            TextBox_shift_viewport_y.Text = "0"
            TextBox_CSF.Text = "1"
            Label_scale_1_to.Text = "Scale 1" & Chr(34) & " to"
            Label_inches.Text = "'"
            TextBox_prefix_excel_data.Text = ""
            TextBox_mtext_height.Text = "8"
            If ComboBox_text_style_mtext.Items.Contains("Arial") = True Then
                ComboBox_text_style_mtext.SelectedIndex = ComboBox_text_style_mtext.Items.IndexOf("Arial")
            End If

            TextBox_minimum_distance_mtext.Text = "35"

            If ComboBox_layers_for_Mtext.Items.Contains("Materials") = True Then
                ComboBox_layers_for_Mtext.SelectedIndex = ComboBox_layers_for_Mtext.Items.IndexOf("Materials")
            End If

            Is_canada = False
        ElseIf RadioButton_prop_band_peneast.Checked = True Then
            TextBox_viewport_Width.Text = "3050"
            TextBox_viewport_Height.Text = "100"
            TextBox_X_MS.Text = "10000"
            TextBox_Y_MS.Text = "-1000"
            TextBox_X_PS.Text = "294.998"
            TextBox_Y_PS.Text = "2250.001"
            TextBox_BAND_SPACING.Text = "250"
            TextBox_shift_viewport_X.Text = "0"
            TextBox_shift_viewport_y.Text = "0"
            TextBox_CSF.Text = "1"
            Label_scale_1_to.Text = "Scale 1" & Chr(34) & " to"
            Label_inches.Text = "'"

            TextBox_textOffset_X.Text = "1.5"
            TextBox_textOffset_Y.Text = "4"
            TextBox_text_height.Text = "8"
            With ComboBox_text_style
                If .Items.Contains("Arial") = True Then
                    .SelectedIndex = .Items.IndexOf("Arial")
                End If
            End With
            TextBox_textwidth.Text = "0.9"
            TextBox_minimum_distance.Text = "150"
            Is_canada = False
        ElseIf RadioButton_land_use_peneast.Checked = True Then
            TextBox_viewport_Width.Text = "3050"
            TextBox_viewport_Height.Text = "25"
            TextBox_X_MS.Text = "50000"
            TextBox_Y_MS.Text = "-1000"
            TextBox_X_PS.Text = "294.998"
            TextBox_Y_PS.Text = "1125.001"
            TextBox_BAND_SPACING.Text = "100"
            TextBox_shift_viewport_X.Text = "0"
            TextBox_shift_viewport_y.Text = "0"
            TextBox_CSF.Text = "1"
            Label_scale_1_to.Text = "Scale 1" & Chr(34) & " to"
            Label_inches.Text = "'"
            TextBox_prefix_excel_data.Text = ""
            TextBox_mtext_height.Text = "8"
            If ComboBox_text_style_mtext.Items.Contains("Arial") = True Then
                ComboBox_text_style_mtext.SelectedIndex = ComboBox_text_style_mtext.Items.IndexOf("Arial")
            End If

            TextBox_minimum_distance_mtext.Text = "50"

            If ComboBox_layers_for_Mtext.Items.Contains("Materials") = True Then
                ComboBox_layers_for_Mtext.SelectedIndex = ComboBox_layers_for_Mtext.Items.IndexOf("Materials")
            End If

            Is_canada = False
        ElseIf RadioButton_class_band_peneast.Checked = True Then
            TextBox_viewport_Width.Text = "3050"
            TextBox_viewport_Height.Text = "50"
            TextBox_X_MS.Text = "60000"
            TextBox_Y_MS.Text = "-1000"
            TextBox_X_PS.Text = "294.998"
            TextBox_Y_PS.Text = "825.001"
            TextBox_BAND_SPACING.Text = "200"
            TextBox_shift_viewport_X.Text = "0"
            TextBox_shift_viewport_y.Text = "0"
            TextBox_CSF.Text = "1"
            Label_scale_1_to.Text = "Scale 1" & Chr(34) & " to"
            Label_inches.Text = "'"
            TextBox_class_width.Text = "0.9"
            TextBox_class_height.Text = "8"
            If ComboBox_class_text_style.Items.Contains("Arial") = True Then
                ComboBox_class_text_style.SelectedIndex = ComboBox_class_text_style.Items.IndexOf("Arial")
            End If

            TextBox_class_min_dist.Text = "50"

            If ComboBox_class_layer.Items.Contains("Materials") = True Then
                ComboBox_class_layer.SelectedIndex = ComboBox_class_layer.Items.IndexOf("Materials")
            End If

            Is_canada = False

        ElseIf RadioButton_station_band_peneast.Checked = True Then
            TextBox_viewport_Width.Text = "3050"
            TextBox_viewport_Height.Text = "275"
            TextBox_X_MS.Text = "20000"
            TextBox_Y_MS.Text = "-1000"
            TextBox_X_PS.Text = "294.998"
            TextBox_Y_PS.Text = "1975.001"
            TextBox_BAND_SPACING.Text = "650"
            TextBox_shift_viewport_X.Text = "0"
            TextBox_shift_viewport_y.Text = "-10"
            TextBox_CSF.Text = "1"
            Label_scale_1_to.Text = "Scale 1" & Chr(34) & " to"
            Label_inches.Text = "'"

            TextBox_Text_height_excel_data.Text = "8"
            With ComboBox_text_styles_excel_data
                If .Items.Contains("Arial") = True Then
                    .SelectedIndex = .Items.IndexOf("Arial")
                End If
            End With

            With ComboBox_layers_deflections
                If .Items.Contains("TEXT_PI") = True Then
                    .SelectedIndex = .Items.IndexOf("TEXT_PI")
                End If
            End With

            With ComboBox_layers_crossings
                If .Items.Contains("Station_Band") = True Then
                    .SelectedIndex = .Items.IndexOf("Station_Band")
                End If
            End With

            TextBox_width_factor_Excel_data.Text = "0.9"
            TextBox_min_dist_excel_data.Text = "30"
            TextBox_rotation_excel_data.Text = "45"

            Is_canada = False

        ElseIf RadioButton_prop_band_Ozark.Checked = True Then
            TextBox_viewport_Width.Text = "34.5"
            TextBox_viewport_Height.Text = "2"
            TextBox_X_MS.Text = "0"
            TextBox_Y_MS.Text = "-1000"
            TextBox_X_PS.Text = "0.5"
            TextBox_Y_PS.Text = "23.5"
            TextBox_BAND_SPACING.Text = "1350"
            TextBox_minimum_distance.Text = "225"
            TextBox_text_height.Text = "24"
            TextBox_textOffset_X.Text = "10"
            TextBox_textwidth.Text = "1"
            TextBox_viewport_SCALE.Text = "300"
            TextBox_shift_viewport_X.Text = "-0.1539"
            TextBox_shift_viewport_y.Text = "0.9967"
            TextBox_CSF.Text = "1"
            Label_scale_1_to.Text = "Scale 1" & Chr(34) & " to"
            Label_inches.Text = "'"
            TextBox_prefix_excel_data.Text = ""
            Is_canada = False
        ElseIf RadioButton_crossing_band_spectra.Checked = True Then
            TextBox_viewport_Width.Text = "30.3943"
            TextBox_viewport_Height.Text = "1.7449"
            TextBox_X_MS.Text = "0"
            TextBox_Y_MS.Text = "-1000"
            TextBox_X_PS.Text = "5.1058"
            TextBox_Y_PS.Text = "20.7651"
            TextBox_BAND_SPACING.Text = "200"
            TextBox_min_dist_excel_data.Text = "10"
            TextBox_Text_height_excel_data.Text = "4"
            TextBox_viewport_SCALE.Text = "50"
            TextBox_width_factor_Excel_data.Text = "0.8"
            TextBox_rotation_excel_data.Text = "45"
            If ComboBox_text_styles_excel_data.Items.Contains("ROMANS") = True Then
                ComboBox_text_styles_excel_data.SelectedIndex = ComboBox_text_styles_excel_data.Items.IndexOf("ROMANS")
            End If
            If ComboBox_layers_deflections.Items.Contains("TEXT_PI") = True Then
                ComboBox_layers_deflections.SelectedIndex = ComboBox_layers_deflections.Items.IndexOf("TEXT_PI")
            End If
            If ComboBox_layers_crossings.Items.Contains("Stationing") = True Then
                ComboBox_layers_crossings.SelectedIndex = ComboBox_layers_crossings.Items.IndexOf("Stationing")
            End If
            TextBox_shift_viewport_X.Text = "0"
            TextBox_shift_viewport_y.Text = "0.8"
            TextBox_CSF.Text = "1"
            Label_scale_1_to.Text = "Scale 1" & Chr(34) & " to"
            Label_inches.Text = "'"
            TextBox_prefix_excel_data.Text = ""
            Is_canada = False
        ElseIf RadioButton_env_band_spectra.Checked = True Then
            TextBox_viewport_Width.Text = "30.3943"
            TextBox_viewport_Height.Text = "1.9946"
            TextBox_X_MS.Text = "10000"
            TextBox_Y_MS.Text = "-1000"
            TextBox_X_PS.Text = "5.1058"
            TextBox_Y_PS.Text = "8.7794"
            TextBox_BAND_SPACING.Text = "250"
            TextBox_min_dist_excel_data.Text = "10"
            TextBox_Text_height_excel_data.Text = "4"
            TextBox_viewport_SCALE.Text = "50"
            TextBox_width_factor_Excel_data.Text = "0.8"
            TextBox_rotation_excel_data.Text = "45"
            If ComboBox_text_styles_excel_data.Items.Contains("ROMANS") = True Then
                ComboBox_text_styles_excel_data.SelectedIndex = ComboBox_text_styles_excel_data.Items.IndexOf("ROMANS")
            End If
            If ComboBox_layers_deflections.Items.Contains("TEXT_PI") = True Then
                ComboBox_layers_deflections.SelectedIndex = ComboBox_layers_deflections.Items.IndexOf("TEXT_PI")
            End If
            If ComboBox_layers_crossings.Items.Contains("Stationing") = True Then
                ComboBox_layers_crossings.SelectedIndex = ComboBox_layers_crossings.Items.IndexOf("Stationing")
            End If
            TextBox_shift_viewport_X.Text = "0"
            TextBox_shift_viewport_y.Text = "0.9569"
            TextBox_CSF.Text = "1"
            Label_scale_1_to.Text = "Scale 1" & Chr(34) & " to"
            Label_inches.Text = "'"
            TextBox_prefix_excel_data.Text = ""
            Is_canada = False
        ElseIf RadioButton_propBand_spectra.Checked = True Then
            TextBox_viewport_Width.Text = "30.3943"
            TextBox_viewport_Height.Text = "0.9894"
            TextBox_X_MS.Text = "5000"
            TextBox_Y_MS.Text = "-1000"
            TextBox_X_PS.Text = "5.1058"
            TextBox_Y_PS.Text = "22.5101"
            TextBox_BAND_SPACING.Text = "150"
            TextBox_minimum_distance.Text = "35"
            TextBox_text_height.Text = "4"
            TextBox_textOffset_X.Text = "1.5"
            TextBox_textOffset_Y.Text = "1.5"
            TextBox_textwidth.Text = "0.8"
            TextBox_viewport_SCALE.Text = "50"
            TextBox_shift_viewport_X.Text = "0"
            TextBox_shift_viewport_y.Text = "0.4848"
            If ComboBox_text_style.Items.Contains("Property") = True Then
                ComboBox_text_style.SelectedIndex = ComboBox_text_style.Items.IndexOf("Property")
            End If

            If ComboBox_layers.Items.Contains("TEXT") = True Then
                ComboBox_layers.SelectedIndex = ComboBox_layers.Items.IndexOf("TEXT")
            End If
            TextBox_CSF.Text = "1"
            Label_scale_1_to.Text = "Scale 1" & Chr(34) & " to"
            TextBox_prefix_excel_data.Text = ""
            Label_inches.Text = "'"
            Is_canada = False
        ElseIf RadioButton_class_location_ETC.Checked = True Then
            TextBox_viewport_Width.Text = "6052.4"
            TextBox_viewport_Height.Text = "60"
            TextBox_X_MS.Text = "60000"
            TextBox_Y_MS.Text = "-1000"
            TextBox_X_PS.Text = "1072.5949"
            TextBox_Y_PS.Text = "1831.6521"
            TextBox_BAND_SPACING.Text = "200"
            TextBox_shift_viewport_X.Text = "0"
            TextBox_shift_viewport_y.Text = "0"
            TextBox_class_height.Text = "16"
            TextBox_Class_End_MP.Text = "C"
            TextBox_Class_Start_MP.Text = "B"
            TextBox_Class.Text = "A"
            TextBox_CSF.Text = "1"
            Label_scale_1_to.Text = "Scale 1" & Chr(34) & " to"
            Label_inches.Text = "'"
            TextBox_prefix_excel_data.Text = ""
            Is_canada = False
        ElseIf RadioButton_watersheed_peneast.Checked = True Then
            TextBox_viewport_Width.Text = "3205.2822"
            TextBox_viewport_Height.Text = "140"
            TextBox_X_MS.Text = "60000"
            TextBox_Y_MS.Text = "-1000"
            TextBox_X_PS.Text = "294.9980"
            TextBox_Y_PS.Text = "985.0010"
            TextBox_BAND_SPACING.Text = "350"
            TextBox_shift_viewport_X.Text = "27.2"
            TextBox_shift_viewport_y.Text = "0"
            TextBox_text_height_water.Text = "8"
            TextBox_width_factor_water.Text = "1"
            TextBox_water_deltaY_station.Text = "78"

            If ComboBox_text_style_water.Items.Contains("CLASS") = True Then
                ComboBox_text_style_water.SelectedIndex = ComboBox_text_style_water.Items.IndexOf("CLASS")
            ElseIf ComboBox_text_style_water.Items.Contains("Arial") = True Then
                ComboBox_text_style_water.SelectedIndex = ComboBox_text_style_water.Items.IndexOf("Arial")
            End If

            If ComboBox_layer_water.Items.Contains("TEXT") = True Then
                ComboBox_layer_water.SelectedIndex = ComboBox_layer_water.Items.IndexOf("TEXT")
            ElseIf ComboBox_layer_water.Items.Contains("CLASS") = True Then
                ComboBox_layer_water.SelectedIndex = ComboBox_layer_water.Items.IndexOf("CLASS")
            End If
            TextBox_CSF.Text = "1"
            Label_scale_1_to.Text = "Scale 1" & Chr(34) & " to"
            Label_inches.Text = "'"
            TextBox_prefix_excel_data.Text = ""
            Is_canada = False
        ElseIf RadioButton_crossing_band_thornbury.Checked = True Then
            TextBox_BAND_SPACING.Text = "1100"
            TextBox_viewport_SCALE.Text = "7.5"
            TextBox_viewport_Height.Text = "63.43"
            TextBox_viewport_Width.Text = "780.78"
            TextBox_X_MS.Text = "-10000"
            TextBox_Y_MS.Text = "-1000"
            TextBox_X_PS.Text = "78.56"
            TextBox_Y_PS.Text = "235.72"
            TextBox_shift_viewport_X.Text = "0"
            TextBox_shift_viewport_y.Text = "24"
            TextBox_CSF.Text = "0.999578"
            RadioButton_xcel_right_left.Checked = True
            RadioButton_Right_left_viewport.Checked = True
            TextBox_width_factor_Excel_data.Text = "1"
            TextBox_rotation_excel_data.Text = "90"
            TextBox_min_dist_excel_data.Text = "50"
            TextBox_Text_height_excel_data.Text = "18.75"
            TextBox_prefix_excel_data.Text = ""

            Label_scale_1_to.Text = "Scale 1000:"
            Label_inches.Text = "x 1000"
            Is_canada = True
        ElseIf RadioButton_HOP_peneast.Checked = True Then
            Is_canada = False
            TextBox_BAND_SPACING.Text = "150"
            TextBox_viewport_SCALE.Text = "25"
            TextBox_viewport_Height.Text = "2.2235"
            TextBox_viewport_Width.Text = "15"
            TextBox_X_MS.Text = "70000"
            TextBox_Y_MS.Text = "-1000"
            TextBox_X_PS.Text = "2.8861"
            TextBox_Y_PS.Text = "18.9951"
            TextBox_shift_viewport_X.Text = "0"
            TextBox_shift_viewport_y.Text = "1"
            TextBox_CSF.Text = "1"
            TextBox_Text_height_excel_data.Text = "2"
            TextBox_min_dist_excel_data.Text = "3"
            TextBox_column_description_excel_data.Text = "B"
            TextBox_column_Station_excel_data.Text = "A"
            If ComboBox_text_styles_excel_data.Items.Contains("Arial") = True Then
                ComboBox_text_styles_excel_data.SelectedIndex = ComboBox_text_styles_excel_data.Items.IndexOf("Arial")
            End If
            TextBox_prefix_excel_data.Text = "STA. "

        Else
            TextBox_viewport_Width.Text = "5350"
            TextBox_viewport_Height.Text = "522.25"
            TextBox_X_MS.Text = "10000"
            TextBox_Y_MS.Text = "-1000"
            TextBox_X_PS.Text = "504.000"
            TextBox_Y_PS.Text = "3816.9963"
            TextBox_BAND_SPACING.Text = "500"
            TextBox_shift_viewport_y.Text = "148.25"
            TextBox_CSF.Text = "1"
            Label_scale_1_to.Text = "Scale 1" & Chr(34) & " to"
            Label_inches.Text = "'"
            TextBox_prefix_excel_data.Text = ""
            Is_canada = False
        End If
    End Sub





    Private Sub Button_class_read_excel_Click(sender As Object, e As EventArgs) Handles Button_class_read_excel.Click
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
                If IsNumeric(TextBox_class_row_Start.Text) = True Then
                    Start1 = CInt(TextBox_class_row_Start.Text)
                End If
                If IsNumeric(TextBox_class_row_End.Text) = True Then
                    End1 = CInt(TextBox_class_row_End.Text)
                End If

                If End1 = 0 Or Start1 = 0 Then
                    Freeze_operations = False
                    Exit Sub
                End If
                If End1 < Start1 Then
                    Freeze_operations = False
                    Exit Sub
                End If

                If IsNothing(Data_table_Matchlines) = True Then
                    MsgBox("No matchlines loaded")
                    Freeze_operations = False
                    Exit Sub
                End If

                If Data_table_Matchlines.Rows.Count < 1 Then
                    MsgBox("No matchlines loaded")
                    Freeze_operations = False
                    Exit Sub
                End If

                Dim Column_sta1 As String = ""
                Column_sta1 = TextBox_Class_Start_MP.Text.ToUpper
                Dim Column_sta2 As String = ""
                Column_sta2 = TextBox_Class_End_MP.Text.ToUpper
                Dim Column_class As String = ""
                Column_class = TextBox_Class.Text.ToUpper
                Dim Column_designF As String = ""
                Column_designF = TextBox_design_factor.Text.ToUpper


                Data_table_class_band_Excel_data = New System.Data.DataTable
                Data_table_class_band_Excel_data.Columns.Add("BEGSTA", GetType(Double))
                Data_table_class_band_Excel_data.Columns.Add("ENDSTA", GetType(Double))
                Data_table_class_band_Excel_data.Columns.Add("CLASS", GetType(String))
                Data_table_class_band_Excel_data.Columns.Add("DESIGN_FACTOR", GetType(String))


                Dim Index_data_table As Double



                For i = Start1 To End1
                    Dim Station_string1 As String = W1.Range(Column_sta1 & i).Value2
                    Dim Station_string2 As String = W1.Range(Column_sta2 & i).Value2
                    Dim Class1 As String = W1.Range(Column_class & i).Value2

                    Dim DesignF As String = W1.Range(Column_designF & i).Value2
                    If DesignF = "" Then DesignF = "xxx"


                    If IsNumeric(Station_string1) = True And IsNumeric(Station_string2) = True And Not Class1 = "" Then
                        If CDbl(Station_string1) >= 0 And CDbl(Station_string2) >= 0 Then
                            Dim Station1 As Double = CDbl(Station_string1)
                            Dim Station2 As Double = CDbl(Station_string2)
                            Dim Adaugat As Boolean = False

                            For j = 0 To Data_table_Matchlines.Rows.Count - 1
                                If IsDBNull(Data_table_Matchlines.Rows(j).Item("BEGSTA")) = False And IsDBNull(Data_table_Matchlines.Rows(j).Item("ENDSTA")) = False Then
                                    Dim M1 As Double = Data_table_Matchlines.Rows(j).Item("BEGSTA")
                                    Dim M2 As Double = Data_table_Matchlines.Rows(j).Item("ENDSTA")
                                    If Station1 * 5280 >= M1 And Station2 * 5280 <= M2 Then
                                        Data_table_class_band_Excel_data.Rows.Add()
                                        Data_table_class_band_Excel_data.Rows(Index_data_table).Item("BEGSTA") = Station1
                                        Data_table_class_band_Excel_data.Rows(Index_data_table).Item("ENDSTA") = Station2
                                        Data_table_class_band_Excel_data.Rows(Index_data_table).Item("CLASS") = Class1
                                        Data_table_class_band_Excel_data.Rows(Index_data_table).Item("DESIGN_FACTOR") = DesignF

                                        Index_data_table = Index_data_table + 1
                                        Adaugat = True
                                        Exit For
                                    End If
                                End If
                            Next

                            If Adaugat = False Then
                                For j = 0 To Data_table_Matchlines.Rows.Count - 1
                                    If IsDBNull(Data_table_Matchlines.Rows(j).Item("BEGSTA")) = False And IsDBNull(Data_table_Matchlines.Rows(j).Item("ENDSTA")) = False Then
                                        Adaugat = False
                                        Dim M1 As Double = Data_table_Matchlines.Rows(j).Item("BEGSTA")
                                        Dim M2 As Double = Data_table_Matchlines.Rows(j).Item("ENDSTA")
                                        If M1 >= Station1 * 5280 And M2 >= Station1 * 5280 And M1 <= Station2 * 5280 And M2 <= Station2 * 5280 Then
                                            Data_table_class_band_Excel_data.Rows.Add()
                                            Data_table_class_band_Excel_data.Rows(Index_data_table).Item("BEGSTA") = M1 / 5280
                                            Data_table_class_band_Excel_data.Rows(Index_data_table).Item("ENDSTA") = M2 / 5280
                                            Data_table_class_band_Excel_data.Rows(Index_data_table).Item("CLASS") = Class1
                                            Data_table_class_band_Excel_data.Rows(Index_data_table).Item("DESIGN_FACTOR") = DesignF
                                            Index_data_table = Index_data_table + 1
                                            Adaugat = True
                                        End If

                                        If Adaugat = False Then
                                            If Station1 * 5280 <= M1 And M1 <= Station2 * 5280 And M2 >= Station2 * 5280 And M2 > Station1 * 5280 Then
                                                Data_table_class_band_Excel_data.Rows.Add()
                                                Data_table_class_band_Excel_data.Rows(Index_data_table).Item("BEGSTA") = M1 / 5280
                                                Data_table_class_band_Excel_data.Rows(Index_data_table).Item("ENDSTA") = Station2
                                                Data_table_class_band_Excel_data.Rows(Index_data_table).Item("CLASS") = Class1
                                                Data_table_class_band_Excel_data.Rows(Index_data_table).Item("DESIGN_FACTOR") = DesignF
                                                Index_data_table = Index_data_table + 1

                                            End If
                                        End If

                                        If Adaugat = False Then
                                            If Station1 * 5280 >= M1 And M2 >= Station1 * 5280 And M1 <= Station2 * 5280 And M2 <= Station2 * 5280 Then
                                                Data_table_class_band_Excel_data.Rows.Add()
                                                Data_table_class_band_Excel_data.Rows(Index_data_table).Item("BEGSTA") = Station1
                                                Data_table_class_band_Excel_data.Rows(Index_data_table).Item("ENDSTA") = M2 / 5280
                                                Data_table_class_band_Excel_data.Rows(Index_data_table).Item("CLASS") = Class1
                                                Data_table_class_band_Excel_data.Rows(Index_data_table).Item("DESIGN_FACTOR") = DesignF
                                                Index_data_table = Index_data_table + 1

                                            End If
                                        End If

                                    End If
                                Next
                            End If



                        End If
                    End If
                Next




                Data_table_class_band_Excel_data = Sort_data_table(Data_table_class_band_Excel_data, "BEGSTA")


                Add_to_clipboard_Data_table(Data_table_class_band_Excel_data)


            Catch ex As Exception
                MsgBox(ex.Message)

            End Try
            Freeze_operations = False
        End If




    End Sub

    Private Sub Button_Insert_class_bands_Click(sender As Object, e As EventArgs) Handles Button_Insert_class_bands.Click

        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Round1 As Integer = 0
            If IsNumeric(TextBox_dec1.Text) = True Then
                Round1 = CInt(TextBox_dec1.Text)
            End If


            Try
                If IsNothing(Data_table_class_band_Excel_data) = False And IsNothing(Data_table_Matchlines) = False Then
                    If Data_table_class_band_Excel_data.Rows.Count > 0 Then
                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor

                        Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)


                                Dim Text_style_table As Autodesk.AutoCAD.DatabaseServices.TextStyleTable = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                Dim Layer_table As Autodesk.AutoCAD.DatabaseServices.LayerTable = Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                                Dim Start_point_for_bands As New Point3d(0, 0, 0)
                                If IsNumeric(TextBox_X_MS.Text) = True And IsNumeric(TextBox_Y_MS.Text) = True Then
                                    Start_point_for_bands = New Point3d(CDbl(TextBox_X_MS.Text), CDbl(TextBox_Y_MS.Text), 0)
                                End If
                                Dim Bands_Y_spacing As Double = 200
                                If IsNumeric(TextBox_BAND_SPACING.Text) = True Then
                                    Bands_Y_spacing = CDbl(TextBox_BAND_SPACING.Text)
                                End If


                                Dim Line_len_view_h As Double = 50
                                If IsNumeric(TextBox_viewport_Height.Text) = True Then
                                    Line_len_view_h = CDbl(TextBox_viewport_Height.Text)
                                End If


                                Dim Min_dist As Double = 230
                                If IsNumeric(TextBox_class_min_dist.Text) = True Then
                                    Min_dist = CDbl(TextBox_class_min_dist.Text)
                                End If



                                Dim Text_offset_no_plot As Double = 1


                                Dim TextHeight As Double = 20
                                If IsNumeric(TextBox_class_height.Text) = True Then
                                    TextHeight = CDbl(TextBox_class_height.Text)
                                End If

                                Dim Xprev As Double = Start_point_for_bands.X
                                Dim Start_ptX As Double

                                Creaza_layer(No_plot, 40, No_plot, False)

                                Dim Nr_previous As Integer = -1
                                Dim Nr_rand As Integer = -1

                                For i = 0 To Data_table_Matchlines.Rows.Count - 1
                                    If IsDBNull(Data_table_Matchlines.Rows(i).Item("BEGSTA")) = False And IsDBNull(Data_table_Matchlines.Rows(i).Item("ENDSTA")) = False Then
                                        Dim Station1 As Double = Data_table_Matchlines.Rows(i).Item("BEGSTA")
                                        Dim Station2 As Double = Data_table_Matchlines.Rows(i).Item("ENDSTA")
                                        Nr_rand = Nr_rand + 1
                                        Dim Point_PS1 As New Point3d(0, 0, 0)
                                        Dim Start1 As New Point3d
                                        Dim End1 As New Point3d
                                        Start1 = New Point3d(Data_table_Matchlines.Rows(i).Item("X1"), Data_table_Matchlines.Rows(i).Item("Y1"), 0)
                                        End1 = New Point3d(Data_table_Matchlines.Rows(i).Item("X2"), Data_table_Matchlines.Rows(i).Item("Y2"), 0)
                                        Dim Viewport_line As New Line(Start1, End1)
                                        Start_ptX = Start_point_for_bands.X + CDbl(TextBox_viewport_Width.Text) / 2 - Viewport_line.Length / 2
                                        Point_PS1 = New Point3d(Start_ptX, Start_point_for_bands.Y - (Nr_rand * Bands_Y_spacing), 0)
                                        Dim Linie1 As New Line(Point_PS1, New Point3d(Point_PS1.X, Point_PS1.Y + Line_len_view_h, 0))
                                        If Not ComboBox_class_layer.Text = "" Then
                                            Linie1.Layer = ComboBox_class_layer.Text
                                        End If
                                        BTrecord.AppendEntity(Linie1)
                                        Trans1.AddNewlyCreatedDBObject(Linie1, True)
                                        Dim Mtext1 As New MText
                                        Mtext1.Location = New Point3d(Point_PS1.X - Text_offset_no_plot, Start_point_for_bands.Y - (Nr_rand * Bands_Y_spacing) + Text_offset_no_plot, 0)
                                        Mtext1.TextHeight = 3
                                        Mtext1.Rotation = PI / 2
                                        Mtext1.Layer = No_plot

                                        Mtext1.Contents = Get_chainage_feet_from_double(Station1 + Get_equation_value(Station1), Round1)
                                        Mtext1.Attachment = AttachmentPoint.BottomLeft
                                        BTrecord.AppendEntity(Mtext1)
                                        Trans1.AddNewlyCreatedDBObject(Mtext1, True)
                                        Dim Band_text As New DBText
                                        Band_text.Layer = No_plot
                                        If RadioButton_left_right_viewport.Checked = True Then
                                            Band_text.Justify = AttachmentPoint.MiddleRight
                                            Band_text.AlignmentPoint = New Point3d(Start_ptX - 75, Start_point_for_bands.Y + Line_len_view_h / 2 - (Nr_rand * Bands_Y_spacing), 0)
                                        Else
                                            Band_text.Justify = AttachmentPoint.MiddleLeft
                                            Band_text.AlignmentPoint = New Point3d(Start_ptX + 75, Start_point_for_bands.Y + Line_len_view_h / 2 - (Nr_rand * Bands_Y_spacing), 0)
                                        End If
                                        Band_text.TextString = CStr(Nr_rand + 1)
                                        Band_text.Height = Line_len_view_h / 3
                                        BTrecord.AppendEntity(Band_text)
                                        Trans1.AddNewlyCreatedDBObject(Band_text, True)
                                    End If
                                Next




                                For i = 0 To Data_table_class_band_Excel_data.Rows.Count - 1
                                    If IsDBNull(Data_table_class_band_Excel_data.Rows(i).Item("BEGSTA")) = False And IsDBNull(Data_table_class_band_Excel_data.Rows(i).Item("ENDSTA")) = False Then
                                        Dim Station1 As Double = Data_table_class_band_Excel_data.Rows(i).Item("BEGSTA")
                                        Dim Station2 As Double = Data_table_class_band_Excel_data.Rows(i).Item("ENDSTA")
                                        Dim Class_string As String = "NO DATA"
                                        Dim DesignF_string As String = "NO DATA"

                                        If IsDBNull(Data_table_class_band_Excel_data.Rows(i).Item("CLASS")) = False Then
                                            Class_string = Data_table_class_band_Excel_data.Rows(i).Item("CLASS")
                                        End If

                                        If IsDBNull(Data_table_class_band_Excel_data.Rows(i).Item("DESIGN_FACTOR")) = False Then
                                            DesignF_string = Data_table_class_band_Excel_data.Rows(i).Item("DESIGN_FACTOR")
                                        End If

                                        Dim Point_PS As New Point3d(0, 0, 0)

                                        If IsNothing(Data_table_Matchlines) = False Then
                                            If Data_table_Matchlines.Rows.Count > 0 Then

                                                Dim Start1 As New Point3d
                                                Dim End1 As New Point3d
                                                Dim Match1 As Double = 0
                                                Dim Match2 As Double = 0

                                                Nr_rand = -1

                                                For j = 0 To Data_table_Matchlines.Rows.Count - 1
                                                    If IsDBNull(Data_table_Matchlines.Rows(j).Item("BEGSTA")) = False And IsDBNull(Data_table_Matchlines.Rows(j).Item("ENDSTA")) = False Then
                                                        Dim M1 As Double = 0
                                                        Dim M2 As Double = 0
                                                        M1 = Data_table_Matchlines.Rows(j).Item("BEGSTA")
                                                        M2 = Data_table_Matchlines.Rows(j).Item("ENDSTA")
                                                        Nr_rand = Nr_rand + 1

                                                        If Station2 * 5280 <= M2 And Station2 * 5280 >= M1 Then



                                                            If IsDBNull(Data_table_Matchlines.Rows(j).Item("X1")) = False And IsDBNull(Data_table_Matchlines.Rows(j).Item("Y1")) = False Then
                                                                If IsDBNull(Data_table_Matchlines.Rows(j).Item("X2")) = False And IsDBNull(Data_table_Matchlines.Rows(j).Item("Y2")) = False Then
                                                                    Start1 = New Point3d(Data_table_Matchlines.Rows(j).Item("X1"), Data_table_Matchlines.Rows(j).Item("Y1"), 0)
                                                                    End1 = New Point3d(Data_table_Matchlines.Rows(j).Item("X2"), Data_table_Matchlines.Rows(j).Item("Y2"), 0)
                                                                    If PolyCL.GetPointAtDist(M1).GetVectorTo(End1).Length < PolyCL.GetPointAtDist(M1).GetVectorTo(Start1).Length Then
                                                                        Dim ttt As New Point3d
                                                                        ttt = Start1
                                                                        Start1 = End1
                                                                        End1 = ttt
                                                                    End If
                                                                    Match2 = M2
                                                                    Match1 = M1



                                                                    Exit For
                                                                End If
                                                            End If

                                                        Else


                                                        End If
                                                    End If
                                                Next






                                                Dim Point_MS As New Point3d
                                                If Station2 * 5280 > PolyCL.Length Then
                                                    Station2 = PolyCL.Length / 5280
                                                End If
                                                Point_MS = PolyCL.GetPointAtDist(Station2 * 5280)

                                                Dim Viewport_line As New Line(Start1, End1)
                                                Dim Dist_from_match As Double = -1.1111
                                                Dim Point_on_line As New Point3d
                                                Point_on_line = Viewport_line.GetClosestPointTo(Point_MS, Vector3d.ZAxis, False)
                                                If IsNothing(Point_on_line) = False Then
                                                    Dist_from_match = Viewport_line.StartPoint.GetVectorTo(Point_on_line).Length
                                                End If
                                                Dim Val1 As Double = Station2 * 5280 - Match1
                                                If Not Dist_from_match = -1.1111 Then
                                                    Val1 = Dist_from_match
                                                End If

                                                Start_ptX = Start_point_for_bands.X + CDbl(TextBox_viewport_Width.Text) / 2 - Viewport_line.Length / 2

                                                If Nr_previous <> Nr_rand Then
                                                    Xprev = Start_ptX
                                                End If

                                                If RadioButton_left_right_viewport.Checked = True Then
                                                    Point_PS = New Point3d(Start_ptX + Val1, Start_point_for_bands.Y - (Nr_rand * Bands_Y_spacing), 0)
                                                    If Point_PS.X - Xprev < Min_dist Then
                                                        Point_PS = New Point3d(Xprev + Min_dist, Start_point_for_bands.Y - (Nr_rand * Bands_Y_spacing), 0)
                                                    End If
                                                Else
                                                    Point_PS = New Point3d(Start_ptX - Val1, Start_point_for_bands.Y - (Nr_rand * Bands_Y_spacing), 0)

                                                    If Xprev - Point_PS.X < Min_dist Then
                                                        Point_PS = New Point3d(Xprev - Min_dist, Start_point_for_bands.Y - (Nr_rand * Bands_Y_spacing), 0)
                                                    End If
                                                End If

                                            End If
                                        End If



                                        Dim Linie1 As New Line(New Point3d(Point_PS.X, Start_point_for_bands.Y - (Nr_rand * Bands_Y_spacing), 0), New Point3d(Point_PS.X, Start_point_for_bands.Y - (Nr_rand * Bands_Y_spacing) + Line_len_view_h, 0))
                                        If Not ComboBox_class_layer.Text = "" Then
                                            Linie1.Layer = ComboBox_class_layer.Text
                                        End If
                                        BTrecord.AppendEntity(Linie1)
                                        Trans1.AddNewlyCreatedDBObject(Linie1, True)


                                        Dim Mtext1 As New MText
                                        Mtext1.Location = New Point3d(Point_PS.X - Text_offset_no_plot, Start_point_for_bands.Y - (Nr_rand * Bands_Y_spacing) + Text_offset_no_plot, 0)
                                        Mtext1.TextHeight = TextHeight / 6
                                        Mtext1.Rotation = PI / 2
                                        Mtext1.Layer = No_plot
                                        Dim Feet1 As Double = Round(Station2 * 5280 + Get_equation_value(Station2 * 5280), Round1)
                                        Dim Mile1 As Double = Round(Feet1 / 5280, 1)
                                        Mtext1.Contents = Get_chainage_feet_from_double(Feet1, Round1) & "(" & Get_String_Rounded(Mile1, 1) & " miles)"
                                        Mtext1.Attachment = AttachmentPoint.BottomLeft

                                        BTrecord.AppendEntity(Mtext1)
                                        Trans1.AddNewlyCreatedDBObject(Mtext1, True)


                                        Dim Y_middle1 As Double = Point_PS.Y + Line_len_view_h / 2 + Line_len_view_h / 4
                                        Dim Y_middle2 As Double = Point_PS.Y + Line_len_view_h / 4

                                        Dim PT_ins11 As New Point3d
                                        If RadioButton_left_right_viewport.Checked = True Then
                                            PT_ins11 = New Point3d(Point_PS.X - (Point_PS.X - Xprev) / 2, Y_middle1, 0)
                                        Else
                                            PT_ins11 = New Point3d(Point_PS.X + (Xprev - Point_PS.X) / 2, Y_middle1, 0)
                                        End If

                                        Dim Mtext_class As New MText
                                        Mtext_class.Contents = Class_string
                                        Mtext_class.TextHeight = TextHeight
                                        Mtext_class.Location = PT_ins11
                                        Mtext_class.Rotation = 0
                                        If Layer_table.Has(ComboBox_class_layer.Text) = True Then
                                            Mtext_class.Layer = ComboBox_class_layer.Text
                                        End If
                                        Mtext_class.Attachment = AttachmentPoint.MiddleCenter
                                        Dim ObjId1 As TextStyleTableRecord
                                        If Text_style_table.Has(ComboBox_class_text_style.Text) = True Then
                                            ObjId1 = Text_style_table(ComboBox_class_text_style.Text).GetObject(OpenMode.ForRead)
                                            Mtext_class.TextStyleId = ObjId1.ObjectId
                                        End If
                                        BTrecord.AppendEntity(Mtext_class)
                                        Trans1.AddNewlyCreatedDBObject(Mtext_class, True)

                                        Dim PT_ins2 As New Point3d
                                        If RadioButton_left_right_viewport.Checked = True Then
                                            PT_ins2 = New Point3d(Point_PS.X - (Point_PS.X - Xprev) / 2, Y_middle2, 0)
                                        Else
                                            PT_ins2 = New Point3d(Point_PS.X + (Xprev - Point_PS.X) / 2, Y_middle2, 0)
                                        End If

                                        Dim Mtext_designF As New MText
                                        Mtext_designF.Contents = DesignF_string
                                        Mtext_designF.TextHeight = TextHeight
                                        Mtext_designF.Location = PT_ins2
                                        Mtext_designF.Rotation = 0
                                        If Layer_table.Has(ComboBox_class_layer.Text) = True Then
                                            Mtext_designF.Layer = ComboBox_class_layer.Text
                                        End If
                                        Mtext_designF.Attachment = AttachmentPoint.MiddleCenter
                                        If Text_style_table.Has(ComboBox_class_text_style.Text) = True Then
                                            Mtext_designF.TextStyleId = ObjId1.ObjectId
                                        End If
                                        BTrecord.AppendEntity(Mtext_designF)
                                        Trans1.AddNewlyCreatedDBObject(Mtext_designF, True)

                                        Xprev = Point_PS.X
                                        Nr_previous = Nr_rand

                                    End If







                                Next



                                Trans1.Commit()
                            End Using
                        End Using


                        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                    End If
                Else

                    MsgBox("You did not have data loaded for matchlines", MsgBoxStyle.Critical, "Dan says...")
                End If


            Catch ex As Exception
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If

    End Sub

    Private Sub TabPage_class_Click(sender As Object, e As EventArgs) Handles TabPage_class.Click
        Incarca_existing_layers_to_combobox(ComboBox_class_layer)

        If ComboBox_class_layer.Items.Count > 0 Then
            If ComboBox_class_layer.Items.Contains("CLASS") = True Then
                ComboBox_class_layer.SelectedIndex = ComboBox_class_layer.Items.IndexOf("CLASS")
            Else
                ComboBox_class_layer.SelectedIndex = 0
            End If
        End If


        Incarca_existing_textstyles_to_combobox(ComboBox_class_text_style)
        If ComboBox_class_text_style.Items.Count > 0 Then
            If ComboBox_class_text_style.Items.Contains("CLASS") = True Then
                ComboBox_class_text_style.SelectedIndex = ComboBox_class_text_style.Items.IndexOf("CLASS")
            Else
                ComboBox_class_text_style.SelectedIndex = 0
            End If
        End If
    End Sub

    Private Sub TabPage_WATER_Click(sender As Object, e As EventArgs) Handles TabPage_water.Click

        Incarca_existing_layers_to_combobox(ComboBox_layer_water)

        If ComboBox_layer_water.Items.Contains("TEXT") = True Then
            ComboBox_layer_water.SelectedIndex = ComboBox_layer_water.Items.IndexOf("TEXT")
        ElseIf ComboBox_layer_water.Items.Contains("CLASS") = True Then
            ComboBox_layer_water.SelectedIndex = ComboBox_layer_water.Items.IndexOf("CLASS")
        Else
            ComboBox_layer_water.SelectedIndex = 0
        End If

        Incarca_existing_textstyles_to_combobox(ComboBox_text_style_water)
        If ComboBox_text_style_water.Items.Count > 0 Then

            If ComboBox_text_style_water.Items.Contains("CLASS") = True Then
                ComboBox_text_style_water.SelectedIndex = ComboBox_text_style_water.Items.IndexOf("CLASS")
            ElseIf ComboBox_text_style_water.Items.Contains("Arial") = True Then
                ComboBox_text_style_water.SelectedIndex = ComboBox_text_style_water.Items.IndexOf("Arial")
            Else
                ComboBox_text_style_water.SelectedIndex = 0
            End If
        End If

    End Sub

    Private Sub Button_mtext_load_OD_Click(sender As Object, e As EventArgs) Handles Button_mtext_load_OD.Click


        If Freeze_operations = False Then
            Freeze_operations = True
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor

            Try
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                Colectie1 = New Specialized.StringCollection

                Editor1.SetImpliedSelection(Empty_array)

                Dim Rezultat_Parc As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select a sample parcel containing object data:"
                Object_Prompt.SingleOnly = True
                Rezultat_Parc = Editor1.GetSelection(Object_Prompt)
                If Not Rezultat_Parc.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Freeze_operations = False
                    Editor1.SetImpliedSelection(Empty_array)
                    Editor1.WriteMessage(vbLf & "Command:")
                    Freeze_operations = False
                    Exit Sub
                End If

                If Rezultat_Parc.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    If IsNothing(Rezultat_Parc) = False Then
                        Using Lock_dwg As DocumentLock = ThisDrawing.LockDocument
                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                                If ComboBox_Mtext_ObjectData.Items.Count = 0 Then ComboBox_Mtext_ObjectData.Items.Add("")

                                For j = 0 To Rezultat_Parc.Value.Count - 1
                                    Dim Ent1 As Entity = Trans1.GetObject(Rezultat_Parc.Value.Item(j).ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

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
                                                        With ComboBox_Mtext_ObjectData
                                                            If .Items.Contains(Field_def1.Name) = False Then
                                                                .Items.Add(Field_def1.Name)
                                                            End If
                                                        End With


                                                    Next
                                                Next
                                            End If
                                        End If
                                    End Using
                                Next


                                With ComboBox_Mtext_ObjectData
                                    If .Items.Count > 0 Then
                                        For i = 0 To .Items.Count - 1

                                            If .Items(i).ToString.ToUpper = "LANDUSE" Then
                                                .SelectedIndex = i
                                                Exit For
                                            Else
                                                .SelectedIndex = 0
                                            End If
                                        Next

                                    End If
                                End With


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


    Private Sub Button_load_data_for_mtext_Click(sender As Object, e As EventArgs) Handles Button_load_data_for_mtext.Click
        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Round1 As Integer = 0
            If IsNumeric(TextBox_dec1.Text) = True Then
                Round1 = CInt(TextBox_dec1.Text)
            End If

            Dim Error_index1 As Integer
            Dim Error_index2 As Integer

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor

            Dim RezultatCL As Autodesk.AutoCAD.EditorInput.PromptEntityResult

            Dim Object_PromptCL As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select centerline:")

            Object_PromptCL.SetRejectMessage(vbLf & "Please select a lightweight polyline or a 3D polyline")
            Object_PromptCL.AddAllowedClass(GetType(Polyline), True)
            Object_PromptCL.AddAllowedClass(GetType(Polyline3d), True)
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            RezultatCL = Editor1.GetEntity(Object_PromptCL)


            If Not RezultatCL.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                MsgBox("NO centerline")
                Editor1.WriteMessage(vbLf & "Command:")
                Freeze_operations = False
                Editor1.SetImpliedSelection(Empty_array)
                Exit Sub
            End If

            If IsNothing(Data_table_Matchlines) = True Then
                MsgBox("Please load your matchlines")
                Editor1.WriteMessage(vbLf & "Command:")
                Freeze_operations = False
                Editor1.SetImpliedSelection(Empty_array)
                Exit Sub
            End If


            Try

                Colectie1 = New Specialized.StringCollection

                Editor1.SetImpliedSelection(Empty_array)
                Label_od_LOADED.Text = "Identified:"

                Dim idarrayEmpty() As ObjectId
                ThisDrawing.Editor.SetImpliedSelection(idarrayEmpty)


                Data_table_Parcels = New System.Data.DataTable
                Data_table_Parcels.Columns.Add("LANDUSE", GetType(String))
                Data_table_Parcels.Columns.Add("X1", GetType(Double))
                Data_table_Parcels.Columns.Add("Y1", GetType(Double))
                Data_table_Parcels.Columns.Add("X2", GetType(Double))
                Data_table_Parcels.Columns.Add("Y2", GetType(Double))
                Data_table_Parcels.Columns.Add("BEGSTA", GetType(Double))
                Data_table_Parcels.Columns.Add("ENDSTA", GetType(Double))






                If RezultatCL.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                    If IsNothing(RezultatCL) = False Then
                        Using Lock_dwg As DocumentLock = ThisDrawing.LockDocument



                            Dim Index_data_parcels As Double = 0

                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                                PolyCL = TryCast(Trans1.GetObject(RezultatCL.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Polyline)

                                PolyCL3D = TryCast(Trans1.GetObject(RezultatCL.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Polyline3d)

                                If IsNothing(PolyCL3D) = False Then
                                    PolyCL = Creaza_polyline_din_polyline3d(Trans1, PolyCL3D)
                                End If



                                If PolyCL.NumberOfVertices < 2 Then
                                    MsgBox("NO centerline")
                                    Editor1.WriteMessage(vbLf & "Command:")
                                    Freeze_operations = False
                                    Editor1.SetImpliedSelection(Empty_array)
                                    Exit Sub
                                End If

                                Dim Poly_CL As New Polyline
                                Poly_CL = PolyCL.Clone
                                Poly_CL.Elevation = 0

                                Dim BTrecord As BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                                For Each OD1 As ObjectId In BTrecord
                                    Dim Ent1 As Entity = TryCast(Trans1.GetObject(OD1, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Entity)



                                    If IsNothing(Ent1) = False Then

                                        Dim ruleaza As Boolean = False
                                        If IsNothing(PolyCL3D) = False Then
                                            If Not OD1 = PolyCL3D.ObjectId Then
                                                ruleaza = True
                                            End If
                                        Else
                                            If Not OD1 = PolyCL.ObjectId Then
                                                ruleaza = True
                                            End If
                                        End If

                                        If ruleaza = True Then
                                            If TypeOf Ent1 Is Polyline Then
                                                Dim Poly_with_object_data As Polyline = Ent1
                                                Dim Poly_pt_intersectie As New Polyline
                                                Poly_pt_intersectie = Poly_with_object_data.Clone
                                                Poly_pt_intersectie.Elevation = 0

                                                Dim COL_INT As New Point3dCollection
                                                COL_INT = Intersect_on_both_operands(Poly_CL, Poly_pt_intersectie)

                                                Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                                                Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables
                                                Dim Id1 As ObjectId = Ent1.ObjectId


                                                If COL_INT.Count > 0 Then

                                                    Dim Data_table_unique As New System.Data.DataTable
                                                    Data_table_unique.Columns.Add("LANDUSE", GetType(String))

                                                    Data_table_unique.Columns.Add("X", GetType(Double))
                                                    Data_table_unique.Columns.Add("Y", GetType(Double))
                                                    Data_table_unique.Columns.Add("STATION", GetType(Double))


                                                    Dim LandUse1 As String = "XXX"




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

                                                                        If ComboBox_Mtext_ObjectData.Text = Field_def1.Name Then
                                                                            If Not Record1(i).StrValue = "" Then
                                                                                LandUse1 = Record1(i).StrValue
                                                                            End If
                                                                        End If

                                                                    Next
                                                                Next
                                                            End If

                                                        End If

                                                    End Using


                                                    Dim Start_end As Integer = 0

                                                    For m = 0 To COL_INT.Count - 1
                                                        Data_table_unique.Rows.Add()
                                                        Data_table_unique.Rows(m).Item("X") = COL_INT(m).X
                                                        Data_table_unique.Rows(m).Item("Y") = COL_INT(m).Y


                                                        Dim Station2D As Double = Round(Poly_CL.GetDistAtPoint(COL_INT(m)), Round1)

                                                        If IsNothing(PolyCL3D) = False Then
                                                            Dim Param2d As Double = PolyCL.GetParameterAtPoint(COL_INT(m))

                                                            Dim Station3D As Double = Round(PolyCL3D.GetDistanceAtParameter(Param2d), Round1)
                                                            Data_table_unique.Rows(m).Item("STATION") = Station3D


                                                        Else

                                                            Data_table_unique.Rows(m).Item("STATION") = Station2D
                                                        End If


                                                        Data_table_unique.Rows(m).Item("LANDUSE") = LandUse1

                                                    Next

                                                    Data_table_unique = Sort_data_table(Data_table_unique, "STATION")




                                                    Dim Nr_magic As Integer = COL_INT.Count

                                                    If Not Floor(Nr_magic / 2) * 2 = Nr_magic Then
                                                        Data_table_unique.Rows.Add()
                                                        Data_table_unique.Rows(Nr_magic).Item("X") = Data_table_unique.Rows(Nr_magic - 1).Item("X")
                                                        Data_table_unique.Rows(Nr_magic).Item("Y") = Data_table_unique.Rows(Nr_magic - 1).Item("Y")
                                                        Data_table_unique.Rows(Nr_magic).Item("STATION") = Data_table_unique.Rows(Nr_magic - 1).Item("STATION") + 1
                                                        Data_table_unique.Rows(Nr_magic).Item("LANDUSE") = Data_table_unique.Rows(Nr_magic - 1).Item("LANDUSE")
                                                    End If


start1:

                                                    For m = 1 To COL_INT.Count - 1
                                                        Dim S As Double = Data_table_unique.Rows(m).Item("STATION")
                                                        Dim Sp As Double = Data_table_unique.Rows(m - 1).Item("STATION")

                                                        If S = Sp Then
                                                            Data_table_unique.Rows(m).Item("STATION") = Data_table_unique.Rows(m).Item("STATION") + 1
                                                            GoTo start1
                                                        End If

                                                    Next


                                                    Data_table_unique = Sort_data_table(Data_table_unique, "STATION")


                                                    For s = 0 To Data_table_unique.Rows.Count - 2 Step 2

                                                        Data_table_Parcels.Rows.Add()
                                                        Data_table_Parcels.Rows(Index_data_parcels).Item("X1") = Data_table_unique.Rows(s).Item("X")
                                                        Data_table_Parcels.Rows(Index_data_parcels).Item("Y1") = Data_table_unique.Rows(s).Item("Y")
                                                        Data_table_Parcels.Rows(Index_data_parcels).Item("X2") = Data_table_unique.Rows(s + 1).Item("X")
                                                        Data_table_Parcels.Rows(Index_data_parcels).Item("Y2") = Data_table_unique.Rows(s + 1).Item("Y")
                                                        Data_table_Parcels.Rows(Index_data_parcels).Item("BEGSTA") = Data_table_unique.Rows(s).Item("STATION")
                                                        Data_table_Parcels.Rows(Index_data_parcels).Item("ENDSTA") = Data_table_unique.Rows(s + 1).Item("STATION")
                                                        Data_table_Parcels.Rows(Index_data_parcels).Item("LANDUSE") = Data_table_unique.Rows(s).Item("LANDUSE")

                                                        Index_data_parcels = Index_data_parcels + 1
                                                    Next

                                                End If
                                            End If
                                        End If
                                    Else


                                    End If
                                Next

                                Trans1.Abort()
                            End Using
                        End Using
                    End If
                End If







                If Data_table_Matchlines.Rows.Count > 0 Then

                    Dim Index_data_table_add As Integer = Data_table_Parcels.Rows.Count
                    Dim Nr_rand As Integer = Data_table_Parcels.Rows.Count


                    For i = 0 To Nr_rand - 1
                        If IsDBNull(Data_table_Parcels.Rows(i).Item("BEGSTA")) = False And IsDBNull(Data_table_Parcels.Rows(i).Item("ENDSTA")) = False Then
                            Error_index1 = i

                            Dim Station1 As Double = Data_table_Parcels.Rows(i).Item("BEGSTA")
                            Dim Station2 As Double = Data_table_Parcels.Rows(i).Item("ENDSTA")



                            Dim Luse1 As String = ""

                            Dim X1 As Double = 0
                            Dim X2 As Double = 0
                            Dim Y1 As Double = 0
                            Dim Y2 As Double = 0

                            If IsDBNull(Data_table_Parcels.Rows(i).Item("LANDUSE")) = False Then
                                Luse1 = Data_table_Parcels.Rows(i).Item("LANDUSE")
                            End If


                            If IsDBNull(Data_table_Parcels.Rows(i).Item("X1")) = False Then
                                X1 = Data_table_Parcels.Rows(i).Item("X1")
                            End If
                            If IsDBNull(Data_table_Parcels.Rows(i).Item("X2")) = False Then
                                X2 = Data_table_Parcels.Rows(i).Item("X2")
                            End If
                            If IsDBNull(Data_table_Parcels.Rows(i).Item("Y1")) = False Then
                                Y1 = Data_table_Parcels.Rows(i).Item("Y1")
                            End If
                            If IsDBNull(Data_table_Parcels.Rows(i).Item("Y2")) = False Then
                                Y2 = Data_table_Parcels.Rows(i).Item("Y2")
                            End If

                            Dim I_start As Integer = 0
                            Dim go_to_add_S1_S2 As Boolean = False




123:





                            For j = I_start To Data_table_Matchlines.Rows.Count - 1
                                If IsDBNull(Data_table_Matchlines.Rows(j).Item("BEGSTA")) = False And IsDBNull(Data_table_Matchlines.Rows(j).Item("ENDSTA")) = False Then


                                    Error_index2 = j
                                    If i = 119 And j = 50 Then
                                        ' MsgBox("investigate")
                                    End If

                                    Dim M1 As Double = Data_table_Matchlines.Rows(j).Item("BEGSTA")
                                    Dim M2 As Double = Data_table_Matchlines.Rows(j).Item("ENDSTA")



                                    If IsNothing(PolyCL3D) = False Then
                                        If M2 > PolyCL3D.Length Then
                                            M2 = PolyCL3D.Length
                                        End If

                                        If Station2 > PolyCL3D.Length Then
                                            Station2 = PolyCL3D.Length
                                        End If
                                    Else
                                        If M2 > PolyCL.Length Then
                                            M2 = PolyCL.Length
                                        End If
                                        If Station2 > PolyCL.Length Then
                                            Station2 = PolyCL.Length
                                        End If
                                    End If



                                    If go_to_add_S1_S2 = True Then
                                        GoTo label_add_S1_S2
                                    End If




                                    'case 1
                                    If M1 <= Station1 And M2 <= Station2 And M1 <= Station2 And M2 >= Station1 Then
                                        Data_table_Parcels.Rows(i).Item("ENDSTA") = M2
                                        Station1 = M2
                                        I_start = j + 1
                                        go_to_add_S1_S2 = True
                                        GoTo 123
                                    End If

                                    ' case 5
                                    If Station1 >= M1 And Station2 <= M2 Then
                                        Exit For
                                    End If




label_add_S1_S2:

                                    ' add S1, S2
                                    If Station1 >= M1 And Station2 <= M2 Then
                                        Data_table_Parcels.Rows.Add()
                                        Data_table_Parcels.Rows(Index_data_table_add).Item("BEGSTA") = Station1
                                        Data_table_Parcels.Rows(Index_data_table_add).Item("ENDSTA") = Station2

                                        If IsNothing(PolyCL3D) = False Then
                                            Data_table_Parcels.Rows(Index_data_table_add).Item("X2") = PolyCL3D.GetPointAtDist(Station2).X
                                            Data_table_Parcels.Rows(Index_data_table_add).Item("Y2") = PolyCL3D.GetPointAtDist(Station2).Y
                                            Data_table_Parcels.Rows(Index_data_table_add).Item("X1") = PolyCL3D.GetPointAtDist(Station1).X
                                            Data_table_Parcels.Rows(Index_data_table_add).Item("Y1") = PolyCL3D.GetPointAtDist(Station1).Y
                                        Else

                                            Data_table_Parcels.Rows(Index_data_table_add).Item("X2") = PolyCL.GetPointAtDist(Station2).X
                                            Data_table_Parcels.Rows(Index_data_table_add).Item("Y2") = PolyCL.GetPointAtDist(Station2).Y
                                            Data_table_Parcels.Rows(Index_data_table_add).Item("X1") = PolyCL.GetPointAtDist(Station1).X
                                            Data_table_Parcels.Rows(Index_data_table_add).Item("Y1") = PolyCL.GetPointAtDist(Station1).Y
                                        End If

                                        Data_table_Parcels.Rows(Index_data_table_add).Item("LANDUSE") = Luse1
                                        Index_data_table_add = Index_data_table_add + 1
                                        Exit For

                                    ElseIf Station1 <= M2 And Station1 >= M1 Then
                                        Data_table_Parcels.Rows.Add()
                                        Data_table_Parcels.Rows(Index_data_table_add).Item("BEGSTA") = Station1
                                        Data_table_Parcels.Rows(Index_data_table_add).Item("ENDSTA") = M2

                                        If IsNothing(PolyCL3D) = False Then
                                            Data_table_Parcels.Rows(Index_data_table_add).Item("X2") = PolyCL3D.GetPointAtDist(M2).X
                                            Data_table_Parcels.Rows(Index_data_table_add).Item("Y2") = PolyCL3D.GetPointAtDist(M2).Y
                                            Data_table_Parcels.Rows(Index_data_table_add).Item("X1") = PolyCL3D.GetPointAtDist(Station1).X
                                            Data_table_Parcels.Rows(Index_data_table_add).Item("Y1") = PolyCL3D.GetPointAtDist(Station1).Y
                                        Else

                                            Data_table_Parcels.Rows(Index_data_table_add).Item("X2") = PolyCL.GetPointAtDist(M2).X
                                            Data_table_Parcels.Rows(Index_data_table_add).Item("Y2") = PolyCL.GetPointAtDist(M2).Y
                                            Data_table_Parcels.Rows(Index_data_table_add).Item("X1") = PolyCL.GetPointAtDist(Station1).X
                                            Data_table_Parcels.Rows(Index_data_table_add).Item("Y1") = PolyCL.GetPointAtDist(Station1).Y
                                        End If

                                        Data_table_Parcels.Rows(Index_data_table_add).Item("LANDUSE") = Luse1

                                        Index_data_table_add = Index_data_table_add + 1


                                        Station1 = M2
                                        I_start = j + 1
                                        go_to_add_S1_S2 = True
                                        GoTo 123

                                    End If

                                End If




                            Next

                        End If
                    Next



                End If



                Data_table_Parcels = Sort_data_table(Data_table_Parcels, "BEGSTA")

                Add_to_clipboard_Data_table(Data_table_Parcels)



                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")

            Catch ex As Exception
                Editor1.WriteMessage(vbLf & "Command:")
                Freeze_operations = False
                MsgBox(ex.Message)

                MsgBox(Error_index1 & vbCrLf & Error_index2)

            End Try


            Freeze_operations = False
        End If
    End Sub


    Private Sub Button_LAND_USE_Click(sender As Object, e As EventArgs) Handles Button_LAND_USE.Click
        If IsNumeric(TextBox_viewport_Height.Text) = False Then
            MsgBox("Not numeric viewport height specified")
            Exit Sub
        End If
        If IsNumeric(TextBox_viewport_SCALE.Text) = False Then
            MsgBox("Not numeric viewport scale specified")
            Exit Sub
        End If
        If IsNumeric(TextBox_X_MS.Text) = False Then
            MsgBox("Not numeric 0+000 X position in modelspace specified")
            Exit Sub
        End If
        If IsNumeric(TextBox_Y_MS.Text) = False Then
            MsgBox("Not numeric 0+000 Y position in modelspace specified")
            Exit Sub
        End If
        If Freeze_operations = False Then
            Freeze_operations = True
            Freeze_operations = True

            Dim Round1 As Integer = 0
            If IsNumeric(TextBox_dec1.Text) = True Then
                Round1 = CInt(TextBox_dec1.Text)
            End If

            Try

                If Data_table_Parcels.Rows.Count > 0 Then
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor

                    Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                            BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)


                            Dim Text_style_table As Autodesk.AutoCAD.DatabaseServices.TextStyleTable = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                            Dim Layer_table As Autodesk.AutoCAD.DatabaseServices.LayerTable = Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                            Dim Start_point_for_bands As New Point3d(CDbl(TextBox_X_MS.Text), CDbl(TextBox_Y_MS.Text), 0)
                            Dim Line_len As Double = Abs(CDbl(TextBox_viewport_Height.Text) * CDbl(TextBox_viewport_SCALE.Text))


                            Dim Bands_Y_spacing As Double = 200
                            Bands_Y_spacing = Ceiling(2.25 * Line_len / 50) * 50

                            If IsNumeric(TextBox_band_spacing_forced.Text) = True Then
                                Bands_Y_spacing = CDbl(TextBox_band_spacing_forced.Text)
                            End If

                            TextBox_BAND_SPACING.Text = Bands_Y_spacing


                            Dim Min_dist As Double = 50
                            If IsNumeric(TextBox_minimum_distance_mtext.Text) = True Then
                                Min_dist = CDbl(TextBox_minimum_distance_mtext.Text)
                            End If
                            Dim Text_offset_no_plot As Double = 1

                            Dim TextHeight As Double = 8
                            If IsNumeric(TextBox_mtext_height.Text) = True Then
                                TextHeight = CDbl(TextBox_mtext_height.Text)
                            End If

                            Dim Station2_prev As Double = 0

                            Dim X2prev As Double = Start_point_for_bands.X



                            Creaza_layer(No_plot, 40, No_plot, False)





                            For i = 0 To Data_table_Parcels.Rows.Count - 1
                                If IsDBNull(Data_table_Parcels.Rows(i).Item("BEGSTA")) = False And IsDBNull(Data_table_Parcels.Rows(i).Item("ENDSTA")) = False Then
                                    Dim Station1 As Double = Data_table_Parcels.Rows(i).Item("BEGSTA")
                                    Dim Station2 As Double = Data_table_Parcels.Rows(i).Item("ENDSTA")

                                    Dim Land_use As String = "XXX"

                                    If IsDBNull(Data_table_Parcels.Rows(i).Item("LANDUSE")) = False Then
                                        Land_use = Data_table_Parcels.Rows(i).Item("LANDUSE")
                                    End If



                                    Dim Viewport_line As New Line(New Point3d(0, 0, 0), New Point3d(0, 0, 0))

                                    Dim Point_B1 As New Point3d
                                    Dim Point_B2 As New Point3d

                                    Dim Dist_from_start1 As Double
                                    Dim Dist_from_start2 As Double

                                    Dim M1 As Double = 0
                                    Dim M2 As Double = 0

                                    Dim Band_number As Integer = -1

                                    For j = 0 To Data_table_Matchlines.Rows.Count - 1
                                        If IsDBNull(Data_table_Matchlines.Rows(j).Item("BEGSTA")) = False And IsDBNull(Data_table_Matchlines.Rows(j).Item("ENDSTA")) = False Then

                                            M1 = Data_table_Matchlines.Rows(j).Item("BEGSTA")
                                            M2 = Data_table_Matchlines.Rows(j).Item("ENDSTA")
                                            Band_number = Band_number + 1
                                            If Round(Station1, Round1) <= Round(M2, Round1) And Round(Station1, Round1) >= Round(M1, Round1) And Round(Station2, Round1) <= Round(M2, Round1) And Round(Station2, Round1) >= Round(M1, Round1) Then
                                                If IsDBNull(Data_table_Matchlines.Rows(j).Item("X1")) = False And IsDBNull(Data_table_Matchlines.Rows(j).Item("Y1")) = False Then
                                                    If IsDBNull(Data_table_Matchlines.Rows(j).Item("X2")) = False And IsDBNull(Data_table_Matchlines.Rows(j).Item("Y2")) = False Then
                                                        Dim Start1 As New Point3d
                                                        Dim End1 As New Point3d

                                                        Start1 = New Point3d(Data_table_Matchlines.Rows(j).Item("X1"), Data_table_Matchlines.Rows(j).Item("Y1"), 0)
                                                        End1 = New Point3d(Data_table_Matchlines.Rows(j).Item("X2"), Data_table_Matchlines.Rows(j).Item("Y2"), 0)

                                                        If IsNothing(PolyCL3D) = False Then

                                                            Start1 = PolyCL.GetClosestPointTo(Start1, Vector3d.ZAxis, False)
                                                            End1 = PolyCL.GetClosestPointTo(End1, Vector3d.ZAxis, False)
                                                            Dim Param1 As Double = PolyCL.GetParameterAtPoint(Start1)
                                                            Dim Param2 As Double = PolyCL.GetParameterAtPoint(End1)

                                                            Dim Point1 As New Point3d
                                                            Point1 = PolyCL3D.GetPointAtParameter(Param1)
                                                            Dim Point2 As New Point3d
                                                            Point2 = PolyCL3D.GetPointAtParameter(Param2)

                                                            If PolyCL3D.GetPointAtDist(M1).GetVectorTo(Point2).Length < PolyCL3D.GetPointAtDist(M1).GetVectorTo(Point1).Length Then
                                                                Dim ttt As New Point3d
                                                                ttt = Start1
                                                                Start1 = End1
                                                                End1 = ttt
                                                            End If

                                                        Else
                                                            If PolyCL.GetPointAtDist(M1).GetVectorTo(End1).Length < PolyCL.GetPointAtDist(M1).GetVectorTo(Start1).Length Then
                                                                Dim ttt As New Point3d
                                                                ttt = Start1
                                                                Start1 = End1
                                                                End1 = ttt
                                                            End If
                                                        End If
                                                        Viewport_line = New Line(Start1, End1)

                                                        If RadioButton_Left_right.Checked = True Then
                                                            Start_point_for_bands = New Point3d(CDbl(TextBox_X_MS.Text) + (CDbl(TextBox_viewport_Width.Text) / 2) * CDbl(TextBox_viewport_SCALE.Text) - Viewport_line.Length / 2, Start_point_for_bands.Y, 0)
                                                        Else
                                                            Start_point_for_bands = New Point3d(CDbl(TextBox_X_MS.Text) - (CDbl(TextBox_viewport_Width.Text) / 2) * CDbl(TextBox_viewport_SCALE.Text) + Viewport_line.Length / 2, Start_point_for_bands.Y, 0)
                                                        End If

                                                        Exit For
                                                    End If
                                                End If





                                            End If



                                        End If
                                    Next

                                    If Viewport_line.Length > 0 Then



                                        If IsNothing(PolyCL3D) = False Then
                                            Dim Point1 As New Point3d
                                            Point1 = PolyCL3D.GetPointAtDist(Station1)

                                            Dim Point2 As New Point3d
                                            Point2 = PolyCL3D.GetPointAtDist(Station2)

                                            Point1 = New Point3d(Point1.X, Point1.Y, 0)
                                            Point2 = New Point3d(Point2.X, Point2.Y, 0)

                                            Dim PointV1 As New Point3d
                                            PointV1 = Viewport_line.GetClosestPointTo(Point1, Vector3d.ZAxis, False)
                                            Dim PointV2 As New Point3d
                                            PointV2 = Viewport_line.GetClosestPointTo(Point2, Vector3d.ZAxis, False)
                                            Dist_from_start1 = PointV1.GetVectorTo(Viewport_line.StartPoint).Length
                                            Dist_from_start2 = PointV2.GetVectorTo(Viewport_line.StartPoint).Length


                                        Else
                                            If Station1 > PolyCL.Length Then
                                                Station1 = PolyCL.Length
                                            End If
                                            Dim Point1 As New Point3d
                                            Point1 = PolyCL.GetPointAtDist(Station1)

                                            Dim Point2 As New Point3d
                                            Point2 = PolyCL.GetPointAtDist(Station2)


                                            Dim PointV1 As New Point3d
                                            PointV1 = Viewport_line.GetClosestPointTo(Point1, Vector3d.ZAxis, False)
                                            Dim PointV2 As New Point3d
                                            PointV2 = Viewport_line.GetClosestPointTo(Point2, Vector3d.ZAxis, False)
                                            Dist_from_start1 = PointV1.GetVectorTo(Viewport_line.StartPoint).Length
                                            Dist_from_start2 = PointV2.GetVectorTo(Viewport_line.StartPoint).Length
                                        End If


                                        Dim Width1 As Double = Dist_from_start2 - Dist_from_start1
                                        If Width1 < Min_dist Then
                                            Width1 = Min_dist
                                        End If

                                        If RadioButton_left_right_viewport.Checked = True Then
                                            Point_B1 = New Point3d(X2prev, Start_point_for_bands.Y - (Band_number * Bands_Y_spacing), 0)
                                            Point_B2 = New Point3d(Point_B1.X + Width1, Point_B1.Y, 0)
                                        Else
                                            Point_B1 = New Point3d(X2prev, Start_point_for_bands.Y - (Band_number * Bands_Y_spacing), 0)
                                            Point_B2 = New Point3d(Point_B1.X - Width1, Point_B1.Y, 0)
                                        End If





                                        If Round(Station1, Round1) = Round(M1, Round1) Then

                                            Dim Band_text As New DBText
                                            Band_text.Layer = No_plot
                                            If RadioButton_left_right_viewport.Checked = True Then

                                                Band_text.Justify = AttachmentPoint.MiddleRight
                                                Band_text.AlignmentPoint = New Point3d(Start_point_for_bands.X - 5 * TextHeight, Start_point_for_bands.Y + Line_len / 2 - (Band_number * Bands_Y_spacing), 0)
                                            Else

                                                Band_text.Justify = AttachmentPoint.MiddleLeft
                                                Band_text.AlignmentPoint = New Point3d(Start_point_for_bands.X + 5 * TextHeight, Start_point_for_bands.Y + Line_len / 2 - (Band_number * Bands_Y_spacing), 0)
                                            End If

                                            Band_text.TextString = CStr(Band_number + 1)
                                            Band_text.Height = 7.5 * TextHeight

                                            BTrecord.AppendEntity(Band_text)
                                            Trans1.AddNewlyCreatedDBObject(Band_text, True)

                                            If RadioButton_left_right_viewport.Checked = True Then
                                                Point_B1 = New Point3d(Start_point_for_bands.X, Start_point_for_bands.Y - (Band_number * Bands_Y_spacing), 0)
                                                Point_B2 = New Point3d(Point_B1.X + Width1, Point_B1.Y, 0)
                                            Else
                                                Point_B1 = New Point3d(Start_point_for_bands.X, Start_point_for_bands.Y - (Band_number * Bands_Y_spacing), 0)
                                                Point_B2 = New Point3d(Point_B1.X - Width1, Point_B1.Y, 0)
                                            End If


                                            Dim Linie1 As New Line(New Point3d(Point_B1.X, Start_point_for_bands.Y - (Band_number * Bands_Y_spacing), 0), New Point3d(Point_B1.X, Start_point_for_bands.Y - (Band_number * Bands_Y_spacing) + Line_len, 0))
                                            If Not ComboBox_layers.Text = "" Then
                                                Linie1.Layer = ComboBox_layers.Text
                                            End If
                                            BTrecord.AppendEntity(Linie1)
                                            Trans1.AddNewlyCreatedDBObject(Linie1, True)

                                            Dim Mtext1 As New MText
                                            Mtext1.Location = New Point3d(Point_B1.X, Start_point_for_bands.Y - (Band_number * Bands_Y_spacing), 0)
                                            Mtext1.TextHeight = 0.5 * TextHeight
                                            Mtext1.Rotation = PI / 2

                                            Mtext1.Layer = No_plot



                                            Dim ContinutMtext1 As String

                                            If CheckBox_use_equation.Checked = True Then

                                                If IsNumeric(TextBox_textwidth.Text) = True Then
                                                    ContinutMtext1 = "{\W" & TextBox_textwidth.Text & ";" & Get_chainage_feet_from_double(Station1 + Get_equation_value(Station1), Round1) & "}"
                                                Else
                                                    ContinutMtext1 = Get_chainage_feet_from_double(Station1 + Get_equation_value(Station1), Round1)
                                                End If

                                            Else
                                                If IsNumeric(TextBox_textwidth.Text) = True Then
                                                    ContinutMtext1 = "{\W" & TextBox_textwidth.Text & ";" & Get_chainage_feet_from_double(Station1, Round1) & "}"
                                                Else
                                                    ContinutMtext1 = Get_chainage_feet_from_double(Station1, Round1)
                                                End If

                                            End If



                                            Mtext1.Contents = ContinutMtext1



                                            Mtext1.Attachment = AttachmentPoint.BottomLeft


                                            BTrecord.AppendEntity(Mtext1)
                                            Trans1.AddNewlyCreatedDBObject(Mtext1, True)



                                        End If 'este de la If Station1 = M1 

                                        Dim Linie2 As New Line(New Point3d(Point_B2.X, Start_point_for_bands.Y - (Band_number * Bands_Y_spacing), 0), New Point3d(Point_B2.X, Start_point_for_bands.Y - (Band_number * Bands_Y_spacing) + Line_len, 0))
                                        If Not ComboBox_layers.Text = "" Then
                                            Linie2.Layer = ComboBox_layers.Text
                                        End If
                                        BTrecord.AppendEntity(Linie2)
                                        Trans1.AddNewlyCreatedDBObject(Linie2, True)

                                        Dim Mtext2 As New MText
                                        Mtext2.Location = New Point3d(Point_B2.X, Start_point_for_bands.Y - (Band_number * Bands_Y_spacing), 0)
                                        Mtext2.TextHeight = 0.5 * TextHeight
                                        Mtext2.Rotation = PI / 2
                                        Mtext2.Layer = No_plot



                                        Dim ContinutMtext2 As String

                                        If CheckBox_use_equation.Checked = True Then

                                            If IsNumeric(TextBox_textwidth.Text) = True Then
                                                ContinutMtext2 = "{\W" & TextBox_textwidth.Text & ";" & Get_chainage_feet_from_double(Station2 + Get_equation_value(Station2), Round1) & "}"
                                            Else
                                                ContinutMtext2 = Get_chainage_feet_from_double(Station2 + Get_equation_value(Station2), Round1)
                                            End If
                                        Else
                                            If IsNumeric(TextBox_textwidth.Text) = True Then
                                                ContinutMtext2 = "{\W" & TextBox_textwidth.Text & ";" & Get_chainage_feet_from_double(Station2, Round1) & "}"
                                            Else
                                                ContinutMtext2 = Get_chainage_feet_from_double(Station2, Round1)
                                            End If

                                        End If



                                        Mtext2.Contents = ContinutMtext2



                                        Mtext2.Attachment = AttachmentPoint.BottomLeft



                                        BTrecord.AppendEntity(Mtext2)
                                        Trans1.AddNewlyCreatedDBObject(Mtext2, True)

                                        Dim Insertion_point As New Point3d
                                        Insertion_point = New Point3d((Point_B1.X + Point_B2.X) / 2, Point_B1.Y + Line_len / 2, 0)


                                        Dim Mtext_ll As New MText
                                        Mtext_ll.Contents = Land_use
                                        Mtext_ll.TextHeight = CDbl(TextBox_mtext_height.Text)
                                        Mtext_ll.Location = Insertion_point
                                        Mtext_ll.Rotation = 0
                                        If Layer_table.Has(ComboBox_layers_for_Mtext.Text) = True Then
                                            Mtext_ll.Layer = ComboBox_layers_for_Mtext.Text
                                        End If
                                        Mtext_ll.Attachment = AttachmentPoint.MiddleCenter
                                        Dim ObjId1 As TextStyleTableRecord
                                        If Text_style_table.Has(ComboBox_text_style_mtext.Text) = True Then
                                            ObjId1 = Text_style_table(ComboBox_text_style_mtext.Text).GetObject(OpenMode.ForRead)
                                            Mtext_ll.TextStyleId = ObjId1.ObjectId
                                        End If
                                        BTrecord.AppendEntity(Mtext_ll)
                                        Trans1.AddNewlyCreatedDBObject(Mtext_ll, True)




                                        X2prev = Point_B2.X



                                        Station2_prev = Station2

                                    End If 'asta e de la viewport length>0

                                End If ' asta e de la If IsDBNull(Data_table_Parcels.Rows(i).Item("BEGSTA")) = False And IsDBNull(Data_table_Parcels.Rows(i).Item("ENDSTA")) = False
                            Next



                            Trans1.Commit()
                        End Using
                    End Using


                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: there are " & Data_table_Parcels.Rows.Count & " rows")
                End If



            Catch ex As Exception
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
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



    Private Sub Button_insert_mleaders_Click(sender As Object, e As EventArgs) Handles Button_insert_mleaders.Click
        If Freeze_operations = False Then
            Freeze_operations = True

            Try

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Dim Start1 As Integer = 0
                Dim End1 As Integer = 0
                If IsNumeric(TextBox_debug_row1.Text) = True Then
                    Start1 = CInt(TextBox_debug_row1.Text)
                End If
                If IsNumeric(TextBox_debug_row2.Text) = True Then
                    End1 = CInt(TextBox_debug_row2.Text)
                End If

                If End1 = 0 Or Start1 = 0 Then
                    Freeze_operations = False
                    Exit Sub
                End If
                If End1 < Start1 Then
                    Freeze_operations = False
                    Exit Sub
                End If
                Dim Column_sta1 As String = TextBox_debug_Col1.Text.ToUpper

                Dim Column_sta2 As String = TextBox_debug_Col2.Text.ToUpper

                Dim Column_formula As String = TextBox_debug_formula.Text.ToUpper

                Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor

                Dim RezultatCL As Autodesk.AutoCAD.EditorInput.PromptEntityResult

                Dim Object_PromptCL As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select centerline:")

                Object_PromptCL.SetRejectMessage(vbLf & "Please select a lightweight polyline or a 3D polyline")
                Object_PromptCL.AddAllowedClass(GetType(Polyline), True)
                Object_PromptCL.AddAllowedClass(GetType(Polyline3d), True)
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                RezultatCL = Editor1.GetEntity(Object_PromptCL)


                If Not RezultatCL.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    MsgBox("NO centerline")
                    Editor1.WriteMessage(vbLf & "Command:")
                    Freeze_operations = False
                    Editor1.SetImpliedSelection(Empty_array)
                    Exit Sub
                End If



                If RezultatCL.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                    If IsNothing(RezultatCL) = False Then
                        Using Lock_dwg As DocumentLock = ThisDrawing.LockDocument

                            Creaza_layer(No_plot, 40, No_plot, False)

                            Dim Index_data_parcels As Double = 0

                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Dim PolyCL As Polyline
                                PolyCL = TryCast(Trans1.GetObject(RezultatCL.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Polyline)
                                Dim PolyCL3D As Polyline3d
                                PolyCL3D = TryCast(Trans1.GetObject(RezultatCL.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Polyline3d)






                                For i = Start1 To End1
                                    Dim Station_string1 As String = W1.Range(Column_sta1 & i).Value2
                                    Dim Station_string2 As String = W1.Range(Column_sta2 & i).Value2
                                    If i < End1 Then
                                        If IsNumeric(Station_string1) = True And IsNumeric(Station_string2) = True Then
                                            W1.Range(Column_formula & i).Formula = "=" & Column_sta2 & i & "-" & Column_sta1 & (i + 1)
                                            If Not W1.Range(Column_formula & i).Value2 = 0 Then
                                                If IsNumeric(Station_string1) = True Then
                                                    Dim Station1 As Double = Abs(CDbl(Station_string1))

                                                    If IsNothing(PolyCL) = False Then
                                                        If PolyCL.Length >= Station1 Then
                                                            Dim Mleader1 As New MLeader
                                                            Mleader1 = Creaza_Mleader_nou_fara_UCS_transform(PolyCL.GetPointAtDist(Station1), Get_chainage_feet_from_double(Station1, 0), 5, 1, 2.5, 1, 10)
                                                            Mleader1.Layer = No_plot
                                                        End If

                                                    End If

                                                    If IsNothing(PolyCL3D) = False Then
                                                        If PolyCL3D.Length >= Station1 Then
                                                            Dim Mleader1 As New MLeader
                                                            Mleader1 = Creaza_Mleader_nou_fara_UCS_transform(PolyCL3D.GetPointAtDist(Station1), Get_chainage_feet_from_double(Station1, 0), 5, 1, 2.5, 1, 10)
                                                            Mleader1.Layer = No_plot
                                                        End If

                                                    End If

                                                End If
                                                If IsNumeric(Station_string2) = True Then
                                                    Dim Station2 As Double = Abs(CDbl(Station_string2))

                                                    If IsNothing(PolyCL) = False Then
                                                        If PolyCL.Length >= Station2 Then
                                                            Dim Mleader1 As New MLeader
                                                            Mleader1 = Creaza_Mleader_nou_fara_UCS_transform(PolyCL.GetPointAtDist(Station2), Get_chainage_feet_from_double(Station2, 0), 5, 1, 2.5, 1, 10)
                                                            Mleader1.Layer = No_plot
                                                        End If

                                                    End If

                                                    If IsNothing(PolyCL3D) = False Then
                                                        If PolyCL3D.Length >= Station2 Then
                                                            Dim Mleader1 As New MLeader
                                                            Mleader1 = Creaza_Mleader_nou_fara_UCS_transform(PolyCL3D.GetPointAtDist(Station2), Get_chainage_feet_from_double(Station2, 0), 5, 1, 2.5, 1, 10)
                                                            Mleader1.Layer = No_plot
                                                        End If

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
                'MsgBox(Data_table_Centerline.Rows.Count)


            Catch ex As Exception
                MsgBox(ex.Message)
                Freeze_operations = False
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


    Private Sub Button_calc_sta_eq_values_Click(sender As Object, e As EventArgs) Handles Button_calc_sta_eq_values.Click
        If Freeze_operations = False Then
            Freeze_operations = True

            Try

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Dim Start1 As Integer = 0
                Dim End1 As Integer = 0
                If IsNumeric(TextBox_m_start.Text) = True Then
                    Start1 = CInt(TextBox_m_start.Text)
                End If
                If IsNumeric(TextBox_m_end.Text) = True Then
                    End1 = CInt(TextBox_m_end.Text)
                End If

                If End1 = 0 Or Start1 = 0 Then
                    Freeze_operations = False
                    Exit Sub
                End If
                If End1 < Start1 Then
                    Freeze_operations = False
                    Exit Sub
                End If

                Dim Round_sta As Integer = 0
                If IsNumeric(TextBox_dec1.Text) = True Then
                    Round_sta = CInt(TextBox_dec1.Text)
                End If

                Dim Column_Sta1 As String = ""
                Column_Sta1 = TextBox_m1.Text.ToUpper
                Dim Column_sta2 As String = ""
                Column_sta2 = TextBox_m2.Text.ToUpper

                Dim Column_Sta11 As String = ""
                Column_Sta11 = TextBox_eq_m1.Text.ToUpper
                Dim Column_sta12 As String = ""
                Column_sta12 = TextBox_eq_m2.Text.ToUpper





                For i = Start1 To End1
                    Dim Station1 As String = W1.Range(Column_Sta1 & i).Value2
                    Dim Station2 As String = W1.Range(Column_sta2 & i).Value2
                    If IsNumeric(Station1) = True Then
                        W1.Range(Column_Sta11 & i).Value2 = Round(Station1 + Get_equation_value(Station1), Round_sta)
                    Else
                        MsgBox("non numerical values on " & Column_Sta1 & i)
                        W1.Range(Column_Sta1 & i).Select()
                        Freeze_operations = False
                        Exit Sub
                    End If

                    If IsNumeric(Station2) = True Then
                        W1.Range(Column_sta12 & i).Value2 = Round(Station2 + Get_equation_value(Station2), Round_sta)
                    Else
                        MsgBox("non numerical values on " & Column_sta2 & i)
                        W1.Range(Column_sta2 & i).Select()
                        Freeze_operations = False
                        Exit Sub
                    End If

                Next





            Catch ex As Exception
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try
        End If
        Freeze_operations = False
    End Sub






    Private Sub Button_water_bands_Click(sender As Object, e As EventArgs) Handles Button_water_bands.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Dim Round1 As Integer = 0
            If IsNumeric(TextBox_dec1.Text) = True Then
                Round1 = CInt(TextBox_dec1.Text)
            End If


            Try
                If IsNothing(Data_table_water_band_Excel_data) = False And IsNothing(Data_table_Matchlines) = False Then
                    If Data_table_water_band_Excel_data.Rows.Count > 0 Then
                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor

                        Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)


                                Dim Text_style_table As Autodesk.AutoCAD.DatabaseServices.TextStyleTable = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                Dim Layer_table As Autodesk.AutoCAD.DatabaseServices.LayerTable = Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                                Dim Start_point_for_bands As New Point3d(0, 0, 0)
                                If IsNumeric(TextBox_X_MS.Text) = True And IsNumeric(TextBox_Y_MS.Text) = True Then
                                    Start_point_for_bands = New Point3d(CDbl(TextBox_X_MS.Text), CDbl(TextBox_Y_MS.Text), 0)
                                End If


                                Dim Bands_Y_spacing As Double = 350




                                Dim Line_len As Double = Abs(CDbl(TextBox_viewport_Height.Text) * CDbl(TextBox_viewport_SCALE.Text))

                                Bands_Y_spacing = Ceiling(2.25 * Line_len / 50) * 50
                                TextBox_BAND_SPACING.Text = Bands_Y_spacing




                                Dim Min_dist As Double = 200
                                If IsNumeric(TextBox_min_dist_water.Text) = True Then
                                    Min_dist = CDbl(TextBox_min_dist_water.Text)
                                End If



                                Dim Text_offset_Y As Double = 1
                                If IsNumeric(TextBox_water_deltaY_station.Text) = True Then
                                    Text_offset_Y = CDbl(TextBox_water_deltaY_station.Text)
                                End If

                                Dim TextHeight As Double = 8
                                If IsNumeric(TextBox_text_height_water.Text) = True Then
                                    TextHeight = CDbl(TextBox_text_height_water.Text)
                                End If

                                Dim Xprev As Double = Start_point_for_bands.X
                                Dim Start_ptX As Double

                                Creaza_layer(No_plot, 40, No_plot, False)

                                Creaza_layer(ComboBox_layer_water.Text, 7, "", True)

                                Dim Nr_previous As Integer = -1
                                Dim Nr_rand As Integer = -1

                                For i = 0 To Data_table_Matchlines.Rows.Count - 1
                                    If IsDBNull(Data_table_Matchlines.Rows(i).Item("BEGSTA")) = False And IsDBNull(Data_table_Matchlines.Rows(i).Item("ENDSTA")) = False Then
                                        Dim Station1 As Double = Data_table_Matchlines.Rows(i).Item("BEGSTA")
                                        Dim Station2 As Double = Data_table_Matchlines.Rows(i).Item("ENDSTA")
                                        Nr_rand = Nr_rand + 1
                                        Dim Point_PS1 As New Point3d(0, 0, 0)
                                        Dim Start1 As New Point3d
                                        Dim End1 As New Point3d
                                        Start1 = New Point3d(Data_table_Matchlines.Rows(i).Item("X1"), Data_table_Matchlines.Rows(i).Item("Y1"), 0)
                                        End1 = New Point3d(Data_table_Matchlines.Rows(i).Item("X2"), Data_table_Matchlines.Rows(i).Item("Y2"), 0)
                                        Dim Viewport_line As New Line(Start1, End1)
                                        Start_ptX = Start_point_for_bands.X + CDbl(TextBox_viewport_Width.Text) / 2 - 50 - Viewport_line.Length / 2
                                        Point_PS1 = New Point3d(Start_ptX, Start_point_for_bands.Y - (Nr_rand * Bands_Y_spacing), 0)
                                        Dim Linie1 As New Line(Point_PS1, New Point3d(Point_PS1.X, Point_PS1.Y + Line_len, 0))
                                        If Not ComboBox_class_layer.Text = "" Then
                                            Linie1.Layer = ComboBox_class_layer.Text
                                        End If
                                        BTrecord.AppendEntity(Linie1)
                                        Trans1.AddNewlyCreatedDBObject(Linie1, True)
                                        Dim Mtext1 As New MText
                                        Mtext1.Location = New Point3d(Point_PS1.X - TextHeight / 4, Start_point_for_bands.Y - (Nr_rand * Bands_Y_spacing) + Text_offset_Y, 0)
                                        Mtext1.TextHeight = TextHeight
                                        Mtext1.Rotation = PI / 2
                                        Mtext1.Layer = ComboBox_layer_water.Text

                                        Mtext1.Contents = Get_chainage_feet_from_double(Station1 + Get_equation_value(Station1), Round1)
                                        Mtext1.Attachment = AttachmentPoint.BottomLeft
                                        BTrecord.AppendEntity(Mtext1)
                                        Trans1.AddNewlyCreatedDBObject(Mtext1, True)
                                        Dim Band_text As New DBText
                                        Band_text.Layer = No_plot
                                        If RadioButton_Left_right.Checked = True Then
                                            Band_text.Justify = AttachmentPoint.MiddleRight
                                            Band_text.AlignmentPoint = New Point3d(Start_ptX - 75, Start_point_for_bands.Y + Line_len / 2 - (Nr_rand * Bands_Y_spacing), 0)
                                        Else
                                            Band_text.Justify = AttachmentPoint.MiddleLeft
                                            Band_text.AlignmentPoint = New Point3d(Start_ptX + 75, Start_point_for_bands.Y + Line_len / 2 - (Nr_rand * Bands_Y_spacing), 0)
                                        End If

                                        Band_text.TextString = CStr(Nr_rand + 1)
                                        Band_text.Height = 7.5 * TextHeight
                                        BTrecord.AppendEntity(Band_text)
                                        Trans1.AddNewlyCreatedDBObject(Band_text, True)
                                    End If
                                Next




                                For i = 0 To Data_table_water_band_Excel_data.Rows.Count - 1
                                    If IsDBNull(Data_table_water_band_Excel_data.Rows(i).Item("BEGSTA")) = False And IsDBNull(Data_table_water_band_Excel_data.Rows(i).Item("ENDSTA")) = False Then
                                        Dim Station1 As Double = Data_table_water_band_Excel_data.Rows(i).Item("BEGSTA")
                                        Dim Station2 As Double = Data_table_water_band_Excel_data.Rows(i).Item("ENDSTA")
                                        Dim Name_string As String = "NO DATA"
                                        Dim Designation_string As String = "NO DATA"
                                        If IsDBNull(Data_table_water_band_Excel_data.Rows(i).Item("NAME")) = False Then
                                            Name_string = Data_table_water_band_Excel_data.Rows(i).Item("NAME")
                                        End If


                                        If IsDBNull(Data_table_water_band_Excel_data.Rows(i).Item("DESIGNATION")) = False Then
                                            Designation_string = Data_table_water_band_Excel_data.Rows(i).Item("DESIGNATION")
                                        End If



                                        Dim Point_PS As New Point3d(0, 0, 0)

                                        If IsNothing(Data_table_Matchlines) = False Then
                                            If Data_table_Matchlines.Rows.Count > 0 Then

                                                Dim Start1 As New Point3d
                                                Dim End1 As New Point3d
                                                Dim Match1 As Double = 0
                                                Dim Match2 As Double = 0

                                                Nr_rand = -1

                                                For j = 0 To Data_table_Matchlines.Rows.Count - 1
                                                    If IsDBNull(Data_table_Matchlines.Rows(j).Item("BEGSTA")) = False And IsDBNull(Data_table_Matchlines.Rows(j).Item("ENDSTA")) = False Then
                                                        Dim M1 As Double = 0
                                                        Dim M2 As Double = 0
                                                        M1 = Data_table_Matchlines.Rows(j).Item("BEGSTA")
                                                        M2 = Data_table_Matchlines.Rows(j).Item("ENDSTA")
                                                        Nr_rand = Nr_rand + 1

                                                        If Station2 <= M2 And Station2 >= M1 Then



                                                            If IsDBNull(Data_table_Matchlines.Rows(j).Item("X1")) = False And IsDBNull(Data_table_Matchlines.Rows(j).Item("Y1")) = False Then
                                                                If IsDBNull(Data_table_Matchlines.Rows(j).Item("X2")) = False And IsDBNull(Data_table_Matchlines.Rows(j).Item("Y2")) = False Then
                                                                    Start1 = New Point3d(Data_table_Matchlines.Rows(j).Item("X1"), Data_table_Matchlines.Rows(j).Item("Y1"), 0)
                                                                    End1 = New Point3d(Data_table_Matchlines.Rows(j).Item("X2"), Data_table_Matchlines.Rows(j).Item("Y2"), 0)
                                                                    If PolyCL.GetPointAtDist(M1).GetVectorTo(End1).Length < PolyCL.GetPointAtDist(M1).GetVectorTo(Start1).Length Then
                                                                        Dim ttt As New Point3d
                                                                        ttt = Start1
                                                                        Start1 = End1
                                                                        End1 = ttt
                                                                    End If
                                                                    Match2 = M2
                                                                    Match1 = M1



                                                                    Exit For
                                                                End If
                                                            End If

                                                        Else


                                                        End If
                                                    End If
                                                Next






                                                Dim Point_MS As New Point3d
                                                If Station2 > PolyCL.Length Then
                                                    Station2 = PolyCL.Length
                                                End If
                                                Point_MS = PolyCL.GetPointAtDist(Station2)

                                                Dim Viewport_line As New Line(Start1, End1)
                                                Dim Dist_from_match As Double = -1.1111
                                                Dim Point_on_line As New Point3d
                                                Point_on_line = Viewport_line.GetClosestPointTo(Point_MS, Vector3d.ZAxis, False)
                                                If IsNothing(Point_on_line) = False Then
                                                    Dist_from_match = Viewport_line.StartPoint.GetVectorTo(Point_on_line).Length
                                                End If
                                                Dim Val1 As Double = Station2 - Match1
                                                If Not Dist_from_match = -1.1111 Then
                                                    Val1 = Dist_from_match
                                                End If

                                                Start_ptX = Start_point_for_bands.X + CDbl(TextBox_viewport_Width.Text) / 2 - 50 - Viewport_line.Length / 2

                                                If Nr_previous <> Nr_rand Then
                                                    Xprev = Start_ptX
                                                End If

                                                If RadioButton_Left_right.Checked = True Then
                                                    Point_PS = New Point3d(Start_ptX + Val1, Start_point_for_bands.Y - (Nr_rand * Bands_Y_spacing), 0)
                                                    If Point_PS.X - Xprev < Min_dist Then
                                                        Point_PS = New Point3d(Xprev + Min_dist, Start_point_for_bands.Y - (Nr_rand * Bands_Y_spacing), 0)
                                                    End If
                                                Else
                                                    Point_PS = New Point3d(Start_ptX - Val1, Start_point_for_bands.Y - (Nr_rand * Bands_Y_spacing), 0)

                                                    If Xprev - Point_PS.X < Min_dist Then
                                                        Point_PS = New Point3d(Xprev - Min_dist, Start_point_for_bands.Y - (Nr_rand * Bands_Y_spacing), 0)
                                                    End If
                                                End If

                                            End If
                                        End If



                                        Dim Linie1 As New Line(New Point3d(Point_PS.X, Start_point_for_bands.Y - (Nr_rand * Bands_Y_spacing), 0), New Point3d(Point_PS.X, Start_point_for_bands.Y - (Nr_rand * Bands_Y_spacing) + Line_len, 0))
                                        If Not ComboBox_layer_water.Text = "" Then
                                            Linie1.Layer = ComboBox_layer_water.Text
                                        End If
                                        BTrecord.AppendEntity(Linie1)
                                        Trans1.AddNewlyCreatedDBObject(Linie1, True)


                                        Dim Mtext_STATION As New MText
                                        Mtext_STATION.Location = New Point3d(Point_PS.X - TextHeight / 4, Start_point_for_bands.Y - (Nr_rand * Bands_Y_spacing) + Text_offset_Y, 0)
                                        Mtext_STATION.TextHeight = TextHeight
                                        Mtext_STATION.Rotation = PI / 2
                                        Mtext_STATION.Layer = Linie1.Layer
                                        Mtext_STATION.Contents = Get_chainage_feet_from_double(Station2 + Get_equation_value(Station2), Round1)
                                        Mtext_STATION.Attachment = AttachmentPoint.BottomLeft

                                        BTrecord.AppendEntity(Mtext_STATION)
                                        Trans1.AddNewlyCreatedDBObject(Mtext_STATION, True)


                                        Dim Y_middle As Double = Point_PS.Y + Line_len / 2
                                        Dim Y_middle1 As Double = Point_PS.Y + Line_len / 4
                                        Dim Y_middle2 As Double = Point_PS.Y + 3 * Line_len / 4

                                        Dim PT_ins1 As New Point3d
                                        PT_ins1 = New Point3d(Point_PS.X - (Point_PS.X - Xprev) / 2, Y_middle1, 0)

                                        Dim PT_ins2 As New Point3d
                                        PT_ins2 = New Point3d(Point_PS.X - (Point_PS.X - Xprev) / 2, Y_middle2, 0)

                                        Dim Mtext_name As New MText
                                        Mtext_name.Contents = Name_string
                                        Mtext_name.TextHeight = TextHeight
                                        Mtext_name.Location = PT_ins2
                                        Mtext_name.Rotation = 0
                                        If Layer_table.Has(ComboBox_layer_water.Text) = True Then
                                            Mtext_name.Layer = ComboBox_layer_water.Text
                                        End If
                                        Mtext_name.Attachment = AttachmentPoint.MiddleCenter



                                        Dim Mtext_designation As New MText
                                        Mtext_designation.Contents = Designation_string
                                        Mtext_designation.TextHeight = TextHeight
                                        Mtext_designation.Location = PT_ins1
                                        Mtext_designation.Rotation = 0
                                        If Layer_table.Has(ComboBox_layer_water.Text) = True Then
                                            Mtext_designation.Layer = ComboBox_layer_water.Text
                                        End If
                                        Mtext_designation.Attachment = AttachmentPoint.MiddleCenter


                                        Dim ObjId1 As TextStyleTableRecord
                                        If Text_style_table.Has(ComboBox_text_style_water.Text) = True Then
                                            ObjId1 = Text_style_table(ComboBox_text_style_water.Text).GetObject(OpenMode.ForRead)
                                            Mtext_name.TextStyleId = ObjId1.ObjectId
                                            Mtext_designation.TextStyleId = ObjId1.ObjectId
                                        End If

                                        BTrecord.AppendEntity(Mtext_name)
                                        Trans1.AddNewlyCreatedDBObject(Mtext_name, True)

                                        BTrecord.AppendEntity(Mtext_designation)
                                        Trans1.AddNewlyCreatedDBObject(Mtext_designation, True)

                                        Xprev = Point_PS.X
                                        Nr_previous = Nr_rand

                                    End If







                                Next



                                Trans1.Commit()
                            End Using
                        End Using


                        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                    End If
                Else

                    MsgBox("You did not have data loaded for matchlines", MsgBoxStyle.Critical, "Dan says...")
                End If


            Catch ex As Exception
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_load_excel_water_data_Click(sender As Object, e As EventArgs) Handles Button_load_excel_water_data.Click

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
                If IsNumeric(TextBox_ROW_start_water.Text) = True Then
                    Start1 = CInt(TextBox_ROW_start_water.Text)
                End If
                If IsNumeric(TextBox_ROW_end_water.Text) = True Then
                    End1 = CInt(TextBox_ROW_end_water.Text)
                End If

                If End1 = 0 Or Start1 = 0 Then
                    Freeze_operations = False
                    Exit Sub
                End If
                If End1 < Start1 Then
                    Freeze_operations = False
                    Exit Sub
                End If

                If IsNothing(Data_table_Matchlines) = True Then
                    MsgBox("No matchlines loaded")
                    Freeze_operations = False
                    Exit Sub
                End If

                If Data_table_Matchlines.Rows.Count < 1 Then
                    MsgBox("No matchlines loaded")
                    Freeze_operations = False
                    Exit Sub
                End If

                Dim Column_sta1 As String = ""
                Column_sta1 = TextBox_column_start_water.Text.ToUpper
                Dim Column_sta2 As String = ""
                Column_sta2 = TextBox_column_end_water.Text.ToUpper
                Dim Column_name As String = ""
                Column_name = TextBox_column_name_water.Text.ToUpper
                Dim Column_designation As String = ""
                Column_designation = TextBox_column_designation_water.Text.ToUpper

                Data_table_water_band_Excel_data = New System.Data.DataTable
                Data_table_water_band_Excel_data.Columns.Add("BEGSTA", GetType(Double))
                Data_table_water_band_Excel_data.Columns.Add("ENDSTA", GetType(Double))
                Data_table_water_band_Excel_data.Columns.Add("NAME", GetType(String))
                Data_table_water_band_Excel_data.Columns.Add("DESIGNATION", GetType(String))


                Dim Index_data_table As Double



                For i = Start1 To End1
                    Dim Station_string1 As String = W1.Range(Column_sta1 & i).Value2
                    Dim Station_string2 As String = W1.Range(Column_sta2 & i).Value2
                    Dim Name1 As String = W1.Range(Column_name & i).Value2
                    Dim Design1 As String = W1.Range(Column_designation & i).Value2
                    If IsNumeric(Station_string1) = True And IsNumeric(Station_string2) = True And Not Name1 = "" And Not Design1 = "" Then
                        If CDbl(Station_string1) >= 0 And CDbl(Station_string2) >= 0 Then
                            Dim Station1 As Double = CDbl(Station_string1)
                            Dim Station2 As Double = CDbl(Station_string2)
                            Dim Adaugat As Boolean = False

                            For j = 0 To Data_table_Matchlines.Rows.Count - 1
                                If IsDBNull(Data_table_Matchlines.Rows(j).Item("BEGSTA")) = False And IsDBNull(Data_table_Matchlines.Rows(j).Item("ENDSTA")) = False Then
                                    Dim M1 As Double = Data_table_Matchlines.Rows(j).Item("BEGSTA")
                                    Dim M2 As Double = Data_table_Matchlines.Rows(j).Item("ENDSTA")
                                    If Station1 >= M1 And Station2 <= M2 Then
                                        Data_table_water_band_Excel_data.Rows.Add()
                                        Data_table_water_band_Excel_data.Rows(Index_data_table).Item("BEGSTA") = Station1
                                        Data_table_water_band_Excel_data.Rows(Index_data_table).Item("ENDSTA") = Station2
                                        Data_table_water_band_Excel_data.Rows(Index_data_table).Item("NAME") = Name1
                                        Data_table_water_band_Excel_data.Rows(Index_data_table).Item("DESIGNATION") = Design1
                                        Index_data_table = Index_data_table + 1
                                        Adaugat = True
                                        Exit For
                                    End If
                                End If
                            Next

                            If Adaugat = False Then
                                For j = 0 To Data_table_Matchlines.Rows.Count - 1
                                    If IsDBNull(Data_table_Matchlines.Rows(j).Item("BEGSTA")) = False And IsDBNull(Data_table_Matchlines.Rows(j).Item("ENDSTA")) = False Then
                                        Adaugat = False
                                        Dim M1 As Double = Data_table_Matchlines.Rows(j).Item("BEGSTA")
                                        Dim M2 As Double = Data_table_Matchlines.Rows(j).Item("ENDSTA")
                                        If M1 >= Station1 And M2 >= Station1 And M1 <= Station2 And M2 <= Station2 Then
                                            Data_table_water_band_Excel_data.Rows.Add()
                                            Data_table_water_band_Excel_data.Rows(Index_data_table).Item("BEGSTA") = M1
                                            Data_table_water_band_Excel_data.Rows(Index_data_table).Item("ENDSTA") = M2
                                            Data_table_water_band_Excel_data.Rows(Index_data_table).Item("NAME") = Name1
                                            Data_table_water_band_Excel_data.Rows(Index_data_table).Item("DESIGNATION") = Design1

                                            Index_data_table = Index_data_table + 1
                                            Adaugat = True
                                        End If

                                        If Adaugat = False Then
                                            If Station1 <= M1 And M1 <= Station2 And M2 >= Station2 And M2 > Station1 Then
                                                Data_table_water_band_Excel_data.Rows.Add()
                                                Data_table_water_band_Excel_data.Rows(Index_data_table).Item("BEGSTA") = M1
                                                Data_table_water_band_Excel_data.Rows(Index_data_table).Item("ENDSTA") = Station2
                                                Data_table_water_band_Excel_data.Rows(Index_data_table).Item("NAME") = Name1
                                                Data_table_water_band_Excel_data.Rows(Index_data_table).Item("DESIGNATION") = Design1
                                                Index_data_table = Index_data_table + 1

                                            End If
                                        End If

                                        If Adaugat = False Then
                                            If Station1 >= M1 And M2 >= Station1 And M1 <= Station2 And M2 <= Station2 Then
                                                Data_table_water_band_Excel_data.Rows.Add()
                                                Data_table_water_band_Excel_data.Rows(Index_data_table).Item("BEGSTA") = Station1
                                                Data_table_water_band_Excel_data.Rows(Index_data_table).Item("ENDSTA") = M2
                                                Data_table_water_band_Excel_data.Rows(Index_data_table).Item("NAME") = Name1
                                                Data_table_water_band_Excel_data.Rows(Index_data_table).Item("DESIGNATION") = Design1
                                                Index_data_table = Index_data_table + 1

                                            End If
                                        End If

                                    End If
                                Next
                            End If



                        End If
                    End If
                Next




                Data_table_water_band_Excel_data = Sort_data_table(Data_table_water_band_Excel_data, "BEGSTA")


                Add_to_clipboard_Data_table(Data_table_water_band_Excel_data)


            Catch ex As Exception
                MsgBox(ex.Message)

            End Try
            Freeze_operations = False
        End If
    End Sub


    Private Sub Button_load_DWG_Click(sender As Object, e As EventArgs) Handles Button_load_DWG.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                Dim FileBrowserDialog1 As New Windows.Forms.OpenFileDialog
                FileBrowserDialog1.Filter = "Drawing Files (*.dwg)|*.dwg|All Files (*.*)|*.*"
                FileBrowserDialog1.Multiselect = True

                If FileBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    For Each file1 In FileBrowserDialog1.FileNames
                        ListBox_DWG.Items.Add(file1)
                    Next

                End If

            Catch ex As System.Exception
                Freeze_operations = False
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_DWG_Insert_viewport_Click(sender As Object, e As EventArgs) Handles Button_DWG_Insert_viewport.Click



        If Freeze_operations = False Then
            Freeze_operations = True


            Try

                If IsNumeric(TextBox_viewport_SCALE.Text) = False Then
                    MsgBox("Please specify the viewport scale!")
                    Freeze_operations = False
                    Exit Sub
                End If
                If IsNumeric(TextBox_viewport_Height.Text) = False Then
                    MsgBox("Please specify the viewport height!")
                    Freeze_operations = False
                    Exit Sub
                End If
                If IsNumeric(TextBox_viewport_Width.Text) = False Then
                    MsgBox("Please specify the viewport width!")
                    Freeze_operations = False
                    Exit Sub
                End If
                If IsNumeric(TextBox_BAND_SPACING.Text) = False Then
                    MsgBox("Please specify the distance between bands!")
                    Freeze_operations = False
                    Exit Sub
                End If
                If IsNumeric(TextBox_X_MS.Text) = False Then
                    MsgBox("Please specify the X of the 0+00 station!")
                    Freeze_operations = False
                    Exit Sub
                End If
                If IsNumeric(TextBox_Y_MS.Text) = False Then
                    MsgBox("Please specify the Y of the 0+00 station!")
                    Freeze_operations = False
                    Exit Sub
                End If

                If IsNumeric(TextBox_X_PS.Text) = False Then
                    MsgBox("Please specify the X of the viewport corner!")
                    Freeze_operations = False
                    Exit Sub
                End If
                If IsNumeric(TextBox_Y_PS.Text) = False Then
                    MsgBox("Please specify the Y of the viewport corner!")
                    Freeze_operations = False
                    Exit Sub
                End If

                If Not ListBox_DWG.Items.Count = ListBox_sheet_numbers.Items.Count Then
                    MsgBox("Please be sure numbers of DWG is equal to the number of bands")
                    Freeze_operations = False
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


                If ListBox_DWG.Items.Count > 0 Then
                    Using lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument

                        Using Trans11 As Transaction = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.TransactionManager.StartTransaction
                            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()




                            For i = 0 To ListBox_DWG.Items.Count - 1
                                Dim Band_index As Integer = CInt(ListBox_sheet_numbers.Items(i))

                                Dim Drawing1 As String = ListBox_DWG.Items(i)
                                If IO.File.Exists(Drawing1) = True Then
                                    Dim Database1 As New Database(False, True)
                                    Database1.ReadDwgFile(Drawing1, IO.FileShare.ReadWrite, True, Nothing)
                                    HostApplicationServices.WorkingDatabase = Database1

                                    Creaza_layer_with_database(Database1, "VP", 4, "VIEWPORT", False)

                                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction
                                        Dim BlockTable1 As BlockTable = Trans1.GetObject(Database1.BlockTableId, OpenMode.ForRead)



                                        Dim BTrecordPS As BlockTableRecord = Trans1.GetObject(BlockTable1(BlockTableRecord.PaperSpace), OpenMode.ForWrite)




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







                                        Trans1.Commit()


                                    End Using


                                    Database1.SaveAs(Drawing1, True, DwgVersion.Current, Database1.SecurityParameters)
                                    Database1.Dispose()
                                    HostApplicationServices.WorkingDatabase = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database
                                End If


                            Next

                            Trans11.Commit()
                        End Using
                    End Using
                End If





            Catch ex As Exception
                MsgBox(ex.Message)

            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: DONE!")
            Freeze_operations = False

        End If
    End Sub

    Private Sub Button_remove_items_list_Click(sender As Object, e As EventArgs) Handles Button_remove_items_list.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            If ListBox_DWG.Items.Count > 0 Then
                If ListBox_DWG.SelectedIndex >= 0 Then
                    ListBox_DWG.Items.RemoveAt((ListBox_DWG.SelectedIndex))
                    If ListBox_sheet_numbers.Items.Count > 0 Then
                        If ListBox_sheet_numbers.Items.Count >= ListBox_DWG.SelectedIndex Then
                            ListBox_sheet_numbers.Items.RemoveAt((ListBox_DWG.SelectedIndex))
                        End If
                    End If
                End If
            End If
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_clear_lists_Click(sender As Object, e As EventArgs) Handles Button_clear_lists.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            If ListBox_DWG.Items.Count > 0 Then
                ListBox_DWG.Items.Clear()
            End If
            If ListBox_sheet_numbers.Items.Count > 0 Then
                ListBox_sheet_numbers.Items.Clear()
            End If
            Freeze_operations = False
        End If
    End Sub
    Private Sub Button_browse_for_block_Click(sender As Object, e As EventArgs) Handles Button_browse_for_block.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                Dim FileBrowserDialog1 As New Windows.Forms.OpenFileDialog
                FileBrowserDialog1.Filter = "Drawing Files (*.dwg)|*.dwg|All Files (*.*)|*.*"


                If FileBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    TextBox_block_name.Text = FileBrowserDialog1.FileName
                End If

            Catch ex As System.Exception
                Freeze_operations = False
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub



    Private Sub Button_insert_block_Click(sender As Object, e As EventArgs) Handles Button_insert_block.Click



        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Error_list As String = ""

            Try

                Try

                    If IsNumeric(TextBox_block_scale.Text) = False Then
                        MsgBox("Please specify the BLOCK scale!")
                        Freeze_operations = False
                        Exit Sub
                    End If
                    If IsNumeric(TextBox_X_block.Text) = False Then
                        MsgBox("Please specify the x coordinate!")
                        Freeze_operations = False
                        Exit Sub
                    End If
                    If IsNumeric(TextBox_Y_block.Text) = False Then
                        MsgBox("Please specify the y coordinate!")
                        Freeze_operations = False
                        Exit Sub
                    End If
                    If IsNumeric(TextBox_Layout_index_START.Text) = False Then
                        MsgBox("Please specify the layout for block insertion!")
                        Freeze_operations = False
                        Exit Sub
                    End If
                    If IsNumeric(TextBox_Layout_index_END.Text) = False Then
                        MsgBox("Please specify the layout for block insertion!")
                        Freeze_operations = False
                        Exit Sub
                    End If


                    Dim Scale1 As Double = CDbl(TextBox_block_scale.Text)

                    Dim Point_ins As New Point3d(CDbl(TextBox_X_block.Text), CDbl(TextBox_Y_block.Text), 0)



                    If Scale1 <= 0 Then
                        MsgBox("Negative values not allowed")
                        Freeze_operations = False
                        Exit Sub
                    End If


                    Dim Column_start As String = TextBox_block_att_value_column_start.Text.ToUpper
                    Dim Column_end As String = TextBox_block_att_value_column_end.Text.ToUpper
                    Dim Column_atr_name As String = TextBox_block_att_name_column.Text.ToUpper

                    Dim Row_with_file_names As Integer = 0
                    If IsNumeric(TextBox_ROW_FILE_NAME.Text) = True Then
                        Row_with_file_names = CInt(TextBox_ROW_FILE_NAME.Text)
                    End If


                    Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                    W1 = Nothing
                    Dim Col_start As Integer = 0
                    Dim Col_end As Integer = 0
                    Dim Col_atr As Integer = 0

                    If Panel_blocks_insertion.Visible = True Then
                        Try
                            W1 = Get_active_worksheet_from_Excel_with_error()
                            If IsNothing(W1) = False Then
                                Col_start = W1.Range(Column_start & "1").Column
                                Col_end = W1.Range(Column_end & "1").Column
                                Col_atr = W1.Range(Column_atr_name & "1").Column
                            End If

                        Catch ex As System.SystemException
                            W1 = Nothing
                            Col_start = 0
                            Col_end = 0
                            Col_atr = 0
                        End Try
                    End If

                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                    If ListBox_DWG.Items.Count > 0 Then
                        Using lock1 As DocumentLock = ThisDrawing.LockDocument

                            Using Trans11 As Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()



                                Dim BlockDwg As String = TextBox_block_name.Text
                                If IO.File.Exists(BlockDwg) = False Then
                                    MsgBox("no dwg file")
                                    Freeze_operations = False
                                    Exit Sub
                                End If

                                Dim Name_of_the_block As String = IO.Path.GetFileNameWithoutExtension(BlockDwg)
                                Dim Layer1 As String = TextBox_block_layer_name.Text

                                If Layer1 = "" Then Layer1 = "0"

                                For i = 0 To ListBox_DWG.Items.Count - 1
                                    Dim Drawing1 As String = ListBox_DWG.Items(i)
                                    If IO.File.Exists(Drawing1) = True Then

                                        Dim Colectie_nume_atribute As New Specialized.StringCollection
                                        Dim Colectie_valori_atribute As New Specialized.StringCollection

                                        If IsNothing(W1) = False Then
                                            Dim Start1 As Integer = 0
                                            Dim End1 As Integer = 0
                                            If IsNumeric(TextBox_blocks_att_row_start.Text) = True Then
                                                Start1 = CInt(TextBox_blocks_att_row_start.Text)
                                            End If
                                            If IsNumeric(TextBox_blocks_att_row_end.Text) = True Then
                                                End1 = CInt(TextBox_blocks_att_row_end.Text)
                                            End If

                                            If Not Row_with_file_names = 0 And Not Col_start = 0 And Not Col_end = 0 And Not Col_end < Col_start And Not Col_atr = 0 And Not Start1 = 0 And Not End1 = 0 And Not End1 < Start1 Then
                                                For j = Col_start To Col_end
                                                    Dim Excel_file As String = W1.Cells(Row_with_file_names, j).Value2
                                                    If Drawing1.ToUpper.Contains(Excel_file.ToUpper) = True Then

                                                        For k = Start1 To End1

                                                            Dim Atr_name As String = ""
                                                            Atr_name = W1.Cells(k, Col_atr).Value2
                                                            Dim Atr_value As String = ""
                                                            Atr_value = W1.Cells(k, j).Value2

                                                            If Not Atr_name = "" Then
                                                                Colectie_nume_atribute.Add(Atr_name)
                                                                Colectie_valori_atribute.Add(Atr_value)
                                                            End If

                                                        Next

                                                        Exit For
                                                    End If
                                                Next
                                            End If
                                        End If


                                        Dim Database1 As New Database(False, True)

                                        Try
                                            Try
                                                Try
                                                    Database1.ReadDwgFile(Drawing1, IO.FileShare.ReadWrite, True, Nothing)
                                                Catch ex As Exception
                                                    MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                    GoTo 123
                                                End Try
                                            Catch ex As IO.IOException
                                                MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                Error_list = Error_list & vbCrLf & Drawing1
                                                GoTo 123
                                            End Try
                                        Catch ex As System.SystemException
                                            MsgBox(Drawing1 & vbCrLf & "could not be open")
                                            Error_list = Error_list & vbCrLf & Drawing1
                                            GoTo 123
                                        End Try


                                        HostApplicationServices.WorkingDatabase = Database1





                                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction
                                            Dim Nume_block_de_sters As String = TextBox_ERASE_BLOCK.Text
                                            Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                                            Dim Layoutdict As DBDictionary

                                            Layoutdict = Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead)
                                            Dim Index_datatable As Integer = 0

                                            If Not TextBox_ERASE_BLOCK.Text = "" Then
                                                Dim BlockTable1 As BlockTable = Trans1.GetObject(Database1.BlockTableId, OpenMode.ForRead)
                                                If BlockTable1.Has(Nume_block_de_sters) = True Then
                                                    For Each entry As DBDictionaryEntry In Layoutdict
                                                        Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                                        If Not Layout1.TabOrder = 0 Then
                                                            Dim BTrecord As BlockTableRecord = Trans1.GetObject(Layout1.BlockTableRecordId, OpenMode.ForWrite)
                                                            For Each Id1 As ObjectId In BTrecord
                                                                Dim Ent1 As Entity = TryCast(Trans1.GetObject(Id1, OpenMode.ForRead), Entity)
                                                                If IsNothing(Ent1) = False Then
                                                                    If TypeOf (Ent1) Is BlockReference Then

                                                                        Dim Block1 As BlockReference = TryCast(Ent1, BlockReference)

                                                                        Dim BlockTrec As BlockTableRecord = Nothing

                                                                        If Block1.IsDynamicBlock = True Then
                                                                            BlockTrec = Trans1.GetObject(Block1.DynamicBlockTableRecord, OpenMode.ForRead)
                                                                        Else
                                                                            BlockTrec = Trans1.GetObject(Block1.BlockTableRecord, OpenMode.ForRead)
                                                                        End If

                                                                        If BlockTrec.Name = Nume_block_de_sters Then
                                                                            Block1.UpgradeOpen()
                                                                            Block1.Erase()
                                                                        End If
                                                                    End If
                                                                End If
                                                            Next
                                                        End If

                                                    Next
                                                    BlockTable1.UpgradeOpen()
                                                    Dim Btr1 As BlockTableRecord = Trans1.GetObject(BlockTable1(Nume_block_de_sters), OpenMode.ForWrite)
                                                    Btr1.Erase()
                                                End If
                                            End If

                                            For Each entry As DBDictionaryEntry In Layoutdict
                                                Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                                If Layout1.TabOrder >= CInt(TextBox_Layout_index_START.Text) And Layout1.TabOrder <= CInt(TextBox_Layout_index_END.Text) Then
                                                    If Not Layer1 = "0" Then
                                                        Creaza_layer_with_database(Database1, Layer1, 7, "", True)
                                                    End If
                                                    Dim BlockTable1 As BlockTable = Trans1.GetObject(Database1.BlockTableId, OpenMode.ForRead)
                                                    Dim BTrecord As BlockTableRecord = Trans1.GetObject(Layout1.BlockTableRecordId, OpenMode.ForWrite)
                                                    'BTrecord = Trans1.GetObject(BlockTable1(BlockTableRecord.PaperSpace), OpenMode.ForWrite)
                                                    InsertBlock_with_multiple_atributes_background(BlockDwg, Name_of_the_block, Database1, Point_ins, Scale1, BTrecord, Layer1, Colectie_nume_atribute, Colectie_valori_atribute)
                                                End If
                                            Next

                                            Trans1.Commit()

                                            Try
                                                Database1.SaveAs(Drawing1, True, DwgVersion.Current, Database1.SecurityParameters)
                                            Catch ex As Exception
                                                Error_list = Error_list & vbCrLf & Drawing1
                                            End Try
                                        End Using

                                        Database1.Dispose()

                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database
                                    End If

123:
                                Next

                                Trans11.Commit()
                            End Using
                        End Using
                    End If


                    If Not Error_list = "" Then
                        MsgBox("There are not saved files. The list of these files is loaded into the clipboard" & vbCrLf &
                               "Use paste <Ctrl+V> command in a text editor program (notepad, excel, word) to see the list")
                        My.Computer.Clipboard.SetText(Error_list)

                    End If


                Catch ex As Exception
                    MsgBox(ex.Message)

                End Try

            Catch ex As System.SystemException
                MsgBox("1: " & ex.Message)
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: DONE!")
            Freeze_operations = False

        End If
    End Sub




    Private Sub Button_load_sheet_from_excel_Click(sender As Object, e As EventArgs) Handles Button_load_band_from_excel.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Dim Start1 As Integer = 0
                Dim End1 As Integer = 0
                If IsNumeric(TextBox_SHEET_ROW_START.Text) = True Then
                    Start1 = CInt(TextBox_SHEET_ROW_START.Text)
                End If
                If IsNumeric(TextBox_SHEET_ROW_END.Text) = True Then
                    End1 = CInt(TextBox_SHEET_ROW_END.Text)
                End If

                If End1 = 0 Or Start1 = 0 Then
                    Freeze_operations = False
                    Exit Sub
                End If

                If End1 < Start1 Then
                    Freeze_operations = False
                    Exit Sub
                End If

                Dim Column_sheet As String = ""
                Column_sheet = TextBox_column_sheet_number.Text.ToUpper

                Dim Column_dwg As String = ""
                Column_dwg = TextBox_column_drawing_number.Text.ToUpper

                ListBox_sheet_numbers.Items.Clear()

                If ListBox_DWG.Items.Count > 0 Then
                    For i = 0 To ListBox_DWG.Items.Count - 1
                        ListBox_sheet_numbers.Items.Add(0)
                    Next
                    If End1 - Start1 + 1 = ListBox_DWG.Items.Count Then
                        For i = Start1 To End1
                            Dim Sheet_string As String = W1.Range(Column_sheet & i).Value2
                            Dim DWG_no As String = W1.Range(Column_dwg & i).Value2
                            If IsNumeric(Sheet_string) = True Then
                                If CInt(Sheet_string) > 0 Then
                                    For j = 0 To ListBox_DWG.Items.Count - 1
                                        Dim DWG_no1 As String = ListBox_DWG.Items(j)
                                        Dim Nume1 As String = System.IO.Path.GetFileName(DWG_no1)
                                        If DWG_no = Nume1 Then
                                            ListBox_sheet_numbers.Items(j) = CInt(Sheet_string)
                                            Exit For
                                        End If
                                    Next
                                End If
                            End If
                        Next
                    End If
                End If


            Catch ex As System.Exception
                Freeze_operations = False
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub


    Private Sub CheckBox_read_from_excel_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_read_from_excel.CheckedChanged
        If CheckBox_read_from_excel.Checked = False Then
            Panel_blocks_insertion.Visible = False
        Else
            Panel_blocks_insertion.Visible = True
        End If
    End Sub

    Private Sub Button_redefine_block_Click(sender As Object, e As EventArgs) Handles Button_redefine_block.Click
        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Error_list As String = ""

            Try

                Try

                    If IsNumeric(TextBox_Layout_index_START.Text) = False Then
                        MsgBox("Please specify the layout for block insertion!")
                        Freeze_operations = False
                        Exit Sub
                    End If
                    If IsNumeric(TextBox_Layout_index_END.Text) = False Then
                        MsgBox("Please specify the layout for block insertion!")
                        Freeze_operations = False
                        Exit Sub
                    End If



                    Dim Column_start As String = TextBox_block_att_value_column_start.Text.ToUpper
                    Dim Column_end As String = TextBox_block_att_value_column_end.Text.ToUpper
                    Dim Column_atr_name As String = TextBox_block_att_name_column.Text.ToUpper

                    Dim Row_with_file_names As Integer = 0
                    If IsNumeric(TextBox_ROW_FILE_NAME.Text) = True Then
                        Row_with_file_names = CInt(TextBox_ROW_FILE_NAME.Text)
                    End If


                    Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                    W1 = Nothing
                    Dim Col_start As Integer = 0
                    Dim Col_end As Integer = 0
                    Dim Col_atr As Integer = 0

                    If Panel_blocks_insertion.Visible = True Then
                        Try
                            W1 = Get_active_worksheet_from_Excel_with_error()
                            If IsNothing(W1) = False Then
                                Col_start = W1.Range(Column_start & "1").Column
                                Col_end = W1.Range(Column_end & "1").Column
                                Col_atr = W1.Range(Column_atr_name & "1").Column
                            End If

                        Catch ex As System.SystemException
                            W1 = Nothing
                            Col_start = 0
                            Col_end = 0
                            Col_atr = 0
                        End Try
                    End If

                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                    If ListBox_DWG.Items.Count > 0 Then
                        Using lock1 As DocumentLock = ThisDrawing.LockDocument

                            Using Trans11 As Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()



                                Dim BlockDwg As String = TextBox_block_name.Text
                                If IO.File.Exists(BlockDwg) = False Then
                                    MsgBox("no dwg file")
                                    Freeze_operations = False
                                    Exit Sub
                                End If

                                Dim Name_of_the_block As String = IO.Path.GetFileNameWithoutExtension(BlockDwg)


                                For i = 0 To ListBox_DWG.Items.Count - 1
                                    Dim Drawing1 As String = ListBox_DWG.Items(i)
                                    If IO.File.Exists(Drawing1) = True Then

                                        Dim Colectie_nume_atribute As New Specialized.StringCollection
                                        Dim Colectie_valori_atribute As New Specialized.StringCollection



                                        Dim Database1 As New Database(False, True)

                                        Try
                                            Try
                                                Try
                                                    Database1.ReadDwgFile(Drawing1, IO.FileShare.ReadWrite, True, Nothing)
                                                Catch ex As Exception
                                                    MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                    GoTo 123
                                                End Try
                                            Catch ex As IO.IOException
                                                MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                Error_list = Error_list & vbCrLf & Drawing1
                                                GoTo 123
                                            End Try
                                        Catch ex As System.SystemException
                                            MsgBox(Drawing1 & vbCrLf & "could not be open")
                                            Error_list = Error_list & vbCrLf & Drawing1
                                            GoTo 123
                                        End Try


                                        HostApplicationServices.WorkingDatabase = Database1





                                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction

                                            Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                                            Dim Layoutdict As DBDictionary

                                            Layoutdict = Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead)
                                            Dim Index_datatable As Integer = 0

                                            Dim Data_table_valori_existente As New System.Data.DataTable
                                            Data_table_valori_existente.Columns.Add("X", GetType(Double))
                                            Data_table_valori_existente.Columns.Add("Y", GetType(Double))
                                            Data_table_valori_existente.Columns.Add("SCALE", GetType(Double))
                                            Data_table_valori_existente.Columns.Add("ROTATION", GetType(Double))
                                            Data_table_valori_existente.Columns.Add("LAYER", GetType(String))
                                            Data_table_valori_existente.Columns.Add("LAYOUT_INDEX", GetType(Integer))

                                            Dim Index1 As Integer = 0

                                            Dim BlockTable1 As BlockTable = Trans1.GetObject(Database1.BlockTableId, OpenMode.ForRead)
                                            If BlockTable1.Has(Name_of_the_block) = True Then
                                                For Each entry As DBDictionaryEntry In Layoutdict
                                                    Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                                    If Not Layout1.TabOrder = 0 Then
                                                        Dim BTrecord As BlockTableRecord = Trans1.GetObject(Layout1.BlockTableRecordId, OpenMode.ForWrite)
                                                        For Each Id1 As ObjectId In BTrecord
                                                            Dim Ent1 As Entity = TryCast(Trans1.GetObject(Id1, OpenMode.ForRead), Entity)
                                                            If IsNothing(Ent1) = False Then
                                                                If TypeOf (Ent1) Is BlockReference Then
                                                                    Dim Block1 As BlockReference = TryCast(Ent1, BlockReference)
                                                                    If IsNothing(Block1) = False Then
                                                                        Dim BlockTrec As BlockTableRecord = Nothing

                                                                        If Block1.IsDynamicBlock = True Then
                                                                            BlockTrec = Trans1.GetObject(Block1.DynamicBlockTableRecord, OpenMode.ForRead)
                                                                        Else
                                                                            BlockTrec = Trans1.GetObject(Block1.BlockTableRecord, OpenMode.ForRead)
                                                                        End If
                                                                        If BlockTrec.Name = Name_of_the_block Then
                                                                            Data_table_valori_existente.Rows.Add()

                                                                            Data_table_valori_existente.Rows(Index1).Item("X") = Block1.Position.X
                                                                            Data_table_valori_existente.Rows(Index1).Item("Y") = Block1.Position.Y
                                                                            Data_table_valori_existente.Rows(Index1).Item("SCALE") = Block1.ScaleFactors.X
                                                                            Data_table_valori_existente.Rows(Index1).Item("ROTATION") = Block1.Rotation
                                                                            Data_table_valori_existente.Rows(Index1).Item("LAYER") = Block1.Layer
                                                                            Data_table_valori_existente.Rows(Index1).Item("LAYOUT_INDEX") = Layout1.TabOrder



                                                                            If Block1.AttributeCollection.Count > 0 Then
                                                                                For Each id As ObjectId In Block1.AttributeCollection
                                                                                    If Not id.IsErased Then
                                                                                        Dim attRef As AttributeReference = DirectCast(Trans1.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), AttributeReference)
                                                                                        Dim Tag As String = attRef.Tag

                                                                                        If Data_table_valori_existente.Columns.Contains("ATR_N" & Tag) = False Then
                                                                                            Data_table_valori_existente.Columns.Add("ATR_N" & Tag, GetType(String))
                                                                                            Data_table_valori_existente.Columns.Add("ATR_V" & Tag, GetType(String))
                                                                                        End If
                                                                                        Data_table_valori_existente.Rows(Index1).Item("ATR_N" & Tag) = Tag

                                                                                        Dim Value1 As String = ""

                                                                                        Dim Value2 As String = ""
                                                                                        If attRef.IsMTextAttribute = False Then

                                                                                            If attRef.HasFields = True Then
                                                                                                Dim extDict As DBDictionary = Trans1.GetObject(attRef.ExtensionDictionary, OpenMode.ForRead)
                                                                                                Dim fldDictName As String = "ACAD_FIELD"
                                                                                                Dim fldEntryName As String = "TEXT"
                                                                                                If extDict.Contains(fldDictName) = True Then
                                                                                                    Dim fldDictId As ObjectId = extDict.GetAt(fldDictName)
                                                                                                    If Not fldDictId = ObjectId.Null Then
                                                                                                        Dim fldDict As DBDictionary = Trans1.GetObject(fldDictId, OpenMode.ForRead)

                                                                                                        If fldDict.Contains(fldEntryName) Then
                                                                                                            Dim fldId As ObjectId = fldDict.GetAt(fldEntryName)
                                                                                                            If Not fldId = ObjectId.Null Then
                                                                                                                Dim obj = Trans1.GetObject(fldId, OpenMode.ForRead)
                                                                                                                Dim Fld As Autodesk.AutoCAD.DatabaseServices.Field
                                                                                                                Fld = TryCast(obj, Autodesk.AutoCAD.DatabaseServices.Field)
                                                                                                                If IsNothing(Fld) = False Then
                                                                                                                    Value2 = Fld.GetFieldCode(FieldCodeFlags.AddMarkers Or FieldCodeFlags.FieldCode)

                                                                                                                    Data_table_valori_existente.Rows(Index1).Item("ATR_V" & Tag) = Value2
                                                                                                                End If
                                                                                                            End If
                                                                                                        End If

                                                                                                    End If
                                                                                                End If
                                                                                            Else
                                                                                                Value1 = attRef.TextString
                                                                                                Data_table_valori_existente.Rows(Index1).Item("ATR_V" & Tag) = Value1
                                                                                            End If
                                                                                        Else

                                                                                            If attRef.HasFields = True Then
                                                                                                Dim extDict As DBDictionary = Trans1.GetObject(attRef.ExtensionDictionary, OpenMode.ForRead)
                                                                                                Dim fldDictName As String = "ACAD_FIELD"
                                                                                                Dim fldEntryName As String = "TEXT"
                                                                                                If extDict.Contains(fldDictName) = True Then
                                                                                                    Dim fldDictId As ObjectId = extDict.GetAt(fldDictName)
                                                                                                    If Not fldDictId = ObjectId.Null Then
                                                                                                        Dim fldDict As DBDictionary = Trans1.GetObject(fldDictId, OpenMode.ForRead)

                                                                                                        If fldDict.Contains(fldEntryName) Then
                                                                                                            Dim fldId As ObjectId = fldDict.GetAt(fldEntryName)
                                                                                                            If Not fldId = ObjectId.Null Then
                                                                                                                Dim obj = Trans1.GetObject(fldId, OpenMode.ForRead)
                                                                                                                Dim Fld As Autodesk.AutoCAD.DatabaseServices.Field
                                                                                                                Fld = TryCast(obj, Autodesk.AutoCAD.DatabaseServices.Field)
                                                                                                                If IsNothing(Fld) = False Then
                                                                                                                    Value2 = Fld.GetFieldCode(FieldCodeFlags.AddMarkers Or FieldCodeFlags.FieldCode)
                                                                                                                    Data_table_valori_existente.Rows(Index1).Item("ATR_V" & Tag) = Value2

                                                                                                                End If
                                                                                                            End If
                                                                                                        End If

                                                                                                    End If
                                                                                                End If
                                                                                            Else
                                                                                                Value1 = attRef.MTextAttribute.Contents
                                                                                                Data_table_valori_existente.Rows(Index1).Item("ATR_V" & Tag) = Value1
                                                                                            End If
                                                                                        End If








                                                                                    End If
                                                                                Next
                                                                            End If
                                                                            Index1 = Index1 + 1


                                                                            Block1.UpgradeOpen()
                                                                            Block1.Erase()
                                                                        End If

                                                                    End If
                                                                End If
                                                            End If
                                                        Next
                                                    End If

                                                Next
                                                BlockTable1.UpgradeOpen()
                                                Dim Btr1 As BlockTableRecord = Trans1.GetObject(BlockTable1(Name_of_the_block), OpenMode.ForWrite)
                                                Btr1.Erase()
                                            End If
                                            If IsNothing(Data_table_valori_existente) = False Then
                                                If Data_table_valori_existente.Rows.Count > 0 Then



                                                    For j = 0 To Data_table_valori_existente.Rows.Count - 1

                                                        Dim X As Double = Data_table_valori_existente.Rows(j).Item("X")
                                                        Dim Y As Double = Data_table_valori_existente.Rows(j).Item("Y")
                                                        Dim Scale1 As Double = Data_table_valori_existente.Rows(j).Item("SCALE")
                                                        Dim Rotation1 As Double = Data_table_valori_existente.Rows(j).Item("ROTATION")
                                                        Dim Layer1 As String = Data_table_valori_existente.Rows(j).Item("LAYER")
                                                        Dim Layout_index As Integer = Data_table_valori_existente.Rows(j).Item("LAYOUT_INDEX")
                                                        If Data_table_valori_existente.Columns.Count > 6 Then
                                                            For k = 6 To Data_table_valori_existente.Columns.Count - 1 Step 2
                                                                Colectie_nume_atribute.Add(Data_table_valori_existente.Rows(j).Item(k))
                                                                Colectie_valori_atribute.Add(Data_table_valori_existente.Rows(j).Item(k + 1))
                                                            Next
                                                        End If
                                                        For Each entry As DBDictionaryEntry In Layoutdict
                                                            Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                                            If Layout1.TabOrder = Layout_index Then

                                                                Dim BTrecord As BlockTableRecord = Trans1.GetObject(Layout1.BlockTableRecordId, OpenMode.ForWrite)

                                                                InsertBlock_with_multiple_atributes_background(BlockDwg, Name_of_the_block, Database1, New Point3d(X, Y, 0), Scale1, BTrecord, Layer1, Colectie_nume_atribute, Colectie_valori_atribute)
                                                            End If
                                                        Next


                                                    Next



                                                End If
                                            End If



                                            Trans1.Commit()

                                            Try
                                                Database1.SaveAs(Drawing1, True, DwgVersion.Current, Database1.SecurityParameters)
                                            Catch ex As Exception
                                                Error_list = Error_list & vbCrLf & Drawing1
                                            End Try
                                        End Using

                                        Database1.Dispose()

                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database
                                    End If

123:
                                Next

                                Trans11.Commit()
                            End Using
                        End Using
                    End If


                    If Not Error_list = "" Then
                        MsgBox("There are not saved files. The list of these files is loaded into the clipboard" & vbCrLf &
                               "Use paste <Ctrl+V> command in a text editor program (notepad, excel, word) to see the list")
                        My.Computer.Clipboard.SetText(Error_list)

                    End If


                Catch ex As Exception
                    MsgBox(ex.Message)

                End Try

            Catch ex As System.SystemException
                MsgBox("1: " & ex.Message)
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: DONE!")
            Freeze_operations = False

        End If
    End Sub




    Private Sub Button_delete_prop_band_old_peneast_Click(sender As Object, e As EventArgs) Handles Button_delete_prop_band_old_peneast.Click
        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Error_list As String = ""

            Dim x1 As Double = 294.998
            Dim y1 As Double = 2236.4206
            Dim x2 As Double = 3519.998
            Dim y2 As Double = 2361.1932

            Try

                Try

                    If IsNumeric(TextBox_Layout_index_START.Text) = False Then
                        MsgBox("Please specify the layout START INDEX!")
                        Freeze_operations = False
                        Exit Sub
                    End If
                    If IsNumeric(TextBox_Layout_index_END.Text) = False Then
                        MsgBox("Please specify the layout END INDEX!")
                        Freeze_operations = False
                        Exit Sub
                    End If

                    If CInt(TextBox_Layout_index_START.Text) = 0 Or CInt(TextBox_Layout_index_END.Text) = 0 Then
                        MsgBox("the layout 0 = modelspace")
                        Freeze_operations = False
                        Exit Sub
                    End If




                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                    If ListBox_DWG.Items.Count > 0 Then
                        Using lock1 As DocumentLock = ThisDrawing.LockDocument

                            Using Trans11 As Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                                For i = 0 To ListBox_DWG.Items.Count - 1
                                    Dim Drawing1 As String = ListBox_DWG.Items(i)
                                    If IO.File.Exists(Drawing1) = True Then

                                        Dim Colectie_nume_atribute As New Specialized.StringCollection
                                        Dim Colectie_valori_atribute As New Specialized.StringCollection



                                        Dim Database1 As New Database(False, True)

                                        Try
                                            Try
                                                Try
                                                    Database1.ReadDwgFile(Drawing1, IO.FileShare.ReadWrite, True, Nothing)
                                                Catch ex As Exception
                                                    MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                    GoTo 123
                                                End Try
                                            Catch ex As IO.IOException
                                                MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                Error_list = Error_list & vbCrLf & Drawing1
                                                GoTo 123
                                            End Try
                                        Catch ex As System.SystemException
                                            MsgBox(Drawing1 & vbCrLf & "could not be open")
                                            Error_list = Error_list & vbCrLf & Drawing1
                                            GoTo 123
                                        End Try


                                        HostApplicationServices.WorkingDatabase = Database1



                                        Dim is_erased As Boolean = False

                                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction

                                            Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                                            Dim Layoutdict As DBDictionary

                                            Layoutdict = Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead)

                                            For Each entry As DBDictionaryEntry In Layoutdict
                                                Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                                If Layout1.TabOrder >= CInt(TextBox_Layout_index_START.Text) And Layout1.TabOrder <= CInt(TextBox_Layout_index_END.Text) And Layout1.TabOrder > 0 Then
                                                    Dim BTrecord As BlockTableRecord = Trans1.GetObject(Layout1.BlockTableRecordId, OpenMode.ForWrite)
                                                    For Each Id1 As ObjectId In BTrecord
                                                        Dim Ent1 As Entity = TryCast(Trans1.GetObject(Id1, OpenMode.ForRead), Entity)
                                                        If IsNothing(Ent1) = False Then
                                                            If TypeOf (Ent1) Is Viewport Then
                                                                Dim Viewport1 As Viewport = TryCast(Ent1, Viewport)
                                                                If IsNothing(Viewport1) = False Then

                                                                    If Viewport1.CenterPoint.X > x1 And Viewport1.CenterPoint.X < x2 And Viewport1.CenterPoint.Y > y1 And Viewport1.CenterPoint.Y < y2 Then
                                                                        Viewport1.UpgradeOpen()
                                                                        Viewport1.Erase()
                                                                        is_erased = True
                                                                    End If

                                                                End If


                                                            End If

                                                            Dim mtext1 As MText = TryCast(Ent1, MText)
                                                            If IsNothing(mtext1) = False Then

                                                                If mtext1.Location.X > x1 And mtext1.Location.X < x2 And mtext1.Location.Y > y1 And mtext1.Location.Y < y2 Then

                                                                    mtext1.UpgradeOpen()
                                                                    mtext1.Erase()
                                                                    is_erased = True

                                                                End If

                                                            End If

                                                            Dim text1 As DBText = TryCast(Ent1, DBText)
                                                            If IsNothing(text1) = False Then

                                                                If text1.Position.X > x1 And text1.Position.X < x2 And text1.Position.Y > y1 And text1.Position.Y < y2 Then

                                                                    text1.UpgradeOpen()
                                                                    text1.Erase()
                                                                    is_erased = True

                                                                End If

                                                            End If

                                                            Dim LINE1 As Line = TryCast(Ent1, Line)
                                                            If IsNothing(LINE1) = False Then
                                                                Dim ptx1 As Double = LINE1.StartPoint.X
                                                                Dim ptx2 As Double = LINE1.EndPoint.X
                                                                Dim pty1 As Double = LINE1.StartPoint.Y
                                                                Dim pty2 As Double = LINE1.EndPoint.Y

                                                                If ptx1 > x1 And ptx1 < x2 And pty1 > y1 And pty1 < y2 And ptx2 > x1 And ptx2 < x2 And pty2 > y1 And pty2 < y2 Then
                                                                    LINE1.UpgradeOpen()
                                                                    LINE1.Erase()
                                                                    is_erased = True
                                                                End If

                                                            End If

                                                            Dim Pline1 As Polyline = TryCast(Ent1, Polyline)
                                                            If IsNothing(Pline1) = False Then
                                                                Dim ptx1 As Double = Pline1.StartPoint.X
                                                                Dim ptx2 As Double = Pline1.EndPoint.X
                                                                Dim pty1 As Double = Pline1.StartPoint.Y
                                                                Dim pty2 As Double = Pline1.EndPoint.Y

                                                                If ptx1 > x1 And ptx1 < x2 And pty1 > y1 And pty1 < y2 And ptx2 > x1 And ptx2 < x2 And pty2 > y1 And pty2 < y2 Then
                                                                    Pline1.UpgradeOpen()
                                                                    Pline1.Erase()
                                                                    is_erased = True
                                                                End If

                                                            End If

                                                        End If
                                                    Next
                                                End If

                                            Next

                                            If is_erased = True Then
                                                Trans1.Commit()

                                                Try
                                                    Database1.SaveAs(Drawing1, True, DwgVersion.Current, Database1.SecurityParameters)
                                                Catch ex As Exception
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                End Try
                                            End If

                                        End Using

                                        Database1.Dispose()

                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database
                                    End If

123:
                                Next

                                Trans11.Commit()
                            End Using
                        End Using
                    End If


                    If Not Error_list = "" Then
                        MsgBox("There are not saved files. The list of these files is loaded into the clipboard" & vbCrLf &
                               "Use paste <Ctrl+V> command in a text editor program (notepad, excel, word) to see the list")
                        My.Computer.Clipboard.SetText(Error_list)

                    End If


                Catch ex As Exception
                    MsgBox(ex.Message)

                End Try

            Catch ex As System.SystemException
                MsgBox("1: " & ex.Message)
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: DONE!")
            Freeze_operations = False

        End If
    End Sub


    Private Sub Button_del_class_viewport_Click(sender As Object, e As EventArgs)
        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Error_list As String = ""

            Dim x1 As Double = 224.998
            Dim x2 As Double = 3519.998
            Dim y1 As Double = 825.001
            Dim y2 As Double = 875.001

            Try

                Try

                    If IsNumeric(TextBox_Layout_index_START.Text) = False Then
                        MsgBox("Please specify the layout START INDEX!")
                        Freeze_operations = False
                        Exit Sub
                    End If
                    If IsNumeric(TextBox_Layout_index_END.Text) = False Then
                        MsgBox("Please specify the layout END INDEX!")
                        Freeze_operations = False
                        Exit Sub
                    End If

                    If CInt(TextBox_Layout_index_START.Text) = 0 Or CInt(TextBox_Layout_index_END.Text) = 0 Then
                        MsgBox("the layout 0 = modelspace")
                        Freeze_operations = False
                        Exit Sub
                    End If




                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                    If ListBox_DWG.Items.Count > 0 Then
                        Using lock1 As DocumentLock = ThisDrawing.LockDocument

                            Using Trans11 As Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                                For i = 0 To ListBox_DWG.Items.Count - 1
                                    Dim Drawing1 As String = ListBox_DWG.Items(i)
                                    If IO.File.Exists(Drawing1) = True Then

                                        Dim Colectie_nume_atribute As New Specialized.StringCollection
                                        Dim Colectie_valori_atribute As New Specialized.StringCollection



                                        Dim Database1 As New Database(False, True)

                                        Try
                                            Try
                                                Try
                                                    Database1.ReadDwgFile(Drawing1, IO.FileShare.ReadWrite, True, Nothing)
                                                Catch ex As Exception
                                                    MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                    GoTo 123
                                                End Try
                                            Catch ex As IO.IOException
                                                MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                Error_list = Error_list & vbCrLf & Drawing1
                                                GoTo 123
                                            End Try
                                        Catch ex As System.SystemException
                                            MsgBox(Drawing1 & vbCrLf & "could not be open")
                                            Error_list = Error_list & vbCrLf & Drawing1
                                            GoTo 123
                                        End Try


                                        HostApplicationServices.WorkingDatabase = Database1



                                        Dim is_erased As Boolean = False

                                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction

                                            Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                                            Dim Layoutdict As DBDictionary

                                            Layoutdict = Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead)

                                            For Each entry As DBDictionaryEntry In Layoutdict
                                                Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                                If Layout1.TabOrder >= CInt(TextBox_Layout_index_START.Text) And Layout1.TabOrder <= CInt(TextBox_Layout_index_END.Text) And Layout1.TabOrder > 0 Then
                                                    Dim BTrecord As BlockTableRecord = Trans1.GetObject(Layout1.BlockTableRecordId, OpenMode.ForWrite)
                                                    For Each Id1 As ObjectId In BTrecord
                                                        Dim Ent1 As Entity = TryCast(Trans1.GetObject(Id1, OpenMode.ForRead), Entity)
                                                        If IsNothing(Ent1) = False Then
                                                            If TypeOf (Ent1) Is Viewport Then
                                                                Dim Viewport1 As Viewport = TryCast(Ent1, Viewport)
                                                                If IsNothing(Viewport1) = False Then

                                                                    If Viewport1.CenterPoint.X > x1 And Viewport1.CenterPoint.X < x2 And Viewport1.CenterPoint.Y > y1 And Viewport1.CenterPoint.Y < y2 Then
                                                                        Viewport1.UpgradeOpen()
                                                                        Viewport1.Erase()
                                                                        is_erased = True
                                                                    End If

                                                                End If


                                                            End If

                                                            Dim mtext1 As MText = TryCast(Ent1, MText)
                                                            If IsNothing(mtext1) = False Then

                                                                If mtext1.Location.X > x1 And mtext1.Location.X < x2 And mtext1.Location.Y > y1 And mtext1.Location.Y < y2 And (mtext1.Layer = "CLASS" Or mtext1.Layer = "DFACTOR") Then

                                                                    mtext1.UpgradeOpen()
                                                                    mtext1.Erase()
                                                                    is_erased = True

                                                                End If

                                                            End If

                                                            Dim text1 As DBText = TryCast(Ent1, DBText)
                                                            If IsNothing(text1) = False Then

                                                                If text1.Position.X > x1 And text1.Position.X < x2 And text1.Position.Y > y1 And text1.Position.Y < y2 And (text1.Layer = "CLASS" Or text1.Layer = "DFACTOR") Then

                                                                    text1.UpgradeOpen()
                                                                    text1.Erase()
                                                                    is_erased = True

                                                                End If

                                                            End If

                                                            Dim LINE1 As Line = TryCast(Ent1, Line)
                                                            If IsNothing(LINE1) = False Then

                                                                If Round(LINE1.Length, 0) = 25 And LINE1.Layer = "DFACTOR" Then

                                                                    LINE1.UpgradeOpen()
                                                                    LINE1.Erase()
                                                                    is_erased = True
                                                                End If

                                                            End If

                                                        End If
                                                    Next
                                                End If

                                            Next

                                            If is_erased = True Then
                                                Trans1.Commit()

                                                Try
                                                    Database1.SaveAs(Drawing1, True, DwgVersion.Current, Database1.SecurityParameters)
                                                Catch ex As Exception
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                End Try
                                            End If

                                        End Using

                                        Database1.Dispose()

                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database
                                    End If

123:
                                Next

                                Trans11.Commit()
                            End Using
                        End Using
                    End If


                    If Not Error_list = "" Then
                        MsgBox("There are not saved files. The list of these files is loaded into the clipboard" & vbCrLf &
                               "Use paste <Ctrl+V> command in a text editor program (notepad, excel, word) to see the list")
                        My.Computer.Clipboard.SetText(Error_list)

                    End If


                Catch ex As Exception
                    MsgBox(ex.Message)

                End Try

            Catch ex As System.SystemException
                MsgBox("1: " & ex.Message)
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: DONE!")
            Freeze_operations = False

        End If
    End Sub

    Private Sub Button_del_graph_viewport_Click(sender As Object, e As EventArgs)
        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Error_list As String = ""

            Dim x1 As Double = 443.5236
            Dim x2 As Double = 3519.998
            Dim y1 As Double = 250.001
            Dim y2 As Double = 800.001

            Try

                Try

                    If IsNumeric(TextBox_Layout_index_START.Text) = False Then
                        MsgBox("Please specify the layout START INDEX!")
                        Freeze_operations = False
                        Exit Sub
                    End If
                    If IsNumeric(TextBox_Layout_index_END.Text) = False Then
                        MsgBox("Please specify the layout END INDEX!")
                        Freeze_operations = False
                        Exit Sub
                    End If

                    If CInt(TextBox_Layout_index_START.Text) = 0 Or CInt(TextBox_Layout_index_END.Text) = 0 Then
                        MsgBox("the layout 0 = modelspace")
                        Freeze_operations = False
                        Exit Sub
                    End If




                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                    If ListBox_DWG.Items.Count > 0 Then
                        Using lock1 As DocumentLock = ThisDrawing.LockDocument

                            Using Trans11 As Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                                For i = 0 To ListBox_DWG.Items.Count - 1
                                    Dim Drawing1 As String = ListBox_DWG.Items(i)
                                    If IO.File.Exists(Drawing1) = True Then

                                        Dim Colectie_nume_atribute As New Specialized.StringCollection
                                        Dim Colectie_valori_atribute As New Specialized.StringCollection



                                        Dim Database1 As New Database(False, True)

                                        Try
                                            Try
                                                Try
                                                    Database1.ReadDwgFile(Drawing1, IO.FileShare.ReadWrite, True, Nothing)
                                                Catch ex As Exception
                                                    MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                    GoTo 123
                                                End Try
                                            Catch ex As IO.IOException
                                                MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                Error_list = Error_list & vbCrLf & Drawing1
                                                GoTo 123
                                            End Try
                                        Catch ex As System.SystemException
                                            MsgBox(Drawing1 & vbCrLf & "could not be open")
                                            Error_list = Error_list & vbCrLf & Drawing1
                                            GoTo 123
                                        End Try


                                        HostApplicationServices.WorkingDatabase = Database1



                                        Dim is_erased As Boolean = False

                                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction

                                            Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                                            Dim Layoutdict As DBDictionary

                                            Layoutdict = Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead)

                                            For Each entry As DBDictionaryEntry In Layoutdict
                                                Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                                If Layout1.TabOrder >= CInt(TextBox_Layout_index_START.Text) And Layout1.TabOrder <= CInt(TextBox_Layout_index_END.Text) And Layout1.TabOrder > 0 Then
                                                    Dim BTrecord As BlockTableRecord = Trans1.GetObject(Layout1.BlockTableRecordId, OpenMode.ForWrite)
                                                    For Each Id1 As ObjectId In BTrecord
                                                        Dim Ent1 As Entity = TryCast(Trans1.GetObject(Id1, OpenMode.ForRead), Entity)
                                                        If IsNothing(Ent1) = False Then
                                                            If TypeOf (Ent1) Is Viewport Then
                                                                Dim Viewport1 As Viewport = TryCast(Ent1, Viewport)
                                                                If IsNothing(Viewport1) = False Then

                                                                    If Viewport1.CenterPoint.X > x1 And Viewport1.CenterPoint.X < x2 And Viewport1.CenterPoint.Y > y1 And Viewport1.CenterPoint.Y < y2 Then
                                                                        Viewport1.UpgradeOpen()
                                                                        Viewport1.Erase()
                                                                        is_erased = True
                                                                    End If

                                                                End If


                                                            End If

                                                            Dim mtext1 As MText = TryCast(Ent1, MText)
                                                            If IsNothing(mtext1) = False Then

                                                                If mtext1.Location.X > x1 And mtext1.Location.X < x2 And mtext1.Location.Y > y1 And mtext1.Location.Y < y2 Then
                                                                    If mtext1.Contents.Contains("STA") = True Or mtext1.Contents = ("ELEVATION") Then
                                                                        mtext1.UpgradeOpen()
                                                                        mtext1.Erase()
                                                                        is_erased = True
                                                                    End If

                                                                End If

                                                            End If

                                                        End If
                                                    Next
                                                End If

                                            Next

                                            If is_erased = True Then
                                                Trans1.Commit()

                                                Try
                                                    Database1.SaveAs(Drawing1, True, DwgVersion.Current, Database1.SecurityParameters)
                                                Catch ex As Exception
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                End Try
                                            End If

                                        End Using

                                        Database1.Dispose()

                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database
                                    End If

123:
                                Next

                                Trans11.Commit()
                            End Using
                        End Using
                    End If


                    If Not Error_list = "" Then
                        MsgBox("There are not saved files. The list of these files is loaded into the clipboard" & vbCrLf &
                               "Use paste <Ctrl+V> command in a text editor program (notepad, excel, word) to see the list")
                        My.Computer.Clipboard.SetText(Error_list)

                    End If


                Catch ex As Exception
                    MsgBox(ex.Message)

                End Try

            Catch ex As System.SystemException
                MsgBox("1: " & ex.Message)
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: DONE!")
            Freeze_operations = False

        End If
    End Sub



    Public Shared Function FindObjectId(text As String, ByRef objId As ObjectId) As String
        Const prefix As String = "%<\_ObjId "
        Const suffix As String = ">%"

        ' Find the location of the prefix string
        Dim preLoc As Integer = text.IndexOf(prefix)
        If preLoc > 0 Then
            ' Find the location of the ID itself
            Dim idLoc As Integer = preLoc + prefix.Length

            ' Get the remaining string
            Dim remains As String = text.Substring(idLoc)

            ' Find the location of the suffix
            Dim sufLoc As Integer = remains.IndexOf(suffix)

            ' Extract the ID string and get the ObjectId
            Dim id As String = remains.Remove(sufLoc)
            objId = New ObjectId(Convert.ToInt32(id))

            ' Return the remainder, to allow extraction
            ' of any remaining IDs
            Return remains.Substring(sufLoc + suffix.Length)
        Else
            objId = ObjectId.Null
            Return ""
        End If
    End Function

    Private Sub Button_find_replace_Click(sender As Object, e As EventArgs) Handles Button_find_replace.Click
        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Error_list As String = ""

            If TextBox_find_text.Text = "" Or TextBox_replace_text.Text = "" Then
                Freeze_operations = False
                Exit Sub
            End If

            Try

                Try

                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                    If ListBox_DWG.Items.Count > 0 Then
                        Using lock1 As DocumentLock = ThisDrawing.LockDocument

                            Using Trans11 As Transaction = ThisDrawing.TransactionManager.StartTransaction

                                For i = 0 To ListBox_DWG.Items.Count - 1
                                    Dim Drawing1 As String = ListBox_DWG.Items(i)
                                    If IO.File.Exists(Drawing1) = True Then

                                        Dim Colectie_nume_atribute As New Specialized.StringCollection
                                        Dim Colectie_valori_atribute As New Specialized.StringCollection



                                        Dim Database1 As New Database(False, True)

                                        Try
                                            Try
                                                Try
                                                    Database1.ReadDwgFile(Drawing1, IO.FileShare.ReadWrite, True, Nothing)
                                                Catch ex As Exception
                                                    MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                    GoTo 123
                                                End Try
                                            Catch ex As IO.IOException
                                                MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                Error_list = Error_list & vbCrLf & Drawing1
                                                GoTo 123
                                            End Try
                                        Catch ex As System.SystemException
                                            MsgBox(Drawing1 & vbCrLf & "could not be open")
                                            Error_list = Error_list & vbCrLf & Drawing1
                                            GoTo 123
                                        End Try


                                        HostApplicationServices.WorkingDatabase = Database1





                                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction

                                            Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                                            Dim Layoutdict As DBDictionary

                                            Layoutdict = Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead)
                                            Dim Index_datatable As Integer = 0


                                            Dim BlockTable1 As BlockTable = Trans1.GetObject(Database1.BlockTableId, OpenMode.ForRead)

                                            For Each entry As DBDictionaryEntry In Layoutdict
                                                Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                                Dim BTrecord As BlockTableRecord = Trans1.GetObject(Layout1.BlockTableRecordId, OpenMode.ForWrite)
                                                For Each Id1 As ObjectId In BTrecord
                                                    Dim Ent1 As Entity = TryCast(Trans1.GetObject(Id1, OpenMode.ForRead), Entity)
                                                    If IsNothing(Ent1) = False Then
                                                        If TypeOf (Ent1) Is DBText Then
                                                            Dim Text1 As DBText = TryCast(Ent1, DBText)
                                                            If Text1.TextString = TextBox_find_text.Text Then
                                                                Dim Mtext1 As New MText
                                                                Mtext1.Contents = TextBox_replace_text.Text
                                                                With Mtext1
                                                                    .TextHeight = Text1.Height
                                                                    .TextStyleId = Text1.TextStyleId
                                                                    .Layer = Text1.Layer
                                                                    .ColorIndex = Text1.ColorIndex
                                                                    .LineWeight = Text1.LineWeight
                                                                    .Location = Text1.Position
                                                                    .Rotation = Text1.Rotation
                                                                    .Attachment = AttachmentPoint.BottomLeft
                                                                    BTrecord.AppendEntity(Mtext1)
                                                                    Trans1.AddNewlyCreatedDBObject(Mtext1, True)
                                                                End With
                                                                Text1.UpgradeOpen()
                                                                Text1.Erase()
                                                            End If

                                                        End If
                                                    End If

                                                Next


                                            Next





                                            Trans1.Commit()

                                            Try
                                                Database1.SaveAs(Drawing1, True, DwgVersion.Current, Database1.SecurityParameters)
                                            Catch ex As Exception
                                                Error_list = Error_list & vbCrLf & Drawing1
                                            End Try
                                        End Using

                                        Database1.Dispose()

                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database
                                    End If

123:
                                Next

                                Trans11.Commit()
                            End Using
                        End Using
                    End If


                    If Not Error_list = "" Then
                        MsgBox("There are not saved files. The list of these files is loaded into the clipboard" & vbCrLf &
                               "Use paste <Ctrl+V> command in a text editor program (notepad, excel, word) to see the list")
                        My.Computer.Clipboard.SetText(Error_list)

                    End If


                Catch ex As Exception
                    MsgBox(ex.Message)

                End Try

            Catch ex As System.SystemException
                MsgBox("1: " & ex.Message)
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: DONE!")
            Freeze_operations = False

        End If
    End Sub

    Private Sub Button_replace_block_Click(sender As Object, e As EventArgs) Handles Button_replace_block.Click
        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Error_list As String = ""

            Try

                Try

                    If IsNumeric(TextBox_Layout_index_START.Text) = False Then
                        MsgBox("Please specify the layout for block insertion!")
                        Freeze_operations = False
                        Exit Sub
                    End If
                    If IsNumeric(TextBox_Layout_index_END.Text) = False Then
                        MsgBox("Please specify the layout for block insertion!")
                        Freeze_operations = False
                        Exit Sub
                    End If



                    Dim Column_start As String = TextBox_block_att_value_column_start.Text.ToUpper
                    Dim Column_end As String = TextBox_block_att_value_column_end.Text.ToUpper
                    Dim Column_atr_name As String = TextBox_block_att_name_column.Text.ToUpper

                    Dim Row_with_file_names As Integer = 0
                    If IsNumeric(TextBox_ROW_FILE_NAME.Text) = True Then
                        Row_with_file_names = CInt(TextBox_ROW_FILE_NAME.Text)
                    End If


                    Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                    W1 = Nothing
                    Dim Col_start As Integer = 0
                    Dim Col_end As Integer = 0
                    Dim Col_atr As Integer = 0

                    If Panel_blocks_insertion.Visible = True Then
                        Try
                            W1 = Get_active_worksheet_from_Excel_with_error()
                            If IsNothing(W1) = False Then
                                Col_start = W1.Range(Column_start & "1").Column
                                Col_end = W1.Range(Column_end & "1").Column
                                Col_atr = W1.Range(Column_atr_name & "1").Column
                            End If

                        Catch ex As System.SystemException
                            W1 = Nothing
                            Col_start = 0
                            Col_end = 0
                            Col_atr = 0
                        End Try
                    End If

                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                    If ListBox_DWG.Items.Count > 0 Then
                        Using lock1 As DocumentLock = ThisDrawing.LockDocument

                            Using Trans11 As Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()



                                Dim BlockDwg As String = TextBox_block_name.Text
                                If IO.File.Exists(BlockDwg) = False Then
                                    MsgBox("no dwg file")
                                    Freeze_operations = False
                                    Exit Sub
                                End If

                                Dim Name_of_the_block As String = IO.Path.GetFileNameWithoutExtension(BlockDwg)
                                Dim Block_de_sters As String = TextBox_ERASE_BLOCK.Text

                                For i = 0 To ListBox_DWG.Items.Count - 1
                                    Dim Drawing1 As String = ListBox_DWG.Items(i)
                                    If IO.File.Exists(Drawing1) = True Then

                                        Dim Colectie_nume_atribute As New Specialized.StringCollection
                                        Dim Colectie_valori_atribute As New Specialized.StringCollection



                                        Dim Database1 As New Database(False, True)

                                        Try
                                            Try
                                                Try
                                                    Database1.ReadDwgFile(Drawing1, IO.FileShare.ReadWrite, True, Nothing)
                                                Catch ex As Exception
                                                    MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                    GoTo 123
                                                End Try
                                            Catch ex As IO.IOException
                                                MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                Error_list = Error_list & vbCrLf & Drawing1
                                                GoTo 123
                                            End Try
                                        Catch ex As System.SystemException
                                            MsgBox(Drawing1 & vbCrLf & "could not be open")
                                            Error_list = Error_list & vbCrLf & Drawing1
                                            GoTo 123
                                        End Try


                                        HostApplicationServices.WorkingDatabase = Database1





                                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction

                                            Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                                            Dim Layoutdict As DBDictionary = Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead)



                                            Dim BlockTable1 As BlockTable = Trans1.GetObject(Database1.BlockTableId, OpenMode.ForRead)
                                            If BlockTable1.Has(Block_de_sters) = True Then
                                                For Each entry As DBDictionaryEntry In Layoutdict
                                                    Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                                    If Not Layout1.TabOrder = 0 Then
                                                        Dim BTrecord As BlockTableRecord = Trans1.GetObject(Layout1.BlockTableRecordId, OpenMode.ForWrite)
                                                        For Each Id1 As ObjectId In BTrecord
                                                            Dim Ent1 As Entity = TryCast(Trans1.GetObject(Id1, OpenMode.ForRead), Entity)
                                                            If IsNothing(Ent1) = False Then
                                                                If TypeOf (Ent1) Is BlockReference Then
                                                                    Dim Block1 As BlockReference = TryCast(Ent1, BlockReference)
                                                                    If IsNothing(Block1) = False Then
                                                                        Dim BlockTrec As BlockTableRecord = Nothing

                                                                        If Block1.IsDynamicBlock = True Then
                                                                            BlockTrec = Trans1.GetObject(Block1.DynamicBlockTableRecord, OpenMode.ForRead)
                                                                        Else
                                                                            BlockTrec = Trans1.GetObject(Block1.BlockTableRecord, OpenMode.ForRead)
                                                                        End If
                                                                        If BlockTrec.Name = Block_de_sters Then
                                                                            Block1.UpgradeOpen()
                                                                            Block1.Erase()
                                                                        End If

                                                                    End If
                                                                End If
                                                            End If
                                                        Next
                                                    End If

                                                Next
                                                BlockTable1.UpgradeOpen()
                                                Dim Btr1 As BlockTableRecord = Trans1.GetObject(BlockTable1(Block_de_sters), OpenMode.ForWrite)
                                                Btr1.Erase()
                                                Trans1.Commit()
                                            End If

                                        End Using

                                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction
                                            Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                                            Dim Layoutdict As DBDictionary = Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead)


                                            For Each entry As DBDictionaryEntry In Layoutdict
                                                Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                                If Layout1.TabOrder >= CInt(TextBox_Layout_index_START.Text) And Layout1.TabOrder <= CInt(TextBox_Layout_index_END.Text) Then

                                                    Dim BTrecord As BlockTableRecord = Trans1.GetObject(Layout1.BlockTableRecordId, OpenMode.ForWrite)

                                                    InsertBlock_with_multiple_atributes_background(BlockDwg, Name_of_the_block, Database1, New Point3d(0, 0, 0), 1, BTrecord, "0", Colectie_nume_atribute, Colectie_valori_atribute)
                                                End If
                                            Next

                                            Trans1.Commit()

                                            Try
                                                Database1.SaveAs(Drawing1, True, DwgVersion.Current, Database1.SecurityParameters)
                                            Catch ex As Exception
                                                Error_list = Error_list & vbCrLf & Drawing1
                                            End Try
                                        End Using

                                        Database1.Dispose()

                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database
                                    End If

123:
                                Next

                                Trans11.Commit()
                            End Using
                        End Using
                    End If


                    If Not Error_list = "" Then
                        MsgBox("There are not saved files. The list of these files is loaded into the clipboard" & vbCrLf &
                               "Use paste <Ctrl+V> command in a text editor program (notepad, excel, word) to see the list")
                        My.Computer.Clipboard.SetText(Error_list)

                    End If


                Catch ex As Exception
                    MsgBox(ex.Message)

                End Try

            Catch ex As System.SystemException
                MsgBox("1: " & ex.Message)
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: DONE!")
            Freeze_operations = False

        End If
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
    Public Function Extract_equation_value(ByVal Station_measured As Double) As Double
        Dim Valoare As Double = 0
        If IsNothing(Data_table_station_equation) = False Then
            If Data_table_station_equation.Rows.Count > 0 Then
                For i = 0 To Data_table_station_equation.Rows.Count - 1
                    If IsDBNull(Data_table_station_equation.Rows(i).Item("STATION_BACK")) = False And IsDBNull(Data_table_station_equation.Rows(i).Item("STATION_AHEAD")) = False Then
                        Dim Station_back As Double = Data_table_station_equation.Rows(i).Item("STATION_BACK")
                        Dim Station_ahead As Double = Data_table_station_equation.Rows(i).Item("STATION_AHEAD")

                        If Station_measured < Station_back Then
                            Exit For
                        End If

                        Valoare = Valoare + Station_back - Station_ahead

                    End If
                Next
            End If


        End If


        Return Valoare
    End Function

    Private Sub Button_del_station_band_Click(sender As Object, e As EventArgs) Handles Button_del_station_band.Click
        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Error_list As String = ""

            Dim x1 As Double = 294.998
            Dim y1 As Double = 1975.001
            Dim x2 As Double = 3519.998
            Dim y2 As Double = 2250.001

            Try

                Try

                    If IsNumeric(TextBox_Layout_index_START.Text) = False Then
                        MsgBox("Please specify the layout START INDEX!")
                        Freeze_operations = False
                        Exit Sub
                    End If
                    If IsNumeric(TextBox_Layout_index_END.Text) = False Then
                        MsgBox("Please specify the layout END INDEX!")
                        Freeze_operations = False
                        Exit Sub
                    End If

                    If CInt(TextBox_Layout_index_START.Text) = 0 Or CInt(TextBox_Layout_index_END.Text) = 0 Then
                        MsgBox("the layout 0 = modelspace")
                        Freeze_operations = False
                        Exit Sub
                    End If




                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                    If ListBox_DWG.Items.Count > 0 Then
                        Using lock1 As DocumentLock = ThisDrawing.LockDocument

                            Using Trans11 As Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                                For i = 0 To ListBox_DWG.Items.Count - 1
                                    Dim Drawing1 As String = ListBox_DWG.Items(i)
                                    If IO.File.Exists(Drawing1) = True Then

                                        Dim Colectie_nume_atribute As New Specialized.StringCollection
                                        Dim Colectie_valori_atribute As New Specialized.StringCollection



                                        Dim Database1 As New Database(False, True)

                                        Try
                                            Try
                                                Try
                                                    Database1.ReadDwgFile(Drawing1, IO.FileShare.ReadWrite, True, Nothing)
                                                Catch ex As Exception
                                                    MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                    GoTo 123
                                                End Try
                                            Catch ex As IO.IOException
                                                MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                Error_list = Error_list & vbCrLf & Drawing1
                                                GoTo 123
                                            End Try
                                        Catch ex As System.SystemException
                                            MsgBox(Drawing1 & vbCrLf & "could not be open")
                                            Error_list = Error_list & vbCrLf & Drawing1
                                            GoTo 123
                                        End Try


                                        HostApplicationServices.WorkingDatabase = Database1



                                        Dim is_erased As Boolean = False

                                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction

                                            Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                                            Dim Layoutdict As DBDictionary

                                            Layoutdict = Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead)

                                            For Each entry As DBDictionaryEntry In Layoutdict
                                                Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                                If Layout1.TabOrder >= CInt(TextBox_Layout_index_START.Text) And Layout1.TabOrder <= CInt(TextBox_Layout_index_END.Text) And Layout1.TabOrder > 0 Then
                                                    Dim BTrecord As BlockTableRecord = Trans1.GetObject(Layout1.BlockTableRecordId, OpenMode.ForWrite)
                                                    For Each Id1 As ObjectId In BTrecord
                                                        Dim Ent1 As Entity = TryCast(Trans1.GetObject(Id1, OpenMode.ForRead), Entity)
                                                        If IsNothing(Ent1) = False Then
                                                            If TypeOf (Ent1) Is Viewport Then
                                                                Dim Viewport1 As Viewport = TryCast(Ent1, Viewport)
                                                                If IsNothing(Viewport1) = False Then

                                                                    If Viewport1.CenterPoint.X > x1 And Viewport1.CenterPoint.X < x2 And Viewport1.CenterPoint.Y > y1 And Viewport1.CenterPoint.Y < y2 Then
                                                                        Viewport1.UpgradeOpen()
                                                                        Viewport1.Erase()
                                                                        is_erased = True
                                                                    End If

                                                                End If


                                                            End If

                                                            Dim mtext1 As MText = TryCast(Ent1, MText)
                                                            If IsNothing(mtext1) = False Then

                                                                If mtext1.Location.X > x1 And mtext1.Location.X < x2 And mtext1.Location.Y > y1 And mtext1.Location.Y < y2 Then

                                                                    mtext1.UpgradeOpen()
                                                                    mtext1.Erase()
                                                                    is_erased = True

                                                                End If

                                                            End If

                                                            Dim text1 As DBText = TryCast(Ent1, DBText)
                                                            If IsNothing(text1) = False Then

                                                                If text1.Position.X > x1 And text1.Position.X < x2 And text1.Position.Y > y1 And text1.Position.Y < y2 Then

                                                                    text1.UpgradeOpen()
                                                                    text1.Erase()
                                                                    is_erased = True

                                                                End If

                                                            End If

                                                            Dim LINE1 As Line = TryCast(Ent1, Line)
                                                            If IsNothing(LINE1) = False Then
                                                                Dim ptx1 As Double = LINE1.StartPoint.X
                                                                Dim ptx2 As Double = LINE1.EndPoint.X
                                                                Dim pty1 As Double = LINE1.StartPoint.Y
                                                                Dim pty2 As Double = LINE1.EndPoint.Y

                                                                If ptx1 > x1 And ptx1 < x2 And pty1 > y1 And pty1 < y2 And ptx2 > x1 And ptx2 < x2 And pty2 > y1 And pty2 < y2 Then
                                                                    LINE1.UpgradeOpen()
                                                                    LINE1.Erase()
                                                                    is_erased = True
                                                                End If

                                                            End If

                                                            Dim Pline1 As Polyline = TryCast(Ent1, Polyline)
                                                            If IsNothing(Pline1) = False Then
                                                                Dim ptx1 As Double = Pline1.StartPoint.X
                                                                Dim ptx2 As Double = Pline1.EndPoint.X
                                                                Dim pty1 As Double = Pline1.StartPoint.Y
                                                                Dim pty2 As Double = Pline1.EndPoint.Y

                                                                If ptx1 > x1 And ptx1 < x2 And pty1 > y1 And pty1 < y2 And ptx2 > x1 And ptx2 < x2 And pty2 > y1 And pty2 < y2 Then
                                                                    Pline1.UpgradeOpen()
                                                                    Pline1.Erase()
                                                                    is_erased = True
                                                                End If

                                                            End If

                                                        End If
                                                    Next
                                                End If

                                            Next

                                            If is_erased = True Then
                                                Trans1.Commit()

                                                Try
                                                    Database1.SaveAs(Drawing1, True, DwgVersion.Current, Database1.SecurityParameters)
                                                Catch ex As Exception
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                End Try
                                            End If

                                        End Using

                                        Database1.Dispose()

                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database
                                    End If

123:
                                Next

                                Trans11.Commit()
                            End Using
                        End Using
                    End If


                    If Not Error_list = "" Then
                        MsgBox("There are not saved files. The list of these files is loaded into the clipboard" & vbCrLf &
                               "Use paste <Ctrl+V> command in a text editor program (notepad, excel, word) to see the list")
                        My.Computer.Clipboard.SetText(Error_list)

                    End If


                Catch ex As Exception
                    MsgBox(ex.Message)

                End Try

            Catch ex As System.SystemException
                MsgBox("1: " & ex.Message)
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: DONE!")
            Freeze_operations = False

        End If
    End Sub

    Private Sub Button_delete_env_blocks_Click(sender As Object, e As EventArgs) Handles Button_delete_env_blocks.Click

        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Error_list As String = ""

            Dim x1 As Double = 294.998
            Dim y1 As Double = 875.001
            Dim x2 As Double = 3519.998
            Dim y2 As Double = 960.001

            Try

                Try

                    If IsNumeric(TextBox_Layout_index_START.Text) = False Then
                        MsgBox("Please specify the layout START INDEX!")
                        Freeze_operations = False
                        Exit Sub
                    End If
                    If IsNumeric(TextBox_Layout_index_END.Text) = False Then
                        MsgBox("Please specify the layout END INDEX!")
                        Freeze_operations = False
                        Exit Sub
                    End If

                    If CInt(TextBox_Layout_index_START.Text) = 0 Or CInt(TextBox_Layout_index_END.Text) = 0 Then
                        MsgBox("the layout 0 = modelspace")
                        Freeze_operations = False
                        Exit Sub
                    End If




                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                    If ListBox_DWG.Items.Count > 0 Then
                        Using lock1 As DocumentLock = ThisDrawing.LockDocument

                            Using Trans11 As Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                                For i = 0 To ListBox_DWG.Items.Count - 1
                                    Dim Drawing1 As String = ListBox_DWG.Items(i)
                                    If IO.File.Exists(Drawing1) = True Then

                                        Dim Colectie_nume_atribute As New Specialized.StringCollection
                                        Dim Colectie_valori_atribute As New Specialized.StringCollection



                                        Dim Database1 As New Database(False, True)

                                        Try
                                            Try
                                                Try
                                                    Database1.ReadDwgFile(Drawing1, IO.FileShare.ReadWrite, True, Nothing)
                                                Catch ex As Exception
                                                    MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                    GoTo 123
                                                End Try
                                            Catch ex As IO.IOException
                                                MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                Error_list = Error_list & vbCrLf & Drawing1
                                                GoTo 123
                                            End Try
                                        Catch ex As System.SystemException
                                            MsgBox(Drawing1 & vbCrLf & "could not be open")
                                            Error_list = Error_list & vbCrLf & Drawing1
                                            GoTo 123
                                        End Try


                                        HostApplicationServices.WorkingDatabase = Database1



                                        Dim is_erased As Boolean = False

                                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction

                                            Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                                            Dim Layoutdict As DBDictionary

                                            Layoutdict = Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead)

                                            For Each entry As DBDictionaryEntry In Layoutdict
                                                Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                                If Layout1.TabOrder >= CInt(TextBox_Layout_index_START.Text) And Layout1.TabOrder <= CInt(TextBox_Layout_index_END.Text) And Layout1.TabOrder > 0 Then
                                                    Dim BTrecord As BlockTableRecord = Trans1.GetObject(Layout1.BlockTableRecordId, OpenMode.ForWrite)
                                                    For Each Id1 As ObjectId In BTrecord
                                                        Dim Ent1 As Entity = TryCast(Trans1.GetObject(Id1, OpenMode.ForRead), Entity)
                                                        If IsNothing(Ent1) = False Then
                                                            If TypeOf (Ent1) Is Viewport Then
                                                                Dim Viewport1 As Viewport = TryCast(Ent1, Viewport)
                                                                If IsNothing(Viewport1) = False Then

                                                                    If Viewport1.CenterPoint.X > x1 And Viewport1.CenterPoint.X < x2 And Viewport1.CenterPoint.Y > y1 And Viewport1.CenterPoint.Y < y2 Then
                                                                        Viewport1.UpgradeOpen()
                                                                        Viewport1.Erase()
                                                                        is_erased = True
                                                                    End If

                                                                End If


                                                            End If

                                                            Dim mtext1 As MText = TryCast(Ent1, MText)
                                                            If IsNothing(mtext1) = False Then

                                                                If mtext1.Location.X > x1 And mtext1.Location.X < x2 And mtext1.Location.Y > y1 And mtext1.Location.Y < y2 Then

                                                                    mtext1.UpgradeOpen()
                                                                    mtext1.Erase()
                                                                    is_erased = True

                                                                End If

                                                            End If

                                                            Dim text1 As DBText = TryCast(Ent1, DBText)
                                                            If IsNothing(text1) = False Then

                                                                If text1.Position.X > x1 And text1.Position.X < x2 And text1.Position.Y > y1 And text1.Position.Y < y2 Then

                                                                    text1.UpgradeOpen()
                                                                    text1.Erase()
                                                                    is_erased = True

                                                                End If

                                                            End If

                                                            Dim LINE1 As Line = TryCast(Ent1, Line)
                                                            If IsNothing(LINE1) = False Then
                                                                Dim ptx1 As Double = LINE1.StartPoint.X
                                                                Dim ptx2 As Double = LINE1.EndPoint.X
                                                                Dim pty1 As Double = LINE1.StartPoint.Y
                                                                Dim pty2 As Double = LINE1.EndPoint.Y

                                                                If ptx1 > x1 And ptx1 < x2 And pty1 > y1 And pty1 < y2 And ptx2 > x1 And ptx2 < x2 And pty2 > y1 And pty2 < y2 Then
                                                                    LINE1.UpgradeOpen()
                                                                    LINE1.Erase()
                                                                    is_erased = True
                                                                End If

                                                            End If

                                                            Dim Pline1 As Polyline = TryCast(Ent1, Polyline)
                                                            If IsNothing(Pline1) = False Then
                                                                Dim ptx1 As Double = Pline1.StartPoint.X
                                                                Dim ptx2 As Double = Pline1.EndPoint.X
                                                                Dim pty1 As Double = Pline1.StartPoint.Y
                                                                Dim pty2 As Double = Pline1.EndPoint.Y

                                                                If ptx1 > x1 And ptx1 < x2 And pty1 > y1 And pty1 < y2 And ptx2 > x1 And ptx2 < x2 And pty2 > y1 And pty2 < y2 Then
                                                                    Pline1.UpgradeOpen()
                                                                    Pline1.Erase()
                                                                    is_erased = True
                                                                End If

                                                            End If

                                                            Dim Block1 As BlockReference = TryCast(Ent1, BlockReference)
                                                            If IsNothing(Block1) = False Then
                                                                Dim ptx1 As Double = Block1.Position.X

                                                                Dim pty1 As Double = Block1.Position.Y


                                                                If ptx1 > x1 And ptx1 < x2 And pty1 > y1 And pty1 < y2 Then
                                                                    Block1.UpgradeOpen()
                                                                    Block1.Erase()
                                                                    is_erased = True
                                                                End If

                                                            End If

                                                        End If
                                                    Next
                                                End If

                                            Next

                                            If is_erased = True Then
                                                Trans1.Commit()

                                                Try
                                                    Database1.SaveAs(Drawing1, True, DwgVersion.Current, Database1.SecurityParameters)
                                                Catch ex As Exception
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                End Try
                                            End If

                                        End Using

                                        Database1.Dispose()

                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database
                                    End If

123:
                                Next

                                Trans11.Commit()
                            End Using
                        End Using
                    End If


                    If Not Error_list = "" Then
                        MsgBox("There are not saved files. The list of these files is loaded into the clipboard" & vbCrLf &
                               "Use paste <Ctrl+V> command in a text editor program (notepad, excel, word) to see the list")
                        My.Computer.Clipboard.SetText(Error_list)

                    End If


                Catch ex As Exception
                    MsgBox(ex.Message)

                End Try

            Catch ex As System.SystemException
                MsgBox("1: " & ex.Message)
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: DONE!")
            Freeze_operations = False

        End If
    End Sub

    Private Sub Button_delete_material_qty_Click(sender As Object, e As EventArgs) Handles Button_delete_material_qty.Click


        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Error_list As String = ""

            Dim x1 As Double = 732.498
            Dim y1 As Double = 51.5788
            Dim x2 As Double = 1207.498
            Dim y2 As Double = 203.001

            Try

                Try

                    If IsNumeric(TextBox_Layout_index_START.Text) = False Then
                        MsgBox("Please specify the layout START INDEX!")
                        Freeze_operations = False
                        Exit Sub
                    End If
                    If IsNumeric(TextBox_Layout_index_END.Text) = False Then
                        MsgBox("Please specify the layout END INDEX!")
                        Freeze_operations = False
                        Exit Sub
                    End If

                    If CInt(TextBox_Layout_index_START.Text) = 0 Or CInt(TextBox_Layout_index_END.Text) = 0 Then
                        MsgBox("the layout 0 = modelspace")
                        Freeze_operations = False
                        Exit Sub
                    End If




                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                    If ListBox_DWG.Items.Count > 0 Then
                        Using lock1 As DocumentLock = ThisDrawing.LockDocument

                            Using Trans11 As Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                                For i = 0 To ListBox_DWG.Items.Count - 1
                                    Dim Drawing1 As String = ListBox_DWG.Items(i)
                                    If IO.File.Exists(Drawing1) = True Then

                                        Dim Colectie_nume_atribute As New Specialized.StringCollection
                                        Dim Colectie_valori_atribute As New Specialized.StringCollection



                                        Dim Database1 As New Database(False, True)

                                        Try
                                            Try
                                                Try
                                                    Database1.ReadDwgFile(Drawing1, IO.FileShare.ReadWrite, True, Nothing)
                                                Catch ex As Exception
                                                    MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                    GoTo 123
                                                End Try
                                            Catch ex As IO.IOException
                                                MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                Error_list = Error_list & vbCrLf & Drawing1
                                                GoTo 123
                                            End Try
                                        Catch ex As System.SystemException
                                            MsgBox(Drawing1 & vbCrLf & "could not be open")
                                            Error_list = Error_list & vbCrLf & Drawing1
                                            GoTo 123
                                        End Try


                                        HostApplicationServices.WorkingDatabase = Database1



                                        Dim is_erased As Boolean = False

                                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction

                                            Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                                            Dim Layoutdict As DBDictionary

                                            Layoutdict = Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead)

                                            For Each entry As DBDictionaryEntry In Layoutdict
                                                Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                                If Layout1.TabOrder >= CInt(TextBox_Layout_index_START.Text) And Layout1.TabOrder <= CInt(TextBox_Layout_index_END.Text) And Layout1.TabOrder > 0 Then
                                                    Dim BTrecord As BlockTableRecord = Trans1.GetObject(Layout1.BlockTableRecordId, OpenMode.ForWrite)
                                                    For Each Id1 As ObjectId In BTrecord
                                                        Dim Ent1 As Entity = TryCast(Trans1.GetObject(Id1, OpenMode.ForRead), Entity)
                                                        If IsNothing(Ent1) = False Then
                                                            If TypeOf (Ent1) Is Viewport Then
                                                                Dim Viewport1 As Viewport = TryCast(Ent1, Viewport)
                                                                If IsNothing(Viewport1) = False Then

                                                                    If Viewport1.CenterPoint.X > x1 And Viewport1.CenterPoint.X < x2 And Viewport1.CenterPoint.Y > y1 And Viewport1.CenterPoint.Y < y2 Then
                                                                        Viewport1.UpgradeOpen()
                                                                        Viewport1.Erase()
                                                                        is_erased = True
                                                                    End If

                                                                End If


                                                            End If

                                                            Dim mtext1 As MText = TryCast(Ent1, MText)
                                                            If IsNothing(mtext1) = False Then

                                                                If mtext1.Location.X > x1 And mtext1.Location.X < x2 And mtext1.Location.Y > y1 And mtext1.Location.Y < y2 Then

                                                                    mtext1.UpgradeOpen()
                                                                    mtext1.Erase()
                                                                    is_erased = True

                                                                End If

                                                            End If

                                                            Dim text1 As DBText = TryCast(Ent1, DBText)
                                                            If IsNothing(text1) = False Then

                                                                If text1.Position.X > x1 And text1.Position.X < x2 And text1.Position.Y > y1 And text1.Position.Y < y2 Then

                                                                    text1.UpgradeOpen()
                                                                    text1.Erase()
                                                                    is_erased = True

                                                                End If

                                                            End If

                                                            Dim LINE1 As Line = TryCast(Ent1, Line)
                                                            If IsNothing(LINE1) = False Then
                                                                Dim ptx1 As Double = LINE1.StartPoint.X
                                                                Dim ptx2 As Double = LINE1.EndPoint.X
                                                                Dim pty1 As Double = LINE1.StartPoint.Y
                                                                Dim pty2 As Double = LINE1.EndPoint.Y

                                                                If ptx1 > x1 And ptx1 < x2 And pty1 > y1 And pty1 < y2 And ptx2 > x1 And ptx2 < x2 And pty2 > y1 And pty2 < y2 Then
                                                                    LINE1.UpgradeOpen()
                                                                    LINE1.Erase()
                                                                    is_erased = True
                                                                End If

                                                            End If

                                                            Dim Pline1 As Polyline = TryCast(Ent1, Polyline)
                                                            If IsNothing(Pline1) = False Then
                                                                Dim ptx1 As Double = Pline1.StartPoint.X
                                                                Dim ptx2 As Double = Pline1.EndPoint.X
                                                                Dim pty1 As Double = Pline1.StartPoint.Y
                                                                Dim pty2 As Double = Pline1.EndPoint.Y

                                                                If ptx1 > x1 And ptx1 < x2 And pty1 > y1 And pty1 < y2 And ptx2 > x1 And ptx2 < x2 And pty2 > y1 And pty2 < y2 Then
                                                                    Pline1.UpgradeOpen()
                                                                    Pline1.Erase()
                                                                    is_erased = True
                                                                End If

                                                            End If


                                                        End If
                                                    Next
                                                End If

                                            Next

                                            If is_erased = True Then
                                                Trans1.Commit()

                                                Try
                                                    Database1.SaveAs(Drawing1, True, DwgVersion.Current, Database1.SecurityParameters)
                                                Catch ex As Exception
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                End Try
                                            End If

                                        End Using

                                        Database1.Dispose()

                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database
                                    End If

123:
                                Next

                                Trans11.Commit()
                            End Using
                        End Using
                    End If


                    If Not Error_list = "" Then
                        MsgBox("There are not saved files. The list of these files is loaded into the clipboard" & vbCrLf &
                               "Use paste <Ctrl+V> command in a text editor program (notepad, excel, word) to see the list")
                        My.Computer.Clipboard.SetText(Error_list)

                    End If


                Catch ex As Exception
                    MsgBox(ex.Message)

                End Try

            Catch ex As System.SystemException
                MsgBox("1: " & ex.Message)
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: DONE!")
            Freeze_operations = False

        End If
    End Sub

    Private Sub Button_access_road__references_Click(sender As Object, e As EventArgs) Handles Button_access_road__references.Click


        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Error_list As String = ""

            Dim x1 As Double = 10.7
            Dim y1 As Double = 0.5
            Dim x2 As Double = 16.825
            Dim y2 As Double = 2.03

            Try

                Try

                    If IsNumeric(TextBox_Layout_index_START.Text) = False Then
                        MsgBox("Please specify the layout START INDEX!")
                        Freeze_operations = False
                        Exit Sub
                    End If
                    If IsNumeric(TextBox_Layout_index_END.Text) = False Then
                        MsgBox("Please specify the layout END INDEX!")
                        Freeze_operations = False
                        Exit Sub
                    End If

                    If CInt(TextBox_Layout_index_START.Text) = 0 Or CInt(TextBox_Layout_index_END.Text) = 0 Then
                        MsgBox("the layout 0 = modelspace")
                        Freeze_operations = False
                        Exit Sub
                    End If




                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                    If ListBox_DWG.Items.Count > 0 Then
                        Using lock1 As DocumentLock = ThisDrawing.LockDocument

                            Using Trans11 As Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                                For i = 0 To ListBox_DWG.Items.Count - 1
                                    Dim Drawing1 As String = ListBox_DWG.Items(i)
                                    If IO.File.Exists(Drawing1) = True Then

                                        Dim Colectie_nume_atribute As New Specialized.StringCollection
                                        Dim Colectie_valori_atribute As New Specialized.StringCollection



                                        Dim Database1 As New Database(False, True)

                                        Try
                                            Try
                                                Try
                                                    Database1.ReadDwgFile(Drawing1, IO.FileShare.ReadWrite, True, Nothing)
                                                Catch ex As Exception
                                                    MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                    GoTo 123
                                                End Try
                                            Catch ex As IO.IOException
                                                MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                Error_list = Error_list & vbCrLf & Drawing1
                                                GoTo 123
                                            End Try
                                        Catch ex As System.SystemException
                                            MsgBox(Drawing1 & vbCrLf & "could not be open")
                                            Error_list = Error_list & vbCrLf & Drawing1
                                            GoTo 123
                                        End Try


                                        HostApplicationServices.WorkingDatabase = Database1



                                        Dim is_erased As Boolean = False

                                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction

                                            Dim Layer_table As Autodesk.AutoCAD.DatabaseServices.LayerTable = Trans1.GetObject(Database1.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                                            Dim Layer_0 As LayerTableRecord = TryCast(Trans1.GetObject(Layer_table("0"), OpenMode.ForRead), LayerTableRecord)
                                            If IsNothing(Layer_0) = False Then
                                                If Not Layer_0.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 7) Then
                                                    Layer_0.UpgradeOpen()
                                                    Layer_0.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 7)
                                                    is_erased = True
                                                End If

                                            End If

                                            Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                                            Dim Layoutdict As DBDictionary

                                            Layoutdict = Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead)

                                            For Each entry As DBDictionaryEntry In Layoutdict
                                                Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                                If Layout1.TabOrder >= CInt(TextBox_Layout_index_START.Text) And Layout1.TabOrder <= CInt(TextBox_Layout_index_END.Text) And Layout1.TabOrder > 0 Then
                                                    Dim BTrecord As BlockTableRecord = Trans1.GetObject(Layout1.BlockTableRecordId, OpenMode.ForWrite)
                                                    For Each Id1 As ObjectId In BTrecord
                                                        Dim Ent1 As Entity = TryCast(Trans1.GetObject(Id1, OpenMode.ForRead), Entity)
                                                        If IsNothing(Ent1) = False Then
                                                            If TypeOf (Ent1) Is Viewport Then
                                                                Dim Viewport1 As Viewport = TryCast(Ent1, Viewport)
                                                                If IsNothing(Viewport1) = False Then

                                                                    If Viewport1.CenterPoint.X > x1 And Viewport1.CenterPoint.X < x2 And Viewport1.CenterPoint.Y > y1 And Viewport1.CenterPoint.Y < y2 Then
                                                                        Viewport1.UpgradeOpen()
                                                                        Viewport1.Erase()
                                                                        is_erased = True
                                                                    End If

                                                                End If


                                                            End If

                                                            Dim mtext1 As MText = TryCast(Ent1, MText)
                                                            If IsNothing(mtext1) = False Then

                                                                If mtext1.Location.X > x1 And mtext1.Location.X < x2 And mtext1.Location.Y > y1 And mtext1.Location.Y < y2 Then

                                                                    mtext1.UpgradeOpen()
                                                                    mtext1.Erase()
                                                                    is_erased = True

                                                                End If

                                                            End If

                                                            Dim text1 As DBText = TryCast(Ent1, DBText)
                                                            If IsNothing(text1) = False Then

                                                                If text1.Position.X > x1 And text1.Position.X < x2 And text1.Position.Y > y1 And text1.Position.Y < y2 Then

                                                                    text1.UpgradeOpen()
                                                                    text1.Erase()
                                                                    is_erased = True

                                                                End If

                                                            End If

                                                            Dim LINE1 As Line = TryCast(Ent1, Line)
                                                            If IsNothing(LINE1) = False Then
                                                                Dim ptx1 As Double = LINE1.StartPoint.X
                                                                Dim ptx2 As Double = LINE1.EndPoint.X
                                                                Dim pty1 As Double = LINE1.StartPoint.Y
                                                                Dim pty2 As Double = LINE1.EndPoint.Y

                                                                If ptx1 > x1 And ptx1 < x2 And pty1 > y1 And pty1 < y2 And ptx2 > x1 And ptx2 < x2 And pty2 > y1 And pty2 < y2 Then
                                                                    LINE1.UpgradeOpen()
                                                                    LINE1.Erase()
                                                                    is_erased = True
                                                                End If

                                                            End If

                                                            Dim Pline1 As Polyline = TryCast(Ent1, Polyline)
                                                            If IsNothing(Pline1) = False Then
                                                                Dim ptx1 As Double = Pline1.StartPoint.X
                                                                Dim ptx2 As Double = Pline1.EndPoint.X
                                                                Dim pty1 As Double = Pline1.StartPoint.Y
                                                                Dim pty2 As Double = Pline1.EndPoint.Y

                                                                If ptx1 > x1 And ptx1 < x2 And pty1 > y1 And pty1 < y2 And ptx2 > x1 And ptx2 < x2 And pty2 > y1 And pty2 < y2 Then
                                                                    Pline1.UpgradeOpen()
                                                                    Pline1.Erase()
                                                                    is_erased = True
                                                                End If

                                                            End If


                                                        End If
                                                    Next
                                                End If

                                            Next

                                            If is_erased = True Then
                                                Trans1.Commit()

                                                Try
                                                    Database1.SaveAs(Drawing1, True, DwgVersion.Current, Database1.SecurityParameters)
                                                Catch ex As Exception
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                End Try
                                            End If

                                        End Using

                                        Database1.Dispose()

                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database
                                    End If

123:
                                Next

                                Trans11.Commit()
                            End Using
                        End Using
                    End If


                    If Not Error_list = "" Then
                        MsgBox("There are not saved files. The list of these files is loaded into the clipboard" & vbCrLf &
                               "Use paste <Ctrl+V> command in a text editor program (notepad, excel, word) to see the list")
                        My.Computer.Clipboard.SetText(Error_list)

                    End If


                Catch ex As Exception
                    MsgBox(ex.Message)

                End Try

            Catch ex As System.SystemException
                MsgBox("1: " & ex.Message)
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: DONE!")
            Freeze_operations = False

        End If
    End Sub
    Private Sub Button_modify_layers_Click(sender As Object, e As EventArgs) Handles Button_modify_layers_peneast.Click


        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Error_list As String = ""


            Try

                Try



                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                    If ListBox_DWG.Items.Count > 0 Then
                        Using lock1 As DocumentLock = ThisDrawing.LockDocument

                            Using Trans11 As Transaction = ThisDrawing.TransactionManager.StartTransaction

                                For i = 0 To ListBox_DWG.Items.Count - 1
                                    Dim Drawing1 As String = ListBox_DWG.Items(i)
                                    If IO.File.Exists(Drawing1) = True Then

                                        Dim Colectie_nume_atribute As New Specialized.StringCollection
                                        Dim Colectie_valori_atribute As New Specialized.StringCollection

                                        Using Database1 As New Database(False, True)
                                            Try
                                                Try
                                                    Try
                                                        Database1.ReadDwgFile(Drawing1, FileOpenMode.OpenForReadAndAllShare, False, "")
                                                        HostApplicationServices.WorkingDatabase = Database1
                                                    Catch ex As Exception
                                                        MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                        Error_list = Error_list & vbCrLf & Drawing1
                                                        GoTo 123
                                                    End Try
                                                Catch ex As IO.IOException
                                                    MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                    GoTo 123
                                                End Try
                                            Catch ex As System.SystemException
                                                MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                Error_list = Error_list & vbCrLf & Drawing1
                                                GoTo 123
                                            End Try





                                            Dim is_erased As Boolean = False


                                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction

                                                Database1.ResolveXrefs(False, True)


                                                Dim Layer_table As Autodesk.AutoCAD.DatabaseServices.LayerTable = Trans1.GetObject(Database1.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                                                Dim Nr1 As Integer = 0
                                                For Each ID1 As ObjectId In Layer_table
                                                    Dim Layer1 As LayerTableRecord = Trans1.GetObject(ID1, OpenMode.ForRead)

                                                    If IsNothing(Layer1) = False Then

                                                        If Layer1.Name.Contains("|E_LS_Rock") = True Then
                                                            Layer1.UpgradeOpen()
                                                            'Layer1.IsFrozen = True
                                                            Layer1.IsOff = True
                                                            is_erased = True
                                                        End If

                                                        If Layer1.Name.Contains("|E_Fea_Wall") = True Then
                                                            Layer1.UpgradeOpen()
                                                            'Layer1.IsFrozen = True
                                                            Layer1.IsOff = True
                                                            is_erased = True
                                                        End If

                                                        Dim Nume1 = "Fea_ErosionCtrl"

                                                        If Layer1.Name.Contains(Nume1) = True And Strings.Right(Layer1.Name, 4) = "Ctrl" Then
                                                            Layer1.UpgradeOpen()
                                                            Layer1.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 10)
                                                            is_erased = True
                                                        End If


                                                    End If
                                                    Nr1 = Nr1 + 1
                                                Next


                                                Dim Name_dict As DBDictionary = Trans1.GetObject(Database1.NamedObjectsDictionaryId, OpenMode.ForRead)
                                                Dim rasterVars As RasterVariables
                                                'get variables dictionary
                                                Dim Image_vars As String = "ACAD_IMAGE_VARS"

                                                If Name_dict.Contains(Image_vars) = True Then

                                                    Dim rastVarsId As ObjectId = Name_dict.GetAt(Image_vars)
                                                    rasterVars = Trans1.GetObject(rastVarsId, OpenMode.ForWrite)
                                                    rasterVars.ImageFrame = FrameSetting.ImageFrameOnNoPlot
                                                    is_erased = True
                                                End If






                                                If is_erased = True Then
                                                    Trans1.Commit()

                                                    Try
                                                        Database1.SaveAs(Drawing1, True, DwgVersion.Current, Database1.SecurityParameters)
                                                    Catch ex As Exception
                                                        Error_list = Error_list & vbCrLf & Drawing1
                                                    End Try
                                                End If

                                            End Using

                                        End Using

                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database
                                    End If

123:
                                Next


                            End Using
                        End Using
                    End If


                    If Not Error_list = "" Then
                        MsgBox("There are not saved files. The list of these files is loaded into the clipboard" & vbCrLf &
                               "Use paste <Ctrl+V> command in a text editor program (notepad, excel, word) to see the list")
                        My.Computer.Clipboard.SetText(Error_list)

                    End If


                Catch ex As Exception
                    MsgBox(ex.Message)

                End Try

            Catch ex As System.SystemException
                MsgBox("1: " & ex.Message)
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: DONE!")
            Freeze_operations = False

        End If
    End Sub

    Private Sub Button_PAS_Click(sender As Object, e As EventArgs) Handles Button_POINT_AT_STATION.Click

        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Empty_array() As ObjectId
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
            Try
                Using lock1 As DocumentLock = ThisDrawing.LockDocument
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
                    Dim Poly2D As Polyline
                    Dim Poly3D As Polyline3d

                    Dim Point_on_poly As New Point3d

                    Dim Dist_from_start_for_zero As Double

                    If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        If IsNothing(Rezultat1) = False Then



                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Rezultat1.Value.Item(0)
                                Dim Ent1 As Entity
                                Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then


                                    Poly2D = Ent1


                                    Dim Point_zero As New Point3d

                                    Point_zero = Poly2D.GetClosestPointTo(Poly2D.StartPoint, Vector3d.ZAxis, False)



                                    Dist_from_start_for_zero = 0

                                    Trans1.Commit()

                                ElseIf TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then


                                    Poly3D = Ent1

                                    Dim Index_Poly As Integer = 0
                                    Poly2D = New Polyline
                                    For Each vId As Autodesk.AutoCAD.DatabaseServices.ObjectId In Poly3D
                                        Dim v3d As Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d = DirectCast(Trans1.GetObject _
                                                (vId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d)

                                        Dim x1 As Double = v3d.Position.X
                                        Dim y1 As Double = v3d.Position.Y
                                        Dim z1 As Double = v3d.Position.Z
                                        Poly2D.AddVertexAt(Index_Poly, New Point2d(x1, y1), 0, 0, 0)
                                        Index_Poly = Index_Poly + 1
                                    Next



                                    Dist_from_start_for_zero = 0

                                    Trans1.Commit()

                                Else
                                    Editor1.WriteMessage("No Polyline")
                                    Editor1.WriteMessage(vbLf & "Command:")
                                    Freeze_operations = False
                                    Exit Sub
                                End If
                            End Using
                        End If
                    End If
1234:
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                        Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Please Pick a point on the same polyline:")
                        Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                        PP1.AllowNone = False
                        Point1 = Editor1.GetPoint(PP1)
                        If Not Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Editor1.SetImpliedSelection(Empty_array)
                            Freeze_operations = False
                            Trans1.Commit()
                            Editor1.SetImpliedSelection(Empty_array)
                            Editor1.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If

                        Dim Distanta_pana_la_xing As Double
                        If IsNothing(Poly2D) = False Then
                            Point_on_poly = Poly2D.GetClosestPointTo(Point1.Value.TransformBy(curent_ucs_matrix), Vector3d.ZAxis, False)
                            Distanta_pana_la_xing = Poly2D.GetDistAtPoint(Point_on_poly)
                        End If

                        If IsNothing(Poly3D) = False Then
                            Dim Point_on_poly2D As New Point3d
                            Point_on_poly2D = Poly2D.GetClosestPointTo(Point1.Value.TransformBy(curent_ucs_matrix), Vector3d.ZAxis, False)
                            Dim param1 As Double = Poly2D.GetParameterAtPoint(Point_on_poly2D)
                            Distanta_pana_la_xing = Poly3D.GetDistanceAtParameter(param1)
                            Point_on_poly = Poly3D.GetPointAtParameter(param1)
                        End If


                        Dim Station1 As Double = Distanta_pana_la_xing - Dist_from_start_for_zero




                        If Dist_from_start_for_zero + Station1 < 0 Then
                            MsgBox("The 0+000 position and your desired station point is not matching.")
                            Freeze_operations = False
                            Exit Sub
                        End If
                        Dim Round1 As Integer = 2
                        If IsNumeric(TextBox_dec1.Text) = True Then
                            Round1 = CInt(TextBox_dec1.Text)
                        End If
                        Dim Chainage_string As String = Get_chainage_feet_from_double(Station1 + Get_equation_value(Station1), Round1)
                        If Chainage_string = "-0+00.00" Then Chainage_string = "0+00.00"

                        Dim Mleader1 As New MLeader

                        If IsNothing(Point_on_poly) = False Then
                            Mleader1 = Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly, Chainage_string, 8, 2.5, 8, 20, 20)
                        End If

                        Trans1.Commit()
                        GoTo 1234

                    End Using
                End Using


            Catch ex As Exception
                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try

        End If

    End Sub

    Private Sub Button_WIPEOUT_OFF1()


        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Error_list As String = ""


            Try

                Try



                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                    If ListBox_DWG.Items.Count > 0 Then
                        Using lock1 As DocumentLock = ThisDrawing.LockDocument

                            Using Trans11 As Transaction = ThisDrawing.TransactionManager.StartTransaction

                                For i = 0 To ListBox_DWG.Items.Count - 1
                                    Dim Drawing1 As String = ListBox_DWG.Items(i)
                                    If IO.File.Exists(Drawing1) = True Then

                                        Dim Colectie_nume_atribute As New Specialized.StringCollection
                                        Dim Colectie_valori_atribute As New Specialized.StringCollection

                                        Using Database1 As New Database(False, True)
                                            Try
                                                Try
                                                    Try
                                                        Database1.ReadDwgFile(Drawing1, FileOpenMode.OpenForReadAndAllShare, False, "")
                                                        HostApplicationServices.WorkingDatabase = Database1
                                                    Catch ex As Exception
                                                        MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                        Error_list = Error_list & vbCrLf & Drawing1
                                                        GoTo 123
                                                    End Try
                                                Catch ex As IO.IOException
                                                    MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                    GoTo 123
                                                End Try
                                            Catch ex As System.SystemException
                                                MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                Error_list = Error_list & vbCrLf & Drawing1
                                                GoTo 123
                                            End Try





                                            Dim is_erased As Boolean = False


                                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction




                                                Dim Name_dict As DBDictionary = Trans1.GetObject(Database1.NamedObjectsDictionaryId, OpenMode.ForRead)
                                                Dim rasterVars As RasterVariables


                                                'get variables dictionary
                                                Dim Image_vars As String = "ACAD_IMAGE_VARS"
                                                Dim Image_vars1 As String = "ACAD_WIPEOUT_VARS"

                                                If Name_dict.Contains(Image_vars) = True Then

                                                    Dim rastVarsId As ObjectId = Name_dict.GetAt(Image_vars)
                                                    rasterVars = Trans1.GetObject(rastVarsId, OpenMode.ForWrite)
                                                    rasterVars.ImageFrame = FrameSetting.ImageFrameOff

                                                    is_erased = True
                                                End If

                                                If Name_dict.Contains(Image_vars1) = True Then

                                                    Dim rastVarsId As ObjectId = Name_dict.GetAt(Image_vars1)

                                                    Dim rasterVars1
                                                    rasterVars1 = Trans1.GetObject(rastVarsId, OpenMode.ForWrite)

                                                    MsgBox(rasterVars1.GetType.GetFields.ToString)

                                                End If




                                                If is_erased = True Then
                                                    Trans1.Commit()

                                                    Try
                                                        Database1.SaveAs(Drawing1, True, DwgVersion.Current, Database1.SecurityParameters)
                                                    Catch ex As Exception
                                                        Error_list = Error_list & vbCrLf & Drawing1
                                                    End Try
                                                End If

                                            End Using

                                        End Using

                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database
                                    End If

123:
                                Next


                            End Using
                        End Using
                    End If


                    If Not Error_list = "" Then
                        MsgBox("There are not saved files. The list of these files is loaded into the clipboard" & vbCrLf &
                               "Use paste <Ctrl+V> command in a text editor program (notepad, excel, word) to see the list")
                        My.Computer.Clipboard.SetText(Error_list)

                    End If


                Catch ex As Exception
                    MsgBox(ex.Message)

                End Try

            Catch ex As System.SystemException
                MsgBox("1: " & ex.Message)
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: DONE!")
            Freeze_operations = False

        End If
    End Sub


    Private Sub Button_ADD_LINE_TO_etc_Click(sender As Object, e As EventArgs) Handles Button_ADD_LINE_TO_etc.Click

        If Freeze_operations = False Then
            Freeze_operations = True




            Try


                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor

                Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)



                        Dim Rezultat_label As Autodesk.AutoCAD.EditorInput.PromptEntityResult

                        Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select a DEFLECTION label:")

                        Object_Prompt1.SetRejectMessage(vbLf & "Please select a an mtext object")
                        Object_Prompt1.AddAllowedClass(GetType(MText), True)

                        Rezultat_label = Editor1.GetEntity(Object_Prompt1)


                        If Not Rezultat_label.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            MsgBox("NO LABEL")
                            Editor1.WriteMessage(vbLf & "Command:")
                            Freeze_operations = False
                            Editor1.SetImpliedSelection(Empty_array)
                            Exit Sub
                        End If
                        Dim Mtext_reference As MText = TryCast(Trans1.GetObject(Rezultat_label.ObjectId, OpenMode.ForRead), MText)




                        If IsNothing(Mtext_reference) = False Then
                            For Each id1 As ObjectId In BTrecord

                                Dim Mtext_defl As MText = TryCast(Trans1.GetObject(id1, OpenMode.ForWrite), MText)
                                If IsNothing(Mtext_defl) = False Then
                                    If Mtext_defl.Layer = Mtext_reference.Layer Then
                                        Mtext_defl.Contents = Mtext_defl.Contents.Replace("\L", "")
                                        Dim Poly1 As New Polyline
                                        Poly1.Layer = Mtext_reference.Layer
                                        Poly1.AddVertexAt(0, New Point2d(Mtext_defl.Location.X + 2, Mtext_defl.Location.Y - 46.7232), 0, 0, 0)
                                        Poly1.AddVertexAt(1, New Point2d(Mtext_defl.Location.X + 2, Mtext_defl.Location.Y + 653.2768), 0, 0, 0)
                                        Poly1.Elevation = 0
                                        BTrecord.AppendEntity(Poly1)
                                        Trans1.AddNewlyCreatedDBObject(Poly1, True)
                                    Else
                                        Mtext_defl.Contents = Mtext_defl.Contents.Replace(" C\L ", " C\\L ")
                                    End If



                                End If

                            Next



                        End If



                        Trans1.Commit()
                    End Using
                End Using


                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: ")




            Catch ex As Exception
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_del_graph_viewport_spire_Click(sender As Object, e As EventArgs) Handles Button_del_graph_viewport_spire.Click

        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Error_list As String = ""

            Dim x1 As Double = 167
            Dim x2 As Double = 7082
            Dim y1 As Double = 748
            Dim y2 As Double = 1537

            Try

                Try

                    If IsNumeric(TextBox_Layout_index_START.Text) = False Then
                        MsgBox("Please specify the layout START INDEX!")
                        Freeze_operations = False
                        Exit Sub
                    End If
                    If IsNumeric(TextBox_Layout_index_END.Text) = False Then
                        MsgBox("Please specify the layout END INDEX!")
                        Freeze_operations = False
                        Exit Sub
                    End If

                    If CInt(TextBox_Layout_index_START.Text) = 0 Or CInt(TextBox_Layout_index_END.Text) = 0 Then
                        MsgBox("the layout 0 = modelspace")
                        Freeze_operations = False
                        Exit Sub
                    End If




                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                    If ListBox_DWG.Items.Count > 0 Then
                        Using lock1 As DocumentLock = ThisDrawing.LockDocument

                            Using Trans11 As Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                                For i = 0 To ListBox_DWG.Items.Count - 1
                                    Dim Drawing1 As String = ListBox_DWG.Items(i)
                                    If IO.File.Exists(Drawing1) = True Then

                                        Dim Colectie_nume_atribute As New Specialized.StringCollection
                                        Dim Colectie_valori_atribute As New Specialized.StringCollection



                                        Dim Database1 As New Database(False, True)

                                        Try
                                            Try
                                                Try
                                                    Database1.ReadDwgFile(Drawing1, IO.FileShare.ReadWrite, True, Nothing)
                                                Catch ex As Exception
                                                    MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                    GoTo 123
                                                End Try
                                            Catch ex As IO.IOException
                                                MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                Error_list = Error_list & vbCrLf & Drawing1
                                                GoTo 123
                                            End Try
                                        Catch ex As System.SystemException
                                            MsgBox(Drawing1 & vbCrLf & "could not be open")
                                            Error_list = Error_list & vbCrLf & Drawing1
                                            GoTo 123
                                        End Try


                                        HostApplicationServices.WorkingDatabase = Database1



                                        Dim is_erased As Boolean = False

                                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction

                                            Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                                            Dim Layoutdict As DBDictionary

                                            Layoutdict = Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead)

                                            For Each entry As DBDictionaryEntry In Layoutdict
                                                Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                                If Layout1.TabOrder >= CInt(TextBox_Layout_index_START.Text) And Layout1.TabOrder <= CInt(TextBox_Layout_index_END.Text) And Layout1.TabOrder > 0 Then
                                                    Dim BTrecord As BlockTableRecord = Trans1.GetObject(Layout1.BlockTableRecordId, OpenMode.ForWrite)


                                                    For Each Id1 As ObjectId In BTrecord
                                                        Dim Ent1 As Entity = TryCast(Trans1.GetObject(Id1, OpenMode.ForRead), Entity)
                                                        If IsNothing(Ent1) = False Then


                                                            Dim text1 As DBText = TryCast(Ent1, DBText)
                                                            If IsNothing(text1) = False Then
                                                                If text1.Position.X > x1 And text1.Position.X < x2 And text1.Position.Y > y1 And text1.Position.Y < y2 Then
                                                                    If text1.TextString.Contains("STA") = True Then
                                                                        text1.UpgradeOpen()
                                                                        text1.Erase()
                                                                        is_erased = True
                                                                    End If
                                                                End If
                                                            End If
                                                            If IsNothing(text1) = False Then
                                                                If text1.Position.X > x1 And text1.Position.X < x2 And text1.Position.Y > y1 And text1.Position.Y < y2 Then
                                                                    If text1.TextString.Contains("STN") = True Then
                                                                        text1.UpgradeOpen()
                                                                        text1.Erase()
                                                                        is_erased = True
                                                                    End If
                                                                End If
                                                            End If

                                                            Dim mtext1 As MText = TryCast(Ent1, MText)
                                                            If IsNothing(mtext1) = False Then
                                                                If mtext1.Location.X > x1 And mtext1.Location.X < x2 And mtext1.Location.Y > y1 And mtext1.Location.Y < y2 Then
                                                                    If mtext1.Contents.Contains("STA") = True Then
                                                                        mtext1.UpgradeOpen()
                                                                        mtext1.Erase()
                                                                        is_erased = True
                                                                    End If
                                                                End If
                                                            End If

                                                            If IsNothing(mtext1) = False Then
                                                                If mtext1.Location.X > x1 And mtext1.Location.X < x2 And mtext1.Location.Y > y1 And mtext1.Location.Y < y2 Then
                                                                    If mtext1.Contents.Contains("STN") = True Then
                                                                        mtext1.UpgradeOpen()
                                                                        mtext1.Erase()
                                                                        is_erased = True
                                                                    End If
                                                                End If
                                                            End If

                                                            Dim Poly1 As Polyline = TryCast(Ent1, Polyline)
                                                            If IsNothing(Poly1) = False And Ent1.Layer = "Pipe" Then
                                                                If Poly1.StartPoint.X > x1 And Poly1.StartPoint.X < x2 And Poly1.StartPoint.Y > y1 And Poly1.StartPoint.Y < y2 And
                                                                    Poly1.EndPoint.X > x1 And Poly1.EndPoint.X < x2 And Poly1.EndPoint.Y > y1 And Poly1.EndPoint.Y < y2 Then
                                                                    Poly1.UpgradeOpen()
                                                                    Poly1.Erase()
                                                                    is_erased = True
                                                                End If
                                                            End If

                                                            Dim Block1 As BlockReference = TryCast(Ent1, BlockReference)
                                                            If IsNothing(Block1) = False Then
                                                                If Block1.Name = "MATCHLINE_MP_Sheet" Or Block1.Name = "Spire_Matchline_Left" Or Block1.Name = "Spire_Matchline_Right" Then
                                                                    If Block1.Position.X > x1 And Block1.Position.X < x2 And Block1.Position.Y > y1 And Block1.Position.Y < y2 Then
                                                                        Block1.UpgradeOpen()
                                                                        Block1.Erase()
                                                                        is_erased = True
                                                                    End If

                                                                End If
                                                            End If

                                                            If TypeOf (Ent1) Is Viewport Then
                                                                Dim Viewport1 As Viewport = TryCast(Ent1, Viewport)
                                                                If IsNothing(Viewport1) = False Then

                                                                    Dim x0 As Double = Viewport1.CenterPoint.X
                                                                    Dim y0 As Double = Viewport1.CenterPoint.Y

                                                                    If x0 > x1 And x0 < x2 And y0 > y1 And y0 < y2 Then

                                                                        Viewport1.UpgradeOpen()
                                                                        Viewport1.Erase()
                                                                        is_erased = True
                                                                    End If



                                                                End If
                                                            End If


                                                        End If
                                                    Next
                                                End If

                                            Next

                                            If is_erased = True Then
                                                Trans1.Commit()

                                                Try
                                                    Database1.SaveAs(Drawing1, True, DwgVersion.Current, Database1.SecurityParameters)
                                                Catch ex As Exception
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                End Try
                                            End If

                                        End Using

                                        Database1.Dispose()

                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database
                                    End If

123:
                                Next

                                Trans11.Commit()
                            End Using
                        End Using
                    End If


                    If Not Error_list = "" Then
                        MsgBox("There are not saved files. The list of these files is loaded into the clipboard" & vbCrLf &
                               "Use paste <Ctrl+V> command in a text editor program (notepad, excel, word) to see the list")
                        My.Computer.Clipboard.SetText(Error_list)

                    End If


                Catch ex As Exception
                    MsgBox(ex.Message)

                End Try

            Catch ex As System.SystemException
                MsgBox("1: " & ex.Message)
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: DONE!")
            Freeze_operations = False

        End If
    End Sub

    Private Sub Button_del_50_mtext_Click(sender As Object, e As EventArgs) Handles Button_duke_energy.Click

        'G:\DukeEnergy\398992_Line_449_Pipeline\Pipeline\Drafting\Alignments\Preliminary
        'G:\DukeEnergy\398992_Line_449_Pipeline\Pipeline\Drafting\Alignments\Preliminary\E&S Control
        'G:\DukeEnergy\398994_Line_448_Pipeline\Pipeline\Drafting\Alignments\Preliminary
        'G:\DukeEnergy\398994_Line_448_Pipeline\Pipeline\Drafting\Alignments\Preliminary\E&S Control


        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Error_list As String = ""

            Dim x1 As Double = 167
            Dim x2 As Double = 7082
            Dim y1 As Double = 748
            Dim y2 As Double = 1537

            Try

                Try

                    If IsNumeric(TextBox_Layout_index_START.Text) = False Then
                        MsgBox("Please specify the layout START INDEX!")
                        Freeze_operations = False
                        Exit Sub
                    End If
                    If IsNumeric(TextBox_Layout_index_END.Text) = False Then
                        MsgBox("Please specify the layout END INDEX!")
                        Freeze_operations = False
                        Exit Sub
                    End If

                    If CInt(TextBox_Layout_index_START.Text) = 0 Or CInt(TextBox_Layout_index_END.Text) = 0 Then
                        MsgBox("the layout 0 = modelspace")
                        Freeze_operations = False
                        Exit Sub
                    End If




                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                    If ListBox_DWG.Items.Count > 0 Then
                        Using lock1 As DocumentLock = ThisDrawing.LockDocument

                            Using Trans11 As Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                                For i = 0 To ListBox_DWG.Items.Count - 1
                                    Dim Drawing1 As String = ListBox_DWG.Items(i)
                                    If IO.File.Exists(Drawing1) = True Then

                                        Dim Colectie_nume_atribute As New Specialized.StringCollection
                                        Dim Colectie_valori_atribute As New Specialized.StringCollection



                                        Dim Database1 As New Database(False, True)

                                        Try
                                            Try
                                                Try
                                                    Database1.ReadDwgFile(Drawing1, IO.FileShare.ReadWrite, True, Nothing)
                                                Catch ex As Exception
                                                    MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                    GoTo 123
                                                End Try
                                            Catch ex As IO.IOException
                                                MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                Error_list = Error_list & vbCrLf & Drawing1
                                                GoTo 123
                                            End Try
                                        Catch ex As System.SystemException
                                            MsgBox(Drawing1 & vbCrLf & "could not be open")
                                            Error_list = Error_list & vbCrLf & Drawing1
                                            GoTo 123
                                        End Try


                                        HostApplicationServices.WorkingDatabase = Database1



                                        Dim is_erased As Boolean = False

                                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction

                                            Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                                            Dim Layoutdict As DBDictionary

                                            Layoutdict = Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead)

                                            For Each entry As DBDictionaryEntry In Layoutdict
                                                Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                                If Layout1.TabOrder >= CInt(TextBox_Layout_index_START.Text) And Layout1.TabOrder <= CInt(TextBox_Layout_index_END.Text) And Layout1.TabOrder > 0 Then
                                                    Dim BTrecord As BlockTableRecord = Trans1.GetObject(Layout1.BlockTableRecordId, OpenMode.ForWrite)


                                                    For Each Id1 As ObjectId In BTrecord
                                                        Dim Ent1 As Entity = TryCast(Trans1.GetObject(Id1, OpenMode.ForRead), Entity)
                                                        If IsNothing(Ent1) = False Then
                                                            Dim text1 As DBText = TryCast(Ent1, DBText)
                                                            If IsNothing(text1) = False Then
                                                                If text1.Layer.ToUpper = "G-ANNO-TEXT" Then
                                                                    If text1.TextString.Contains("50'") = True Then
                                                                        text1.UpgradeOpen()
                                                                        text1.Erase()
                                                                        is_erased = True
                                                                    End If
                                                                End If
                                                            End If

                                                            Dim mtext1 As MText = TryCast(Ent1, MText)
                                                            If IsNothing(mtext1) = False Then
                                                                If mtext1.Layer.ToUpper = "G-ANNO-TEXT" Then
                                                                    If mtext1.Text.Contains("50'") = True Then
                                                                        mtext1.UpgradeOpen()
                                                                        mtext1.Erase()
                                                                        is_erased = True
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    Next
                                                End If
                                            Next

                                            If is_erased = True Then
                                                Trans1.Commit()

                                                Try
                                                    Database1.SaveAs(Drawing1, True, DwgVersion.Current, Database1.SecurityParameters)
                                                Catch ex As Exception
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                End Try
                                            End If

                                        End Using

                                        Database1.Dispose()

                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database
                                    End If

123:
                                Next

                                Trans11.Commit()
                            End Using
                        End Using
                    End If


                    If Not Error_list = "" Then
                        MsgBox("There are not saved files. The list of these files is loaded into the clipboard" & vbCrLf &
                               "Use paste <Ctrl+V> command in a text editor program (notepad, excel, word) to see the list")
                        My.Computer.Clipboard.SetText(Error_list)

                    End If


                Catch ex As Exception
                    MsgBox(ex.Message)

                End Try

            Catch ex As System.SystemException
                MsgBox("1: " & ex.Message)
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: DONE!")
            Freeze_operations = False

        End If
    End Sub

    Private Sub Button_DELETE_VIEWPORTS_SPIRE_Click(sender As Object, e As EventArgs) Handles Button_DELETE_VIEWPORTS_SPIRE.Click



        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Error_list As String = ""

            Dim x1 As Double = 148.48
            Dim x2 As Double = 3515.19
            Dim y1 As Double = 252.3
            Dim y2 As Double = 796.28

            Try

                Try

                    If IsNumeric(TextBox_Layout_index_START.Text) = False Then
                        MsgBox("Please specify the layout START INDEX!")
                        Freeze_operations = False
                        Exit Sub
                    End If
                    If IsNumeric(TextBox_Layout_index_END.Text) = False Then
                        MsgBox("Please specify the layout END INDEX!")
                        Freeze_operations = False
                        Exit Sub
                    End If

                    If CInt(TextBox_Layout_index_START.Text) = 0 Or CInt(TextBox_Layout_index_END.Text) = 0 Then
                        MsgBox("the layout 0 = modelspace")
                        Freeze_operations = False
                        Exit Sub
                    End If




                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                    If ListBox_DWG.Items.Count > 0 Then
                        Using lock1 As DocumentLock = ThisDrawing.LockDocument

                            Using Trans11 As Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                                For i = 0 To ListBox_DWG.Items.Count - 1
                                    Dim Drawing1 As String = ListBox_DWG.Items(i)
                                    If IO.File.Exists(Drawing1) = True Then

                                        Dim Colectie_nume_atribute As New Specialized.StringCollection
                                        Dim Colectie_valori_atribute As New Specialized.StringCollection



                                        Dim Database1 As New Database(False, True)

                                        Try
                                            Try
                                                Try
                                                    Database1.ReadDwgFile(Drawing1, IO.FileShare.ReadWrite, True, Nothing)
                                                Catch ex As Exception
                                                    MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                    GoTo 123
                                                End Try
                                            Catch ex As IO.IOException
                                                MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                Error_list = Error_list & vbCrLf & Drawing1
                                                GoTo 123
                                            End Try
                                        Catch ex As System.SystemException
                                            MsgBox(Drawing1 & vbCrLf & "could not be open")
                                            Error_list = Error_list & vbCrLf & Drawing1
                                            GoTo 123
                                        End Try


                                        HostApplicationServices.WorkingDatabase = Database1



                                        Dim is_erased As Boolean = False

                                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction

                                            Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                                            Dim Layoutdict As DBDictionary

                                            Layoutdict = Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead)

                                            For Each entry As DBDictionaryEntry In Layoutdict
                                                Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                                If Layout1.TabOrder >= CInt(TextBox_Layout_index_START.Text) And Layout1.TabOrder <= CInt(TextBox_Layout_index_END.Text) And Layout1.TabOrder > 0 Then
                                                    Dim BTrecord As BlockTableRecord = Trans1.GetObject(Layout1.BlockTableRecordId, OpenMode.ForWrite)
                                                    For Each Id1 As ObjectId In BTrecord
                                                        Dim Ent1 As Entity = TryCast(Trans1.GetObject(Id1, OpenMode.ForRead), Entity)
                                                        If IsNothing(Ent1) = False Then


                                                            If TypeOf (Ent1) Is Viewport Then
                                                                Dim Viewport1 As Viewport = TryCast(Ent1, Viewport)
                                                                If IsNothing(Viewport1) = False Then

                                                                    If Viewport1.CenterPoint.X > x1 And Viewport1.CenterPoint.X < x2 And Viewport1.CenterPoint.Y > y1 And Viewport1.CenterPoint.Y < y2 Then
                                                                        Viewport1.UpgradeOpen()
                                                                        Viewport1.Erase()
                                                                        is_erased = True
                                                                    End If

                                                                End If


                                                            End If

                                                            If TypeOf (Ent1) Is MText Then
                                                                Dim mt As MText = TryCast(Ent1, MText)
                                                                If IsNothing(mt) = False Then

                                                                    If mt.Location.X > x1 And mt.Location.X < x2 And mt.Location.Y > y1 And mt.Location.Y < y2 And mt.Text.Contains("ELEVATION") = False Then
                                                                        mt.UpgradeOpen()
                                                                        mt.Erase()
                                                                        is_erased = True
                                                                    End If

                                                                End If


                                                            End If
                                                            If TypeOf (Ent1) Is DBText Then
                                                                Dim txt As DBText = TryCast(Ent1, DBText)
                                                                If IsNothing(txt) = False Then

                                                                    If txt.Position.X > x1 And txt.Position.X < x2 And txt.Position.Y > y1 And txt.Position.Y < y2 And txt.TextString.Contains("ELEVATION") = False Then
                                                                        txt.UpgradeOpen()
                                                                        txt.Erase()
                                                                        is_erased = True
                                                                    End If

                                                                End If


                                                            End If


                                                        End If
                                                    Next
                                                End If

                                            Next

                                            If is_erased = True Then
                                                Trans1.Commit()

                                                Try
                                                    Database1.SaveAs(Drawing1, True, DwgVersion.Current, Database1.SecurityParameters)
                                                Catch ex As Exception
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                End Try
                                            End If

                                        End Using

                                        Database1.Dispose()

                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database
                                    End If

123:
                                Next

                                Trans11.Commit()
                            End Using
                        End Using
                    End If


                    If Not Error_list = "" Then
                        MsgBox("There are not saved files. The list of these files is loaded into the clipboard" & vbCrLf &
                               "Use paste <Ctrl+V> command in a text editor program (notepad, excel, word) to see the list")
                        My.Computer.Clipboard.SetText(Error_list)

                    End If


                Catch ex As Exception
                    MsgBox(ex.Message)

                End Try

            Catch ex As System.SystemException
                MsgBox("1: " & ex.Message)
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: DONE!")
            Freeze_operations = False

        End If
    End Sub



    Private Sub Button_syncronize_bands_Click(sender As Object, e As EventArgs) Handles Button_adjust_bands.Click





        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Error_list As String = ""

            Dim x1 As Double = 3630
            Dim x2 As Double = 3650
            Dim y1 As Double = 4590
            Dim y2 As Double = 4610

            Try

                Try

                    If IsNumeric(TextBox_Layout_index_START.Text) = False Then
                        MsgBox("Please specify the layout START INDEX!")
                        Freeze_operations = False
                        Exit Sub
                    End If
                    If IsNumeric(TextBox_Layout_index_END.Text) = False Then
                        MsgBox("Please specify the layout END INDEX!")
                        Freeze_operations = False
                        Exit Sub
                    End If

                    If CInt(TextBox_Layout_index_START.Text) = 0 Or CInt(TextBox_Layout_index_END.Text) = 0 Then
                        MsgBox("the layout 0 = modelspace")
                        Freeze_operations = False
                        Exit Sub
                    End If




                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                    If ListBox_DWG.Items.Count > 0 Then
                        Using lock1 As DocumentLock = ThisDrawing.LockDocument

                            Using Trans11 As Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                                For i = 0 To ListBox_DWG.Items.Count - 1
                                    Dim Drawing1 As String = ListBox_DWG.Items(i)
                                    If IO.File.Exists(Drawing1) = True Then

                                        Dim Colectie_nume_atribute As New Specialized.StringCollection
                                        Dim Colectie_valori_atribute As New Specialized.StringCollection



                                        Dim Database1 As New Database(False, True)

                                        Try
                                            Try
                                                Try
                                                    Database1.ReadDwgFile(Drawing1, IO.FileShare.ReadWrite, True, Nothing)
                                                Catch ex As Exception
                                                    MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                    GoTo 123
                                                End Try
                                            Catch ex As IO.IOException
                                                MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                Error_list = Error_list & vbCrLf & Drawing1
                                                GoTo 123
                                            End Try
                                        Catch ex As System.SystemException
                                            MsgBox(Drawing1 & vbCrLf & "could not be open")
                                            Error_list = Error_list & vbCrLf & Drawing1
                                            GoTo 123
                                        End Try


                                        HostApplicationServices.WorkingDatabase = Database1



                                        Dim is_erased As Boolean = False

                                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction

                                            Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                                            Dim Layoutdict As DBDictionary

                                            Layoutdict = Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead)

                                            For Each entry As DBDictionaryEntry In Layoutdict
                                                Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                                If Layout1.TabOrder >= CInt(TextBox_Layout_index_START.Text) And Layout1.TabOrder <= CInt(TextBox_Layout_index_END.Text) And Layout1.TabOrder > 0 Then
                                                    Dim BTrecord As BlockTableRecord = Trans1.GetObject(Layout1.BlockTableRecordId, OpenMode.ForWrite)
                                                    For Each Id1 As ObjectId In BTrecord
                                                        Dim Ent1 As Entity = TryCast(Trans1.GetObject(Id1, OpenMode.ForRead), Entity)
                                                        If IsNothing(Ent1) = False Then


                                                            If TypeOf (Ent1) Is Viewport Then
                                                                Dim Viewport1 As Viewport = TryCast(Ent1, Viewport)
                                                                If IsNothing(Viewport1) = False Then

                                                                    If Viewport1.CenterPoint.X > x1 And Viewport1.CenterPoint.X < x2 And Viewport1.CenterPoint.Y > y1 And Viewport1.CenterPoint.Y < y2 Then
                                                                        Viewport1.UpgradeOpen()
                                                                        Viewport1.Erase()
                                                                        is_erased = True
                                                                    End If

                                                                End If


                                                            End If



                                                        End If
                                                    Next
                                                End If

                                            Next

                                            If is_erased = True Then
                                                Trans1.Commit()

                                                Try
                                                    Database1.SaveAs(Drawing1, True, DwgVersion.Current, Database1.SecurityParameters)
                                                Catch ex As Exception
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                End Try
                                            End If

                                        End Using

                                        Database1.Dispose()

                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database
                                    End If

123:
                                Next

                                Trans11.Commit()
                            End Using
                        End Using
                    End If


                    If Not Error_list = "" Then
                        MsgBox("There are not saved files. The list of these files is loaded into the clipboard" & vbCrLf &
                               "Use paste <Ctrl+V> command in a text editor program (notepad, excel, word) to see the list")
                        My.Computer.Clipboard.SetText(Error_list)

                    End If


                Catch ex As Exception
                    MsgBox(ex.Message)

                End Try

            Catch ex As System.SystemException
                MsgBox("1: " & ex.Message)
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: DONE!")
            Freeze_operations = False

        End If
    End Sub

    Private Sub Button_write_page_to_excel_Click(sender As Object, e As EventArgs) Handles Button_write_page_to_excel.Click
        If Freeze_operations = False Then
            Freeze_operations = True

            Try

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Dim Start1 As Integer = 0
                Dim End1 As Integer = 0
                If IsNumeric(TextBox_debug_row1.Text) = True Then
                    Start1 = CInt(TextBox_debug_row1.Text)
                End If
                If IsNumeric(TextBox_debug_row2.Text) = True Then
                    End1 = CInt(TextBox_debug_row2.Text)
                End If

                If End1 = 0 Or Start1 = 0 Then
                    Freeze_operations = False
                    Exit Sub
                End If
                If End1 < Start1 Then
                    Freeze_operations = False
                    Exit Sub
                End If
                Dim Column_sta1 As String = "A"





                If IsNothing(Data_table_Matchlines) = False Then

                    If Data_table_Matchlines.Rows.Count > 0 Then







                        For i = Start1 To End1
                            Dim Station_string1 As String = W1.Range(Column_sta1 & i).Value2


                            If IsNumeric(Station_string1) = True Then
                                For j = 0 To Data_table_Matchlines.Rows.Count - 1
                                    Dim Station1 As Double = CDbl(Station_string1)
                                    Dim M1 As Double = Data_table_Matchlines.Rows(j).Item("BEGSTA")
                                    Dim M2 As Double = Data_table_Matchlines.Rows(j).Item("ENDSTA")
                                    If Station1 >= M1 And Station1 <= M2 Then
                                        W1.Range("E" & i).Value2 = "Alignment NO " & (j + 1).ToString
                                        W1.Range("F" & i).Value2 = Get_chainage_feet_from_double(M1 + Get_equation_value(M1), 0)
                                        W1.Range("G" & i).Value2 = Get_chainage_feet_from_double(M2 + Get_equation_value(M2), 0)
                                        Exit For
                                    End If


                                Next

                            End If

                        Next


                    End If
                End If
                MsgBox("done")


            Catch ex As Exception
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_rename_band_number_label_Click(sender As Object, e As EventArgs) Handles Button_rename_band_number_label.Click

        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Dim Start1 As Integer = 0
                Dim End1 As Integer = 0
                If IsNumeric(TextBox_SHEET_ROW_START.Text) = True Then
                    Start1 = CInt(TextBox_SHEET_ROW_START.Text)
                End If
                If IsNumeric(TextBox_SHEET_ROW_END.Text) = True Then
                    End1 = CInt(TextBox_SHEET_ROW_END.Text)
                End If

                If End1 = 0 Or Start1 = 0 Then
                    Freeze_operations = False
                    Exit Sub
                End If

                If End1 < Start1 Then
                    Freeze_operations = False
                    Exit Sub
                End If

                Dim Column_band As String = ""
                Column_band = TextBox_column_sheet_number.Text.ToUpper

                Dim Column_dwg As String = ""
                Column_dwg = TextBox_column_drawing_number.Text.ToUpper

                Dim dt1 As New System.Data.DataTable
                dt1.Columns.Add("DWG", GetType(String))
                dt1.Columns.Add("NO", GetType(Integer))


                For i = Start1 To End1
                    Dim Band_no As String = W1.Range(Column_band & i).Value2
                    Dim DWG_no As String = W1.Range(Column_dwg & i).Value2
                    If IsNumeric(Band_no) = True Then
                        dt1.Rows.Add()
                        dt1.Rows(dt1.Rows.Count - 1).Item("DWG") = DWG_no
                        dt1.Rows(dt1.Rows.Count - 1).Item("NO") = CInt(Band_no)


                    End If
                Next

                If dt1.Rows.Count > 0 Then

                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor


                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Colectie1 = New Specialized.StringCollection

                    Editor1.SetImpliedSelection(Empty_array)

                    Dim Rezultat_text As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt.MessageForAdding = vbLf & "Select band labels:"
                    Object_Prompt.SingleOnly = False
                    Rezultat_text = Editor1.GetSelection(Object_Prompt)
                    If Not Rezultat_text.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Freeze_operations = False
                        Editor1.SetImpliedSelection(Empty_array)
                        Editor1.WriteMessage(vbLf & "Command:")
                        Freeze_operations = False
                        Exit Sub
                    End If

                    If Rezultat_text.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        If IsNothing(Rezultat_text) = False Then
                            Using Lock_dwg As DocumentLock = ThisDrawing.LockDocument
                                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                                    Dim BTrecord As BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                                    For j = 0 To Rezultat_text.Value.Count - 1
                                        Dim Ent1 As Entity = Trans1.GetObject(Rezultat_text.Value.Item(j).ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                        If TypeOf Ent1 Is DBText Then
                                            Dim Text1 As DBText = Ent1

                                            If IsNumeric(Text1.TextString) = True Then
                                                Dim dwg_band As Integer = CInt(Text1.TextString)
                                                For i = 0 To dt1.Rows.Count - 1
                                                    Dim band As Integer = dt1.Rows(i).Item("NO")
                                                    If band = dwg_band Then
                                                        Text1.UpgradeOpen()
                                                        Text1.TextString = dt1.Rows(i).Item("DWG") & " - band no " & Text1.TextString
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


                End If






            Catch ex As System.Exception
                Freeze_operations = False
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_delete_viewport_Click(sender As Object, e As EventArgs) Handles Button_delete_viewport.Click
        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Error_list As String = ""
            Try




                Try

                    If IsNumeric(TextBox_Layout_index_START.Text) = False Then
                        MsgBox("Please specify the layout START INDEX!")
                        Freeze_operations = False
                        Exit Sub
                    ElseIf CInt(TextBox_Layout_index_START.Text) < 1 Then
                        MsgBox("Please specify the a paperspace layout index")
                        Freeze_operations = False
                        Exit Sub
                    End If
                    If IsNumeric(TextBox_Layout_index_END.Text) = False Then
                        MsgBox("Please specify the layout END INDEX!")
                        Freeze_operations = False
                        Exit Sub
                    ElseIf CInt(TextBox_Layout_index_END.Text) < 1 Then
                        MsgBox("Please specify the a paperspace layout index")
                        Freeze_operations = False
                        Exit Sub
                    End If


                    Dim H_del As Double = 0
                    Dim W_del As Double = 0
                    Dim Layer1 As String = TextBox_DELlayer.Text
                    Dim X_del As Double = 0
                    Dim Y_del As Double = 0

                    If IsNumeric(TextBox_DELh.Text) = False Then
                        MsgBox("Please specify the height")
                        Freeze_operations = False
                        Exit Sub
                    Else
                        H_del = CDbl(TextBox_DELh.Text)
                    End If
                    If IsNumeric(TextBox_DELw.Text) = False Then
                        MsgBox("Please specify the width")
                        Freeze_operations = False
                        Exit Sub
                    Else
                        W_del = CDbl(TextBox_DELw.Text)
                    End If

                    If IsNumeric(TextBox_DELx.Text) = False Then
                        MsgBox("Please specify the x")
                        Freeze_operations = False
                        Exit Sub
                    Else
                        X_del = CDbl(TextBox_DELx.Text)
                    End If
                    If IsNumeric(TextBox_DELy.Text) = False Then
                        MsgBox("Please specify the y")
                        Freeze_operations = False
                        Exit Sub
                    Else
                        Y_del = CDbl(TextBox_DELy.Text)
                    End If


                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                    If ListBox_DWG.Items.Count > 0 Then
                        Using lock1 As DocumentLock = ThisDrawing.LockDocument

                            Using Trans11 As Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                                For i = 0 To ListBox_DWG.Items.Count - 1
                                    Dim Drawing1 As String = ListBox_DWG.Items(i)
                                    If IO.File.Exists(Drawing1) = True Then

                                        Dim Colectie_nume_atribute As New Specialized.StringCollection
                                        Dim Colectie_valori_atribute As New Specialized.StringCollection



                                        Dim Database1 As New Database(False, True)

                                        Try
                                            Try
                                                Try
                                                    Database1.ReadDwgFile(Drawing1, IO.FileShare.ReadWrite, True, Nothing)
                                                Catch ex As Exception
                                                    MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                    GoTo 123
                                                End Try
                                            Catch ex As IO.IOException
                                                MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                Error_list = Error_list & vbCrLf & Drawing1
                                                GoTo 123
                                            End Try
                                        Catch ex As System.SystemException
                                            MsgBox(Drawing1 & vbCrLf & "could not be open")
                                            Error_list = Error_list & vbCrLf & Drawing1
                                            GoTo 123
                                        End Try


                                        HostApplicationServices.WorkingDatabase = Database1



                                        Dim is_erased As Boolean = False

                                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction

                                            Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                                            Dim Layoutdict As DBDictionary

                                            Layoutdict = Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead)

                                            For Each entry As DBDictionaryEntry In Layoutdict
                                                Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                                If Layout1.TabOrder >= CInt(TextBox_Layout_index_START.Text) And Layout1.TabOrder <= CInt(TextBox_Layout_index_END.Text) And Layout1.TabOrder > 0 Then
                                                    Dim BTrecord As BlockTableRecord = Trans1.GetObject(Layout1.BlockTableRecordId, OpenMode.ForWrite)
                                                    For Each Id1 As ObjectId In BTrecord
                                                        Dim Ent1 As Entity = TryCast(Trans1.GetObject(Id1, OpenMode.ForRead), Entity)
                                                        If IsNothing(Ent1) = False Then
                                                            If TypeOf (Ent1) Is Viewport Then
                                                                Dim Viewport1 As Viewport = TryCast(Ent1, Viewport)
                                                                If IsNothing(Viewport1) = False Then
                                                                    Dim hw As Boolean = False
                                                                    If CheckBox_use_h_w.Checked = False Then
                                                                        If Round(Viewport1.Width, 2) = Round(W_del, 2) And Round(Viewport1.Height, 2) = Round(H_del, 2) Then
                                                                            hw = True
                                                                        End If
                                                                    Else
                                                                        hw = True
                                                                    End If
                                                                    If hw = True And Viewport1.Layer = Layer1 Then
                                                                        If Viewport1.CenterPoint.X > X_del - 5 & Viewport1.CenterPoint.X < X_del + 5 &
                                                                                Viewport1.CenterPoint.Y > Y_del - 5 & Viewport1.CenterPoint.Y < Y_del + 5 Then

                                                                            Viewport1.UpgradeOpen()
                                                                            Viewport1.Erase()
                                                                            is_erased = True
                                                                        End If

                                                                    End If

                                                                End If
                                                            End If
                                                        End If
                                                    Next
                                                End If

                                            Next

                                            If is_erased = True Then
                                                Trans1.Commit()

                                                Try
                                                    Database1.SaveAs(Drawing1, True, DwgVersion.Current, Database1.SecurityParameters)
                                                Catch ex As Exception
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                End Try
                                            End If

                                        End Using

                                        Database1.Dispose()

                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database
                                    End If

123:
                                Next

                                Trans11.Commit()
                            End Using
                        End Using
                    End If


                    If Not Error_list = "" Then
                        MsgBox("There are not saved files. The list of these files is loaded into the clipboard" & vbCrLf &
                               "Use paste <Ctrl+V> command in a text editor program (notepad, excel, word) to see the list")
                        My.Computer.Clipboard.SetText(Error_list)

                    End If


                Catch ex As Exception
                    MsgBox(ex.Message)

                End Try

            Catch ex As System.Exception
                MsgBox(ex.Message)
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: DONE!")
            Freeze_operations = False

        End If
    End Sub
End Class

'Editor1.CurrentUserCoordinateSystem = WCS_align()
'Dim GraphicsManager As Autodesk.AutoCAD.GraphicsSystem.Manager = ThisDrawing.GraphicsManager
'Dim View0 As Autodesk.AutoCAD.GraphicsSystem.View = GraphicsManager.GetGsView(CShort(Application.GetSystemVariable("CVPORT")), True) ' acad 2013
'Dim View0 As Autodesk.AutoCAD.GraphicsSystem.View = GraphicsManager.GetCurrentAcGsView(CShort(Application.GetSystemVariable("CVPORT"))) ' acad 2015
'Len1 = Viewport1.Width
'Height1 = Viewport1.Height