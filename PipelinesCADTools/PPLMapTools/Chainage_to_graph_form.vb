Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Public Class Chainage_to_graph_form
    Dim Colectie_butoane As New Specialized.StringCollection
    Dim Freeze_operations As Boolean = False
    Dim Data_table_station_equation As System.Data.DataTable
    Dim Data_table_Profile3D As System.Data.DataTable

    Private Sub Chainage_to_graph_form_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Incarca_existing_Blocks_with_attributes_to_combobox(ComboBox_blocks)
        With TextBox_kp
            .Select()
        End With
        If ComboBox_blocks.Items.Count > 1 Then
            ComboBox_blocks.SelectedIndex = 1
        End If
        Incarca_existing_layers_to_combobox(ComboBox_BLOCK_LAYER)
        Incarca_existing_layers_to_combobox(ComboBox_Rectangle_LAYER)
        If ComboBox_BLOCK_LAYER.Items.Contains("TEXT") = True Then
            ComboBox_BLOCK_LAYER.SelectedIndex = ComboBox_BLOCK_LAYER.Items.IndexOf("TEXT")
        End If

        If Environment.UserName.ToUpper = "POP70694" Or Environment.UserName.ToUpper = "MOR72937" Or Environment.UserName.ToUpper = "PAN71158" Then
            Panel_3d.Visible = True
        Else
            Panel_3d.Visible = False
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

    Public Function Load_from_Excel_to_data_table(ByVal Start1 As Double, ByVal End1 As Double) As System.Data.DataTable

        Try



            Dim Table_data1 As New System.Data.DataTable

            Table_data1.Columns.Add("Chainage", GetType(String))
            Table_data1.Columns.Add("Description1", GetType(String))
            Table_data1.Columns.Add("Description2", GetType(String))
            Table_data1.Columns.Add("Chainage_down", GetType(String))
            Table_data1.Columns.Add("Description1_down", GetType(String))
            Table_data1.Columns.Add("Description2_down", GetType(String))

            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
            W1 = Get_active_worksheet_from_Excel()
            Dim Index1 As Double = 0
            For i = Start1 To End1
                Table_data1.Rows.Add()
                If Not TextBox_kp.Text = "" Then
                    Dim VALOARE As String = W1.Range(TextBox_kp.Text.ToUpper & i).Value
                    If Not Replace(VALOARE, " ", "") = "" Then
                        Table_data1.Rows.Item(Index1).Item("Chainage") = VALOARE
                    End If
                End If
                If Not TextBox_DESCR1.Text = "" Then
                    Dim VALOARE As String = W1.Range(TextBox_DESCR1.Text.ToUpper & i).Value
                    If Not Replace(VALOARE, " ", "") = "" Then
                        Table_data1.Rows.Item(Index1).Item("Description1") = VALOARE
                    End If
                End If
                If Not TextBox_DESCR2.Text = "" Then
                    Dim VALOARE As String = W1.Range(TextBox_DESCR2.Text.ToUpper & i).Value
                    If Not Replace(VALOARE, " ", "") = "" Then
                        Table_data1.Rows.Item(Index1).Item("Description2") = VALOARE
                    End If
                End If

                If Not TextBox_ch_down.Text = "" Then
                    Dim VALOARE As String = W1.Range(TextBox_ch_down.Text.ToUpper & i).Value
                    If Not Replace(VALOARE, " ", "") = "" Then
                        Table_data1.Rows.Item(Index1).Item("Chainage_down") = VALOARE
                    End If
                End If
                If Not TextBox_descr1_down.Text = "" Then
                    Dim VALOARE As String = W1.Range(TextBox_descr1_down.Text.ToUpper & i).Value
                    If Not Replace(VALOARE, " ", "") = "" Then
                        Table_data1.Rows.Item(Index1).Item("Description1_down") = VALOARE
                    End If
                End If
                If Not TextBox_descr2_down.Text = "" Then
                    Dim VALOARE As String = W1.Range(TextBox_descr2_down.Text.ToUpper & i).Value
                    If Not Replace(VALOARE, " ", "") = "" Then
                        Table_data1.Rows.Item(Index1).Item("Description2_down") = VALOARE
                    End If
                End If

                Index1 = Index1 + 1

            Next



            Return Table_data1


        Catch ex As System.Exception
            MsgBox(ex.Message)
        End Try
    End Function

    Private Sub Button_insert_blocks_to_graph_Click(sender As System.Object, e As System.EventArgs) Handles Button_insert_blocks_to_graph.Click
        Try


            If Val(TextBox_row_start.Text) < 1 Or IsNumeric(TextBox_row_start.Text) = False Then
                With TextBox_row_start
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify start row")

                Exit Sub
            End If

            If Val(TextBox_row_end.Text) < 1 Or IsNumeric(TextBox_row_end.Text) = False Then
                With TextBox_row_end
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify end row")

                Exit Sub
            End If

            If Val(TextBox_row_end.Text) < Val(TextBox_row_start.Text) Then
                With TextBox_row_end
                    .Text = ""
                    .Focus()
                End With
                MsgBox("End row smaller than start row")

                Exit Sub
            End If
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            If Freeze_operations = False Then

                Freeze_operations = True


                Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument

                    Dim Table_data1 As New System.Data.DataTable
                    Table_data1 = Load_from_Excel_to_data_table(Val(TextBox_row_start.Text), Val(TextBox_row_end.Text))




                    ' Dim k As Double = 1
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
                        Dim start1 As Double = CDbl(TextBox_row_start.Text)
                        Dim end1 As Double = CDbl(TextBox_row_end.Text)


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
                        Dim Block_scale As Double
                        If IsNumeric(TextBox_block_scale.Text) = True Then
                            Block_scale = CDbl(TextBox_block_scale.Text)
                        Else
                            MsgBox("SPECIFIED BLOCK SCALE IS NOT NUMERIC!")
                            Freeze_operations = False
                            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                            Freeze_operations = False
                            Exit Sub
                        End If



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

                        Dim Chainage_cunoscuta As Double = -100000

                        Dim Rezultat_poly_graph As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                        Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_Prompt1.MessageForAdding = vbLf & "Select the Graph polyline:"
                        Object_Prompt1.SingleOnly = True
                        Rezultat_poly_graph = Editor1.GetSelection(Object_Prompt1)

                        If Not Rezultat_poly_graph.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Editor1.SetImpliedSelection(Empty_array)
                            Freeze_operations = False
                            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                            Freeze_operations = False
                            Exit Sub
                        End If

                        Dim Rezultat__elev_lines As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                        Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_Prompt2.MessageForAdding = vbLf & "Select the top and bottom of graph:"
                        Object_Prompt2.SingleOnly = False
                        Rezultat__elev_lines = Editor1.GetSelection(Object_Prompt2)

                        Dim Ymin, Ymax As Double
                        Dim Assigned As Boolean = False

                        If Rezultat__elev_lines.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            For i = 1 To Rezultat__elev_lines.Value.Count
                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Rezultat__elev_lines.Value.Item(i - 1)
                                Dim Ent1 As Entity
                                Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                                    Dim Linie1 As Line = Ent1
                                    If Abs(Linie1.StartPoint.Y - Linie1.EndPoint.Y) < 0.02 Then
                                        If i = 1 Then
                                            Assigned = True
                                            Ymin = Linie1.StartPoint.Y
                                            Ymax = Linie1.StartPoint.Y
                                        Else
                                            If Linie1.StartPoint.Y > Ymax Then Ymax = Linie1.StartPoint.Y
                                            If Linie1.StartPoint.Y < Ymin Then Ymin = Linie1.StartPoint.Y
                                        End If
                                    End If
                                End If
                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                    Dim PLinie1 As Polyline = Ent1
                                    If PLinie1.StartPoint.Y = PLinie1.EndPoint.Y And PLinie1.NumberOfVertices = 2 Then
                                        If i = 1 And Assigned = False Then
                                            Ymin = PLinie1.StartPoint.Y
                                            Ymax = PLinie1.StartPoint.Y
                                        Else
                                            If PLinie1.StartPoint.Y > Ymax Then Ymax = PLinie1.StartPoint.Y
                                            If PLinie1.StartPoint.Y < Ymin Then Ymin = PLinie1.StartPoint.Y
                                        End If
                                    End If
                                End If
                            Next
                        Else
                            Ymin = 0
                            Ymax = 10
                        End If






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
                            If IsNumeric(Replace(mText_cunoscut_chainage.Text, "+", "")) = True Then Chainage_cunoscuta = CDbl(Replace(mText_cunoscut_chainage.Text, "+", ""))
                        End If

                        If TypeOf Ent3_chainage Is Autodesk.AutoCAD.DatabaseServices.MText Then
                            mText_cunoscut_chainage = Ent3_chainage
                            If IsNumeric(Replace(mText_cunoscut_chainage.Text, "+", "")) = True Then Chainage_cunoscuta = CDbl(Replace(mText_cunoscut_chainage.Text, "+", ""))
                        End If

                        If TypeOf Ent2_chainage Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                            Text_cunoscut_chainage = Ent2_chainage
                            If IsNumeric(Replace(Text_cunoscut_chainage.TextString, "+", "")) = True Then Chainage_cunoscuta = CDbl(Replace(Text_cunoscut_chainage.TextString, "+", ""))
                        End If

                        If TypeOf Ent3_chainage Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                            Text_cunoscut_chainage = Ent3_chainage
                            If IsNumeric(Replace(Text_cunoscut_chainage.TextString, "+", "")) = True Then Chainage_cunoscuta = CDbl(Replace(Text_cunoscut_chainage.TextString, "+", ""))
                        End If

                        If Chainage_cunoscuta = -100000 Then
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

                        Creaza_layer("NO PLOT", 40, "", False)

                        Dim Poly_graph As Polyline

                        For j = 0 To Rezultat_poly_graph.Value.Count - 1
                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat_poly_graph.Value.Item(j)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                Poly_graph = Ent1
                            ElseIf TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline2d Then
                                Poly_graph = New Polyline
                                Dim Poly2d As Polyline2d = Ent1
                                Dim Indx As Integer = 0
                                For Each Id1 As ObjectId In Poly2d
                                    Dim vertex2d As Vertex2d = Trans1.GetObject(Id1, OpenMode.ForRead)
                                    Poly_graph.AddVertexAt(Indx, New Point2d(vertex2d.Position.X, vertex2d.Position.Y), 0, 0, 0)
                                    Indx = Indx + 1
                                Next

                            End If

                        Next
                        If IsNothing(Poly_graph) = True Then
                            Freeze_operations = False
                            MsgBox("no polyline selected")
                            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                            Freeze_operations = False
                            Exit Sub
                        End If

                        Dim Elevatia_cunoscuta As Double = -100000
                        Dim V_EXAG As Double = 1
                        Dim Y_elev_cunoscut As Double = 0


                        If Not ComboBox_atrib_elev_down.Text = "" Or Not ComboBox_atrib_elev_up.Text = "" Then
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


                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat_vert.Value.Item(0)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                            Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj2 = Rezultat_vert.Value.Item(1)
                            Dim Ent2 As Entity
                            Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)

                            Dim mText_cunoscut As Autodesk.AutoCAD.DatabaseServices.MText
                            Dim Text_cunoscut As Autodesk.AutoCAD.DatabaseServices.DBText

                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                                mText_cunoscut = Ent1
                                If IsNumeric(Replace(mText_cunoscut.Text, "'", "")) = True Then Elevatia_cunoscuta = CDbl(Replace(mText_cunoscut.Text, "'", ""))
                            End If

                            If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                                mText_cunoscut = Ent2
                                If IsNumeric(Replace(mText_cunoscut.Text, "'", "")) = True Then Elevatia_cunoscuta = CDbl(Replace(mText_cunoscut.Text, "'", ""))
                            End If

                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                                Text_cunoscut = Ent1
                                If IsNumeric(Replace(Text_cunoscut.TextString, "'", "")) = True Then Elevatia_cunoscuta = CDbl(Replace(Text_cunoscut.TextString, "'", ""))
                            End If

                            If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                                Text_cunoscut = Ent2
                                If IsNumeric(Replace(Text_cunoscut.TextString, "'", "")) = True Then Elevatia_cunoscuta = CDbl(Replace(Text_cunoscut.TextString, "'", ""))
                            End If

                            Dim Linia_Elev_cunoscuta As Autodesk.AutoCAD.DatabaseServices.Line
                            Dim polylinia_elev_cunoscuta As Autodesk.AutoCAD.DatabaseServices.Polyline

                            Dim x0e1, y0e1, x0e2, y0e2 As Double

                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                                Linia_Elev_cunoscuta = Ent1
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

                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                polylinia_elev_cunoscuta = Ent1

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


                        End If



                        For i = 0 To Table_data1.Rows.Count - 1
                            Dim Colectie_atr_name As New Specialized.StringCollection
                            Dim Colectie_atr_value As New Specialized.StringCollection


                            If IsDBNull(Table_data1.Rows.Item(i).Item("Description1")) = False Then
                                If Not TextBox_DESCR1.Text = "" Then
                                    If Not ComboBox_atrib_description_1.Text = "" Then
                                        Colectie_atr_value.Add(Table_data1.Rows.Item(i).Item("Description1"))
                                        Colectie_atr_name.Add(ComboBox_atrib_description_1.Text)
                                    End If
                                End If

                            End If


                            If IsDBNull(Table_data1.Rows.Item(i).Item("Description2")) = False Then
                                If Not TextBox_DESCR2.Text = "" Then
                                    If Not ComboBox_atrib_description_2.Text = "" Then
                                        Colectie_atr_value.Add(Table_data1.Rows.Item(i).Item("Description2"))
                                        Colectie_atr_name.Add(ComboBox_atrib_description_2.Text)
                                    End If
                                End If

                            End If

                            If IsDBNull(Table_data1.Rows.Item(i).Item("Description1_down")) = False Then
                                If Not TextBox_descr1_down.Text = "" Then
                                    If Not ComboBox_atrib_description_1_down.Text = "" Then
                                        Colectie_atr_value.Add(Table_data1.Rows.Item(i).Item("Description1_down"))
                                        Colectie_atr_name.Add(ComboBox_atrib_description_1_down.Text)
                                    End If
                                End If

                            End If


                            If IsDBNull(Table_data1.Rows.Item(i).Item("Description2_down")) = False Then
                                If Not TextBox_descr2_down.Text = "" Then
                                    If Not ComboBox_atrib_description_2_down.Text = "" Then
                                        Colectie_atr_value.Add(Table_data1.Rows.Item(i).Item("Description2_down"))
                                        Colectie_atr_name.Add(ComboBox_atrib_description_2_down.Text)
                                    End If
                                End If

                            End If

                            If IsDBNull(Table_data1.Rows.Item(i).Item("Chainage")) = False Then
                                If Not TextBox_kp.Text = "" Then
                                    If Not ComboBox_atrib_chainage.Text = "" Then
                                        Dim Text1 As String = Table_data1.Rows.Item(i).Item("Chainage")


                                        If IsNumeric(Text1) = True Then
                                            Dim Nr1 As Double = CDbl(Text1)
                                            If CheckBox_use_equation.Checked = True Then
                                                Nr1 = Nr1 + Get_equation_value(Nr1)
                                            End If

                                            If CheckBox_USA.Checked = True Then
                                                Text1 = Get_chainage_feet_from_double(Nr1, 0)
                                            Else
                                                Text1 = Get_chainage_from_double(Nr1, 1)
                                            End If

                                        End If
                                        Colectie_atr_value.Add(Text1)
                                        Colectie_atr_name.Add(ComboBox_atrib_chainage.Text)
                                    End If
                                End If

                            End If

                            If IsDBNull(Table_data1.Rows.Item(i).Item("Chainage_down")) = False Then
                                If Not TextBox_ch_down.Text = "" Then
                                    If Not ComboBox_atrib_chainage_down.Text = "" Then
                                        Dim Text1 As String = Table_data1.Rows.Item(i).Item("Chainage_down")


                                        If IsNumeric(Text1) = True Then
                                            Dim Nr1 As Double = CDbl(Text1)
                                            If CheckBox_use_equation.Checked = True Then
                                                Nr1 = Nr1 + Get_equation_value(Nr1)
                                            End If

                                            If CheckBox_USA.Checked = True Then
                                                Text1 = Get_chainage_feet_from_double(Nr1, 0)
                                            Else
                                                Text1 = Get_chainage_from_double(Nr1, 1)
                                            End If

                                        End If
                                        Colectie_atr_value.Add(Text1)
                                        Colectie_atr_name.Add(ComboBox_atrib_chainage_down.Text)
                                    End If
                                End If

                            End If




                            Dim Block_name_string As String = ComboBox_blocks.Text


                            'de aici sunt bLOCURI

                            Dim Valoare_chainage As Double
                            If IsNumeric(Replace(Table_data1.Rows.Item(i).Item("Chainage"), "+", "")) = True Then
                                Valoare_chainage = CDbl(Replace(Table_data1.Rows.Item(i).Item("Chainage"), "+", ""))
                                Dim x As Double
                                If Radio_LEFT_R.Checked = True Then
                                    x = x0_sta1 + (Valoare_chainage - Chainage_cunoscuta) * H_EXAG
                                Else
                                    x = x0_sta1 - (Valoare_chainage - Chainage_cunoscuta) * H_EXAG
                                End If
                                Dim Linie1 As New Line
                                Linie1.StartPoint = New Point3d(x, Ymin, 0)
                                Linie1.EndPoint = New Point3d(x, Ymax, 0)
                                Linie1.Layer = "NO PLOT"
                                Linie1.ColorIndex = 7
                                Dim Descr As String = " "
                                If IsDBNull(Table_data1.Rows.Item(i).Item("Description1")) = False Then
                                    Descr = "-" & Table_data1.Rows.Item(i).Item("Description1")
                                    If IsDBNull(Table_data1.Rows.Item(i).Item("Description2")) = False Then
                                        Descr = Descr & " " & Table_data1.Rows.Item(i).Item("Description2")
                                    End If
                                End If


                                Dim MTEXT1 As New MText


                                If CheckBox_use_equation.Checked = True Then
                                    Valoare_chainage = Valoare_chainage + Get_equation_value(Valoare_chainage)
                                End If
                                Dim String1 As String
                                If CheckBox_USA.Checked = True Then
                                    String1 = Get_chainage_feet_from_double(Valoare_chainage, 0)
                                Else
                                    String1 = Get_chainage_from_double(Valoare_chainage, 1)

                                End If
                                MTEXT1.Contents = String1 & "-" & Descr
                                MTEXT1.Location = New Point3d(x, Ymin + 0.25, 0)
                                MTEXT1.Rotation = PI / 2
                                MTEXT1.Layer = "NO PLOT"
                                MTEXT1.TextHeight = 0.5
                                MTEXT1.ColorIndex = 7
                                MTEXT1.Attachment = AttachmentPoint.MiddleLeft

                                Dim Colectie_pt As New Point3dCollection
                                Linie1.IntersectWith(Poly_graph, Intersect.ExtendThis, Colectie_pt, IntPtr.Zero, IntPtr.Zero)
                                If Colectie_pt.Count > 0 Then
                                    If Rezultat__elev_lines.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                        BTrecord.AppendEntity(Linie1)
                                        Trans1.AddNewlyCreatedDBObject(Linie1, True)
                                        BTrecord.AppendEntity(MTEXT1)
                                        Trans1.AddNewlyCreatedDBObject(MTEXT1, True)
                                    End If

                                    Dim Point1 As New Point3d
                                    Point1 = Colectie_pt(0)

                                    If Not ComboBox_atrib_elev_up.Text = "" Or Not ComboBox_atrib_elev_down.Text = "" Then
                                        If Not Elevatia_cunoscuta = -100000 Then
                                            Dim Round1 As Integer = 1
                                            If IsNumeric(TextBox_rounding_elevation.Text) = True Then
                                                Round1 = CInt(TextBox_rounding_elevation.Text)
                                            End If
                                            Dim Elev1 As Double = Elevatia_cunoscuta + (Point1.Y - Y_elev_cunoscut) / V_EXAG
                                            Dim Elevation_string As String = TextBox_Elevation_prefix.Text & Get_String_Rounded(Elev1, Round1) & TextBox_Elevation_suffix.Text
                                            If Not ComboBox_atrib_elev_up.Text = "" Then
                                                Colectie_atr_name.Add(ComboBox_atrib_elev_up.Text)
                                                Colectie_atr_value.Add(Elevation_string)
                                            End If
                                            If Not ComboBox_atrib_elev_down.Text = "" Then
                                                Colectie_atr_name.Add(ComboBox_atrib_elev_down.Text)
                                                Colectie_atr_value.Add(Elevation_string)
                                            End If

                                        End If
                                    End If



                                    InsertBlock_with_multiple_atributes(Block_name_string & ".dwg", Block_name_string, Point1, Block_scale, BTrecord, ComboBox_BLOCK_LAYER.Text, Colectie_atr_name, Colectie_atr_value)


                                Else
                                    W1.Range(TextBox_kp.Text.ToUpper & (i + start1)).Interior.ColorIndex = 3

                                    If MsgBox("No intersection at cell " & TextBox_kp.Text.ToUpper & (i + start1) & vbCrLf & "Do you want to continue?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                                        If MsgBox("Do you want to add all blocks up to here?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then Trans1.Commit()
                                        Freeze_operations = False
                                        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                        Exit Sub
                                    End If

                                End If


                            Else
                                W1.Range(TextBox_kp.Text.ToUpper & (i + start1)).Interior.ColorIndex = 3

                                If MsgBox("Not valid chainage at cell " & TextBox_kp.Text.ToUpper & (i + start1) & vbCrLf & "Do you want to continue?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                                    If MsgBox("Do you want to add all blocks up to here?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then Trans1.Commit()
                                    Freeze_operations = False
                                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                    Exit Sub
                                End If

                            End If

                            'asta e de la INSERT BLOCKS
                        Next

                        Trans1.Commit()
                        ' asta e de la tranzactie
                    End Using





                    Freeze_operations = False



                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    ' asta e de la lock
                End Using
                Freeze_operations = False
            End If
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

        Catch ex As System.Exception
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

            Freeze_operations = False
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub ComboBox_blocks_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ComboBox_blocks.SelectedIndexChanged
        Try
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Using LOCK1 As DocumentLock = ThisDrawing.LockDocument
                If ComboBox_atrib_description_1.Items.Count > 0 Then ComboBox_atrib_description_1.Items.Clear()
                If ComboBox_atrib_description_2.Items.Count > 0 Then ComboBox_atrib_description_2.Items.Clear()
                If ComboBox_atrib_chainage.Items.Count > 0 Then ComboBox_atrib_chainage.Items.Clear()
                If ComboBox_atrib_description_1_down.Items.Count > 0 Then ComboBox_atrib_description_1_down.Items.Clear()
                If ComboBox_atrib_description_2_down.Items.Count > 0 Then ComboBox_atrib_description_2_down.Items.Clear()
                If ComboBox_atrib_chainage_down.Items.Count > 0 Then ComboBox_atrib_chainage_down.Items.Clear()
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
                                        ComboBox_atrib_description_2.Items.Add(attDefinition1.Tag)
                                        ComboBox_atrib_chainage.Items.Add(attDefinition1.Tag)
                                        ComboBox_atrib_description_1_down.Items.Add(attDefinition1.Tag)
                                        ComboBox_atrib_description_2_down.Items.Add(attDefinition1.Tag)
                                        ComboBox_atrib_chainage_down.Items.Add(attDefinition1.Tag)
                                        ComboBox_atrib_elev_up.Items.Add(attDefinition1.Tag)
                                        ComboBox_atrib_elev_down.Items.Add(attDefinition1.Tag)
                                    End If
                                End If
                            Next
                        End If
                    End If
                End Using ' asta e de la trans1
            End Using
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Catch ex As System.Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub TextBox_kp_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_kp.KeyDown, TextBox_column_Stationp2e.KeyDown, TextBox_column_x.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            With TextBox_DESCR1
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub

    Private Sub TextBox_DESCR1_KeyDown(sender As Object, e As Windows.Forms.KeyEventArgs) Handles TextBox_DESCR1.KeyDown, TextBox_column_Elevationp2e.KeyDown, TextBox_column_y.KeyDown, TextBox_Elevation_suffix.KeyDown, TextBox_Elevation_prefix.KeyDown, TextBox_rounding_elevation.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            With TextBox_DESCR2
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub

    Private Sub TextBox_DESCR2_KeyDown(sender As Object, e As Windows.Forms.KeyEventArgs) Handles TextBox_DESCR2.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            With TextBox_ch_down
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub

    Private Sub TextBox_ch_down_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_ch_down.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            With TextBox_descr1_down
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub

    Private Sub TextBox_descr1_down_KeyDown(sender As Object, e As Windows.Forms.KeyEventArgs) Handles TextBox_descr1_down.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            With TextBox_descr2_down
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub

    Private Sub TextBox_descr2_down_KeyDown(sender As Object, e As Windows.Forms.KeyEventArgs) Handles TextBox_descr2_down.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            With TextBox_row_start
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub

    Private Sub TextBox_row_start_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_row_start.KeyDown, TextBox_start_row_p2e.KeyDown, TextBox_end_row_p2e.KeyDown, TextBox_interval.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            With TextBox_row_end
                .SelectAll()
                .Focus()
            End With

        End If
    End Sub

    Private Sub Panel_graph_Click(sender As Object, e As EventArgs) Handles Panel_graph.Click, Panel_layer1.Click
        If ComboBox_blocks.Items.Count > 0 Then ComboBox_blocks.SelectedIndex = 0
        Incarca_existing_Blocks_with_attributes_to_combobox(ComboBox_blocks)
        Incarca_existing_layers_to_combobox(ComboBox_BLOCK_LAYER)
        If ComboBox_BLOCK_LAYER.Items.Contains("TEXT") = True Then
            ComboBox_BLOCK_LAYER.SelectedIndex = ComboBox_BLOCK_LAYER.Items.IndexOf("TEXT")
        End If
        Incarca_existing_layers_to_combobox(ComboBox_Rectangle_LAYER)
    End Sub


    Private Sub Button_add_all_blocks_Click(sender As Object, e As EventArgs) Handles Button_add_all_blocks.Click
        Incarca_existing_Blocks_to_combobox(ComboBox_blocks)
    End Sub

    Private Sub Button_fix_block_Click(sender As Object, e As EventArgs) Handles Button_fix_block.Click
        Try


            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Freeze_operations = True
            Me.Refresh()
            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1

                ' Dim k As Double = 1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)



                    For Each iD1 As ObjectId In BTrecord
                        Dim Ent1 As Entity = TryCast(Trans1.GetObject(iD1, OpenMode.ForRead), Entity)
                        If IsNothing(Ent1) = False Then
                            If TypeOf (Ent1) Is BlockReference Then
                                Dim Block1 As BlockReference = Ent1

                                If Block1.AttributeCollection.Count > 0 Then
                                    Block1.UpgradeOpen()
                                    Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                    Dim First_line As String = "FIRST_LINE"
                                    Dim Second_line As String = "SECOND_LINE"
                                    Dim Chainage As String = "CHAINAGE"
                                    Dim First_line_D As String = "FST_LINE_D"
                                    Dim Second_line_D As String = "SEC_LINE_D"
                                    Dim Chainage_D As String = "CHAINAGE_D"

                                    Dim First_line_V As String = ""
                                    Dim Second_line_V As String = ""
                                    Dim Chainage_V As String = ""
                                    Dim First_line_D_V As String = ""
                                    Dim Second_line_D_V As String = ""
                                    Dim Chainage_D_V As String = ""


                                    For Each id In attColl
                                        Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForWrite)

                                        If attref.Tag.ToUpper = First_line Then
                                            Dim Continut As String = attref.TextString
                                            If Not Replace(Continut, " ", "") = "" Then
                                                First_line_V = Continut
                                            End If
                                        End If
                                        If attref.Tag.ToUpper = Second_line Then
                                            Dim Continut As String = attref.TextString
                                            If Not Replace(Continut, " ", "") = "" Then
                                                Second_line_V = Continut
                                            End If
                                        End If
                                        If attref.Tag.ToUpper = Chainage Then
                                            Dim Continut As String = attref.TextString
                                            If Not Replace(Continut, " ", "") = "" Then
                                                Chainage_V = Continut
                                            End If
                                        End If
                                        If attref.Tag.ToUpper = First_line_D Then
                                            Dim Continut As String = attref.TextString
                                            If Not Replace(Continut, " ", "") = "" Then
                                                First_line_D_V = Continut
                                            End If
                                        End If
                                        If attref.Tag.ToUpper = Second_line_D Then
                                            Dim Continut As String = attref.TextString
                                            If Not Replace(Continut, " ", "") = "" Then
                                                Second_line_D_V = Continut
                                            End If
                                        End If
                                        If attref.Tag.ToUpper = Chainage_D Then
                                            Dim Continut As String = attref.TextString
                                            If Not Replace(Continut, " ", "") = "" Then
                                                Chainage_D_V = Continut
                                            End If
                                        End If
                                    Next


                                    For Each id In attColl
                                        Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                        If attref.Tag.ToUpper = First_line Then
                                            Dim Continut As String = attref.TextString
                                            If Not Replace(First_line_D_V, " ", "") = "" Then
                                                If Continut = "" Then
                                                    attref.TextString = First_line_D_V
                                                End If
                                            End If



                                        End If
                                        If attref.Tag.ToUpper = Second_line Then
                                            Dim Continut As String = attref.TextString
                                            If Not Replace(Second_line_D_V, " ", "") = "" Then
                                                If Continut = "" Then
                                                    attref.TextString = Second_line_D_V
                                                End If
                                            End If
                                        End If
                                        If attref.Tag.ToUpper = Chainage Then
                                            Dim Continut As String = attref.TextString
                                            If Not Replace(Chainage_D_V, " ", "") = "" Then
                                                If Continut = "" Then
                                                    attref.TextString = Chainage_D_V
                                                End If
                                            End If
                                        End If
                                        If attref.Tag.ToUpper = First_line_D Then
                                            Dim Continut As String = attref.TextString
                                            If Not Replace(First_line_V, " ", "") = "" Then
                                                If Continut = "" Then
                                                    attref.TextString = First_line_V
                                                End If
                                            End If
                                        End If
                                        If attref.Tag.ToUpper = Second_line_D Then
                                            Dim Continut As String = attref.TextString
                                            If Not Replace(Second_line_V, " ", "") = "" Then
                                                If Continut = "" Then
                                                    attref.TextString = Second_line_V
                                                End If
                                            End If
                                        End If
                                        If attref.Tag.ToUpper = Chainage_D Then
                                            Dim Continut As String = attref.TextString
                                            If Not Replace(Chainage_V, " ", "") = "" Then
                                                If Continut = "" Then
                                                    attref.TextString = Chainage_V

                                                End If
                                            End If
                                        End If
                                    Next


                                End If



                            End If


                        End If




                    Next
                    Trans1.Commit()
                End Using
            End Using
        Catch ex As System.Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_profile_to_excel_Click(sender As Object, e As EventArgs) Handles Button_profile_to_excel.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
            W1 = Get_active_worksheet_from_Excel()
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                    Dim start1 As Double = CDbl(TextBox_start_row_p2e.Text)
                    Dim Column_sta As String = TextBox_column_Stationp2e.Text.ToUpper
                    Dim Column_elev As String = TextBox_column_Elevationp2e.Text.ToUpper


                    Dim Rezultat_hor As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_prompt_hor As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_prompt_hor.MessageForAdding = vbLf & "Select a known vertical line (STATION) and the label for it:" & vbCrLf & "If your selection contains only a line means you selected 0+000 station"

                    Object_prompt_hor.SingleOnly = False
                    Rezultat_hor = Editor1.GetSelection(Object_prompt_hor)


                    Dim Rezultat_vert As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_prompt_vert As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_prompt_vert.MessageForAdding = vbLf & "Select a known horizontal line (ELEVATION) and the label for it:"

                    Object_prompt_vert.SingleOnly = False
                    Rezultat_vert = Editor1.GetSelection(Object_prompt_vert)

                    Dim Curent_UCS As Matrix3d = Editor1.CurrentUserCoordinateSystem

                    Dim Rezultat_horiz_exag As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify Horizontal Exaggeration:")
                    Rezultat_horiz_exag.DefaultValue = 1
                    Rezultat_horiz_exag.AllowNone = True
                    Dim Rezultat_h_exag As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(Rezultat_horiz_exag)

                    Dim Horizontal_exag As Double = Rezultat_h_exag.Value
                    If Horizontal_exag = 0 Then Horizontal_exag = 1

                    Dim Rezultat_vert_exag As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify Vertical Exaggeration:")
                    Rezultat_vert_exag.DefaultValue = 1
                    Rezultat_vert_exag.AllowNone = True
                    Dim Rezultat_v_exag As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(Rezultat_vert_exag)

                    Dim Vertical_exag As Double = Rezultat_v_exag.Value
                    If Vertical_exag = 0 Then Vertical_exag = 1




                    If Rezultat_hor.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK And _
                        Rezultat_vert.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK And _
                        Rezultat_h_exag.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK And _
                        Rezultat_v_exag.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        Dim Elevatia_cunoscuta As Double = -100000
                        Dim Distanta_de_la_zero1 As Double = -100000
                        Dim Chainage_cunoscuta As Double = -100000

                        Dim mText_cunoscut As Autodesk.AutoCAD.DatabaseServices.MText
                        Dim Text_cunoscut As Autodesk.AutoCAD.DatabaseServices.DBText
                        Dim Linia_cunoscuta As Autodesk.AutoCAD.DatabaseServices.Line
                        Dim polyLinia_cunoscuta As Autodesk.AutoCAD.DatabaseServices.Polyline

                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_vert.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                        Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj2 = Rezultat_vert.Value.Item(1)
                        Dim Ent2 As Entity
                        Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)

                        Dim Obj3 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj3 = Rezultat_hor.Value.Item(0)
                        Dim Ent3 As Entity
                        Ent3 = Obj3.ObjectId.GetObject(OpenMode.ForRead)

                        Dim Obj4 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Dim Ent4 As Entity

                        If Rezultat_hor.Value.Count > 1 Then
                            Obj4 = Rezultat_hor.Value.Item(1)
                            Ent4 = Obj4.ObjectId.GetObject(OpenMode.ForRead)
                        End If





                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                            mText_cunoscut = Ent1
                            If IsNumeric(Replace(mText_cunoscut.Text, "'", "")) = True Then Elevatia_cunoscuta = CDbl(Replace(mText_cunoscut.Text, "'", ""))
                        End If

                        If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                            mText_cunoscut = Ent2
                            If IsNumeric(Replace(mText_cunoscut.Text, "'", "")) = True Then Elevatia_cunoscuta = CDbl(Replace(mText_cunoscut.Text, "'", ""))
                        End If

                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                            Text_cunoscut = Ent1
                            If IsNumeric(Replace(Text_cunoscut.TextString, "'", "")) = True Then Elevatia_cunoscuta = CDbl(Replace(Text_cunoscut.TextString, "'", ""))
                        End If

                        If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                            Text_cunoscut = Ent2
                            If IsNumeric(Replace(Text_cunoscut.TextString, "'", "")) = True Then Elevatia_cunoscuta = CDbl(Replace(Text_cunoscut.TextString, "'", ""))
                        End If

                        If Elevatia_cunoscuta = -100000 Then
                            Freeze_operations = False
                            Editor1.SetImpliedSelection(Empty_array)
                            MsgBox("No elevation")
                            Exit Sub
                        End If

                        Dim Rezultat_poly_graph As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                        Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_Prompt1.MessageForAdding = vbLf & "Select the Graph polyline:"
                        Object_Prompt1.SingleOnly = True
                        Rezultat_poly_graph = Editor1.GetSelection(Object_Prompt1)

                        If Not Rezultat_poly_graph.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Editor1.SetImpliedSelection(Empty_array)
                            Freeze_operations = False
                            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If

                        Dim Ent_poly As Entity = Trans1.GetObject(Rezultat_poly_graph.Value(0).ObjectId, OpenMode.ForRead)
                        If TypeOf Ent_poly Is Polyline Then
                            Dim x01, y01, x02, y02, dist1, a1 As Double
                            Dim Poly_graph As Polyline = Ent_poly
                            Dim index_excel As Integer = start1


                            For i = 0 To Poly_graph.NumberOfVertices - 1
                                Dim Point1 As Point3d
                                Point1 = Poly_graph.GetPoint3dAt(i)



                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                                    Linia_cunoscuta = Ent1
                                    x01 = Linia_cunoscuta.StartPoint.X
                                    y01 = Linia_cunoscuta.StartPoint.Y
                                    x02 = Linia_cunoscuta.EndPoint.X
                                    y02 = Linia_cunoscuta.EndPoint.Y
                                    If Abs(y01 - y02) > 0.001 Then
                                        Editor1.SetImpliedSelection(Empty_array)
                                        Freeze_operations = False
                                        Exit Sub
                                    End If
                                    dist1 = ((Point1.X - x01) ^ 2 + (Point1.Y - y01) ^ 2) ^ 0.5
                                    a1 = Abs(Point1.X - x01)
                                    Distanta_de_la_zero1 = (dist1 ^ 2 - a1 ^ 2) ^ 0.5

                                End If


                                If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                                    Linia_cunoscuta = Ent2
                                    x01 = Linia_cunoscuta.StartPoint.X
                                    y01 = Linia_cunoscuta.StartPoint.Y
                                    x02 = Linia_cunoscuta.EndPoint.X
                                    y02 = Linia_cunoscuta.EndPoint.Y
                                    If Abs(y01 - y02) > 0.001 Then
                                        Editor1.SetImpliedSelection(Empty_array)
                                        Freeze_operations = False
                                        Exit Sub
                                    End If
                                    dist1 = ((Point1.X - x01) ^ 2 + (Point1.Y - y01) ^ 2) ^ 0.5
                                    a1 = Abs(Point1.X - x01)
                                    Distanta_de_la_zero1 = (dist1 ^ 2 - a1 ^ 2) ^ 0.5

                                End If

                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                    polyLinia_cunoscuta = Ent1

                                    x01 = polyLinia_cunoscuta.StartPoint.X
                                    y01 = polyLinia_cunoscuta.StartPoint.Y
                                    x02 = polyLinia_cunoscuta.EndPoint.X
                                    y02 = polyLinia_cunoscuta.EndPoint.Y
                                    If Abs(y01 - y02) > 0.001 Then
                                        Editor1.SetImpliedSelection(Empty_array)
                                        Freeze_operations = False
                                        Exit Sub
                                    End If
                                    dist1 = ((Point1.X - x01) ^ 2 + (Point1.Y - y01) ^ 2) ^ 0.5
                                    a1 = Abs(Point1.X - x01)
                                    Distanta_de_la_zero1 = (dist1 ^ 2 - a1 ^ 2) ^ 0.5

                                End If

                                If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                    polyLinia_cunoscuta = Ent2

                                    x01 = polyLinia_cunoscuta.StartPoint.X
                                    y01 = polyLinia_cunoscuta.StartPoint.Y
                                    x02 = polyLinia_cunoscuta.EndPoint.X
                                    y02 = polyLinia_cunoscuta.EndPoint.Y
                                    If Abs(y01 - y02) > 0.001 Then
                                        Editor1.SetImpliedSelection(Empty_array)
                                        Freeze_operations = False
                                        Exit Sub
                                    End If
                                    dist1 = ((Point1.X - x01) ^ 2 + (Point1.Y - y01) ^ 2) ^ 0.5
                                    a1 = Abs(Point1.X - x01)
                                    Distanta_de_la_zero1 = (dist1 ^ 2 - a1 ^ 2) ^ 0.5

                                End If

                                If Distanta_de_la_zero1 = -100000 Then
                                    Editor1.SetImpliedSelection(Empty_array)
                                    Freeze_operations = False
                                    Exit Sub
                                End If

                                Distanta_de_la_zero1 = Distanta_de_la_zero1 / Vertical_exag



                                If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                                    mText_cunoscut = Ent3
                                    Dim numar_fara_plus As String = Replace(mText_cunoscut.Text, "+", "")
                                    If IsNumeric(numar_fara_plus) = True Then Chainage_cunoscuta = CDbl(numar_fara_plus)
                                End If

                                If TypeOf Ent4 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                                    mText_cunoscut = Ent4
                                    Dim numar_fara_plus As String = Replace(mText_cunoscut.Text, "+", "")
                                    If IsNumeric(numar_fara_plus) = True Then Chainage_cunoscuta = CDbl(numar_fara_plus)
                                End If

                                If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                                    Text_cunoscut = Ent3
                                    Dim numar_fara_plus As String = Replace(Text_cunoscut.TextString, "+", "")
                                    If IsNumeric(numar_fara_plus) = True Then Chainage_cunoscuta = CDbl(numar_fara_plus)
                                End If

                                If TypeOf Ent4 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                                    Text_cunoscut = Ent4
                                    Dim numar_fara_plus As String = Replace(Text_cunoscut.TextString, "+", "")
                                    If IsNumeric(numar_fara_plus) = True Then Chainage_cunoscuta = CDbl(numar_fara_plus)
                                End If

                                If Not Distanta_de_la_zero1 = -100000 Then
                                    If Rezultat_hor.Value.Count = 1 Then Chainage_cunoscuta = 0
                                End If


                                If Chainage_cunoscuta = -100000 Then
                                    Freeze_operations = False
                                    Editor1.SetImpliedSelection(Empty_array)
                                    Exit Sub
                                End If
                                Dim x03, y03, x04, y04 As Double
                                Dim Chainage_at_point As Double

                                If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                                    Linia_cunoscuta = Ent3
                                    x03 = Linia_cunoscuta.StartPoint.X
                                    y03 = Linia_cunoscuta.StartPoint.Y
                                    x04 = Linia_cunoscuta.EndPoint.X
                                    y04 = Linia_cunoscuta.EndPoint.Y
                                    If Abs(x03 - x04) > 0.001 Then
                                        Editor1.SetImpliedSelection(Empty_array)
                                        Freeze_operations = False
                                        Exit Sub
                                    End If
                                    If Point1.X < x03 Then
                                        Chainage_at_point = Chainage_cunoscuta - Abs(x03 - Point1.X) / Horizontal_exag
                                    Else
                                        Chainage_at_point = Chainage_cunoscuta + Abs(x03 - Point1.X) / Horizontal_exag
                                    End If
                                End If


                                If TypeOf Ent4 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                                    Linia_cunoscuta = Ent4
                                    x03 = Linia_cunoscuta.StartPoint.X
                                    y03 = Linia_cunoscuta.StartPoint.Y
                                    x04 = Linia_cunoscuta.EndPoint.X
                                    y04 = Linia_cunoscuta.EndPoint.Y
                                    If Abs(x03 - x04) > 0.001 Then
                                        Editor1.SetImpliedSelection(Empty_array)
                                        Freeze_operations = False
                                        Exit Sub
                                    End If
                                    If Point1.X < x03 Then
                                        Chainage_at_point = Chainage_cunoscuta - Abs(x03 - Point1.X) / Horizontal_exag
                                    Else
                                        Chainage_at_point = Chainage_cunoscuta + Abs(x03 - Point1.X) / Horizontal_exag
                                    End If


                                End If

                                If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                    polyLinia_cunoscuta = Ent3
                                    x03 = polyLinia_cunoscuta.StartPoint.X
                                    y03 = polyLinia_cunoscuta.StartPoint.Y
                                    x04 = polyLinia_cunoscuta.EndPoint.X
                                    y04 = polyLinia_cunoscuta.EndPoint.Y
                                    If Abs(x03 - x04) > 0.001 Then
                                        Editor1.SetImpliedSelection(Empty_array)
                                        Freeze_operations = False
                                        Exit Sub
                                    End If
                                    If Point1.X < x03 Then
                                        Chainage_at_point = Chainage_cunoscuta - Abs(x03 - Point1.X) / Horizontal_exag
                                    Else
                                        Chainage_at_point = Chainage_cunoscuta + Abs(x03 - Point1.X) / Horizontal_exag
                                    End If
                                End If

                                If TypeOf Ent4 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                    polyLinia_cunoscuta = Ent4
                                    x03 = polyLinia_cunoscuta.StartPoint.X
                                    y03 = polyLinia_cunoscuta.StartPoint.Y
                                    x04 = polyLinia_cunoscuta.EndPoint.X
                                    y04 = polyLinia_cunoscuta.EndPoint.Y
                                    If Abs(x03 - x04) > 0.001 Then
                                        Editor1.SetImpliedSelection(Empty_array)
                                        Freeze_operations = False
                                        Exit Sub
                                    End If
                                    If Point1.X < x03 Then
                                        Chainage_at_point = Chainage_cunoscuta - Abs(x03 - Point1.X) / Horizontal_exag
                                    Else
                                        Chainage_at_point = Chainage_cunoscuta + Abs(x03 - Point1.X) / Horizontal_exag
                                    End If
                                End If

                                Dim Elevation_at_point As Double

                                If Point1.Y < y01 Then
                                    Elevation_at_point = Elevatia_cunoscuta - Distanta_de_la_zero1
                                Else
                                    Elevation_at_point = Elevatia_cunoscuta + Distanta_de_la_zero1
                                End If

                                W1.Range(Column_sta & index_excel).Value2 = Round(Chainage_at_point, 2)

                                W1.Range(Column_elev & index_excel).Value2 = Round(Elevation_at_point, 2)

                                index_excel = index_excel + 1

                            Next
                        End If






                    End If



                End Using
            End Using


            Freeze_operations = False
        End If
    End Sub




    Private Sub Button_poly_to_excel_Click(sender As Object, e As EventArgs) Handles Button_poly_to_excel.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
                Dim Empty_array() As ObjectId
                Editor1.SetImpliedSelection(Empty_array)
                Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                        Dim start1 As Double = CDbl(TextBox_start_row_p2e.Text)
                        Dim Column_sta As String = TextBox_column_Stationp2e.Text.ToUpper

                        Dim Column_X As String = TextBox_column_x.Text.ToUpper
                        Dim Column_Y As String = TextBox_column_y.Text.ToUpper

                        Dim Rezultat_poly As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                        Dim Object_prompt_poly As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_prompt_poly.MessageForAdding = vbLf & "Select the polyline"

                        Object_prompt_poly.SingleOnly = True
                        Rezultat_poly = Editor1.GetSelection(Object_prompt_poly)
                        If Rezultat_poly.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Dim PolyCL As Autodesk.AutoCAD.DatabaseServices.Polyline



                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat_poly.Value.Item(0)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)






                            Dim index_excel As Integer = CInt(TextBox_start_row_p2e.Text)
                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                PolyCL = Ent1
                                For i = 0 To PolyCL.NumberOfVertices - 1
                                    Dim Point1 As Point3d
                                    Point1 = PolyCL.GetPoint3dAt(i)





                                    W1.Range(Column_X & index_excel).Value2 = Round(Point1.X, 2)

                                    W1.Range(Column_Y & index_excel).Value2 = Round(Point1.Y, 2)

                                    index_excel = index_excel + 1

                                Next

                            End If









                        End If



                    End Using
                End Using


            Catch ex As System.Exception
                MsgBox(ex.Message)
            End Try


            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_writeXY_Click(sender As Object, e As EventArgs) Handles Button_read_STA_WRITE_X_Y.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                Try
                    Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                    W1 = Get_active_worksheet_from_Excel()
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)
                    Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim start1 As Double = 100
                            Dim end1 As Double = 99
                            If IsNumeric(TextBox_start_row_p2e.Text) = True Then
                                start1 = CDbl(TextBox_start_row_p2e.Text)
                            End If
                            If IsNumeric(TextBox_end_row_p2e.Text) = True Then
                                end1 = CDbl(TextBox_end_row_p2e.Text)
                            End If

                            If start1 > end1 Then
                                MsgBox("Start is bigger than end row")
                                Freeze_operations = False
                                Exit Sub
                            End If


                            Dim Column_sta As String = TextBox_column_Stationp2e.Text.ToUpper
                            Dim Column_ELEV As String = TextBox_column_Elevationp2e.Text.ToUpper
                            Dim Column_X As String = TextBox_column_x.Text.ToUpper
                            Dim Column_Y As String = TextBox_column_y.Text.ToUpper
                            If CheckBox_header.Checked = True And start1 > 1 Then
                                If Not Column_sta = "" Then
                                    W1.Range(Column_sta & start1 - 1).Value2 = "STATION"
                                End If
                                If Not Column_ELEV = "" Then
                                    W1.Range(Column_ELEV & start1 - 1).Value2 = "ELEVATION"
                                End If
                                If Not Column_X = "" Then
                                    W1.Range(Column_X & start1 - 1).Value2 = "EASTING"
                                End If
                                If Not Column_Y = "" Then
                                    W1.Range(Column_Y & start1 - 1).Value2 = "NORTHTING"
                                End If
                            End If



                            Dim Rezultat_poly As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                            Dim Object_prompt_poly As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                            Object_prompt_poly.MessageForAdding = vbLf & "Select the polyline"

                            Object_prompt_poly.SingleOnly = True
                            Rezultat_poly = Editor1.GetSelection(Object_prompt_poly)
                            If Rezultat_poly.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Dim PolyCL As Autodesk.AutoCAD.DatabaseServices.Polyline



                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Rezultat_poly.Value.Item(0)
                                Dim Ent1 As Entity
                                Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)






                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                    PolyCL = Ent1

                                    For i = start1 To end1
                                        Dim Point1 As Point3d
                                        Dim Station1 As String = Replace(W1.Range(Column_sta & i).Value2, "+", "")
                                        If IsNumeric(Station1) = True Then
                                            Dim Sta1 As Double = CDbl(Station1)
                                            If Sta1 > PolyCL.Length Then
                                                If Abs(Sta1 - PolyCL.Length) < 1 Then
                                                    Sta1 = PolyCL.Length
                                                Else
                                                    MsgBox("The polyline is shorter than the Station!")
                                                    Freeze_operations = False
                                                    W1.Range(Column_sta & i).Interior.ColorIndex = 7
                                                    W1.Range(Column_sta & i).Select()
                                                    Exit Sub
                                                End If
                                            End If

                                            Point1 = PolyCL.GetPointAtDist(Sta1)
                                            W1.Range(Column_X & i).Value2 = Round(Point1.X, 2)
                                            W1.Range(Column_Y & i).Value2 = Round(Point1.Y, 2)
                                        Else
                                            W1.Range(Column_sta & i).Interior.ColorIndex = 7
                                        End If




                                    Next

                                End If









                            End If



                        End Using
                    End Using


                Catch ex As System.Exception
                    MsgBox(ex.Message)
                End Try


            Catch ex As System.SystemException
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



            Catch ex As System.Exception
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try
            Freeze_operations = False
        End If

    End Sub

    Private Sub CheckBox_USA_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_USA.CheckedChanged
        If CheckBox_USA.Checked = True Then
            Label_st1.Text = "Station attribute"
            Label_st2.Text = "Station attribute" & vbCrLf & "down"
        Else
            Label_st1.Text = "Chainage attribute"
            Label_st2.Text = "Chainage attribute" & vbCrLf & "down"
        End If
    End Sub

    Private Sub Button_draw_rectang_Click(sender As Object, e As EventArgs) Handles Button_draw_rectang.Click
        If Freeze_operations = False Then
            Freeze_operations = True

            Try

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Dim Start1 As Integer = 0
                Dim End1 As Integer = 0
                If IsNumeric(TextBox_rect_row_start.Text) = True Then
                    Start1 = CInt(TextBox_rect_row_start.Text)
                End If
                If IsNumeric(TextBox_rect_row_end.Text) = True Then
                    End1 = CInt(TextBox_rect_row_end.Text)
                End If

                If End1 = 0 Or Start1 = 0 Then
                    Freeze_operations = False
                    Exit Sub
                End If
                If End1 < Start1 Then
                    Freeze_operations = False
                    Exit Sub
                End If
                Dim Column_start As String = ""
                Column_start = TextBox_rect_col_start.Text.ToUpper
                Dim Column_end As String = ""
                Column_end = TextBox_rect_col_end.Text.ToUpper

                Dim Column_label As String = ""
                Column_label = TextBox_rect_label.Text.ToUpper



                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                Freeze_operations = True


                Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument




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
                            Exit Sub
                        End If

                        If Rezultat_hline.Value.Count <> 2 Then
                            MsgBox("Your selection contains " & Rezultat_hline.Value.Count & " objects")
                            Editor1.SetImpliedSelection(Empty_array)
                            Freeze_operations = False
                            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If

                        Dim Chainage_cunoscuta As Double = -100000



                        Dim Rezultat__elev_lines As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                        Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_Prompt2.MessageForAdding = vbLf & "Select the top and bottom of graph:"
                        Object_Prompt2.SingleOnly = False
                        Rezultat__elev_lines = Editor1.GetSelection(Object_Prompt2)

                        Dim Ymin, Ymax As Double
                        Dim Assigned As Boolean = False

                        If Rezultat__elev_lines.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            For i = 1 To Rezultat__elev_lines.Value.Count
                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Rezultat__elev_lines.Value.Item(i - 1)
                                Dim Ent1 As Entity
                                Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                                    Dim Linie1 As Line = Ent1
                                    If Abs(Linie1.StartPoint.Y - Linie1.EndPoint.Y) < 0.02 Then
                                        If i = 1 Then
                                            Assigned = True
                                            Ymin = Linie1.StartPoint.Y
                                            Ymax = Linie1.StartPoint.Y
                                        Else
                                            If Linie1.StartPoint.Y > Ymax Then Ymax = Linie1.StartPoint.Y
                                            If Linie1.StartPoint.Y < Ymin Then Ymin = Linie1.StartPoint.Y
                                        End If
                                    End If
                                End If
                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                    Dim PLinie1 As Polyline = Ent1
                                    If PLinie1.StartPoint.Y = PLinie1.EndPoint.Y And PLinie1.NumberOfVertices = 2 Then
                                        If i = 1 And Assigned = False Then
                                            Ymin = PLinie1.StartPoint.Y
                                            Ymax = PLinie1.StartPoint.Y
                                        Else
                                            If PLinie1.StartPoint.Y > Ymax Then Ymax = PLinie1.StartPoint.Y
                                            If PLinie1.StartPoint.Y < Ymin Then Ymin = PLinie1.StartPoint.Y
                                        End If
                                    End If
                                End If
                            Next
                        Else
                            Ymin = 0
                            Ymax = 10
                        End If






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
                            If IsNumeric(Replace(mText_cunoscut_chainage.Text, "+", "")) = True Then Chainage_cunoscuta = CDbl(Replace(mText_cunoscut_chainage.Text, "+", ""))
                        End If

                        If TypeOf Ent3_chainage Is Autodesk.AutoCAD.DatabaseServices.MText Then
                            mText_cunoscut_chainage = Ent3_chainage
                            If IsNumeric(Replace(mText_cunoscut_chainage.Text, "+", "")) = True Then Chainage_cunoscuta = CDbl(Replace(mText_cunoscut_chainage.Text, "+", ""))
                        End If

                        If TypeOf Ent2_chainage Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                            Text_cunoscut_chainage = Ent2_chainage
                            If IsNumeric(Replace(Text_cunoscut_chainage.TextString, "+", "")) = True Then Chainage_cunoscuta = CDbl(Replace(Text_cunoscut_chainage.TextString, "+", ""))
                        End If

                        If TypeOf Ent3_chainage Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                            Text_cunoscut_chainage = Ent3_chainage
                            If IsNumeric(Replace(Text_cunoscut_chainage.TextString, "+", "")) = True Then Chainage_cunoscuta = CDbl(Replace(Text_cunoscut_chainage.TextString, "+", ""))
                        End If

                        If Chainage_cunoscuta = -100000 Then
                            MsgBox("Chainage datum not numeric")
                            Editor1.SetImpliedSelection(Empty_array)
                            Freeze_operations = False
                            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If


                        Dim Linia_cunoscuta As Autodesk.AutoCAD.DatabaseServices.Line
                        Dim x01, x02 As Double
                        Dim y01, y02 As Double

                        If TypeOf Ent2_chainage Is Autodesk.AutoCAD.DatabaseServices.Line Then
                            Linia_cunoscuta = Ent2_chainage
                            x01 = Linia_cunoscuta.StartPoint.X 'TransformBy(UCS_CURENT).X
                            x02 = Linia_cunoscuta.EndPoint.X 'TransformBy(UCS_CURENT).X
                            y01 = Linia_cunoscuta.StartPoint.Y 'TransformBy(UCS_CURENT).Y
                            y02 = Linia_cunoscuta.EndPoint.Y ' TransformBy(UCS_CURENT).Y
                            If Abs(x01 - x02) > 0.001 Then
                                MsgBox("Vertical line you selected is not vertical")
                                Editor1.SetImpliedSelection(Empty_array)
                                Freeze_operations = False
                                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                Exit Sub
                            End If

                        End If

                        If TypeOf Ent3_chainage Is Autodesk.AutoCAD.DatabaseServices.Line Then
                            Linia_cunoscuta = Ent3_chainage
                            x01 = Linia_cunoscuta.StartPoint.X 'TransformBy(UCS_CURENT).X
                            x02 = Linia_cunoscuta.EndPoint.X 'TransformBy(UCS_CURENT).X
                            y01 = Linia_cunoscuta.StartPoint.Y 'TransformBy(UCS_CURENT).Y
                            y02 = Linia_cunoscuta.EndPoint.Y 'TransformBy(UCS_CURENT).Y
                            If Abs(x01 - x02) > 0.001 Then
                                MsgBox("Vertical line you selected is not vertical")
                                Editor1.SetImpliedSelection(Empty_array)
                                Freeze_operations = False
                                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                Exit Sub
                            End If

                        End If

                        Creaza_layer("NO PLOT", 40, "", False)































                        For i = Start1 To End1
                            Dim Station_start As String = W1.Range(Column_start & i).Value2
                            Dim Station_end As String = W1.Range(Column_end & i).Value2
                            Dim rect_label As String = W1.Range(Column_label & i.ToString).Value2

                            If IsNumeric(Station_end) = True And IsNumeric(Station_start) = True And rect_label <> "" Then
                                Dim Sta1 As Double = CDbl(Station_start)
                                Dim Sta2 As Double = CDbl(Station_end)

                                Dim x1 As Double
                                If Radio_LEFT_R.Checked = True Then
                                    x1 = x01 + (Sta1 - Chainage_cunoscuta) * H_EXAG
                                Else
                                    x1 = x01 - (Sta1 - Chainage_cunoscuta) * H_EXAG
                                End If

                                Dim x2 As Double
                                If Radio_LEFT_R.Checked = True Then
                                    x2 = x01 + (Sta2 - Chainage_cunoscuta) * H_EXAG
                                Else
                                    x2 = x01 - (Sta2 - Chainage_cunoscuta) * H_EXAG
                                End If


                                Dim Vlength As Double = Abs(Ymax - Ymin)
                                Dim Poly1 As New Polyline
                                Poly1.Layer = ComboBox_Rectangle_LAYER.Text


                                Poly1.AddVertexAt(0, New Point2d(x1, Ymin - Vlength / 6), 0, 0, 0)
                                Poly1.AddVertexAt(1, New Point2d(x1, Ymax + Vlength / 6), 0, 0, 0)
                                Poly1.AddVertexAt(2, New Point2d(x2, Ymax + Vlength / 6), 0, 0, 0)
                                Poly1.AddVertexAt(3, New Point2d(x2, Ymin - Vlength / 6), 0, 0, 0)
                                Poly1.Closed = True
                                Poly1.Elevation = 0
                                BTrecord.AppendEntity(Poly1)
                                Trans1.AddNewlyCreatedDBObject(Poly1, True)


                                Dim Mtext1 As New MText
                                Mtext1.Location = New Point3d(x1, Ymax + Vlength / 6 + 2, 0)
                                Mtext1.Layer = ComboBox_Rectangle_LAYER.Text
                                Mtext1.TextHeight = 8
                                Mtext1.Contents = rect_label

                                BTrecord.AppendEntity(Mtext1)
                                Trans1.AddNewlyCreatedDBObject(Mtext1, True)
                            Else
                                MsgBox("non numerical values on row " & i)
                                W1.Rows(i).select()
                                Freeze_operations = False
                                Exit Sub

                            End If
                        Next



                        'MsgBox(Data_table_Centerline.Rows.Count)
                        Trans1.Commit()
                        ' asta e de la tranzactie
                    End Using




                    ' asta e de la lock
                End Using


                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")


            Catch ex As System.Exception
                MsgBox(ex.Message)
                Freeze_operations = False
            End Try
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_load_profile_curve_Click(sender As System.Object, e As System.EventArgs) Handles Button_load_profile_curve.Click
        Try

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor

            If Freeze_operations = False Then

                Freeze_operations = True


                Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument



                    Dim Sta_ref As Double = 0
                    If IsNumeric(TextBox_Station_reference.Text) = True Then
                        Sta_ref = CDbl(TextBox_Station_reference.Text)
                    End If
                    Dim Interval As Double = 0
                    If IsNumeric(TextBox_interval.Text) = True Then
                        Interval = CDbl(TextBox_interval.Text)
                    End If

                    If Interval = 0 Then
                        MsgBox("please specify the interval")
                        Freeze_operations = False
                        Exit Sub
                    End If

                    ' Dim k As Double = 1
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction



                        Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)



                        Dim Result_point0 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                        Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Please specify the reference point position:")

                        PP1.AllowNone = False
                        Result_point0 = Editor1.GetPoint(PP1)
                        If Result_point0.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                            Freeze_operations = False
                            Exit Sub
                        End If






                        Dim Empty_array() As ObjectId
                        Editor1.SetImpliedSelection(Empty_array)

                        Dim UCS_CURENT As Matrix3d = Editor1.CurrentUserCoordinateSystem


                        '****************************************************************************************





                        Dim Rezultat_hlineSCALE As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify Horizontal Exaggeration:")
                        Rezultat_hlineSCALE.DefaultValue = 1
                        Rezultat_hlineSCALE.AllowNone = True
                        Dim Rezultat_hline44 As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(Rezultat_hlineSCALE)

                        Dim H_EXAG As Double = Rezultat_hline44.Value
                        If H_EXAG = 0 Then H_EXAG = 1



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


                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_vert.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                        Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj2 = Rezultat_vert.Value.Item(1)
                        Dim Ent2 As Entity
                        Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)

                        Dim mText_cunoscut As Autodesk.AutoCAD.DatabaseServices.MText
                        Dim Text_cunoscut As Autodesk.AutoCAD.DatabaseServices.DBText

                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                            mText_cunoscut = Ent1
                            If IsNumeric(Replace(mText_cunoscut.Text, "'", "")) = True Then Elevatia_cunoscuta = CDbl(Replace(mText_cunoscut.Text, "'", ""))
                        End If

                        If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                            mText_cunoscut = Ent2
                            If IsNumeric(Replace(mText_cunoscut.Text, "'", "")) = True Then Elevatia_cunoscuta = CDbl(Replace(mText_cunoscut.Text, "'", ""))
                        End If

                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                            Text_cunoscut = Ent1
                            If IsNumeric(Replace(Text_cunoscut.TextString, "'", "")) = True Then Elevatia_cunoscuta = CDbl(Replace(Text_cunoscut.TextString, "'", ""))
                        End If

                        If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                            Text_cunoscut = Ent2
                            If IsNumeric(Replace(Text_cunoscut.TextString, "'", "")) = True Then Elevatia_cunoscuta = CDbl(Replace(Text_cunoscut.TextString, "'", ""))
                        End If

                        Dim Linia_Elev_cunoscuta As Autodesk.AutoCAD.DatabaseServices.Line
                        Dim polylinia_elev_cunoscuta As Autodesk.AutoCAD.DatabaseServices.Polyline

                        Dim x0e1, y0e1, x0e2, y0e2 As Double

                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                            Linia_Elev_cunoscuta = Ent1
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

                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                            polylinia_elev_cunoscuta = Ent1

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







                        Dim Rezultat_poly_graph As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                        Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_Prompt1.MessageForAdding = vbLf & "Select the Graph profile objects:"
                        Object_Prompt1.SingleOnly = False
                        Rezultat_poly_graph = Editor1.GetSelection(Object_Prompt1)

                        If Not Rezultat_poly_graph.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Editor1.SetImpliedSelection(Empty_array)
                            Freeze_operations = False
                            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                            Freeze_operations = False
                            Exit Sub
                        End If


                        Data_table_Profile3D = New System.Data.DataTable
                        Data_table_Profile3D.Columns.Add("StaH", GetType(Double))
                        Data_table_Profile3D.Columns.Add("Elev", GetType(Double))

                        For j = 0 To Rezultat_poly_graph.Value.Count - 1

                            Dim Ent_profile As Entity
                            Ent_profile = Trans1.GetObject(Rezultat_poly_graph.Value.Item(j).ObjectId, OpenMode.ForRead)


                            If TypeOf Ent_profile Is Autodesk.AutoCAD.DatabaseServices.Line Then
                                Dim Linie1 As Line = Ent_profile
                                Dim Xstart As Double = Linie1.StartPoint.X
                                Dim Xend As Double = Linie1.EndPoint.X
                                If Xstart > Xend Then
                                    Dim t As Double = Xstart
                                    Xstart = Xend
                                    Xend = t
                                End If

                                Dim LIne_int0 As New Line(New Point3d(Xstart, -1000000000, 0), New Point3d(Xstart, 1000000000, 0))
                                Dim col_int0 As New Point3dCollection
                                col_int0 = Intersect_on_both_operands(LIne_int0, Linie1)
                                If col_int0.Count = 1 Then
                                    Data_table_Profile3D.Rows.Add()
                                    Data_table_Profile3D.Rows(Data_table_Profile3D.Rows.Count - 1).Item("StaH") = Sta_ref - (Result_point0.Value.X - col_int0(0).X) / H_EXAG
                                    Data_table_Profile3D.Rows(Data_table_Profile3D.Rows.Count - 1).Item("Elev") = Elevatia_cunoscuta + (col_int0(0).Y - Y_elev_cunoscut) / V_EXAG
                                End If



                                Dim LIne_int2 As New Line(New Point3d(Xend, -1000000000, 0), New Point3d(Xend, 1000000000, 0))
                                Dim col_int2 As New Point3dCollection
                                col_int2 = Intersect_on_both_operands(LIne_int2, Linie1)
                                If col_int2.Count = 1 Then
                                    Data_table_Profile3D.Rows.Add()
                                    Data_table_Profile3D.Rows(Data_table_Profile3D.Rows.Count - 1).Item("StaH") = Sta_ref - (Result_point0.Value.X - col_int2(0).X) / H_EXAG
                                    Data_table_Profile3D.Rows(Data_table_Profile3D.Rows.Count - 1).Item("Elev") = Elevatia_cunoscuta + (col_int2(0).Y - Y_elev_cunoscut) / V_EXAG
                                End If





                            ElseIf TypeOf Ent_profile Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                Dim Poly1 As Polyline = Ent_profile


                                For i = 0 To Poly1.NumberOfVertices - 1

                                    If Poly1.GetSegmentType(i) = SegmentType.Line Then
                                        Dim Line1 As LineSegment3d = Poly1.GetLineSegmentAt(i)
                                        Dim Xstart As Double = Line1.StartPoint.X


                                        Dim LIne_int0 As New Line(New Point3d(Xstart, -1000000000, 0), New Point3d(Xstart, 1000000000, 0))
                                        Dim col_int0 As New Point3dCollection
                                        col_int0 = Intersect_on_both_operands(LIne_int0, Poly1)
                                        If col_int0.Count = 1 Then
                                            Data_table_Profile3D.Rows.Add()
                                            Data_table_Profile3D.Rows(Data_table_Profile3D.Rows.Count - 1).Item("StaH") = Sta_ref - (Result_point0.Value.X - col_int0(0).X) / H_EXAG
                                            Data_table_Profile3D.Rows(Data_table_Profile3D.Rows.Count - 1).Item("Elev") = Elevatia_cunoscuta + (col_int0(0).Y - Y_elev_cunoscut) / V_EXAG
                                        End If

                                        If i = Poly1.NumberOfVertices - 1 Then
                                            Dim Xend As Double = Line1.EndPoint.X
                                            Dim LIne_int2 As New Line(New Point3d(Xend, -1000000000, 0), New Point3d(Xend, 1000000000, 0))
                                            Dim col_int2 As New Point3dCollection
                                            col_int2 = Intersect_on_both_operands(LIne_int2, Poly1)
                                            If col_int2.Count = 1 Then
                                                Data_table_Profile3D.Rows.Add()
                                                Data_table_Profile3D.Rows(Data_table_Profile3D.Rows.Count - 1).Item("StaH") = Sta_ref - (Result_point0.Value.X - col_int2(0).X) / H_EXAG
                                                Data_table_Profile3D.Rows(Data_table_Profile3D.Rows.Count - 1).Item("Elev") = Elevatia_cunoscuta + (col_int2(0).Y - Y_elev_cunoscut) / V_EXAG
                                            End If

                                        End If
                                    End If

                                    If Poly1.GetSegmentType(i) = SegmentType.Arc Then
                                        Dim Arc1 As CircularArc2d = Poly1.GetArcSegment2dAt(i)
                                        Dim Xstart As Double = Arc1.StartPoint.X
                                        Dim Xend As Double = Arc1.EndPoint.X
                                        Dim LIne_int0 As New Line(New Point3d(Xstart, -1000000000, 0), New Point3d(Xstart, 1000000000, 0))
                                        Dim col_int0 As New Point3dCollection
                                        col_int0 = Intersect_on_both_operands(LIne_int0, Poly1)
                                        If col_int0.Count = 1 Then
                                            Data_table_Profile3D.Rows.Add()
                                            Data_table_Profile3D.Rows(Data_table_Profile3D.Rows.Count - 1).Item("StaH") = Sta_ref - (Result_point0.Value.X - col_int0(0).X) / H_EXAG
                                            Data_table_Profile3D.Rows(Data_table_Profile3D.Rows.Count - 1).Item("Elev") = Elevatia_cunoscuta + (col_int0(0).Y - Y_elev_cunoscut) / V_EXAG
                                        End If
                                        Dim No_vert_int_line As Integer = Floor(Poly1.GetDistanceAtParameter(i + 1) - Poly1.GetDistanceAtParameter(i) / Interval)
                                        For k = 1 To No_vert_int_line
                                            Dim LIne_int1 As New Line(New Point3d(Xstart + k * Interval, -1000000000, 0), New Point3d(Xstart + k * Interval, 1000000000, 0))
                                            Dim col_int As New Point3dCollection
                                            col_int = Intersect_on_both_operands(LIne_int1, Poly1)
                                            If col_int.Count = 1 Then
                                                Data_table_Profile3D.Rows.Add()
                                                Data_table_Profile3D.Rows(Data_table_Profile3D.Rows.Count - 1).Item("StaH") = Sta_ref - (Result_point0.Value.X - col_int(0).X) / H_EXAG
                                                Data_table_Profile3D.Rows(Data_table_Profile3D.Rows.Count - 1).Item("Elev") = Elevatia_cunoscuta + (col_int(0).Y - Y_elev_cunoscut) / V_EXAG



                                            End If




                                        Next

                                        Dim LIne_int2 As New Line(New Point3d(Xend, -1000000000, 0), New Point3d(Xend, 1000000000, 0))
                                        Dim col_int2 As New Point3dCollection
                                        col_int2 = Intersect_on_both_operands(LIne_int2, Poly1)
                                        If col_int2.Count = 1 Then
                                            Data_table_Profile3D.Rows.Add()
                                            Data_table_Profile3D.Rows(Data_table_Profile3D.Rows.Count - 1).Item("StaH") = Sta_ref - (Result_point0.Value.X - col_int2(0).X) / H_EXAG
                                            Data_table_Profile3D.Rows(Data_table_Profile3D.Rows.Count - 1).Item("Elev") = Elevatia_cunoscuta + (col_int2(0).Y - Y_elev_cunoscut) / V_EXAG
                                        End If

                                    End If

                                Next












                            ElseIf TypeOf Ent_profile Is Autodesk.AutoCAD.DatabaseServices.Curve Then
                                Dim Curva1 As Curve = Ent_profile
                                Dim Xstart As Double = Curva1.StartPoint.X
                                Dim Xend As Double = Curva1.EndPoint.X
                                If Xstart > Xend Then
                                    Dim t As Double = Xstart
                                    Xstart = Xend
                                    Xend = t
                                End If

                                Dim LIne_int0 As New Line(New Point3d(Xstart, -1000000000, 0), New Point3d(Xstart, 1000000000, 0))
                                Dim col_int0 As New Point3dCollection
                                col_int0 = Intersect_on_both_operands(LIne_int0, Curva1)
                                If col_int0.Count = 1 Then
                                    Data_table_Profile3D.Rows.Add()
                                    Data_table_Profile3D.Rows(Data_table_Profile3D.Rows.Count - 1).Item("StaH") = Sta_ref - (Result_point0.Value.X - col_int0(0).X) / H_EXAG
                                    Data_table_Profile3D.Rows(Data_table_Profile3D.Rows.Count - 1).Item("Elev") = Elevatia_cunoscuta + (col_int0(0).Y - Y_elev_cunoscut) / V_EXAG
                                End If


                                Dim No_vert_int_line As Integer = Floor((Xend - Xstart) / Interval)

                                For k = 1 To No_vert_int_line
                                    Dim LIne_int1 As New Line(New Point3d(Xstart + k * Interval, -1000000000, 0), New Point3d(Xstart + k * Interval, 1000000000, 0))
                                    Dim col_int As New Point3dCollection
                                    col_int = Intersect_on_both_operands(LIne_int1, Curva1)
                                    If col_int.Count = 1 Then
                                        Data_table_Profile3D.Rows.Add()
                                        Data_table_Profile3D.Rows(Data_table_Profile3D.Rows.Count - 1).Item("StaH") = Sta_ref - (Result_point0.Value.X - col_int(0).X) / H_EXAG
                                        Data_table_Profile3D.Rows(Data_table_Profile3D.Rows.Count - 1).Item("Elev") = Elevatia_cunoscuta + (col_int(0).Y - Y_elev_cunoscut) / V_EXAG



                                    End If




                                Next

                                Dim LIne_int2 As New Line(New Point3d(Xend, -1000000000, 0), New Point3d(Xend, 1000000000, 0))
                                Dim col_int2 As New Point3dCollection
                                col_int2 = Intersect_on_both_operands(LIne_int2, Curva1)
                                If col_int2.Count = 1 Then
                                    Data_table_Profile3D.Rows.Add()
                                    Data_table_Profile3D.Rows(Data_table_Profile3D.Rows.Count - 1).Item("StaH") = Sta_ref - (Result_point0.Value.X - col_int2(0).X) / H_EXAG
                                    Data_table_Profile3D.Rows(Data_table_Profile3D.Rows.Count - 1).Item("Elev") = Elevatia_cunoscuta + (col_int2(0).Y - Y_elev_cunoscut) / V_EXAG
                                End If



                            End If





                        Next







                        Trans1.Commit()
                        ' asta e de la tranzactie
                    End Using


                    Data_table_Profile3D = Sort_data_table(Data_table_Profile3D, "StaH")

                    Transfer_datatable_to_new_excel_spreadsheet(Data_table_Profile3D)


                    Freeze_operations = False



                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    ' asta e de la lock
                End Using
                Freeze_operations = False
            End If
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

        Catch ex As System.Exception
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

            Freeze_operations = False
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_draw3d_poly_Click(sender As Object, e As EventArgs) Handles Button_draw3d_poly.Click
        Try

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor

            If IsNothing(Data_table_Profile3D) = True Then
                MsgBox("No data loaded")
                Exit Sub
            End If
            If Data_table_Profile3D.Rows.Count = 0 Then
                MsgBox("No data loaded")
                Exit Sub
            End If

            If Freeze_operations = False Then

                Freeze_operations = True


                Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument



                    Dim Sta_ref As Double = 0
                    If IsNumeric(TextBox_Station_reference.Text) = True Then
                        Sta_ref = CDbl(TextBox_Station_reference.Text)
                    End If


                    ' Dim k As Double = 1
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction



                        Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)



                        Dim Result_point0 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                        Dim PP0 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Please specify the reference point position:")

                        PP0.AllowNone = False
                        Result_point0 = Editor1.GetPoint(PP0)
                        If Result_point0.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                            Freeze_operations = False
                            Exit Sub
                        End If

                        Dim Result_point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                        Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Please specify the alignment first point")

                        PP1.AllowNone = False
                        Result_point1 = Editor1.GetPoint(PP1)
                        If Result_point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                            Freeze_operations = False
                            Exit Sub
                        End If

                        Dim Result_point2 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                        Dim PP2 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Please specify the alignment second point:")
                        PP2.UseBasePoint = True
                        PP2.BasePoint = Result_point1.Value
                        PP2.AllowNone = False
                        Result_point2 = Editor1.GetPoint(PP2)
                        If Result_point2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                            Freeze_operations = False
                            Exit Sub
                        End If



                        Dim Empty_array() As ObjectId
                        Editor1.SetImpliedSelection(Empty_array)

                        Dim UCS_CURENT As Matrix3d = Editor1.CurrentUserCoordinateSystem


                        '****************************************************************************************
                        Dim p0 As New Point3d(Result_point0.Value.X, Result_point0.Value.Y, 0)
                        Dim p1 As New Point3d(Result_point1.Value.X, Result_point1.Value.Y, 0)
                        Dim p2 As New Point3d(Result_point2.Value.X, Result_point2.Value.Y, 0)

                        Dim Poly1 As New Polyline
                        Poly1.AddVertexAt(0, New Point2d(0, 0), 0, 0, 0)
                        Poly1.AddVertexAt(1, New Point2d(10000000, 0), 0, 0, 0)




                        Poly1.TransformBy(Matrix3d.Displacement(Poly1.GetPointAtParameter(0.5).GetVectorTo(p0)))
                        Poly1.TransformBy(Matrix3d.Rotation(GET_Bearing_rad(p1.X, p1.Y, p2.X, p2.Y), Vector3d.ZAxis, p0))
                        Dim Point0_on_poly As New Point3d

                        Point0_on_poly = Poly1.GetClosestPointTo(p0, Vector3d.ZAxis, False)

                        Dim Sta0 As Double = Poly1.GetDistAtPoint(Point0_on_poly)


                        Dim Poly3D As New Polyline3d
                        Poly3D.SetDatabaseDefaults()
                        BTrecord.AppendEntity(Poly3D)
                        Trans1.AddNewlyCreatedDBObject(Poly3D, True)


                        For i = 0 To Data_table_Profile3D.Rows.Count - 1

                            Dim Sta As Double = Data_table_Profile3D.Rows(i).Item("StaH")

                            Dim Vertex_new As New PolylineVertex3d(New Point3d(Poly1.GetPointAtDist(Sta0 + Sta - Sta_ref).X, Poly1.GetPointAtDist(Sta0 + Sta - Sta_ref).Y, Data_table_Profile3D.Rows(i).Item("Elev")))
                            Poly3D.AppendVertex(Vertex_new)
                            Trans1.AddNewlyCreatedDBObject(Vertex_new, True)

                        Next






                        Trans1.Commit()
                        ' asta e de la tranzactie
                    End Using







                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    ' asta e de la lock
                End Using
                Freeze_operations = False
            End If


        Catch ex As System.Exception
            Freeze_operations = False
            MsgBox(ex.Message)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

    End Sub

    Private Sub Button_adjust_labels_on_profile_poly_Click(sender As Object, e As EventArgs) Handles Button_adjust_labels_on_profile_poly.Click
        Try

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor



            If Freeze_operations = False Then

                Freeze_operations = True


                Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument





                    ' Dim k As Double = 1
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction



                        Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)







                        Dim Empty_array() As ObjectId
                        Editor1.SetImpliedSelection(Empty_array)

                        Dim UCS_CURENT As Matrix3d = Editor1.CurrentUserCoordinateSystem

                        Dim Rezultat0 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                        Dim Object_Prompt0 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_Prompt0.MessageForAdding = vbLf & "Select profile graph polyline:"

                        Object_Prompt0.SingleOnly = True

                        Rezultat0 = Editor1.GetSelection(Object_Prompt0)


                        If Rezultat0.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Editor1.WriteMessage(vbLf & "Command:")
                            Freeze_operations = False
                            Exit Sub
                        End If

                        Dim Poly0 As Polyline = TryCast(Trans1.GetObject(Rezultat0.Value(0).ObjectId, OpenMode.ForRead), Polyline)

                        If IsNothing(Poly0) = False Then
                            Editor1.SetImpliedSelection(Empty_array)


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







                            For i = 0 To Rezultat1.Value.Count - 1
                                Dim Ent1 As Entity = Trans1.GetObject(Rezultat1.Value(i).ObjectId, OpenMode.ForWrite)
                                If TypeOf Ent1 Is BlockReference Then
                                    Dim block1 As BlockReference = Ent1
                                    Dim Poly1 As New Polyline
                                    Poly1.AddVertexAt(0, New Point2d(block1.Position.X, -1000000), 0, 0, 0)
                                    Poly1.AddVertexAt(1, New Point2d(block1.Position.X, 1000000), 0, 0, 0)
                                    Poly1.Elevation = Poly0.Elevation
                                    Dim Colint As New Point3dCollection
                                    Colint = Intersect_on_both_operands(Poly0, Poly1)
                                    If Colint.Count > 0 Then
                                        block1.TransformBy(Matrix3d.Displacement(block1.Position.GetVectorTo(Colint(0))))



                                    End If



                                End If

                                If TypeOf Ent1 Is MLeader Then
                                    Dim Mleader1 As MLeader = Ent1
                                    Dim Poly1 As New Polyline


                                    Poly1.AddVertexAt(0, New Point2d(Mleader1.GetFirstVertex(0).X, -1000000), 0, 0, 0)
                                    Poly1.AddVertexAt(1, New Point2d(Mleader1.GetFirstVertex(0).X, 1000000), 0, 0, 0)
                                    Poly1.Elevation = Poly0.Elevation
                                    Dim Colint As New Point3dCollection
                                    Colint = Intersect_on_both_operands(Poly0, Poly1)
                                    If Colint.Count > 0 Then
                                        Mleader1.TransformBy(Matrix3d.Displacement(Mleader1.GetFirstVertex(0).GetVectorTo(Colint(0))))
                                    End If



                                End If


                            Next
                        End If










                        Trans1.Commit()
                        ' asta e de la tranzactie
                    End Using







                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    ' asta e de la lock
                End Using
                Freeze_operations = False
            End If


        Catch ex As System.Exception
            Freeze_operations = False
            MsgBox(ex.Message)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

    End Sub


    Private Sub Button_W_blocks_to_xl_Click(sender As Object, e As EventArgs) Handles Button_W_blocks_to_xl.Click
        Try

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor

            If Freeze_operations = False Then

                Freeze_operations = True
                Dim Data_table_Blocks As System.Data.DataTable = New System.Data.DataTable
                Data_table_Blocks.Columns.Add("StaH", GetType(Double))
                Data_table_Blocks.Columns.Add("BlockName", GetType(String))

                Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument




                    ' Dim k As Double = 1
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction



                        Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)









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

                        Dim Station_cunoscuta As Double = -100000

                        


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
                            If IsNumeric(Replace(mText_cunoscut_chainage.Text, "+", "")) = True Then Station_cunoscuta = CDbl(Replace(mText_cunoscut_chainage.Text, "+", ""))
                        End If

                        If TypeOf Ent3_chainage Is Autodesk.AutoCAD.DatabaseServices.MText Then
                            mText_cunoscut_chainage = Ent3_chainage
                            If IsNumeric(Replace(mText_cunoscut_chainage.Text, "+", "")) = True Then Station_cunoscuta = CDbl(Replace(mText_cunoscut_chainage.Text, "+", ""))
                        End If

                        If TypeOf Ent2_chainage Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                            Text_cunoscut_chainage = Ent2_chainage
                            If IsNumeric(Replace(Text_cunoscut_chainage.TextString, "+", "")) = True Then Station_cunoscuta = CDbl(Replace(Text_cunoscut_chainage.TextString, "+", ""))
                        End If

                        If TypeOf Ent3_chainage Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                            Text_cunoscut_chainage = Ent3_chainage
                            If IsNumeric(Replace(Text_cunoscut_chainage.TextString, "+", "")) = True Then Station_cunoscuta = CDbl(Replace(Text_cunoscut_chainage.TextString, "+", ""))
                        End If

                        If Station_cunoscuta = -100000 Then
                            MsgBox("Chainage datum not numeric")
                            Editor1.SetImpliedSelection(Empty_array)
                            Freeze_operations = False
                            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                            Freeze_operations = False
                            Exit Sub
                        End If


                        Dim Linia_cunoscuta As Autodesk.AutoCAD.DatabaseServices.Line
                        Dim X_sta_cunoscut, x0_sta2 As Double
                        Dim y0_sta1, y0_sta2 As Double

                        If TypeOf Ent2_chainage Is Autodesk.AutoCAD.DatabaseServices.Line Then
                            Linia_cunoscuta = Ent2_chainage
                            X_sta_cunoscut = Linia_cunoscuta.StartPoint.X 'TransformBy(UCS_CURENT).X
                            x0_sta2 = Linia_cunoscuta.EndPoint.X 'TransformBy(UCS_CURENT).X
                            y0_sta1 = Linia_cunoscuta.StartPoint.Y 'TransformBy(UCS_CURENT).Y
                            y0_sta2 = Linia_cunoscuta.EndPoint.Y ' TransformBy(UCS_CURENT).Y
                            If Abs(X_sta_cunoscut - x0_sta2) > 0.001 Then
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
                            X_sta_cunoscut = Linia_cunoscuta.StartPoint.X 'TransformBy(UCS_CURENT).X
                            x0_sta2 = Linia_cunoscuta.EndPoint.X 'TransformBy(UCS_CURENT).X
                            y0_sta1 = Linia_cunoscuta.StartPoint.Y 'TransformBy(UCS_CURENT).Y
                            y0_sta2 = Linia_cunoscuta.EndPoint.Y 'TransformBy(UCS_CURENT).Y
                            If Abs(X_sta_cunoscut - x0_sta2) > 0.001 Then
                                MsgBox("Vertical line you selected is not vertical")
                                Editor1.SetImpliedSelection(Empty_array)
                                Freeze_operations = False
                                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                Freeze_operations = False
                                Exit Sub
                            End If

                        End If




                        Dim Rezultat_blocks As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                        Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_Prompt1.MessageForAdding = vbLf & "Select the Blocks:"
                        Object_Prompt1.SingleOnly = False
                        Rezultat_blocks = Editor1.GetSelection(Object_Prompt1)

                        If Not Rezultat_blocks.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Editor1.SetImpliedSelection(Empty_array)
                            Freeze_operations = False
                            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                            Freeze_operations = False
                            Exit Sub
                        End If




                        For j = 0 To Rezultat_blocks.Value.Count - 1

                            Dim Ent_profile As Entity
                            Ent_profile = Trans1.GetObject(Rezultat_blocks.Value.Item(j).ObjectId, OpenMode.ForRead)


                            If TypeOf Ent_profile Is Autodesk.AutoCAD.DatabaseServices.BlockReference Then
                                Dim Block1 As BlockReference = Ent_profile

                                Dim BlockTrec As BlockTableRecord = Nothing
                                Dim Nume1 As String = "xxx"
                                If Block1.IsDynamicBlock = True Then
                                    BlockTrec = Trans1.GetObject(Block1.DynamicBlockTableRecord, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                    Nume1 = BlockTrec.Name
                                Else
                                    BlockTrec = Trans1.GetObject(Block1.BlockTableRecord, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                    Nume1 = BlockTrec.Name
                                End If

                                Data_table_Blocks.Rows.Add()
                                Data_table_Blocks.Rows(Data_table_Blocks.Rows.Count - 1).Item("StaH") = Station_cunoscuta - (X_sta_cunoscut - Block1.Position.X) / H_EXAG
                                Data_table_Blocks.Rows(Data_table_Blocks.Rows.Count - 1).Item("BlockName") = Nume1







                            End If





                        Next







                        Trans1.Commit()
                        ' asta e de la tranzactie
                    End Using




                    Transfer_datatable_to_new_excel_spreadsheet(Data_table_blocks)


                    Freeze_operations = False



                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    ' asta e de la lock
                End Using
                Freeze_operations = False
            End If
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

        Catch ex As System.Exception
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

            Freeze_operations = False
            MsgBox(ex.Message)
        End Try
    End Sub
End Class