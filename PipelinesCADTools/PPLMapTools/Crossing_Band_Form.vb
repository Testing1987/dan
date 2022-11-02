Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class Crossing_Band_Form
    Dim Colectie1 As New Specialized.StringCollection
    Dim Data_table_Centerline As System.Data.DataTable
    Dim Empty_array() As ObjectId
    Dim Poly1 As Polyline

    Private Sub Crossing_Band_Form_Load(sender As Object, e As EventArgs) Handles Me.Load
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

        Incarca_existing_textstyles_to_combobox(ComboBox_text_style)

        If ComboBox_text_style.Items.Count > 0 Then
            If ComboBox_text_style.Items.Contains("ALIGNDB") = True Then
                ComboBox_text_style.SelectedIndex = ComboBox_text_style.Items.IndexOf("ALIGNDB")
            Else
                ComboBox_text_style.SelectedIndex = 0
            End If
        End If

        Button_load_From_excel.Visible = False
        Button_draw.Visible = False

    End Sub
    Private Sub Panel_design_param_Click(sender As Object, e As EventArgs) Handles Panel_design_param.Click, Panel_REFRESH.Click
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

        Incarca_existing_textstyles_to_combobox(ComboBox_text_style)

        If ComboBox_text_style.Items.Count > 0 Then
            If ComboBox_text_style.Items.Contains("ALIGNDB") = True Then
                ComboBox_text_style.SelectedIndex = ComboBox_text_style.Items.IndexOf("ALIGNDB")
            Else
                ComboBox_text_style.SelectedIndex = 0
            End If
        End If
    End Sub
    Private Sub Button_load_CL_Click(sender As Object, e As EventArgs) Handles Button_load_CL.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor


        Editor1.SetImpliedSelection(Empty_array)
        Try
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Colectie1 = New Specialized.StringCollection
            ascunde_butoanele_pentru_forms(Me, Colectie1)

            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptEntityResult

            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select centerline:")

            Object_Prompt.SetRejectMessage(vbLf & "Please select a lightweight polyline")
            Object_Prompt.AddAllowedClass(GetType(Polyline), True)

            Rezultat1 = Editor1.GetEntity(Object_Prompt)


            If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                Editor1.SetImpliedSelection(Empty_array)
                Exit Sub
            End If

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                If IsNothing(Rezultat1) = False Then
                    Button_load_From_excel.Visible = False
                    Button_draw.Visible = False
                    Using Lock_dwg As DocumentLock = ThisDrawing.LockDocument
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim Ent1 As Entity = Trans1.GetObject(Rezultat1.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                            Poly1 = Ent1
                            Data_table_Centerline = New System.Data.DataTable
                            Data_table_Centerline.Columns.Add("X", GetType(Double))
                            Data_table_Centerline.Columns.Add("Y", GetType(Double))
                            Data_table_Centerline.Columns.Add("STATION", GetType(Double))
                            Data_table_Centerline.Columns.Add("DEFLECTION", GetType(Double))
                            Data_table_Centerline.Columns.Add("DEFLECTION_DMS", GetType(String))
                            Data_table_Centerline.Columns.Add("DESCRIPTION", GetType(String))


                            Dim Planul_curent As Plane = New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)

                            For i = 0 To Poly1.NumberOfVertices - 1
                                Data_table_Centerline.Rows.Add()
                                Dim X As Double = Poly1.GetPointAtParameter(i).X
                                Dim Y As Double = Poly1.GetPointAtParameter(i).Y
                                Dim Sta As Double = Poly1.GetDistanceAtParameter(i)

                                Data_table_Centerline.Rows(i).Item("X") = X
                                Data_table_Centerline.Rows(i).Item("Y") = y
                                Data_table_Centerline.Rows(i).Item("STATION") = Sta

                                If CheckBox_CALC_DEFLECTIONS.Checked = True Then
                                    If i > 0 And i < Poly1.NumberOfVertices - 1 Then

                                        Dim vector1 As Vector3d = Poly1.GetPoint3dAt(i - 1).GetVectorTo(Poly1.GetPoint3dAt(i))
                                        If vector1.Length < 0.01 Then
                                            Dim K As Double = 2
                                            Do While vector1.Length < 0.01
                                                If i - K >= 0 Then
                                                    vector1 = Poly1.GetPoint3dAt(i - K).GetVectorTo(Poly1.GetPoint3dAt(i))
                                                Else
                                                    Exit Do
                                                End If
                                                K = K + 1
                                            Loop
                                        End If

                                        Dim vector2 As Vector3d = Poly1.GetPoint3dAt(i).GetVectorTo(Poly1.GetPoint3dAt(i + 1))
                                        If vector2.Length < 0.01 Then
                                            Dim K As Double = 2
                                            Do While vector2.Length < 0.01
                                                If i + K <= Poly1.NumberOfVertices - 1 Then
                                                    vector2 = Poly1.GetPoint3dAt(i).GetVectorTo(Poly1.GetPoint3dAt(i + K))
                                                Else
                                                    Exit Do
                                                End If
                                                K = K + 1
                                            Loop
                                        End If

                                        Dim Bearing1 As Double = (vector1.AngleOnPlane(Planul_curent)) * 180 / PI
                                        Dim Bearing2 As Double = (vector2.AngleOnPlane(Planul_curent)) * 180 / PI

                                        Dim Angle1 As Double = (vector2.GetAngleTo(vector1)) * 180 / PI

                                        Dim LT_RT As String = ""


                                        If Bearing1 < 180 Then
                                            If Bearing2 < Bearing1 + 180 And Bearing2 > Bearing1 Then
                                                LT_RT = " LT"
                                            Else
                                                LT_RT = " RT"
                                            End If
                                        Else
                                            If Bearing2 < Bearing1 And Bearing2 > Bearing1 - 180 Then
                                                LT_RT = " RT"
                                            Else
                                                LT_RT = " LT"
                                            End If

                                        End If

                                        Dim Angle_string As String = Floor(Angle1) & "°"
                                        Dim Minute As Double = (Angle1 - Floor(Angle1)) * 60
                                        Dim Minute_string As String = Floor(Minute) & "'"
                                        Dim Second As Integer = Round((Minute - Floor(Minute)) * 60, 0)

                                        If Minute = 60 Then
                                            Angle_string = Floor(Angle1 + 1) & "°"
                                            Minute_string = "0'"
                                        End If


                                        Dim Second_string As String = Second.ToString & Chr(34)

                                        If Second = 60 Then
                                            Minute_string = Floor(Minute + 1) & "'"
                                            Second_string = "0" & Chr(34)
                                        End If

                                        Dim Angle1_DMS As String = Angle_string & Minute_string & Second_string & LT_RT
                                        Data_table_Centerline.Rows(i).Item("DEFLECTION") = Angle1
                                        Data_table_Centerline.Rows(i).Item("DEFLECTION_DMS") = Angle1_DMS

                                    End If
                                End If




                            Next
                            Button_draw.Visible = False
                            Button_load_From_excel.Visible = True
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
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub


    Private Sub Button_load_From_excel_Click(sender As Object, e As EventArgs) Handles Button_load_From_excel.Click
        Try
            ascunde_butoanele_pentru_forms(Me, Colectie1)
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
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                Exit Sub
            End If
            If End1 < Start1 Then
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                Exit Sub
            End If
            Dim Column_sta As String = ""
            Column_sta = TextBox_column_Station.Text.ToUpper
            Dim Column_descr As String = ""
            Column_descr = TextBox_column_description.Text.ToUpper


            'If Data_table_Centerline.Rows.Count > 0 Then
            'Dim NR As Integer = Data_table_Centerline.Rows.Count
            'For i = 0 To NR - 1
            'If i <= NR - 1 Then
            'If IsDBNull(Data_table_Centerline.Rows(i).Item("DEFLECTION")) = True Then
            'Data_table_Centerline.Rows(i).Delete()
            'i = i - 1
            'NR = NR - 1
            'End If
            'End If
            'Next
            'End If


            Dim Index_data_table As Double
            If Data_table_Centerline.Rows.Count > 0 Then
                Index_data_table = Data_table_Centerline.Rows.Count
                For i = Start1 To End1
                    Dim Station_string As String = W1.Range(Column_sta & i).Value
                    Dim Descriptie As String = W1.Range(Column_descr & i).Value
                    If IsNumeric(Station_string) = True And Not CDbl(Station_string) <= 0 And Not Descriptie = "" Then
                        Data_table_Centerline.Rows.Add()
                        Data_table_Centerline.Rows(Index_data_table).Item("STATION") = CDbl(Station_string)
                        Data_table_Centerline.Rows(Index_data_table).Item("DESCRIPTION") = Descriptie
                        Index_data_table = Index_data_table + 1
                    End If
                Next
            End If

            Data_table_Centerline = Sort_data_table(Data_table_Centerline, "STATION")

            'MsgBox(Data_table_Centerline.Rows.Count)

            Button_draw.Visible = True

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub

    Private Sub Button_draw_Click(sender As Object, e As EventArgs) Handles Button_draw.Click
        Try
            If IsNothing(Data_table_Centerline) = False Then
                If Data_table_Centerline.Rows.Count > 0 Then
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

                        ' Dim k As Double = 1
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            Dim BTrecordPS As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord

                            BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            BTrecordPS = Trans1.GetObject(BlockTable_data1(BlockTableRecord.PaperSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)


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
                                    Editor1.CurrentUserCoordinateSystem = WCS_align()
                                    'Dim GraphicsManager As Autodesk.AutoCAD.GraphicsSystem.Manager = ThisDrawing.GraphicsManager
                                    'Dim View0 As Autodesk.AutoCAD.GraphicsSystem.View = GraphicsManager.GetGsView(CShort(Application.GetSystemVariable("CVPORT")), True) ' acad 2013
                                    'Dim View0 As Autodesk.AutoCAD.GraphicsSystem.View = GraphicsManager.GetCurrentAcGsView(CShort(Application.GetSystemVariable("CVPORT"))) ' acad 2015
                                    Angle1 = Viewport1.TwistAngle

                                    'Len1 = Viewport1.Width
                                    'Height1 = Viewport1.Height
                                    Scale1 = Viewport1.CustomScale
                                    Centru_MS = Application.GetSystemVariable("VIEWCTR") 'View0.Target
                                    Centru_PS = Viewport1.CenterPoint
                                    Exista_viewport = True
                                    Editor1.SwitchToPaperSpace()
                                    Exit For
                                End If
                            Next

                            Dim Point_MS As New Point3d
                            Dim Point_PS As New Point3d

                            If Exista_viewport = True Then

                                Dim Match1 As Double = 0
                                If IsNumeric(Replace(TextBox_start.Text, "+", "")) = True Then Match1 = CDbl(Replace(TextBox_start.Text, "+", ""))
                                Dim Match2 As Double = 0
                                If IsNumeric(Replace(TextBox_end.Text, "+", "")) = True Then Match2 = CDbl(Replace(TextBox_end.Text, "+", ""))

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

                                If IsNumeric(TextBox_Y.Text) = True Then
                                    Y1 = CDbl(TextBox_Y.Text)
                                End If


                                Dim Chainage_previous As Double = -1
                                Dim x_previous As Double = -100000000
                                Dim x_BLOCK As Double = 0


                                Dim Min_dist As Double = 45

                                If IsNumeric(TextBox_minimum_distance.Text) = True Then
                                    Min_dist = CDbl(TextBox_minimum_distance.Text)
                                End If


                                Dim TextHeight As Double = 16
                                If IsNumeric(TextBox_text_height.Text) = True Then
                                    TextHeight = CDbl(TextBox_text_height.Text)
                                End If

                                Dim Mtext_rotation As Double = PI / 2



                                For i = 0 To Data_table_Centerline.Rows.Count - 1
                                    If IsDBNull(Data_table_Centerline.Rows(i).Item("STATION")) = False Then
                                        Dim Station As Double = Data_table_Centerline.Rows(i).Item("STATION")

                                        If Station >= Match1 And Station <= Match2 Then

                                            Dim Descriptie As String = ""
                                            Dim Layer_xing As String = ""
                                            Dim TextWidth As Double = 0.8
                                            If IsNumeric(TextBox_textwidth.Text) = True Then
                                                TextWidth = CDbl(TextBox_textwidth.Text)
                                            End If

                                            'Dim CL As String = "℄"

                                            If IsDBNull(Data_table_Centerline.Rows(i).Item("DESCRIPTION")) = False Then

                                                If Data_table_Centerline.Rows(i).Item("DESCRIPTION").ToString.Contains("?") = True Then
                                                    'Data_table_Centerline.Rows(i).Item("DESCRIPTION") = Replace(Data_table_Centerline.Rows(i).Item("DESCRIPTION"), "?", CL)
                                                End If

                                                Descriptie = "{\W" & TextWidth & ";" & Get_chainage_feet_from_double(Station, 0) & " " & Data_table_Centerline.Rows(i).Item("DESCRIPTION") & "}"
                                                Layer_xing = ComboBox_layers_crossings.Text
                                            Else
                                                If IsDBNull(Data_table_Centerline.Rows(i).Item("DEFLECTION_DMS")) = False Then
                                                    Descriptie = "{\W" & TextWidth & ";\L" & Get_chainage_feet_from_double(Station, 0) & " P.I. ~ " & Data_table_Centerline.Rows(i).Item("DEFLECTION_DMS") & "}"
                                                    Layer_xing = ComboBox_layers_deflections.Text
                                                End If
                                            End If


                                            If Station <= Poly1.Length And Not Descriptie = "" Then
                                                Point_MS = Poly1.GetPointAtDist(Station)
                                                Point_PS = New Point3d(Centru_PS.X - (Centru_MS.X - Point_MS.X) * Scale1, Centru_PS.Y - (Centru_MS.Y - Point_MS.Y) * Scale1, 0)
                                                Point_PS = Point_PS.TransformBy(Matrix3d.Rotation(Angle1, Vector3d.ZAxis, Centru_PS))


                                                If RadioButton_Left_right.Checked = True Then
                                                    If Point_PS.X - x_previous < Min_dist Then
                                                        Point_PS = New Point3d(x_previous + Min_dist, Point_PS.Y, 0)
                                                    End If
                                                Else
                                                    If x_previous - Point_PS.X < Min_dist Then
                                                        Point_PS = New Point3d(x_previous - Min_dist, Point_PS.Y, 0)
                                                    End If
                                                End If



                                                Dim Mtext As New MText
                                                Mtext.Location = New Point3d(Point_PS.X, Y1, 0)

                                                Mtext.TextHeight = TextHeight
                                                Mtext.Rotation = Mtext_rotation
                                                Mtext.Layer = Layer_xing
                                                Mtext.Contents = Descriptie

                                                Mtext.Attachment = AttachmentPoint.BottomLeft
                                                Dim ObjId1 As TextStyleTableRecord

                                                If Text_style_table.Has(ComboBox_text_style.Text) = True Then
                                                    ObjId1 = Text_style_table(ComboBox_text_style.Text).GetObject(OpenMode.ForRead)
                                                    Mtext.TextStyleId = ObjId1.ObjectId
                                                End If
                                                BTrecordPS.AppendEntity(Mtext)
                                                Trans1.AddNewlyCreatedDBObject(Mtext, True)

                                                x_previous = Point_PS.X
                                            End If
                                        End If
                                    End If
                                Next
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
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
    End Sub
End Class