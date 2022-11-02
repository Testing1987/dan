Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports ACSMCOMPONENTS20Lib

Public Class Alignment_Sheet_cutter

    Dim Vw_scale As Double = 1
    Dim Rotatie As Double






    Dim Layer_name_Main_Viewport As String = "VP"

    Dim Layer_name_Blocks As String = "0NORTH"



    Dim Rotatie_originala As Double
    Dim Vw_height As Double = 4.2
    Dim Vw_width As Double = 7
    Dim Vw_CenX As Double = 8.5
    Dim Vw_CenY As Double = 7.4
    Dim Match_distance As Double = 5280

    Dim VW_TARGET_X As String = "VW_TARGET_X"
    Dim VW_TARGET_Y As String = "VW_TARGET_Y"
    Dim VW_TWIST As String = "VW_TWIST"
    Dim VW_CUST_SCALE As String = "VW_CUST_SCALE"
    Dim Length_of_viewport As String = "LENGTH"
    Dim Height_of_viewport As String = "HEIGHT"

    Dim NEW_NAME As String = "NEW_NAME"
    Dim STATION_COLUMN As String = "STATION"

    Dim Punct_insertie_north_arrow As New Point3d(2.7722595, 8.9239174, 0)

    Dim Scale_factor As Double = 0
    Dim View_rotation As Double = 0
    Dim Text_height As Double = 0.08

    Dim Data_table_Viewport_data As System.Data.DataTable
    Dim Freeze_operations As Boolean = False

    Dim Data_table_matchline As System.Data.DataTable

    Private Sub sheet_cutter_form_Load(sender As Object, e As EventArgs) Handles Me.Load







    End Sub



    Private Sub Button_generate_Platt_Click(sender As Object, e As EventArgs) Handles Button_generate_Platt.Click

        If Freeze_operations = False Then
            Freeze_operations = True



            If TextBox_Output_Directory.Text = "" Then
                MsgBox("Please specify the output folder")
                Freeze_operations = False
                Exit Sub
            End If
            If Not Strings.Right(TextBox_Output_Directory.Text, 1) = "\" Then
                TextBox_Output_Directory.Text = TextBox_Output_Directory.Text & "\"
            End If

            If TextBox_dwt_template.Text = "" Then
                MsgBox("Please specify the template file")
                Freeze_operations = False
                Exit Sub
            End If
            If Not Strings.Right(TextBox_dwt_template.Text, 3).ToUpper = "DWT" Then
                MsgBox("Please specify the template file")
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





            Dim Empty_array() As ObjectId

            Dim BaseMap_drawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument





            If IO.File.Exists(TextBox_dwt_template.Text) = True Then
                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = BaseMap_drawing.Editor
                Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                Try









                    Dim Data_table_cu_valori As New System.Data.DataTable
                    Dim Index_data_table_valori As Double = 0



                    Data_table_cu_valori.Columns.Add(VW_TARGET_X, GetType(Double))
                    Data_table_cu_valori.Columns.Add(VW_TARGET_Y, GetType(Double))
                    Data_table_cu_valori.Columns.Add(VW_TWIST, GetType(Double))
                    Data_table_cu_valori.Columns.Add(VW_CUST_SCALE, GetType(Double))
                    Data_table_cu_valori.Columns.Add(NEW_NAME, GetType(String))


                    Dim Start1 As String = TextBox_start_number.Text



                    Using lock1 As DocumentLock = BaseMap_drawing.LockDocument
                        Creaza_layer(Layer_name_Main_Viewport, 5, Layer_name_Main_Viewport, False)

                        Using Trans_basemap As Autodesk.AutoCAD.DatabaseServices.Transaction = BaseMap_drawing.TransactionManager.StartTransaction
                            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            BlockTable1 = BaseMap_drawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            Dim BTrecord_MS_basemap As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans_basemap.GetObject(BlockTable1(BlockTableRecord.ModelSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)




1254:


                            Vw_scale = 1

                            Dim PromptPointRezult2 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                            Dim PP2 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select first point:")
                            PP2.AllowNone = True

                            PromptPointRezult2 = Editor1.GetPoint(PP2)
                            If PromptPointRezult2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Dim PromptPointRezult3 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                                Dim PP3 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select second point:")
                                PP3.AllowNone = True
                                PP3.BasePoint = PromptPointRezult2.Value
                                PP3.UseBasePoint = True
                                PromptPointRezult3 = Editor1.GetPoint(PP3)


                                If PromptPointRezult3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                    Dim x1 As Double = PromptPointRezult2.Value.X
                                    Dim y1 As Double = PromptPointRezult2.Value.Y
                                    Dim x2 As Double = PromptPointRezult3.Value.X
                                    Dim y2 As Double = PromptPointRezult3.Value.Y
                                    Rotatie = GET_Bearing_rad(x1, y1, x2, y2)


                                End If




                            End If






                            Dim PromptPointRezult1 As Autodesk.AutoCAD.EditorInput.PromptPointResult

                            Dim Jig1 As New Jig_rectangle_viewport_SHEET_CUTTER




                            PromptPointRezult1 = Jig1.StartJig(Vw_width, Vw_height, True)

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





                            Using Trans_basemap2 As Autodesk.AutoCAD.DatabaseServices.Transaction = BaseMap_drawing.TransactionManager.StartTransaction
                                Dim Poly1 As New Autodesk.AutoCAD.DatabaseServices.Polyline
                                Poly1.AddVertexAt(0, New Point2d(pointM.X - 0.5 * Vw_width / Vw_scale, pointM.Y - 0.5 * Vw_height / Vw_scale), 0, 0, 0)
                                Poly1.AddVertexAt(1, New Point2d(pointM.X + 0.5 * Vw_width / Vw_scale, pointM.Y - 0.5 * Vw_height / Vw_scale), 0, 0, 0)
                                Poly1.AddVertexAt(2, New Point2d(pointM.X + 0.5 * Vw_width / Vw_scale, pointM.Y + 0.5 * Vw_height / Vw_scale), 0, 0, 0)
                                Poly1.AddVertexAt(3, New Point2d(pointM.X - 0.5 * Vw_width / Vw_scale, pointM.Y + 0.5 * Vw_height / Vw_scale), 0, 0, 0)
                                Poly1.Closed = True


                                Poly1.TransformBy(Matrix3d.Rotation(Rotatie, Vector3d.ZAxis, pointM))


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






                                        GoTo 1254
                                    Case MsgBoxResult.Yes

                                        Data_table_cu_valori.Rows.Add()

                                        Data_table_cu_valori.Rows(Index_data_table_valori).Item(VW_TWIST) = 2 * PI - Rotatie


                                        Data_table_cu_valori.Rows(Index_data_table_valori).Item(VW_CUST_SCALE) = Vw_scale
                                        Data_table_cu_valori.Rows(Index_data_table_valori).Item(VW_TARGET_X) = pointM.X
                                        Data_table_cu_valori.Rows(Index_data_table_valori).Item(VW_TARGET_Y) = pointM.Y
                                        Data_table_cu_valori.Rows(Index_data_table_valori).Item(NEW_NAME) = TextBox_NEW_NAME_PREFIX.Text & Start1

                                        Start1 = INCREASE_STRING_BY_ONE(Start1)


                                End Select ' msgbox YES



                                Trans_basemap2.Dispose()
                            End Using







                            Index_data_table_valori = Index_data_table_valori + 1





                            If MsgBox("Please pick the next sheet", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                                GoTo 1254

                            End If






                            Trans_basemap.Commit()

                        End Using
                    End Using



                    If IsNothing(Data_table_cu_valori) = False Then
                        If Data_table_cu_valori.Rows.Count > 0 Then



                            Dim Index_existent As Integer = 1


                            If Not Strings.Right(TextBox_Output_Directory.Text, 1) = "\" Then
                                TextBox_Output_Directory.Text = TextBox_Output_Directory.Text & "\"
                            End If

                            If System.IO.Directory.Exists(TextBox_Output_Directory.Text) = True Then




                                For s = 0 To Data_table_cu_valori.Rows.Count - 1
                                    Dim New_plat_name As String = ""

                                    New_plat_name = Data_table_cu_valori.Rows(s).Item(NEW_NAME)

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






                                    Dim Fisier_nou As String = TextBox_Output_Directory.Text & New_plat_name & ".dwg"


                                    Dim DocumentManager1 As DocumentCollection = Application.DocumentManager
                                    Dim New_doc As Document = DocumentCollectionExtension.Add(DocumentManager1, TextBox_dwt_template.Text)
                                    DocumentManager1.MdiActiveDocument = New_doc
                                    Using lock2 As DocumentLock = New_doc.LockDocument
                                        Using Trans_New_doc As Autodesk.AutoCAD.DatabaseServices.Transaction = New_doc.TransactionManager.StartTransaction
                                            Creaza_layer(Layer_name_Main_Viewport, 40, Layer_name_Main_Viewport, False)

                                            Creaza_layer(Layer_name_Blocks, 7, Layer_name_Blocks, True)

                                            Trans_New_doc.Commit()
                                        End Using
                                    End Using



                                    Dim Anno_scale_name1 As String = ""

                                    Dim DWG_units As Integer

                                    Using lock2 As DocumentLock = New_doc.LockDocument

                                        Using Trans_New_doc As Autodesk.AutoCAD.DatabaseServices.Transaction = New_doc.TransactionManager.StartTransaction
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


                                            Dim Colectie_atr_name As New Specialized.StringCollection
                                            Dim Colectie_atr_value As New Specialized.StringCollection

                                            If Not TextBox_north_arrow.Text = "" Then
                                                Dim Block_north As Autodesk.AutoCAD.DatabaseServices.BlockReference
                                                If IsNumeric(TextBox_north_arrow_Big_X.Text) = True And IsNumeric(TextBox_north_arrow_Big_y.Text) = True Then
                                                    Punct_insertie_north_arrow = New Point3d(CDbl(TextBox_north_arrow_Big_X.Text), CDbl(TextBox_north_arrow_Big_y.Text), 0)
                                                End If

                                                Dim Nume_North_arrow As String = TextBox_north_arrow.Text
                                                Dim Scaleblock As Double = 1
                                                If IsNumeric(TextBox_blockScale.Text) = True Then
                                                    Scaleblock = CDbl(TextBox_blockScale.Text)
                                                End If

                                                Block_north = InsertBlock_with_multiple_atributes("", Nume_North_arrow, Punct_insertie_north_arrow, Scaleblock, BTrecord_new_doc_PS, Layer_name_Blocks, Colectie_atr_name, Colectie_atr_value)
                                                Block_north.Rotation = Data_table_cu_valori.Rows(s).Item(VW_TWIST)
                                            End If

                                            Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current

                                            Dim Tilemode1 As Integer = Application.GetSystemVariable("TILEMODE")
                                            Dim CVport1 As Integer = Application.GetSystemVariable("CVPORT")

                                            If (Tilemode1 = 0 And Not CVport1 = 1) Then
                                                New_doc.Editor.SwitchToPaperSpace()
                                            End If
                                            If Tilemode1 = 1 Then
                                                Application.SetSystemVariable("TILEMODE", 0)
                                            End If

                                            Dim Layoutdict As DBDictionary = Trans_New_doc.GetObject(New_doc.Database.LayoutDictionaryId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                            Dim Layout1 As Layout = Trans_New_doc.GetObject(LayoutManager1.GetLayoutId(LayoutManager1.CurrentLayout), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                                            Layout1.LayoutName = New_plat_name


                                            Trans_New_doc.Commit()
                                        End Using



                                        New_doc.Database.SaveAs(Fisier_nou, True, DwgVersion.Current, BaseMap_drawing.Database.SecurityParameters)
                                    End Using
                                    DocumentExtension.CloseAndDiscard(New_doc)
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
                MsgBox("TEMPLATE " & TextBox_dwt_template.Text & " DOES NOT EXIST")
            End If
            Freeze_operations = False
        End If
    End Sub

    Public Function INCREASE_STRING_BY_ONE(ByVal string1 As String) As String
        Dim string2 As String = string1
        If IsNumeric(string1) = True Then
            Dim Nr As Integer = Abs(CInt(string1)) + 1
            string2 = Nr.ToString
            If string1.Length = 2 And Nr < 10 Then
                string2 = "0" & Nr.ToString
            End If
            If string1.Length = 3 And Nr < 10 Then
                string2 = "00" & Nr.ToString
            End If
            If string1.Length = 3 And Nr < 100 And Nr >= 10 Then
                string2 = "0" & Nr.ToString
            End If

            If string1.Length = 4 And Nr < 10 Then
                string2 = "000" & Nr.ToString
            End If
            If string1.Length = 4 And Nr < 100 And Nr >= 10 Then
                string2 = "00" & Nr.ToString
            End If
            If string1.Length = 4 And Nr < 1000 And Nr >= 100 Then
                string2 = "0" & Nr.ToString
            End If

            If string1.Length = 5 And Nr < 10 Then
                string2 = "0000" & Nr.ToString
            End If
            If string1.Length = 5 And Nr < 100 And Nr >= 10 Then
                string2 = "000" & Nr.ToString
            End If
            If string1.Length = 5 And Nr < 1000 And Nr >= 100 Then
                string2 = "00" & Nr.ToString
            End If

            If string1.Length = 5 And Nr < 10000 And Nr >= 1000 Then
                string2 = "0" & Nr.ToString
            End If
        End If
        Return string2

    End Function










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
    Private Sub Button_templates_output_folder_Click(sender As Object, e As EventArgs) Handles Button_templates_output_folder.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                Dim FolderBrowserDialog1 As New Windows.Forms.FolderBrowserDialog
                FolderBrowserDialog1.ShowNewFolderButton = False
                If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    TextBox_templates_Output_Directory.Text = FolderBrowserDialog1.SelectedPath
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
    Private Sub Button_templates_drawing_template_Click(sender As Object, e As EventArgs) Handles Button_templates_drawing_template.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                Dim FileBrowserDialog1 As New Windows.Forms.OpenFileDialog
                FileBrowserDialog1.Filter = "Template Files (*.dwg)|*.dwg|All Files (*.*)|*.*"
                If FileBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    TextBox_templates_dwt_template.Text = FileBrowserDialog1.FileName
                End If
            Catch ex As System.Exception
                Freeze_operations = False
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_rename_layout_Click(sender As Object, e As EventArgs) Handles Button_rename_layout.Click
        If Freeze_operations = False Then
            Freeze_operations = True








            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                    Dim Colectie_nume_layout As New Specialized.StringCollection


                    Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current

                    Dim Layoutdict As DBDictionary = Trans1.GetObject(ThisDrawing.Database.LayoutDictionaryId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                    For Each entry As DBDictionaryEntry In Layoutdict
                        If Not entry.Key.ToUpper = "MODEL" Then
                            Colectie_nume_layout.Add(entry.Key)
                        End If
                    Next


                    'For i = 0 To Colectie_nume_layout.Count - 1
                    'If Colectie_nume_layout(i).ToUpper = "Residential Site Specific".ToUpper Then
                    ' Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(Colectie_nume_layout(i)), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                    'Layout1.LayoutName = IO.Path.GetFileNameWithoutExtension(ThisDrawing.Database.OriginalFileName)
                    'End If
                    'Next
                    Dim Tilemode1 As Integer = Application.GetSystemVariable("TILEMODE")
                    Dim CVport1 As Integer = Application.GetSystemVariable("CVPORT")

                    If (Tilemode1 = 0 And Not CVport1 = 1) Then
                        ThisDrawing.Editor.SwitchToPaperSpace()
                    End If
                    If Tilemode1 = 1 Then
                        Application.SetSystemVariable("TILEMODE", 0)
                    End If

                    If Colectie_nume_layout.Contains(IO.Path.GetFileNameWithoutExtension(ThisDrawing.Database.OriginalFileName)) = False Then
                        Dim Layout2 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(LayoutManager1.CurrentLayout), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                        Layout2.LayoutName = IO.Path.GetFileNameWithoutExtension(ThisDrawing.Database.OriginalFileName)
                    Else
                        MsgBox("Please rename your layout manually - there is already a layout named as the file name")
                    End If



                    Trans1.Commit()
                End Using
            End Using





            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_PLACE_VIEWPORTS_Click(sender As Object, e As EventArgs) Handles Button_PLACE_VIEWPORTS.Click


        If Freeze_operations = False Then
            Freeze_operations = True






            If IsNumeric(TextBox_matchline_length.Text) = True Then
                Match_distance = CDbl(TextBox_matchline_length.Text)
            End If

            If IsNumeric(TextBox_main_viewport_height.Text) = True Then
                Vw_height = CDbl(TextBox_main_viewport_height.Text)
            End If


            Dim Empty_array() As ObjectId

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Try
                Dim Index1 As Double = 0





                Dim Poly2D As Polyline


                Using lock1 As DocumentLock = ThisDrawing.LockDocument
                    Creaza_layer(Layer_name_Main_Viewport, 5, Layer_name_Main_Viewport, False)

                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                        Dim Prompt_optionsCL As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select Centerline:")
                        Prompt_optionsCL.SetRejectMessage(vbLf & "You did not selected a polyline")
                        Prompt_optionsCL.AddAllowedClass(GetType(Polyline), True)

                        Dim Rezultat_CL As Autodesk.AutoCAD.EditorInput.PromptEntityResult = Editor1.GetEntity(Prompt_optionsCL)
                        If Not Rezultat_CL.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Freeze_operations = False
                            Editor1.SetImpliedSelection(Empty_array)
                            Editor1.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If
                        Poly2D = Trans1.GetObject(Rezultat_CL.ObjectId, OpenMode.ForRead)


                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        'Dim BTrecord_MS As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(BlockTable1(BlockTableRecord.ModelSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                        Dim dist1 As Double = 0



                        Data_table_matchline = New System.Data.DataTable
                        Data_table_matchline.Columns.Add("OBJECT_ID", GetType(ObjectId))
                        Data_table_matchline.Columns.Add("M1", GetType(Double))
                        Data_table_matchline.Columns.Add("M2", GetType(Double))



                        Dim dist2 As Double = dist1 + Match_distance
                        Dim Ultimul As Boolean = False
                        Dim Colorindex As Integer = 1

                        If Colorindex = 3233 Then
                            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                            W1 = Get_NEW_worksheet_from_Excel()
                        End If

                        Dim Este_primul As Boolean = True
                        Dim Last_pt As New Point3d



123:
                        Dim Point1 As New Point3d
                        Point1 = Poly2D.GetPointAtDist(dist1)
                        Dim Point2 As New Point3d
                        Point2 = Poly2D.GetPointAtDist(dist2)



                        Dim Poly1r As New Autodesk.AutoCAD.DatabaseServices.Polyline

                        Poly1r = creaza_rectangle_viewport(Point1, Point2, Colorindex)
                        Poly1r.Layer = Layer_name_Main_Viewport

                        Dim Col_int As New Point3dCollection
                        Col_int = Intersect_on_both_operands(Poly2D, Poly1r)
                        If Col_int.Count = 2 Then



                            BTrecord.AppendEntity(Poly1r)
                            Trans1.AddNewlyCreatedDBObject(Poly1r, True)
                            Trans1.TransactionManager.QueueForGraphicsFlush()

                            Data_table_matchline.Rows.Add()
                            Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("OBJECT_ID") = Poly1r.ObjectId
                            Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("M1") = Round(dist1, 0)
                            Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("M2") = Round(dist2, 0)


                            dist1 = dist2
                            dist2 = dist2 + Match_distance

                            Colorindex = Colorindex + 1
                            If Colorindex > 7 Then Colorindex = 1

                            If Ultimul = False Then
                                If Poly2D.Length < dist2 Then
                                    dist2 = Poly2D.Length
                                    Ultimul = True
                                End If
                                Este_primul = True
                                GoTo 123
                            End If

                        Else

                            Dim Pointm1 As New Point3d

                            If dist1 > 0 Then
                                Dim Result_point_m1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                                Dim PP1m As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Please pick start location:")


                                If Este_primul = True Then

                                    PP1m.AllowNone = False
                                    Result_point_m1 = Editor1.GetPoint(PP1m)


                                    If Result_point_m1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                        Trans1.Commit()
                                        GoTo end1
                                    End If
                                    Pointm1 = Result_point_m1.Value
                                    Last_pt = Poly2D.GetClosestPointTo(Pointm1, Vector3d.ZAxis, False)
                                End If
                            End If

                            If dist1 = 0 Then
                                Last_pt = Poly2D.GetPointAtParameter(0)
                                Pointm1 = Last_pt
                            End If


                            Dim Result_point_m2 As Autodesk.AutoCAD.EditorInput.PromptPointResult

                            Dim Jig2 As New Jig_rectangle_viewport_SHEET_CUTTER_manual_pt2
                            Result_point_m2 = Jig2.StartJig(Vw_scale, Vw_width, Vw_height, Poly2D, Last_pt, 10, Match_distance)


                            If Result_point_m2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                Trans1.Commit()
                                GoTo end1
                            End If


                            Dim dist1m As Double
                            If Este_primul = True Then
                                dist1m = Poly2D.GetDistAtPoint(Poly2D.GetClosestPointTo(Pointm1, Vector3d.ZAxis, False))
                            Else
                                dist1m = Poly2D.GetDistAtPoint(Poly2D.GetClosestPointTo(Last_pt, Vector3d.ZAxis, False))
                            End If

                            Last_pt = Poly2D.GetClosestPointTo(Result_point_m2.Value, Vector3d.ZAxis, False)

                            Dim dist2m As Double = Poly2D.GetDistAtPoint(Last_pt)



                            If dist1m > dist2m Then
                                Dim t As Double = dist1m
                                dist1m = dist2m
                                dist2m = t
                            End If

                            Dim Point1m As New Point3d
                            Point1m = Poly2D.GetPointAtDist(dist1m)
                            Dim Point2m As New Point3d
                            Point2m = Poly2D.GetPointAtDist(dist2m)

                            Dim Poly1rm As New Autodesk.AutoCAD.DatabaseServices.Polyline

                            Poly1rm = creaza_rectangle_viewport(Point1m, Point2m, Colorindex)
                            Poly1rm.Layer = Layer_name_Main_Viewport

                            BTrecord.AppendEntity(Poly1rm)
                            Trans1.AddNewlyCreatedDBObject(Poly1rm, True)


                            Dim Line1 As New Line(Poly1rm.GetPointAtParameter(2), Poly1rm.GetPointAtParameter(3))
                            Line1.TransformBy(Matrix3d.Scaling(10000, Line1.StartPoint))
                            Line1.TransformBy(Matrix3d.Scaling(10000, Line1.EndPoint))

                            Dim Jig1 As New Jig_rectangle_viewport_SHEET_CUTTER_manual_up_down(Point2m, Line1)
                            Jig1.AddEntity(Poly1rm)
                            Dim jigRes As Autodesk.AutoCAD.EditorInput.PromptResult = ThisDrawing.Editor.Drag(Jig1)
                            If jigRes.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Jig1.TransformEntities()
                            End If

                            Trans1.TransactionManager.QueueForGraphicsFlush()

                            If Este_primul = True Then
                                If Data_table_matchline.Rows.Count > 0 Then

                                    Dim M1_p As Double = Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("M1")
                                    Dim M2_p As Double = Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("M2")
                                    Dim ob_id As ObjectId = Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("OBJECT_ID")
                                    Dim PolyR As Polyline = Trans1.GetObject(ob_id, OpenMode.ForWrite)

                                    Dim Point01 As New Point3d
                                    Point01 = Poly2D.GetPointAtDist(M1_p)
                                    Dim Point02 As New Point3d
                                    Point02 = Poly2D.GetPointAtDist(M2_p)
                                    Dim Poly0r As New Autodesk.AutoCAD.DatabaseServices.Polyline

                                    Poly0r = creaza_rectangle_viewport(Point01, Point1m, PolyR.ColorIndex)
                                    Poly0r.Layer = Layer_name_Main_Viewport

                                    BTrecord.AppendEntity(Poly0r)
                                    Trans1.AddNewlyCreatedDBObject(Poly0r, True)
                                    Trans1.TransactionManager.QueueForGraphicsFlush()

                                    PolyR.Erase()
                                    Trans1.TransactionManager.QueueForGraphicsFlush()

                                    Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("M2") = Round(dist1m, 0)
                                    Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("OBJECT_ID") = Poly0r.ObjectId
                                End If

                            End If


                            Data_table_matchline.Rows.Add()
                            Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("M1") = Round(dist1m, 0)
                            Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("M2") = Round(dist2m, 0)
                            Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("OBJECT_ID") = Poly1rm.ObjectId

                            Colorindex = Colorindex + 1
                            If Colorindex > 7 Then Colorindex = 1
                            Este_primul = False

                            dist1 = dist2m
                            dist2 = dist2m + Match_distance
                            If Round(dist1, 0) = Round(Poly2D.Length, 0) Then
                                GoTo 124
                            End If
                            If Round(dist2, 0) > Poly2D.Length Then
                                dist2 = Poly2D.Length
                                GoTo 124
                            End If

                            GoTo 123



                        End If
124:                    Editor1.WriteMessage(vbLf & "Command:")

                        Trans1.Commit()
                    End Using
                End Using

end1:

                If IsNothing(Data_table_matchline) = False Then
                    If Data_table_matchline.Rows.Count > 0 Then
                        Add_to_clipboard_Data_table(Data_table_matchline)
                    End If
                End If

            Catch ex As Exception
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try

            Freeze_operations = False
        End If
    End Sub

    Private Function creaza_rectangle_viewport(ByVal Point1 As Point3d, ByVal Point2 As Point3d, ByVal cid As Integer) As Polyline

        Dim Line1R As New Line(Point1, Point2)
        Dim Point_distR As New Point3d
        If Line1R.Length > Vw_height / Vw_scale Then
            Point_distR = Line1R.GetPointAtDist(Vw_height / Vw_scale)
            Line1R.EndPoint = Point_distR
        Else
            Line1R.TransformBy(Matrix3d.Scaling((Vw_height / Vw_scale) / Line1R.Length, Line1R.StartPoint))
        End If

        Line1R.TransformBy(Matrix3d.Rotation(PI / 2, Vector3d.ZAxis, Point1))
        Dim Point_middler As New Point3d((Point1.X + Line1R.EndPoint.X) / 2, (Point1.Y + Line1R.EndPoint.Y) / 2, 0)

        Line1R.TransformBy(Matrix3d.Displacement(Point_middler.GetVectorTo(Point1)))
        Dim Pt1r As New Point3d
        Pt1r = Line1R.StartPoint
        Dim Pt2r As New Point3d
        Pt2r = Line1R.EndPoint
        Line1R.TransformBy(Matrix3d.Displacement(Point1.GetVectorTo(Point2)))

        Dim Pt4r As New Point3d
        Pt4r = Line1R.StartPoint
        Dim Pt3r As New Point3d
        Pt3r = Line1R.EndPoint

        Dim Poly1r As New Autodesk.AutoCAD.DatabaseServices.Polyline
        Poly1r.AddVertexAt(0, New Point2d(Pt1r.X, Pt1r.Y), 0, 0, 0)
        Poly1r.AddVertexAt(1, New Point2d(Pt2r.X, Pt2r.Y), 0, 0, 0)
        Poly1r.AddVertexAt(2, New Point2d(Pt3r.X, Pt3r.Y), 0, 0, 0)
        Poly1r.AddVertexAt(3, New Point2d(Pt4r.X, Pt4r.Y), 0, 0, 0)
        Poly1r.Closed = True
        Poly1r.ColorIndex = cid

        Return Poly1r

    End Function





    Private Sub Button_read_templates_Click(sender As Object, e As EventArgs) Handles Button_read_templates.Click



        If Freeze_operations = False Then
            Freeze_operations = True







            Dim Empty_array() As ObjectId

            Dim This_drawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = This_drawing.Editor
            Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Try


                Data_table_Viewport_data = New System.Data.DataTable
                Data_table_Viewport_data.Columns.Add(Length_of_viewport, GetType(Double))
                Data_table_Viewport_data.Columns.Add(Height_of_viewport, GetType(Double))
                Data_table_Viewport_data.Columns.Add(VW_TARGET_X, GetType(Double))
                Data_table_Viewport_data.Columns.Add(VW_TARGET_Y, GetType(Double))
                Data_table_Viewport_data.Columns.Add(VW_TWIST, GetType(Double))
                Data_table_Viewport_data.Columns.Add(VW_CUST_SCALE, GetType(Double))

                Data_table_Viewport_data.Columns.Add(NEW_NAME, GetType(String))





                Using lock1 As DocumentLock = This_drawing.LockDocument


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = This_drawing.TransactionManager.StartTransaction


                        Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                        Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_Prompt.MessageForAdding = vbLf & "Select Templates"

                        Object_Prompt.SingleOnly = False

                        Rezultat1 = Editor1.GetSelection(Object_Prompt)


                        If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Editor1.WriteMessage(vbLf & "Command:")
                            Freeze_operations = False
                            Editor1.SetImpliedSelection(Empty_array)
                            Exit Sub
                        End If

                        Vw_scale = 1

                        If IsNumeric(TextBox_template_viewport_scale.Text) = True Then
                            Vw_scale = 1 / CDbl(TextBox_template_viewport_scale.Text)
                        End If

                        Dim Index1 As Integer = 0
                        Dim Start1 As String = TextBox_start_number.Text

                        For i = 0 To Rezultat1.Value.Count - 1
                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(i)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                Dim Poly1 As Polyline
                                Poly1 = Ent1





                                If Poly1.NumberOfVertices >= 4 Then
                                    Dim Point1 As New Point3d(0, 0, 0)
                                    Dim Point2 As New Point3d(0, 0, 0)

                                    Dim L1 As Double = 0

                                    Data_table_Viewport_data.Rows.Add()
                                    Data_table_Viewport_data.Rows(Index1).Item(VW_TARGET_X) = (Poly1.GetPointAtParameter(0).X + Poly1.GetPointAtParameter(2).X) / 2
                                    Data_table_Viewport_data.Rows(Index1).Item(VW_TARGET_Y) = (Poly1.GetPointAtParameter(0).Y + Poly1.GetPointAtParameter(2).Y) / 2
                                    Data_table_Viewport_data.Rows(Index1).Item(VW_CUST_SCALE) = Vw_scale


                                    If Not TextBox_OBJECT_DATA_FIELD_NAME.Text = "" Then
                                        Dim Id1 As ObjectId = Ent1.ObjectId
                                        Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                                        Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables
                                        Using Records1 As Autodesk.Gis.Map.ObjectData.Records = Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                            If IsNothing(Records1) = False Then
                                                If Records1.Count > 0 Then
                                                    Dim Record1 As Autodesk.Gis.Map.ObjectData.Record


                                                    For Each Record1 In Records1
                                                        Dim Tabla1 As Autodesk.Gis.Map.ObjectData.Table
                                                        Tabla1 = Tables1(Record1.TableName)


                                                        Dim Field_defs1 As Autodesk.Gis.Map.ObjectData.FieldDefinitions
                                                        Field_defs1 = Tabla1.FieldDefinitions

                                                        For j = 0 To Record1.Count - 1
                                                            Dim Field_def1 As Autodesk.Gis.Map.ObjectData.FieldDefinition
                                                            Field_def1 = Field_defs1(j)
                                                            If Field_def1.Name.ToUpper = TextBox_OBJECT_DATA_FIELD_NAME.Text.ToUpper() Then


                                                                Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                                                Valoare_record1 = Record1(j)
                                                                Start1 = Valoare_record1.StrValue
                                                                If IsNumeric(Start1) = True Then
                                                                    Dim Nr1 As Integer = CInt(Start1)
                                                                    If Nr1 < 10 Then
                                                                        Start1 = "00" & Start1
                                                                    End If
                                                                    If Nr1 < 100 And Nr1 >= 10 Then
                                                                        Start1 = "0" & Start1
                                                                    End If


                                                                End If


                                                                Data_table_Viewport_data.Rows(Index1).Item(NEW_NAME) = TextBox_templates_file_prefix.Text & Start1
                                                                Exit For


                                                            End If


                                                        Next
                                                    Next

                                                End If
                                            End If
                                        End Using
                                    Else

                                        Data_table_Viewport_data.Rows(Index1).Item(NEW_NAME) = TextBox_templates_file_prefix.Text & Start1
                                        Start1 = INCREASE_STRING_BY_ONE(Start1)
                                    End If

                                    Dim H1 As Double = 0




                                    For j = 0 To Poly1.NumberOfVertices - 2
                                        Dim Point01 As New Point3d
                                        Point01 = Poly1.GetPoint3dAt(j)
                                        Dim Point02 As New Point3d
                                        Point02 = Poly1.GetPoint3dAt(j + 1)
                                        Dim L2 As Double = Point01.GetVectorTo(Point02).Length

                                        If L1 < L2 Then
                                            L1 = L2
                                            If Point01.X < Point02.X Then
                                                Point1 = Point01
                                                Point2 = Point02
                                            Else
                                                Point1 = Point02
                                                Point2 = Point01
                                            End If
                                            Data_table_Viewport_data.Rows(Index1).Item(Length_of_viewport) = L1 * Vw_scale
                                        Else
                                            If L2 > 0 And H1 = 0 Then
                                                H1 = L2
                                                Data_table_Viewport_data.Rows(Index1).Item(Height_of_viewport) = L2 * Vw_scale
                                            End If
                                        End If


                                    Next

                                    Data_table_Viewport_data.Rows(Index1).Item(VW_TWIST) = 2 * PI - GET_Bearing_rad(Point1.X, Point1.Y, Point2.X, Point2.Y)

                                    Index1 = Index1 + 1
                                End If




                            End If
                        Next






                        Trans1.Commit()

                    End Using


1254:


                End Using

end1:

                Transfer_datatable_to_new_excel_spreadsheet(Data_table_Viewport_data)

            Catch ex As Exception
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try

            Freeze_operations = False
        End If
    End Sub


    Private Sub Button_templates_Generate_sheet_Click(sender As Object, e As EventArgs) Handles Button_templates_Generate_sheet.Click


        If Freeze_operations = False Then
            Freeze_operations = True



            If TextBox_templates_Output_Directory.Text = "" Then
                MsgBox("Please specify the output folder")
                Freeze_operations = False
                Exit Sub
            End If
            If Not Strings.Right(TextBox_templates_Output_Directory.Text, 1) = "\" Then
                TextBox_templates_Output_Directory.Text = TextBox_templates_Output_Directory.Text & "\"
            End If

            If TextBox_templates_dwt_template.Text = "" Then
                MsgBox("Please specify the template file")
                Freeze_operations = False
                Exit Sub
            End If
            If Not Strings.Right(TextBox_templates_dwt_template.Text, 3).ToUpper = "DWG" Then
                MsgBox("Please specify the template file")
                Freeze_operations = False
                Exit Sub
            End If


            If IsNumeric(TextBox_TEMPLATES_main_viewport_center_X.Text) = True Then
                Vw_CenX = CDbl(TextBox_TEMPLATES_main_viewport_center_X.Text)
            End If
            If IsNumeric(TextBox_TEMPLATES_main_viewport_center_y.Text) = True Then
                Vw_CenY = CDbl(TextBox_TEMPLATES_main_viewport_center_y.Text)
            End If






            Dim Empty_array() As ObjectId

            Dim BaseMap_drawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument





            If IO.File.Exists(TextBox_templates_dwt_template.Text) = True Then
                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = BaseMap_drawing.Editor
                Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                Try
















                    If IsNothing(Data_table_Viewport_data) = False Then
                        If Data_table_Viewport_data.Rows.Count > 0 Then



                            Dim Index_existent As Integer = 1


                            If Not Strings.Right(TextBox_templates_Output_Directory.Text, 1) = "\" Then
                                TextBox_templates_Output_Directory.Text = TextBox_templates_Output_Directory.Text & "\"
                            End If

                            If System.IO.Directory.Exists(TextBox_templates_Output_Directory.Text) = True Then




                                For i = 0 To Data_table_Viewport_data.Rows.Count - 1
                                    Dim New_plat_name As String = ""

                                    If IsDBNull(Data_table_Viewport_data.Rows(i).Item(NEW_NAME)) = False And _
                                        IsDBNull(Data_table_Viewport_data.Rows(i).Item(Height_of_viewport)) = False And
                                        IsDBNull(Data_table_Viewport_data.Rows(i).Item(Length_of_viewport)) = False And _
                                        IsDBNull(Data_table_Viewport_data.Rows(i).Item(VW_TARGET_X)) = False And _
                                        IsDBNull(Data_table_Viewport_data.Rows(i).Item(VW_TARGET_Y)) = False And _
                                        IsDBNull(Data_table_Viewport_data.Rows(i).Item(VW_CUST_SCALE)) = False Then

                                        New_plat_name = Data_table_Viewport_data.Rows(i).Item(NEW_NAME)

                                        Dim Fisierul_exista As Boolean = True
                                        Do Until Fisierul_exista = False
                                            If System.IO.File.Exists(TextBox_templates_Output_Directory.Text & New_plat_name & ".dwg") = True Then
                                                Dim Fisierul_indexat_exista As Boolean = True
                                                Do Until Fisierul_indexat_exista = False
                                                    If System.IO.File.Exists(TextBox_templates_Output_Directory.Text & New_plat_name & "_" & Index_existent.ToString & ".dwg") = True Then
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






                                        Dim Fisier_nou As String = TextBox_templates_Output_Directory.Text & New_plat_name & ".dwg"


                                        Dim DocumentManager1 As DocumentCollection = Application.DocumentManager
                                        Dim New_doc As Document = DocumentCollectionExtension.Add(DocumentManager1, TextBox_templates_dwt_template.Text)
                                        DocumentManager1.MdiActiveDocument = New_doc
                                        HostApplicationServices.WorkingDatabase = New_doc.Database
                                        Using lock2 As DocumentLock = New_doc.LockDocument
                                            Using Trans_New_doc As Autodesk.AutoCAD.DatabaseServices.Transaction = New_doc.TransactionManager.StartTransaction
                                                Creaza_layer(Layer_name_Main_Viewport, 40, Layer_name_Main_Viewport, False)


                                                Trans_New_doc.Commit()
                                            End Using
                                        End Using



                                        Dim Anno_scale_name1 As String = ""

                                        Dim DWG_units As Integer

                                        Using lock2 As DocumentLock = New_doc.LockDocument

                                            Using Trans_New_doc As Autodesk.AutoCAD.DatabaseServices.Transaction = New_doc.TransactionManager.StartTransaction
                                                Dim ocm As ObjectContextManager = New_doc.Database.ObjectContextManager
                                                Dim occ As ObjectContextCollection

                                                If IsNothing(ocm) = False Then
                                                    occ = ocm.GetContextCollection("ACDB_ANNOTATIONSCALES")
                                                End If

                                                Anno_scale_name1 = "1" & Chr(34) & "=" & TextBox_template_viewport_scale.Text.ToString & "'"
                                                DWG_units = 1 / Data_table_Viewport_data.Rows(i).Item(VW_CUST_SCALE)

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

                                                Dim H1 As Double = Data_table_Viewport_data.Rows(i).Item(Height_of_viewport)
                                                Dim W1 As Double = Data_table_Viewport_data.Rows(i).Item(Length_of_viewport)

                                                Dim Viewport1 As New Viewport
                                                Viewport1.SetDatabaseDefaults()
                                                Viewport1.CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(Vw_CenX + W1 / 2, Vw_CenY + H1 / 2, 0) ' asta e pozitia viewport in paper space
                                                Viewport1.Height = H1
                                                Viewport1.Width = W1
                                                Viewport1.Layer = Layer_name_Main_Viewport
                                                Viewport1.ViewDirection = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis
                                                Viewport1.ViewTarget = New Point3d(Data_table_Viewport_data.Rows(i).Item(VW_TARGET_X), Data_table_Viewport_data.Rows(i).Item(VW_TARGET_Y), 0) ' asta e pozitia viewport in MODEL space
                                                Viewport1.ViewCenter = Autodesk.AutoCAD.Geometry.Point2d.Origin
                                                Viewport1.TwistAngle = Data_table_Viewport_data.Rows(i).Item(VW_TWIST) ' asta e PT TWIST



                                                BTrecord_new_doc_PS.AppendEntity(Viewport1)
                                                Trans_New_doc.AddNewlyCreatedDBObject(Viewport1, True)

                                                Viewport1.On = True
                                                Viewport1.CustomScale = Data_table_Viewport_data.Rows(i).Item(VW_CUST_SCALE)

                                                Viewport1.Locked = True

                                                Dim Colectie_atr_name As New Specialized.StringCollection
                                                Dim Colectie_atr_value As New Specialized.StringCollection

                                                If Not TextBox_north_arrow.Text = "" Then
                                                    Dim Block_north As Autodesk.AutoCAD.DatabaseServices.BlockReference
                                                    If IsNumeric(TextBox_north_arrow_Big_X.Text) = True And IsNumeric(TextBox_north_arrow_Big_y.Text) = True Then
                                                        Punct_insertie_north_arrow = New Point3d(CDbl(TextBox_north_arrow_Big_X.Text), CDbl(TextBox_north_arrow_Big_y.Text), 0)
                                                    End If

                                                    Dim Nume_North_arrow As String = TextBox_north_arrow.Text
                                                    Dim Scaleblock As Double = 1
                                                    If IsNumeric(TextBox_blockScale.Text) = True Then
                                                        Scaleblock = CDbl(TextBox_blockScale.Text)
                                                    End If

                                                    Block_north = InsertBlock_with_multiple_atributes("", Nume_North_arrow, Punct_insertie_north_arrow, Scaleblock, BTrecord_new_doc_PS, Layer_name_Blocks, Colectie_atr_name, Colectie_atr_value)
                                                    Block_north.Rotation = Data_table_Viewport_data.Rows(i).Item(VW_TWIST)
                                                End If


                                                Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current

                                                Dim Tilemode1 As Integer = Application.GetSystemVariable("TILEMODE")
                                                Dim CVport1 As Integer = Application.GetSystemVariable("CVPORT")

                                                If (Tilemode1 = 0 And Not CVport1 = 1) Then
                                                    New_doc.Editor.SwitchToPaperSpace()
                                                End If
                                                If Tilemode1 = 1 Then
                                                    Application.SetSystemVariable("TILEMODE", 0)
                                                End If

                                                Dim Layoutdict As DBDictionary = Trans_New_doc.GetObject(New_doc.Database.LayoutDictionaryId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                                Dim Layout1 As Layout = Trans_New_doc.GetObject(LayoutManager1.GetLayoutId(LayoutManager1.CurrentLayout), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                                                Layout1.LayoutName = New_plat_name


                                                Trans_New_doc.Commit()
                                            End Using



                                            New_doc.Database.SaveAs(Fisier_nou, True, DwgVersion.Current, BaseMap_drawing.Database.SecurityParameters)
                                        End Using
                                        DocumentExtension.CloseAndDiscard(New_doc)
                                    Else
                                        Dim DEBUG As String
                                        DEBUG = "INVESTIGATE"

                                    End If
                                Next
                            End If 'If System.IO.Directory.Exists(TextBox_template_Output_Directory.Text) = True Then

                        End If
                    End If



                    MsgBox("You are done")
                Catch ex As Exception
                    Freeze_operations = False
                    MsgBox(ex.Message)
                End Try
            Else
                MsgBox("TEMPLATE " & TextBox_templates_dwt_template.Text & " DOES NOT EXIST")
            End If
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_read_rectangle_Click(sender As Object, e As EventArgs) Handles Button_read_rectangle.Click

        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Empty_array() As ObjectId
            Dim This_drawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = This_drawing.Editor
            Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Try

                Data_table_Viewport_data = New System.Data.DataTable
                Data_table_Viewport_data.Columns.Add(Length_of_viewport, GetType(Double))
                Data_table_Viewport_data.Columns.Add(Height_of_viewport, GetType(Double))
                Data_table_Viewport_data.Columns.Add(VW_TARGET_X, GetType(Double))
                Data_table_Viewport_data.Columns.Add(VW_TARGET_Y, GetType(Double))
                Data_table_Viewport_data.Columns.Add(VW_TWIST, GetType(Double))
                Data_table_Viewport_data.Columns.Add(VW_CUST_SCALE, GetType(Double))

                Data_table_Viewport_data.Columns.Add(NEW_NAME, GetType(String))





                Using lock1 As DocumentLock = This_drawing.LockDocument


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = This_drawing.TransactionManager.StartTransaction


                        Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                        Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_Prompt.MessageForAdding = vbLf & "Select Template Rectangle"

                        Object_Prompt.SingleOnly = True

                        Rezultat1 = Editor1.GetSelection(Object_Prompt)


                        If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Editor1.WriteMessage(vbLf & "Command:")
                            Freeze_operations = False
                            Editor1.SetImpliedSelection(Empty_array)
                            Exit Sub
                        End If



                        Dim Index1 As Integer = 0
                        Dim Start1 As String = TextBox_start_number.Text

                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                            Dim Poly1 As Polyline
                            Poly1 = Ent1
                            If Poly1.NumberOfVertices >= 4 Then

                                Dim OLD_OSnap As Integer = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE")

                                Dim NEW_OSnap As Integer = Autodesk.AutoCAD.EditorInput.ObjectSnapMasks.Near

                                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap)

                                Dim Result_point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                                Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Specify viewport top left side point:")

                                PP1.AllowNone = False
                                Result_point1 = Editor1.GetPoint(PP1)
                                If Result_point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                    Freeze_operations = False
                                    Exit Sub
                                End If


                                Dim Result_point2 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                                Dim PP2 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Specify viewport top right side point:")

                                PP2.AllowNone = False
                                PP2.UseBasePoint = True
                                PP2.BasePoint = Result_point1.Value
                                Result_point2 = Editor1.GetPoint(PP2)
                                If Result_point2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                    Freeze_operations = False
                                    Exit Sub
                                End If

                                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)


                                Dim Point1 As New Point3d(0, 0, 0)
                                Dim Point2 As New Point3d(0, 0, 0)

                                Dim L1 As Double = 0

                                Data_table_Viewport_data.Rows.Add()
                                Data_table_Viewport_data.Rows(Index1).Item(VW_TARGET_X) = (Poly1.GetPointAtParameter(0).X + Poly1.GetPointAtParameter(2).X) / 2
                                Data_table_Viewport_data.Rows(Index1).Item(VW_TARGET_Y) = (Poly1.GetPointAtParameter(0).Y + Poly1.GetPointAtParameter(2).Y) / 2
                                Data_table_Viewport_data.Rows(Index1).Item(VW_CUST_SCALE) = 1

                                Dim Pt1 As New Point3d()
                                Pt1 = Poly1.GetClosestPointTo(Result_point1.Value, False)
                                Dim Pt2 As New Point3d()
                                Pt2 = Poly1.GetClosestPointTo(Result_point2.Value, False)

                                Dim Param1 As Double = Poly1.GetParameterAtPoint(Pt1)
                                Dim Param2 As Double = Poly1.GetParameterAtPoint(Pt2)

                                If Param1 = Param2 Then
                                    Freeze_operations = False
                                    Exit Sub
                                End If




                                If Param1 >= 3 & Param2 >= 3 Then

                                    Point1 = Poly1.GetPointAtParameter(3)
                                    Point2 = Poly1.GetPointAtParameter(0)

                                Else

                                    Point1 = Poly1.GetPointAtParameter(Floor(Param1))
                                    Point2 = Poly1.GetPointAtParameter(Floor(Param1) + 1)

                                End If
                                L1 = Point1.GetVectorTo(Point2).Length

                                Dim H1 As Double = 0

                                If Floor(Param1) = 0 Or Floor(Param1) = 2 Then
                                    Point1 = Poly1.GetPointAtParameter(1)
                                    Point2 = Poly1.GetPointAtParameter(2)
                                    H1 = Point1.GetVectorTo(Point2).Length
                                Else
                                    Point1 = Poly1.GetPointAtParameter(0)
                                    Point2 = Poly1.GetPointAtParameter(1)
                                    H1 = Point1.GetVectorTo(Point2).Length
                                End If
                                Data_table_Viewport_data.Rows(Index1).Item(Length_of_viewport) = L1
                                Data_table_Viewport_data.Rows(Index1).Item(Height_of_viewport) = H1



                                Data_table_Viewport_data.Rows(Index1).Item(VW_TWIST) = 2 * PI - GET_Bearing_rad(Result_point1.Value.X, Result_point1.Value.Y, Result_point2.Value.X, Result_point2.Value.Y)

                                Index1 = Index1 + 1
                            End If




                        End If







                        Trans1.Commit()

                    End Using


1254:


                End Using

end1:

                If IsNothing(Data_table_Viewport_data) = False Then
                    If Data_table_Viewport_data.Rows.Count > 0 Then
                        Add_to_clipboard_Data_table(Data_table_Viewport_data)
                    End If
                End If

            Catch ex As Exception
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try

            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_adjust_viewport_Click(sender As Object, e As EventArgs) Handles Button_adjust_viewport.Click

        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Empty_array() As ObjectId

            Dim This_drawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = This_drawing.Editor
            Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Try

                If Data_table_Viewport_data.Rows.Count > 0 Then
                    Dim Index_existent As Integer = 1
                    For i = 0 To Data_table_Viewport_data.Rows.Count - 1
                        Dim New_plat_name As String = ""
                        If IsDBNull(Data_table_Viewport_data.Rows(i).Item(Height_of_viewport)) = False And _
                            IsDBNull(Data_table_Viewport_data.Rows(i).Item(Length_of_viewport)) = False And _
                            IsDBNull(Data_table_Viewport_data.Rows(i).Item(VW_TARGET_X)) = False And _
                            IsDBNull(Data_table_Viewport_data.Rows(i).Item(VW_TARGET_Y)) = False Then

                            Vw_scale = 1

                            If IsNumeric(TextBox_adjust_viewport_scale.Text) = True Then
                                Vw_scale = 1 / CDbl(TextBox_adjust_viewport_scale.Text)
                            End If

                            Dim Anno_scale_name1 As String = ""
                            Dim DWG_units As Integer
                            Using lock2 As DocumentLock = This_drawing.LockDocument
                                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = This_drawing.TransactionManager.StartTransaction
                                    Dim ocm As ObjectContextManager = This_drawing.Database.ObjectContextManager
                                    Dim occ As ObjectContextCollection

                                    If IsNothing(ocm) = False Then
                                        occ = ocm.GetContextCollection("ACDB_ANNOTATIONSCALES")
                                    End If

                                    Anno_scale_name1 = "1" & Chr(34) & "=" & TextBox_template_viewport_scale.Text.ToString & "'"
                                    DWG_units = 1 / Vw_scale

                                    If IsNothing(occ) = False Then
                                        Dim asc As New AnnotationScale
                                        asc.Name = Anno_scale_name1
                                        asc.PaperUnits = 1
                                        asc.DrawingUnits = DWG_units
                                        If occ.HasContext(asc.Name) = False Then
                                            occ.AddContext(asc)
                                        End If
                                    End If

                                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                                    Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                                    Object_Prompt.MessageForAdding = vbLf & "Select Viewport"
                                    Object_Prompt.SingleOnly = True
                                    Rezultat1 = Editor1.GetSelection(Object_Prompt)

                                    If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                        Editor1.WriteMessage(vbLf & "Command:")
                                        Freeze_operations = False
                                        Editor1.SetImpliedSelection(Empty_array)
                                        Exit Sub
                                    End If

                                    Dim vpId As ObjectId
                                    If TypeOf Trans1.GetObject(Rezultat1.Value(0).ObjectId, OpenMode.ForRead) Is Polyline Then
                                        vpId = LayoutManager.Current.GetNonRectangularViewportIdFromClipId(Rezultat1.Value(0).ObjectId)
                                        If vpId = Nothing Then
                                            Freeze_operations = False
                                            Exit Sub
                                        End If
                                    Else
                                        vpId = Rezultat1.Value(0).ObjectId
                                    End If


                                    Dim Viewport1 As Viewport = TryCast(Trans1.GetObject(vpId, OpenMode.ForWrite), Viewport)
                                    If IsNothing(Viewport1) = False Then
                                        Dim H1 As Double = Data_table_Viewport_data.Rows(i).Item(Height_of_viewport) * Vw_scale
                                        Dim W1 As Double = Data_table_Viewport_data.Rows(i).Item(Length_of_viewport) * Vw_scale
                                        Viewport1.Height = H1
                                        Viewport1.Width = W1
                                        Viewport1.ViewDirection = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis
                                        Viewport1.ViewTarget = New Point3d(Data_table_Viewport_data.Rows(i).Item(VW_TARGET_X), Data_table_Viewport_data.Rows(i).Item(VW_TARGET_Y), 0) ' asta e pozitia viewport in MODEL space
                                        Viewport1.ViewCenter = Autodesk.AutoCAD.Geometry.Point2d.Origin
                                        Viewport1.TwistAngle = Data_table_Viewport_data.Rows(i).Item(VW_TWIST) ' asta e PT TWIST
                                        Viewport1.On = True
                                        Viewport1.CustomScale = Vw_scale
                                        Viewport1.Locked = True
                                        Trans1.Commit()
                                    End If
                                End Using
                            End Using
                        Else
                            Dim DEBUG As String
                            DEBUG = "INVESTIGATE"
                        End If
                    Next



                End If



            Catch ex As Exception
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try

            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_adjust_rectangle_Click(sender As Object, e As EventArgs) Handles Button_adjust_rectangle.Click
        If Freeze_operations = False Then
            Freeze_operations = True


            If IsNumeric(TextBox_main_viewport_height.Text) = True Then
                Vw_height = CDbl(TextBox_main_viewport_height.Text)
            End If


            Dim Empty_array() As ObjectId

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Try

                Dim Poly2D As Polyline


                Using lock1 As DocumentLock = ThisDrawing.LockDocument

                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                        Dim Prompt_optionsCL As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select Centerline:")
                        Prompt_optionsCL.SetRejectMessage(vbLf & "You did not selected a polyline")
                        Prompt_optionsCL.AddAllowedClass(GetType(Polyline), True)

                        Dim Rezultat_CL As Autodesk.AutoCAD.EditorInput.PromptEntityResult = Editor1.GetEntity(Prompt_optionsCL)
                        If Not Rezultat_CL.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Freeze_operations = False
                            Editor1.SetImpliedSelection(Empty_array)
                            Editor1.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If
                        Poly2D = Trans1.GetObject(Rezultat_CL.ObjectId, OpenMode.ForRead)


                        Dim Prompt_optionsrec As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select rectangle:")
                        Prompt_optionsrec.SetRejectMessage(vbLf & "You did not selected a polyline")
                        Prompt_optionsrec.AddAllowedClass(GetType(Polyline), True)

                        Dim Rezultat_rec As Autodesk.AutoCAD.EditorInput.PromptEntityResult = Editor1.GetEntity(Prompt_optionsrec)
                        If Not Rezultat_rec.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Freeze_operations = False
                            Editor1.SetImpliedSelection(Empty_array)
                            Editor1.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If

                        Dim Rect_0 As Polyline = Trans1.GetObject(Rezultat_rec.ObjectId, OpenMode.ForWrite)


                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        'Dim BTrecord_MS As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(BlockTable1(BlockTableRecord.ModelSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                        Dim Obj_id_old As ObjectId = Rect_0.ObjectId
                        If IsNothing(Data_table_matchline) = False Then
                            If Data_table_matchline.Rows.Count > 0 Then
                                Dim Index0 As Integer = -1
                                For i = 0 To Data_table_matchline.Rows.Count - 1
                                    If Data_table_matchline.Rows(i).Item("OBJECT_ID") = Obj_id_old Then
                                        Index0 = i
                                        Exit For
                                    End If
                                Next

                                If Not Index0 = -1 Then




                                    Dim Line1 As New Line(Rect_0.GetPointAtParameter(2), Rect_0.GetPointAtParameter(3))
                                    Line1.TransformBy(Matrix3d.Scaling(10000, Line1.StartPoint))
                                    Line1.TransformBy(Matrix3d.Scaling(10000, Line1.EndPoint))

                                    Dim Jig1 As New Jig_rectangle_viewport_SHEET_CUTTER_manual_up_down(Poly2D.GetPointAtDist(Data_table_matchline.Rows(Index0).Item("M2")), Line1)
                                    Jig1.AddEntity(Rect_0)
                                    Dim jigRes As Autodesk.AutoCAD.EditorInput.PromptResult = ThisDrawing.Editor.Drag(Jig1)
                                    If jigRes.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                        Jig1.TransformEntities()
                                    End If

                                    Trans1.TransactionManager.QueueForGraphicsFlush()





                                End If

                            End If
                        End If




                        Trans1.Commit()
                        Editor1.WriteMessage(vbLf & "Command:")
                    End Using
                End Using

end1:

                If IsNothing(Data_table_matchline) = False Then
                    If Data_table_matchline.Rows.Count > 0 Then
                        Add_to_clipboard_Data_table(Data_table_matchline)



                    End If
                End If

            Catch ex As Exception
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try

            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_manual_rectangles()
        If Freeze_operations = False Then
            Freeze_operations = True






            If IsNumeric(TextBox_matchline_length.Text) = True Then
                Match_distance = CDbl(TextBox_matchline_length.Text)
            End If

            If IsNumeric(TextBox_main_viewport_height.Text) = True Then
                Vw_height = CDbl(TextBox_main_viewport_height.Text)
            End If


            Dim Empty_array() As ObjectId

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Try
                Dim Index1 As Double = 0

                Dim Poly2D As Polyline

                Using lock1 As DocumentLock = ThisDrawing.LockDocument
                    Creaza_layer(Layer_name_Main_Viewport, 5, Layer_name_Main_Viewport, False)

                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim Prompt_optionsCL As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select Centerline:")
                        Prompt_optionsCL.SetRejectMessage(vbLf & "You did not selected a polyline")
                        Prompt_optionsCL.AddAllowedClass(GetType(Polyline), True)

                        Dim Rezultat_CL As Autodesk.AutoCAD.EditorInput.PromptEntityResult = Editor1.GetEntity(Prompt_optionsCL)
                        If Not Rezultat_CL.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Freeze_operations = False
                            Editor1.SetImpliedSelection(Empty_array)
                            Editor1.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If

                        Poly2D = Trans1.GetObject(Rezultat_CL.ObjectId, OpenMode.ForRead)

                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        'Dim BTrecord_MS As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(BlockTable1(BlockTableRecord.ModelSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                        Dim Colorindex As Integer = 1
                        Dim Este_primul As Boolean = True
                        Dim Last_pt As New Point3d


123:

                        Dim Result_point_m1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                        Dim PP1m As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Please pick start location:")
                        If Este_primul = True Then

                            PP1m.AllowNone = False
                            Result_point_m1 = Editor1.GetPoint(PP1m)


                            If Result_point_m1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                Trans1.Commit()
                                GoTo end1
                            End If
                            Last_pt = Poly2D.GetClosestPointTo(Result_point_m1.Value, Vector3d.ZAxis, False)
                        End If


                        Dim Result_point_m2 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                        Dim PP2m As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Please pick end location:")

                        PP2m.AllowNone = False
                        PP2m.UseBasePoint = True
                        PP2m.BasePoint = Last_pt
                        Result_point_m2 = Editor1.GetPoint(PP2m)
                        If Result_point_m2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                            Trans1.Commit()
                            GoTo end1
                        End If

                        Dim dist1 As Double

                        If Este_primul = True Then
                            dist1 = Poly2D.GetDistAtPoint(Poly2D.GetClosestPointTo(Result_point_m1.Value, Vector3d.ZAxis, False))
                        Else
                            dist1 = Poly2D.GetDistAtPoint(Poly2D.GetClosestPointTo(Last_pt, Vector3d.ZAxis, False))
                        End If

                        Last_pt = Poly2D.GetClosestPointTo(Result_point_m2.Value, Vector3d.ZAxis, False)

                        Dim dist2 As Double = Poly2D.GetDistAtPoint(Last_pt)



                        If dist1 > dist2 Then
                            Dim t As Double = dist1
                            dist1 = dist2
                            dist2 = t
                        End If

                        Dim Point1 As New Point3d
                        Point1 = Poly2D.GetPointAtDist(dist1)
                        Dim Point2 As New Point3d
                        Point2 = Poly2D.GetPointAtDist(dist2)

                        Dim Poly1r As New Autodesk.AutoCAD.DatabaseServices.Polyline

                        Poly1r = creaza_rectangle_viewport(Point1, Point2, Colorindex)
                        Poly1r.Layer = Layer_name_Main_Viewport

                        BTrecord.AppendEntity(Poly1r)
                        Trans1.AddNewlyCreatedDBObject(Poly1r, True)
                        Trans1.TransactionManager.QueueForGraphicsFlush()

                        If Este_primul = True Then
                            Dim M1_p As Double = Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("M1")
                            Dim M2_p As Double = Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("M2")
                            Dim ob_id As ObjectId = Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("OBJECT_ID")
                            Dim PolyR As Polyline = Trans1.GetObject(ob_id, OpenMode.ForWrite)

                            Dim Point01 As New Point3d
                            Point01 = Poly2D.GetPointAtDist(M1_p)
                            Dim Point02 As New Point3d
                            Point02 = Poly2D.GetPointAtDist(M2_p)
                            Dim Poly0r As New Autodesk.AutoCAD.DatabaseServices.Polyline

                            Poly0r = creaza_rectangle_viewport(Point01, Point1, PolyR.ColorIndex)
                            Poly0r.Layer = Layer_name_Main_Viewport

                            BTrecord.AppendEntity(Poly0r)
                            Trans1.AddNewlyCreatedDBObject(Poly0r, True)
                            Trans1.TransactionManager.QueueForGraphicsFlush()

                            PolyR.Erase()
                            Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("M2") = Round(dist1, 0)
                            Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("OBJECT_ID") = Poly0r.ObjectId
                        End If




                        Data_table_matchline.Rows.Add()
                        Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("M1") = Round(dist1, 0)
                        Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("M2") = Round(dist2, 0)
                        Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("OBJECT_ID") = Poly1r.ObjectId

                        Colorindex = Colorindex + 1
                        If Colorindex > 7 Then Colorindex = 1
                        Este_primul = False
                        GoTo 123

                    End Using
                End Using

end1:
                Editor1.WriteMessage(vbLf & "Command:")
                Editor1.SetImpliedSelection(Empty_array)
                If IsNothing(Data_table_matchline) = False Then
                    If Data_table_matchline.Rows.Count > 0 Then
                        Add_to_clipboard_Data_table(Data_table_matchline)
                    End If
                End If

            Catch ex As Exception
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try

            Freeze_operations = False
        End If
    End Sub


    Private Sub Button_rectangles_2Pts_Click(sender As Object, e As EventArgs) Handles Button_rectangles_2Pts.Click


        If Freeze_operations = False Then
            Freeze_operations = True






            If IsNumeric(TextBox_matchline_length.Text) = True Then
                Match_distance = CDbl(TextBox_matchline_length.Text)
            End If

            If IsNumeric(TextBox_main_viewport_height.Text) = True Then
                Vw_height = CDbl(TextBox_main_viewport_height.Text)
            End If


            Dim Empty_array() As ObjectId

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Try
                Dim Index1 As Double = 0





                Dim Poly2D As Polyline


                Using lock1 As DocumentLock = ThisDrawing.LockDocument
                    Creaza_layer(Layer_name_Main_Viewport, 5, Layer_name_Main_Viewport, False)

                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                        Dim Prompt_optionsCL As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select Centerline:")
                        Prompt_optionsCL.SetRejectMessage(vbLf & "You did not selected a polyline")
                        Prompt_optionsCL.AddAllowedClass(GetType(Polyline), True)

                        Dim Rezultat_CL As Autodesk.AutoCAD.EditorInput.PromptEntityResult = Editor1.GetEntity(Prompt_optionsCL)
                        If Not Rezultat_CL.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Freeze_operations = False
                            Editor1.SetImpliedSelection(Empty_array)
                            Editor1.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If
                        Poly2D = Trans1.GetObject(Rezultat_CL.ObjectId, OpenMode.ForRead)


                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        'Dim BTrecord_MS As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(BlockTable1(BlockTableRecord.ModelSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                        Dim dist1 As Double = 0



                        Data_table_matchline = New System.Data.DataTable
                        Data_table_matchline.Columns.Add("OBJECT_ID", GetType(ObjectId))
                        Data_table_matchline.Columns.Add("M1", GetType(Double))
                        Data_table_matchline.Columns.Add("M2", GetType(Double))



                        Dim dist2 As Double = dist1 + Match_distance
                        Dim Ultimul As Boolean = False
                        Dim Colorindex As Integer = 1

                        If Colorindex = 3233 Then
                            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                            W1 = Get_NEW_worksheet_from_Excel()
                        End If

                        Dim Este_primul As Boolean = True
                        Dim Last_pt As New Point3d



123:
                        Dim Point1 As New Point3d
                        Point1 = Poly2D.GetPointAtDist(dist1)
                        Dim Point2 As New Point3d
                        Point2 = Poly2D.GetPointAtDist(dist2)



                        Dim Poly1r As New Autodesk.AutoCAD.DatabaseServices.Polyline

                        Poly1r = creaza_rectangle_viewport(Point1, Point2, Colorindex)
                        Poly1r.Layer = Layer_name_Main_Viewport

                        Dim Col_int As New Point3dCollection
                        Col_int = Intersect_on_both_operands(Poly2D, Poly1r)
                        If Col_int.Count = 2 Then



                            BTrecord.AppendEntity(Poly1r)
                            Trans1.AddNewlyCreatedDBObject(Poly1r, True)
                            Trans1.TransactionManager.QueueForGraphicsFlush()

                            Data_table_matchline.Rows.Add()
                            Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("OBJECT_ID") = Poly1r.ObjectId
                            Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("M1") = Round(dist1, 0)
                            Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("M2") = Round(dist2, 0)


                            dist1 = dist2
                            dist2 = dist2 + Match_distance

                            Colorindex = Colorindex + 1
                            If Colorindex > 7 Then Colorindex = 1

                            If Ultimul = False Then
                                If Poly2D.Length < dist2 Then
                                    dist2 = Poly2D.Length
                                    Ultimul = True
                                End If
                                Este_primul = True
                                GoTo 123
                            End If

                        Else

                            Dim Result_point_m1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                            Dim PP1m As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Please pick start location:")
                            If Este_primul = True Then

                                PP1m.AllowNone = False
                                Result_point_m1 = Editor1.GetPoint(PP1m)


                                If Result_point_m1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                    Trans1.Commit()
                                    GoTo end1
                                End If
                                Last_pt = Poly2D.GetClosestPointTo(Result_point_m1.Value, Vector3d.ZAxis, False)
                            End If


                            Dim Result_point_m2 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                            Dim PP2m As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Please pick end location:")

                            PP2m.AllowNone = False
                            PP2m.UseBasePoint = True
                            PP2m.BasePoint = Last_pt
                            Result_point_m2 = Editor1.GetPoint(PP2m)
                            If Result_point_m2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                Trans1.Commit()
                                GoTo end1
                            End If


                            Dim dist1m As Double
                            If Este_primul = True Then
                                dist1m = Poly2D.GetDistAtPoint(Poly2D.GetClosestPointTo(Result_point_m1.Value, Vector3d.ZAxis, False))
                            Else
                                dist1m = Poly2D.GetDistAtPoint(Poly2D.GetClosestPointTo(Last_pt, Vector3d.ZAxis, False))
                            End If

                            Last_pt = Poly2D.GetClosestPointTo(Result_point_m2.Value, Vector3d.ZAxis, False)

                            Dim dist2m As Double = Poly2D.GetDistAtPoint(Last_pt)



                            If dist1m > dist2m Then
                                Dim t As Double = dist1m
                                dist1m = dist2m
                                dist2m = t
                            End If

                            Dim Point1m As New Point3d
                            Point1m = Poly2D.GetPointAtDist(dist1m)
                            Dim Point2m As New Point3d
                            Point2m = Poly2D.GetPointAtDist(dist2m)

                            Dim Poly1rm As New Autodesk.AutoCAD.DatabaseServices.Polyline

                            Poly1rm = creaza_rectangle_viewport(Point1m, Point2m, Colorindex)
                            Poly1rm.Layer = Layer_name_Main_Viewport

                            BTrecord.AppendEntity(Poly1rm)
                            Trans1.AddNewlyCreatedDBObject(Poly1rm, True)
                            Trans1.TransactionManager.QueueForGraphicsFlush()

                            If Este_primul = True Then
                                Dim M1_p As Double = Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("M1")
                                Dim M2_p As Double = Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("M2")
                                Dim ob_id As ObjectId = Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("OBJECT_ID")
                                Dim PolyR As Polyline = Trans1.GetObject(ob_id, OpenMode.ForWrite)

                                Dim Point01 As New Point3d
                                Point01 = Poly2D.GetPointAtDist(M1_p)
                                Dim Point02 As New Point3d
                                Point02 = Poly2D.GetPointAtDist(M2_p)
                                Dim Poly0r As New Autodesk.AutoCAD.DatabaseServices.Polyline

                                Poly0r = creaza_rectangle_viewport(Point01, Point1m, PolyR.ColorIndex)
                                Poly0r.Layer = Layer_name_Main_Viewport

                                BTrecord.AppendEntity(Poly0r)
                                Trans1.AddNewlyCreatedDBObject(Poly0r, True)
                                Trans1.TransactionManager.QueueForGraphicsFlush()

                                PolyR.Erase()
                                Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("M2") = Round(dist1m, 0)
                                Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("OBJECT_ID") = Poly0r.ObjectId
                            End If




                            Data_table_matchline.Rows.Add()
                            Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("M1") = Round(dist1m, 0)
                            Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("M2") = Round(dist2m, 0)
                            Data_table_matchline.Rows(Data_table_matchline.Rows.Count - 1).Item("OBJECT_ID") = Poly1rm.ObjectId

                            Colorindex = Colorindex + 1
                            If Colorindex > 7 Then Colorindex = 1
                            Este_primul = False

                            dist1 = dist2m
                            dist2 = dist2m + Match_distance
                            If Round(dist1, 0) = Round(Poly2D.Length, 0) Then
                                GoTo 124
                            End If
                            If Round(dist2, 0) > dist2 = Poly2D.Length Then
                                GoTo 124
                            End If

                            GoTo 123



                        End If
124:                    Editor1.WriteMessage(vbLf & "Command:")

                        Trans1.Commit()
                    End Using
                End Using

end1:

                If IsNothing(Data_table_matchline) = False Then
                    If Data_table_matchline.Rows.Count > 0 Then
                        Add_to_clipboard_Data_table(Data_table_matchline)
                    End If
                End If

            Catch ex As Exception
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try

            Freeze_operations = False
        End If
    End Sub
End Class


