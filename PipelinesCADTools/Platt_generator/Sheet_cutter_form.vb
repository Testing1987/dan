Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports ACSMCOMPONENTS20Lib
Public Class Sheet_cutter_form

    Public Shared Vw_scale As Double
    Public Shared Rotatie As Double





    Dim Layer_name_Main_Viewport As String = "VP"

    Dim Layer_name_Blocks As String = "0NORTH"



    Dim Rotatie_originala As Double
    Dim Vw_height As Double = 4.2
    Dim Vw_width As Double = 7
    Dim Vw_CenX As Double = 8.5
    Dim Vw_CenY As Double = 7.4

    Dim VW_TARGET_X As String = "VW_TARGET_X"
    Dim VW_TARGET_Y As String = "VW_TARGET_Y"
    Dim VW_TWIST As String = "VW_TWIST"
    Dim VW_CUST_SCALE As String = "VW_CUST_SCALE"
    Dim NEW_NAME As String = "NEW_NAME"
    Dim STATION1 As String = "STATION1"
    Dim STATION2 As String = "STATION2"
    Dim Punct_insertie_north_arrow As New Point3d(2.7722595, 8.9239174, 0)

    Dim Scale_factor As Double = 0
    Dim View_rotation As Double = 0
    Dim Text_height As Double = 0.08


    Dim Elevatia_cunoscuta As Double = -100000
    Dim Chainage_cunoscuta As Double = -100000
    Dim Chainage_Y As Double
    Dim Elevation_X1 As Double
    Dim Elevation_X2 As Double
    Dim Point_cunoscut As New Point3d
    Dim Data_table_poly As System.Data.DataTable

    Dim Freeze_operations As Boolean = False



    Private Sub Platt_Generator_form_Load(sender As Object, e As EventArgs) Handles Me.Load
        HScrollBar_rotate.Minimum = -180
        HScrollBar_rotate.Maximum = 180
        HScrollBar_rotate.Value = 0






    End Sub

    Private Sub MyForm_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown



        If e.KeyCode = Windows.Forms.Keys.Left Then
            HScrollBar_rotate.Value = HScrollBar_rotate.Value - 1
            TextBox_rotation_ammount.Text = CInt(TextBox_rotation_ammount.Text) - 1
        End If
        If e.KeyCode = Windows.Forms.Keys.Right Then
            HScrollBar_rotate.Value = HScrollBar_rotate.Value + 1
            TextBox_rotation_ammount.Text = CInt(TextBox_rotation_ammount.Text) + 1
        End If

    End Sub

    Private Sub Button_generate_sheet_Click(sender As Object, e As EventArgs) Handles Button_generate_Sheet.Click

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

            If IsNumeric(TextBox_Layout_index_main.Text) = False Then
                MsgBox("Please specify the main viewport layout index")
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
                    Data_table_cu_valori.Columns.Add(STATION1, GetType(Double))
                    Data_table_cu_valori.Columns.Add(STATION2, GetType(Double))

                    Dim Start1 As String = TextBox_start_number.Text



                    Using lock1 As DocumentLock = BaseMap_drawing.LockDocument
                        Creaza_layer(Layer_name_Main_Viewport, 5, Layer_name_Main_Viewport, False)

                        Using Trans_basemap As Autodesk.AutoCAD.DatabaseServices.Transaction = BaseMap_drawing.TransactionManager.StartTransaction
                            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            BlockTable1 = BaseMap_drawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            Dim BTrecord_MS_basemap As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans_basemap.GetObject(BlockTable1(BlockTableRecord.ModelSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptEntityResult

                            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select centerline:")
                            Object_Prompt.SetRejectMessage(vbLf & "Please select a lightweight polyline or a 3d polyline")
                            Object_Prompt.AddAllowedClass(GetType(Polyline), True)
                            Object_Prompt.AddAllowedClass(GetType(Polyline3d), True)


                            Rezultat1 = Editor1.GetEntity(Object_Prompt)

                            Dim Poly2d As Polyline
                            Dim Poly3d As Polyline3d


                            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                If IsNothing(Rezultat1) = False Then
                                    Dim Ent1 As Entity = TryCast(Trans_basemap.GetObject(Rezultat1.ObjectId, OpenMode.ForRead), Entity)
                                    If IsNothing(Ent1) = False Then
                                        If TypeOf Ent1 Is Polyline Then
                                            Poly2d = Ent1
                                            Poly2d.UpgradeOpen()
                                            Poly2d.Elevation = 0
                                        End If
                                        If TypeOf Ent1 Is Polyline3d Then
                                            Poly3d = Ent1

                                            Dim Index2d As Integer = 0
                                            Poly2d = New Polyline


                                            For Each Id1 As ObjectId In Poly3d
                                                Dim vertex1 As PolylineVertex3d
                                                vertex1 = Trans_basemap.GetObject(Id1, OpenMode.ForRead)
                                                Poly2d.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)
                                            Next
                                            Poly2d.Elevation = 0





                                        End If
                                    End If
                                End If
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


                            If CheckBox_PICK_ROTATION.Checked = True Then
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

                            End If




                            Dim PromptPointRezult1 As Autodesk.AutoCAD.EditorInput.PromptPointResult

                            Dim Jig1 As New Jig_rectangle_viewport_SHEET_CUTTER




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





                            Using Trans_basemap2 As Autodesk.AutoCAD.DatabaseServices.Transaction = BaseMap_drawing.TransactionManager.StartTransaction
                                Dim Poly1 As New Autodesk.AutoCAD.DatabaseServices.Polyline
                                Poly1.AddVertexAt(0, New Point2d(pointM.X - 0.5 * Vw_width / Vw_scale, pointM.Y - 0.5 * Vw_height / Vw_scale), 0, 0, 0)
                                Poly1.AddVertexAt(1, New Point2d(pointM.X + 0.5 * Vw_width / Vw_scale, pointM.Y - 0.5 * Vw_height / Vw_scale), 0, 0, 0)
                                Poly1.AddVertexAt(2, New Point2d(pointM.X + 0.5 * Vw_width / Vw_scale, pointM.Y + 0.5 * Vw_height / Vw_scale), 0, 0, 0)
                                Poly1.AddVertexAt(3, New Point2d(pointM.X - 0.5 * Vw_width / Vw_scale, pointM.Y + 0.5 * Vw_height / Vw_scale), 0, 0, 0)
                                Poly1.Closed = True
                                Poly1.Elevation = 0
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


                                        GoTo 1254
                                    Case MsgBoxResult.Yes

                                        Data_table_cu_valori.Rows.Add()
                                        If CheckBox_rotate_to_north.Checked = False Then
                                            Data_table_cu_valori.Rows(Index_data_table_valori).Item(VW_TWIST) = 2 * PI - Rotatie
                                        Else
                                            Data_table_cu_valori.Rows(Index_data_table_valori).Item(VW_TWIST) = 0
                                        End If

                                        Data_table_cu_valori.Rows(Index_data_table_valori).Item(VW_CUST_SCALE) = Vw_scale
                                        Data_table_cu_valori.Rows(Index_data_table_valori).Item(VW_TARGET_X) = pointM.X
                                        Data_table_cu_valori.Rows(Index_data_table_valori).Item(VW_TARGET_Y) = pointM.Y
                                        Data_table_cu_valori.Rows(Index_data_table_valori).Item(NEW_NAME) = TextBox_NEW_NAME_PREFIX.Text & Start1
                                        If IsNothing(Poly3d) = False Then
                                            Dim Col_int As New Point3dCollection
                                            Col_int = Intersect_on_both_operands(Poly1, Poly2d)
                                            If Col_int.Count = 2 Then
                                                Dim Param1 As Double = Poly2d.GetParameterAtPoint(Poly2d.GetClosestPointTo(Col_int(0), Vector3d.ZAxis, True))
                                                Dim Param2 As Double = Poly2d.GetParameterAtPoint(Poly2d.GetClosestPointTo(Col_int(1), Vector3d.ZAxis, True))
                                                If Param1 > Param2 Then
                                                    Dim T As Double = Param1
                                                    Param1 = Param2
                                                    Param2 = T
                                                End If
                                                Dim Sta1 As Double = Poly3d.GetDistanceAtParameter(Param1)
                                                Dim Sta2 As Double = Poly3d.GetDistanceAtParameter(Param2)
                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(STATION1) = Sta1
                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(STATION2) = Sta2

                                            End If

                                            If Col_int.Count = 1 Then
                                                Dim Param1 As Double = Poly2d.GetParameterAtPoint(Poly2d.GetClosestPointTo(Col_int(0), Vector3d.ZAxis, True))
                                                Dim Sta1 As Double = Poly3d.GetDistanceAtParameter(Param1)
                                                Dim Sta2 As Double
                                                If Poly3d.Length - Sta1 > Sta1 Then
                                                    Sta2 = 0
                                                Else
                                                    Sta2 = Poly3d.Length
                                                End If

                                                If Sta1 > Sta2 Then
                                                    Dim T As Double = Sta1
                                                    Sta1 = Sta2
                                                    Sta2 = T
                                                End If

                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(STATION1) = Sta1
                                                Data_table_cu_valori.Rows(Index_data_table_valori).Item(STATION2) = Sta2

                                            End If

                                        Else
                                            If IsNothing(Poly2d) = False Then
                                                Dim Col_int As New Point3dCollection
                                                Col_int = Intersect_on_both_operands(Poly1, Poly2d)
                                                If Col_int.Count = 2 Then

                                                    Dim Sta1 As Double = Poly2d.GetDistAtPoint(Poly2d.GetClosestPointTo(Col_int(0), Vector3d.ZAxis, True))
                                                    Dim Sta2 As Double = Poly2d.GetDistAtPoint(Poly2d.GetClosestPointTo(Col_int(1), Vector3d.ZAxis, True))
                                                    If Sta1 > Sta2 Then
                                                        Dim T As Double = Sta1
                                                        Sta1 = Sta2
                                                        Sta2 = T
                                                    End If
                                                    Data_table_cu_valori.Rows(Index_data_table_valori).Item(STATION1) = Sta1
                                                    Data_table_cu_valori.Rows(Index_data_table_valori).Item(STATION2) = Sta2
                                                End If

                                                If Col_int.Count = 1 Then

                                                    Dim Sta1 As Double = Poly2d.GetDistAtPoint(Poly2d.GetClosestPointTo(Col_int(0), Vector3d.ZAxis, True))
                                                    Dim Sta2 As Double
                                                    If Poly2d.Length - Sta1 > Sta1 Then
                                                        Sta2 = 0
                                                    Else
                                                        Sta2 = Poly2d.Length
                                                    End If

                                                    If Sta1 > Sta2 Then
                                                        Dim T As Double = Sta1
                                                        Sta1 = Sta2
                                                        Sta2 = T
                                                    End If

                                                    Data_table_cu_valori.Rows(Index_data_table_valori).Item(STATION1) = Sta1
                                                    Data_table_cu_valori.Rows(Index_data_table_valori).Item(STATION2) = Sta2

                                                End If


                                            End If
                                        End If


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
                                    Dim New_dwg_name As String = ""

                                    New_dwg_name = Data_table_cu_valori.Rows(s).Item(NEW_NAME)

                                    Dim Fisierul_exista As Boolean = True
                                    Do Until Fisierul_exista = False
                                        If System.IO.File.Exists(TextBox_Output_Directory.Text & New_dwg_name & ".dwg") = True Then
                                            Dim Fisierul_indexat_exista As Boolean = True
                                            Do Until Fisierul_indexat_exista = False
                                                If System.IO.File.Exists(TextBox_Output_Directory.Text & New_dwg_name & "_" & Index_existent.ToString & ".dwg") = True Then
                                                    Index_existent = Index_existent + 1
                                                Else
                                                    New_dwg_name = New_dwg_name & "_" & Index_existent.ToString
                                                    Fisierul_indexat_exista = False
                                                End If
                                            Loop
                                        Else
                                            Fisierul_exista = False
                                        End If
                                    Loop






                                    Dim Fisier_nou As String = TextBox_Output_Directory.Text & New_dwg_name & ".dwg"


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
                                            Dim AnnoScale1 As AnnotationScale

                                            If IsNothing(occ) = False Then
                                                AnnoScale1 = New AnnotationScale
                                                AnnoScale1.Name = Anno_scale_name1
                                                AnnoScale1.PaperUnits = 1
                                                AnnoScale1.DrawingUnits = DWG_units
                                                If occ.HasContext(AnnoScale1.Name) = False Then
                                                    occ.AddContext(AnnoScale1)
                                                Else
                                                    AnnoScale1 = occ.GetContext(AnnoScale1.Name)
                                                End If

                                            End If

                                            Dim BlockTable_new_doc As BlockTable = Trans_New_doc.GetObject(New_doc.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                                            Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                                            Dim Layoutdict As DBDictionary = Trans_New_doc.GetObject(New_doc.Database.LayoutDictionaryId, OpenMode.ForRead)

                                            Dim Tilemode1 As Integer = Application.GetSystemVariable("TILEMODE")
                                            Dim CVport1 As Integer = Application.GetSystemVariable("CVPORT")

                                            If (Tilemode1 = 0 And Not CVport1 = 1) Then
                                                New_doc.Editor.SwitchToPaperSpace()
                                            End If
                                            If Tilemode1 = 1 Then
                                                Application.SetSystemVariable("TILEMODE", 0)
                                            End If

                                            For Each entry As DBDictionaryEntry In Layoutdict
                                                Dim Layout1 As Layout = Trans_New_doc.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                                If Layout1.TabOrder = CInt(TextBox_Layout_index_main.Text) And Not Layout1.TabOrder = 0 Then
                                                    Layout1.LayoutName = New_dwg_name
                                                    LayoutManager1.CurrentLayout = Layout1.LayoutName


                                                    Dim BTrecord_new_doc_PS As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans_New_doc.GetObject(Layout1.BlockTableRecordId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
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
                                                    Viewport1.AnnotationScale = AnnoScale1
                                                    Viewport1.Locked = True


                                                    Dim Colectie_atr_name As New Specialized.StringCollection
                                                    Dim Colectie_atr_value As New Specialized.StringCollection





                                                    If Not TextBox_north_arrow.Text = "" Then
                                                        Dim Btr1 As BlockTableRecord

                                                        Try
                                                            Btr1 = TryCast(Trans_New_doc.GetObject(BlockTable_new_doc(TextBox_north_arrow.Text), OpenMode.ForRead), BlockTableRecord)

                                                            If IsNothing(Btr1) = False Then
                                                                If Not Btr1 = Nothing Then
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
                                                            End If
                                                        Catch ex As Exception

                                                        End Try




                                                    End If
                                                End If
                                            Next

                                            If IsNumeric(TextBox_Layout_index_profile.Text) = True Then
                                                If IsNothing(Data_table_poly) = False Then
                                                    For Each entry As DBDictionaryEntry In Layoutdict
                                                        Dim Layout1 As Layout = Trans_New_doc.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                                        If Layout1.TabOrder = CInt(TextBox_Layout_index_profile.Text) And Not Layout1.TabOrder = 0 Then
                                                            If Data_table_poly.Rows.Count > 0 Then
                                                                If IsDBNull(Data_table_cu_valori.Rows(s).Item(STATION1)) = False And IsDBNull(Data_table_cu_valori.Rows(s).Item(STATION2)) = False Then


                                                                    If IsNumeric(TextBox_H1.Text) = False Then
                                                                        MsgBox("Please specify the HEIGHT!")
                                                                        Freeze_operations = False
                                                                        Exit Sub
                                                                    End If
                                                                    If IsNumeric(TextBox_H2.Text) = False Then
                                                                        MsgBox("Please specify the HEIGHT!")
                                                                        Freeze_operations = False
                                                                        Exit Sub
                                                                    End If
                                                                    If IsNumeric(TextBox_W3.Text) = False Then
                                                                        MsgBox("Please specify the WIDTH!")
                                                                        Freeze_operations = False
                                                                        Exit Sub
                                                                    End If
                                                                    If IsNumeric(TextBox_W2.Text) = False Then
                                                                        MsgBox("Please specify the WIDTH!")
                                                                        Freeze_operations = False
                                                                        Exit Sub
                                                                    End If
                                                                    If IsNumeric(TextBox_x.Text) = False Then
                                                                        MsgBox("Please specify the X COORDINATE!")
                                                                        Freeze_operations = False
                                                                        Exit Sub
                                                                    End If
                                                                    If IsNumeric(TextBox_y.Text) = False Then
                                                                        MsgBox("Please specify the Y COORDINATE!")
                                                                        Freeze_operations = False
                                                                        Exit Sub
                                                                    End If
                                                                    If IsNumeric(TextBox_H_SCALE.Text) = False Then
                                                                        MsgBox("Please specify the HORIZONTAL SCALE!")
                                                                        Freeze_operations = False
                                                                        Exit Sub
                                                                    End If
                                                                    If IsNumeric(TextBox_V_SCALE.Text) = False Then
                                                                        MsgBox("Please specify the VERTICAL SCALE!")
                                                                        Freeze_operations = False
                                                                        Exit Sub
                                                                    End If

                                                                    Dim Sta1 As Double = Data_table_cu_valori.Rows(s).Item(STATION1)
                                                                    Dim Sta2 As Double = Data_table_cu_valori.Rows(s).Item(STATION2)
                                                                    Dim Hscale As Double = 1 / CDbl(TextBox_H_SCALE.Text)
                                                                    Dim Vscale As Double = 1 / CDbl(TextBox_V_SCALE.Text)

                                                                    Dim H1 As Double = CDbl(TextBox_H1.Text)
                                                                    Dim H2 As Double = CDbl(TextBox_H2.Text)
                                                                    Dim W1 As Double = CDbl(TextBox_W3.Text)
                                                                    Dim W2 As Double = CDbl(TextBox_W2.Text)
                                                                    Dim x As Double = CDbl(TextBox_x.Text)
                                                                    Dim y As Double = CDbl(TextBox_y.Text)

                                                                    If Hscale <= 0 Or Vscale <= 0 Or H1 <= 0 Or H2 <= 0 Or W1 <= 0 Or W2 <= 0 Then
                                                                        MsgBox("Negative values not allowed")
                                                                        Freeze_operations = False
                                                                        Exit Sub
                                                                    End If

                                                                    LayoutManager1.CurrentLayout = Layout1.LayoutName

                                                                    Dim BTrecordPS As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                                                    BTrecordPS = Trans_New_doc.GetObject(Layout1.BlockTableRecordId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                                                                    Dim Point0 As Point3d
                                                                    Dim Point1 As Point3d
                                                                    Dim Point2 As Point3d
                                                                    Dim Point3 As Point3d

                                                                    If RadioButton_left_right.Checked = True Then
                                                                        Point0 = New Point3d(Point_cunoscut.X - Chainage_cunoscuta * Hscale, Point_cunoscut.Y - Elevatia_cunoscuta * Vscale, 0)
                                                                        Point1 = New Point3d(Point0.X + Sta1 * Hscale, Point0.Y + Elevatia_cunoscuta * Vscale, 0)
                                                                        Point2 = New Point3d(Point0.X + Sta2 * Hscale, Point0.Y + Elevatia_cunoscuta * Vscale, 0)
                                                                    Else
                                                                        Point0 = New Point3d(Point_cunoscut.X + Chainage_cunoscuta * Hscale, Point_cunoscut.Y - Elevatia_cunoscuta * Vscale, 0)
                                                                        Point1 = New Point3d(Point0.X - Sta1 * Hscale, Point0.Y + Elevatia_cunoscuta * Vscale, 0)
                                                                        Point2 = New Point3d(Point0.X - Sta2 * Hscale, Point0.Y + Elevatia_cunoscuta * Vscale, 0)
                                                                    End If
                                                                    Point3 = New Point3d((Point1.X + Point2.X) / 2, (Point1.Y + Point2.Y) / 2, 0)

                                                                    Dim Poly1 As New Polyline
                                                                    For i = 0 To Data_table_poly.Rows.Count - 1
                                                                        Poly1.AddVertexAt(i, New Point2d(Data_table_poly.Rows(i).Item("X"), Data_table_poly.Rows(i).Item("Y")), 0, 0, 0)
                                                                    Next

                                                                    Dim Linie1 As New Line(New Point3d(Point1.X, Point1.Y - 100000, 0), New Point3d(Point1.X, Point1.Y + 100000, 0))
                                                                    Dim Linie2 As New Line(New Point3d(Point2.X, Point2.Y - 100000, 0), New Point3d(Point2.X, Point2.Y + 100000, 0))

                                                                    Dim ColInt1 As New Point3dCollection
                                                                    Dim ColInt2 As New Point3dCollection
                                                                    Poly1.IntersectWith(Linie1, Intersect.OnBothOperands, ColInt1, IntPtr.Zero, IntPtr.Zero)
                                                                    Poly1.IntersectWith(Linie2, Intersect.OnBothOperands, ColInt2, IntPtr.Zero, IntPtr.Zero)

                                                                    If ColInt1.Count > 0 And ColInt2.Count > 0 Then
                                                                        Point3 = New Point3d((Point1.X + Point2.X) / 2, (ColInt1(0).Y + ColInt2(0).Y) / 2, 0)
                                                                    End If

                                                                    Dim DeltaY2 As Double

                                                                    If IsNumeric(TextBox_deltaY_Vw2.Text) = True Then
                                                                        DeltaY2 = CDbl(TextBox_shiftY_viewport.Text)
                                                                    End If
                                                                    If Not DeltaY2 = 0 Then
                                                                        Point3 = New Point3d(Point3.X, Point3.Y + DeltaY2, 0)
                                                                    End If

                                                                    Dim Point4 As New Point3d(Point3.X - Abs(W1 - W2) * (1000 / Hscale) / 2, Chainage_Y, 0)

                                                                    If Elevation_X1 < Elevation_X2 Then
                                                                        Dim Temp1 As Double = Elevation_X1
                                                                        Elevation_X1 = Elevation_X2
                                                                        Elevation_X2 = Temp1
                                                                    End If

                                                                    Dim Point5 As New Point3d(Elevation_X2, Point3.Y, 0)
                                                                    Dim Point6 As New Point3d(Elevation_X1, Point3.Y, 0)

                                                                    Dim DeltaX1, DeltaX2, DeltaY As Double
                                                                    If IsNumeric(TextBox_deltaY_Vw2.Text) = True Then
                                                                        DeltaY = CDbl(TextBox_deltaY_Vw2.Text)
                                                                    End If
                                                                    If Not DeltaY = 0 Then
                                                                        Point4 = New Point3d(Point4.X, Point4.Y + DeltaY, 0)
                                                                    End If



                                                                    If IsNumeric(TextBox_delta_x1.Text) = True Then
                                                                        DeltaX1 = CDbl(TextBox_delta_x1.Text)
                                                                    End If
                                                                    If IsNumeric(TextBox_delta_x2.Text) = True Then
                                                                        DeltaX2 = CDbl(TextBox_delta_x2.Text)
                                                                    End If


                                                                    If Not DeltaX2 = 0 Then
                                                                        Point5 = New Point3d(Point5.X + DeltaX2, Point5.Y, 0)
                                                                    End If
                                                                    If Not DeltaX1 = 0 Then
                                                                        Point6 = New Point3d(Point6.X + DeltaX1, Point6.Y, 0)
                                                                    End If


                                                                    Dim ExtraL As Double = 0


                                                                    Dim L1 As Double = (Abs(Point1.X - Point2.X) + 2 * ExtraL) * Data_table_cu_valori.Rows(s).Item(VW_CUST_SCALE)

                                                                    Dim Spacing1 As Double = 0


                                                                    Dim Spacing2 As Double = 0


                                                                    x = x - (L1 / 2 + (W2 + W1) / 2 + (Spacing1 + Spacing2) / 2)



                                                                    Dim Viewport1 As New Viewport
                                                                    Viewport1.SetDatabaseDefaults()
                                                                    Viewport1.CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(x + L1 / 2 + (W2 + W1) / 2 + (Spacing1 + Spacing2) / 2, y + H1 + H2 / 2, 0) ' asta e pozitia viewport in paper space
                                                                    Viewport1.Height = H2
                                                                    Viewport1.Width = L1
                                                                    Viewport1.Layer = Layer_name_Main_Viewport
                                                                    'Viewport1.ColorIndex = 1
                                                                    Viewport1.ViewDirection = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis
                                                                    Viewport1.ViewTarget = Point3 ' asta e pozitia viewport in MODEL space
                                                                    Viewport1.ViewCenter = Autodesk.AutoCAD.Geometry.Point2d.Origin
                                                                    Viewport1.TwistAngle = 0 ' asta e PT TWIST

                                                                    BTrecordPS.AppendEntity(Viewport1)
                                                                    Trans_New_doc.AddNewlyCreatedDBObject(Viewport1, True)

                                                                    Viewport1.On = True
                                                                    Viewport1.CustomScale = Data_table_cu_valori.Rows(s).Item(VW_CUST_SCALE)
                                                                    Viewport1.AnnotationScale = AnnoScale1
                                                                    Viewport1.Locked = True

                                                                    Dim Viewport2 As New Viewport
                                                                    Viewport2.SetDatabaseDefaults()
                                                                    Viewport2.CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(x + L1 / 2 + (W2 + W1) / 2 + (Spacing1 + Spacing2) / 2, y + H1 / 2, 0) ' asta e pozitia viewport in paper space
                                                                    Viewport2.Height = H1
                                                                    Viewport2.Width = L1
                                                                    Viewport2.Layer = Layer_name_Main_Viewport
                                                                    'Viewport2.ColorIndex = 2
                                                                    Viewport2.ViewDirection = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis
                                                                    Viewport2.ViewTarget = Point4 ' asta e pozitia viewport in MODEL space
                                                                    Viewport2.ViewCenter = Autodesk.AutoCAD.Geometry.Point2d.Origin
                                                                    Viewport2.TwistAngle = 0 ' asta e PT TWIST

                                                                    BTrecordPS.AppendEntity(Viewport2)
                                                                    Trans_New_doc.AddNewlyCreatedDBObject(Viewport2, True)

                                                                    Viewport2.On = True
                                                                    Viewport2.CustomScale = Data_table_cu_valori.Rows(s).Item(VW_CUST_SCALE)
                                                                    Viewport2.AnnotationScale = AnnoScale1
                                                                    Viewport2.Locked = True

                                                                    Dim Viewport3 As New Viewport
                                                                    Viewport3.SetDatabaseDefaults()
                                                                    Viewport3.CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(x + W2 / 2, y + H1 + H2 / 2, 0) ' asta e pozitia viewport in paper space
                                                                    Viewport3.Height = H2
                                                                    Viewport3.Width = W2
                                                                    Viewport3.Layer = Layer_name_Main_Viewport
                                                                    'Viewport3.ColorIndex = 3
                                                                    Viewport3.ViewDirection = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis
                                                                    Viewport3.ViewTarget = Point5 ' asta e pozitia viewport in MODEL space
                                                                    Viewport3.ViewCenter = Autodesk.AutoCAD.Geometry.Point2d.Origin
                                                                    Viewport3.TwistAngle = 0 ' asta e PT TWIST

                                                                    BTrecordPS.AppendEntity(Viewport3)
                                                                    Trans_New_doc.AddNewlyCreatedDBObject(Viewport3, True)

                                                                    Viewport3.On = True
                                                                    Viewport3.CustomScale = Data_table_cu_valori.Rows(s).Item(VW_CUST_SCALE)
                                                                    Viewport3.AnnotationScale = AnnoScale1
                                                                    Viewport3.Locked = True

                                                                    Dim Viewport4 As New Viewport
                                                                    Viewport4.SetDatabaseDefaults()
                                                                    Viewport4.CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(x + W2 + L1 + W1 / 2 + Spacing1 + Spacing2, y + H1 + H2 / 2, 0) ' asta e pozitia viewport in paper space
                                                                    Viewport4.Height = H2
                                                                    Viewport4.Width = W1
                                                                    Viewport4.Layer = Layer_name_Main_Viewport
                                                                    'Viewport4.ColorIndex = 4
                                                                    Viewport4.ViewDirection = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis
                                                                    Viewport4.ViewTarget = Point6 ' asta e pozitia viewport in MODEL space
                                                                    Viewport4.ViewCenter = Autodesk.AutoCAD.Geometry.Point2d.Origin
                                                                    Viewport4.TwistAngle = 0 ' asta e PT TWIST

                                                                    BTrecordPS.AppendEntity(Viewport4)
                                                                    Trans_New_doc.AddNewlyCreatedDBObject(Viewport4, True)

                                                                    Viewport4.On = True
                                                                    Viewport4.CustomScale = Data_table_cu_valori.Rows(s).Item(VW_CUST_SCALE)
                                                                    Viewport4.AnnotationScale = AnnoScale1
                                                                    Viewport4.Locked = True






                                                                End If
                                                            End If
                                                        End If
                                                    Next
                                                    End If
                                            End If
                                            Trans_New_doc.Commit()
                                        End Using



                                        New_doc.Database.SaveAs(Fisier_nou, True, DwgVersion.Current, BaseMap_drawing.Database.SecurityParameters)
                                    End Using
                                    DocumentExtension.CloseAndDiscard(New_doc)
                                Next
                            End If 'If System.IO.Directory.Exists(TextBox_Output_Directory.Text) = True Then

                        End If
                    End If



                    MsgBox("You are done", , "residential")
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


    Private Sub Button_draw_fence_Click(sender As Object, e As EventArgs) Handles Button_draw_fence.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            If IsNumeric(TextBox_buffer_dist.Text) = False Or _
                    IsNumeric(TextBox_offset.Text) = False Or _
                    IsNumeric(TextBox_max_dist.Text) = False Or _
                    IsNumeric(TextBox_arrow_size.Text) = False Or _
                    IsNumeric(TextBox_text_size.Text) = False Then
                Freeze_operations = False
                MsgBox("Non numerical value")
                Exit Sub
            End If

            Dim Distanta As Double = CDbl(TextBox_buffer_dist.Text)
            Dim Empty_array() As ObjectId
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor


            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Try
                Using lock1 As DocumentLock = ThisDrawing.LockDocument
                    Dim RezultatCL As Autodesk.AutoCAD.EditorInput.PromptEntityResult
                    Dim Object_PromptCL As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select centerline:")
                    RezultatCL = Editor1.GetEntity(Object_PromptCL)
                    If RezultatCL.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Editor1.WriteMessage(vbLf & "Command:")
                        Freeze_operations = False
                        Exit Sub
                    End If

                    Dim PolyCL As Polyline
                    If RezultatCL.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        If IsNothing(RezultatCL) = False Then
                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Dim Ent1 As Entity
                                Ent1 = Trans1.GetObject(RezultatCL.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                    PolyCL = Trans1.GetObject(RezultatCL.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                Else
                                    Editor1.WriteMessage(vbLf & "No Polyline")
                                    Editor1.WriteMessage(vbLf & "Command:")
                                    Freeze_operations = False
                                    Exit Sub
                                End If
                            End Using
                        End If
                    End If

                    Dim Rezultat_str As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_str As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_str.MessageForAdding = vbLf & "Select structures:"

                    Object_Prompt_str.SingleOnly = False
                    Rezultat_str = Editor1.GetSelection(Object_Prompt_str)

                    Dim Rezultat_ws As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt_ws As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_ws.MessageForAdding = vbLf & "Select existing workspace:"

                    Object_Prompt_ws.SingleOnly = False
                    Rezultat_ws = Editor1.GetSelection(Object_Prompt_ws)

                    Dim colectie_linii_si_cercuri As New DBObjectCollection
                    Dim colectie_poly_str As New DBObjectCollection
                    Dim Poly_before_offset As New Polyline
                    Dim Point_collection As New Point3dCollection
                    Dim WS_collection As New DBObjectCollection
                    Dim Continua1 As Boolean = False
                    Dim Continua2 As Boolean = False

                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        If IsNothing(PolyCL) = False And Rezultat_str.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK And Rezultat_ws.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            For i = 0 To Rezultat_str.Value.Count - 1
                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Rezultat_str.Value.Item(i)
                                Dim Ent1 As Entity
                                Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                                If TypeOf Ent1 Is Polyline Then
                                    Dim poly_str As Polyline = Ent1
                                    colectie_poly_str.Add(Ent1)
                                    For j = 0 To poly_str.NumberOfVertices - 1
                                        Dim Pt1 As New Point3d
                                        Pt1 = poly_str.GetPoint3dAt(j)
                                        If Point_collection.Contains(Pt1) = False Then
                                            Point_collection.Add(Pt1)
                                        End If
                                    Next

                                End If
                            Next
                            For i = 0 To Rezultat_ws.Value.Count - 1
                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Rezultat_ws.Value.Item(i)
                                Dim Ent1 As Entity
                                Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                                If TypeOf Ent1 Is Polyline Then
                                    WS_collection.Add(Ent1)
                                End If
                            Next
                            Continua1 = True
                            Trans1.Commit()
                        End If
                    End Using
                    Dim Pmax As New Point3d
                    Dim Pmin As New Point3d
                    If Continua1 = True Then
                        Using Trans2 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            If Point_collection.Count > 1 And WS_collection.Count > 0 Then
                                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans2.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                                Pmax = PolyCL.GetClosestPointTo(Point_collection(0), Vector3d.ZAxis, False)
                                Pmin = PolyCL.GetClosestPointTo(Point_collection(0), Vector3d.ZAxis, False)
                                Dim Chainage_max As Double = PolyCL.GetDistAtPoint(Pmax)
                                Dim Chainage_min As Double = PolyCL.GetDistAtPoint(Pmin)

                                For i = 1 To Point_collection.Count - 1
                                    Dim P1 As New Point3d
                                    P1 = PolyCL.GetClosestPointTo(Point_collection(i), Vector3d.ZAxis, False)
                                    Dim Chainage1 As Double = PolyCL.GetDistAtPoint(P1)
                                    If Chainage1 > Chainage_max Then
                                        Pmax = P1
                                        Chainage_max = Chainage1
                                    End If
                                    If Chainage1 < Chainage_min Then
                                        Pmin = P1
                                        Chainage_min = Chainage1
                                    End If
                                Next
                                If Chainage_max <> Chainage_min Then
                                    Dim Punct_min As New Point3d
                                    Dim Punct_max As New Point3d

                                    If Chainage_min - Distanta >= 0 And Chainage_max + Distanta <= PolyCL.Length Then
                                        Creaza_layer("NO PLOT", 40, "NO PLOT", False)

                                        Punct_min = PolyCL.GetPointAtDist(Chainage_min - Distanta)
                                        Dim Circle1 As New Circle(Punct_min, Vector3d.ZAxis, 10)

                                        Circle1.Layer = "NO PLOT"
                                        Circle1.LineWeight = LineWeight.LineWeight000
                                        Circle1.ColorIndex = 256

                                        BTrecord.AppendEntity(Circle1)
                                        Trans2.AddNewlyCreatedDBObject(Circle1, True)
                                        colectie_linii_si_cercuri.Add(Circle1)

                                        Punct_max = PolyCL.GetPointAtDist(Chainage_max + Distanta)
                                        Dim Circle2 As New Circle(Punct_max, Vector3d.ZAxis, 10)

                                        Circle2.Layer = "NO PLOT"
                                        Circle2.LineWeight = LineWeight.LineWeight000
                                        Circle2.ColorIndex = 256

                                        BTrecord.AppendEntity(Circle2)
                                        Trans2.AddNewlyCreatedDBObject(Circle2, True)
                                        colectie_linii_si_cercuri.Add(Circle2)

                                        For i = 0 To WS_collection.Count - 1
                                            Dim PolyWS As Polyline = WS_collection(i)
                                            Dim P_WS_MAX As New Point3d
                                            P_WS_MAX = PolyWS.GetClosestPointTo(Punct_max, Vector3d.ZAxis, False)
                                            Dim P_WS_MIN As New Point3d
                                            P_WS_MIN = PolyWS.GetClosestPointTo(Punct_min, Vector3d.ZAxis, False)

                                            Dim Linie_max As New Line(Punct_max, P_WS_MAX)
                                            Dim Linie_min As New Line(Punct_min, P_WS_MIN)


                                            Dim UNGHIMAX As Double
                                            Dim UNGHIMIN As Double

                                            Dim Param_max As Double = PolyWS.GetParameterAtPoint(P_WS_MAX)

                                            Dim LinieWS_max As New Line(PolyWS.GetPointAtParameter(Floor(Param_max)), PolyWS.GetPointAtParameter(Ceiling(Param_max)))
                                            UNGHIMAX = Linie_max.StartPoint.GetVectorTo(Linie_max.EndPoint).GetAngleTo(LinieWS_max.StartPoint.GetVectorTo(LinieWS_max.EndPoint))

                                            If Abs(Round(UNGHIMAX * 180 / PI, 0)) = 90 Then
                                                Dim Linie_max1 As New Line
                                                Dim Linie_max2 As New Line


                                                Linie_max.TransformBy(Matrix3d.Scaling(100, Punct_max))
                                                Dim Colint_max As New Point3dCollection
                                                Linie_max.IntersectWith(PolyWS, Intersect.OnBothOperands, Colint_max, IntPtr.Zero, IntPtr.Zero)

                                                If Colint_max.Count > 1 Then
                                                    For j = 0 To Colint_max.Count - 1
                                                        If IsNothing(Linie_max1) = True Then
                                                            Linie_max1 = New Line(Punct_max, Colint_max(j))

                                                        Else
                                                            If Not Linie_max1.Length = 0 Then
                                                                Linie_max2 = New Line(Punct_max, Colint_max(j))

                                                            Else
                                                                Linie_max1 = New Line(Punct_max, Colint_max(j))

                                                            End If


                                                        End If
                                                    Next
                                                    If IsNothing(Linie_max1) = False And IsNothing(Linie_max2) = False Then
                                                        If Linie_max1.Length > Linie_max2.Length Then
                                                            Linie_max1.Layer = "NO PLOT"

                                                            Linie_max1.LineWeight = LineWeight.LineWeight000

                                                            Linie_max1.ColorIndex = 256

                                                            BTrecord.AppendEntity(Linie_max1)
                                                            Trans2.AddNewlyCreatedDBObject(Linie_max1, True)
                                                            colectie_linii_si_cercuri.Add(Linie_max1)
                                                        Else

                                                            Linie_max2.Layer = "NO PLOT"

                                                            Linie_max2.LineWeight = LineWeight.LineWeight000

                                                            Linie_max2.ColorIndex = 256

                                                            BTrecord.AppendEntity(Linie_max2)

                                                            Trans2.AddNewlyCreatedDBObject(Linie_max2, True)
                                                            colectie_linii_si_cercuri.Add(Linie_max2)
                                                        End If

                                                    End If
                                                Else
                                                    Linie_max.TransformBy(Matrix3d.Scaling(0.01, Punct_max))
                                                    Linie_max.Layer = "NO PLOT"

                                                    Linie_max.LineWeight = LineWeight.LineWeight000

                                                    Linie_max.ColorIndex = 256

                                                    BTrecord.AppendEntity(Linie_max)
                                                    Trans2.AddNewlyCreatedDBObject(Linie_max, True)
                                                    colectie_linii_si_cercuri.Add(Linie_max)
                                                End If

                                            End If



                                            Dim Param_min As Double = PolyWS.GetParameterAtPoint(P_WS_MIN)

                                            Dim LinieWS_min As New Line(PolyWS.GetPointAtParameter(Floor(Param_min)), PolyWS.GetPointAtParameter(Ceiling(Param_min)))
                                            UNGHIMIN = Linie_min.StartPoint.GetVectorTo(Linie_min.EndPoint).GetAngleTo(LinieWS_min.StartPoint.GetVectorTo(LinieWS_min.EndPoint))

                                            If Abs(Round(UNGHIMIN * 180 / PI, 0)) = 90 Then
                                                Dim Linie_min1 As New Line
                                                Dim Linie_min2 As New Line


                                                Linie_min.TransformBy(Matrix3d.Scaling(100, Punct_min))
                                                Dim Colint_min As New Point3dCollection
                                                Linie_min.IntersectWith(PolyWS, Intersect.OnBothOperands, Colint_min, IntPtr.Zero, IntPtr.Zero)

                                                If Colint_min.Count > 1 Then
                                                    For j = 0 To Colint_min.Count - 1
                                                        If IsNothing(Linie_min1) = True Then
                                                            Linie_min1 = New Line(Punct_min, Colint_min(j))

                                                        Else
                                                            If Not Linie_min1.Length = 0 Then
                                                                Linie_min2 = New Line(Punct_min, Colint_min(j))

                                                            Else
                                                                Linie_min1 = New Line(Punct_min, Colint_min(j))

                                                            End If


                                                        End If
                                                    Next
                                                    If IsNothing(Linie_min1) = False And IsNothing(Linie_min2) = False Then
                                                        If Linie_min1.Length > Linie_min2.Length Then
                                                            Linie_min1.Layer = "NO PLOT"

                                                            Linie_min1.LineWeight = LineWeight.LineWeight000

                                                            Linie_min1.ColorIndex = 256

                                                            BTrecord.AppendEntity(Linie_min1)
                                                            Trans2.AddNewlyCreatedDBObject(Linie_min1, True)
                                                            colectie_linii_si_cercuri.Add(Linie_min1)
                                                        Else

                                                            Linie_min2.Layer = "NO PLOT"

                                                            Linie_min2.LineWeight = LineWeight.LineWeight000

                                                            Linie_min2.ColorIndex = 256
                                                            BTrecord.AppendEntity(Linie_min2)
                                                            Trans2.AddNewlyCreatedDBObject(Linie_min2, True)
                                                            colectie_linii_si_cercuri.Add(Linie_min2)
                                                        End If

                                                    End If
                                                Else
                                                    Linie_min.TransformBy(Matrix3d.Scaling(0.01, Punct_min))

                                                    Linie_min.Layer = "NO PLOT"

                                                    Linie_min.LineWeight = LineWeight.LineWeight000

                                                    Linie_min.ColorIndex = 256
                                                    BTrecord.AppendEntity(Linie_min)
                                                    Trans2.AddNewlyCreatedDBObject(Linie_min, True)
                                                    colectie_linii_si_cercuri.Add(Linie_min)
                                                End If

                                            End If
                                        Next
                                        Trans2.TransactionManager.QueueForGraphicsFlush()
                                        Trans2.Commit()
                                        Continua2 = True
                                    End If
                                End If
                            End If
                        End Using
                    End If





                    If Continua2 = True Then
                        Using Trans3 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans3.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                            Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                            Dim Colectie_pt As New Point3dCollection
                            Dim Ask_for_point As Boolean = True
                            Dim Counter_for_pt As Integer = 1

                            Dim NEW_OSnap, Old_OSnap As Integer
                            Old_OSnap = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE")

                            NEW_OSnap = Autodesk.AutoCAD.EditorInput.ObjectSnapMasks.End + Autodesk.AutoCAD.EditorInput.ObjectSnapMasks.Intersection

                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap)
                            Do Until Ask_for_point = False
                                Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select " & Counter_for_pt & " point:")
                                PP1.AllowNone = True
                                If Counter_for_pt > 1 Then
                                    PP1.UseBasePoint = True
                                    PP1.BasePoint = Colectie_pt(Counter_for_pt - 2)
                                End If
                                Point1 = Editor1.GetPoint(PP1)
                                If Not Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                    Ask_for_point = False
                                Else
                                    Colectie_pt.Add(Point1.Value)
                                    Counter_for_pt = Counter_for_pt + 1
                                End If
                            Loop
                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", Old_OSnap)

                            If Colectie_pt.Count > 1 Then
                                For i = 0 To Colectie_pt.Count - 1
                                    Poly_before_offset.AddVertexAt(i, New Point2d(Colectie_pt(i).X, Colectie_pt(i).Y), 0, 2, 2)
                                Next
                                BTrecord.AppendEntity(Poly_before_offset)
                                Trans3.AddNewlyCreatedDBObject(Poly_before_offset, True)

                                Dim punct_pt_dir_offset As New Point3d((Pmax.X + Pmin.X) / 2, (Pmax.Y + Pmin.Y) / 2, (Pmax.Z + Pmin.Z) / 2)
                                Dim Directie_pt_offset As Double = Directie_offset(Poly_before_offset, punct_pt_dir_offset)

                                Dim Object_colection1 As Autodesk.AutoCAD.DatabaseServices.DBObjectCollection = Poly_before_offset.GetOffsetCurves(CDbl(TextBox_offset.Text) * Directie_pt_offset)

                                Dim Poly_after_offset As New Polyline

                                Poly_after_offset = Object_colection1(0)
                                Dim Este_in_afara As Boolean = False

                                For i = 0 To colectie_poly_str.Count - 1
                                    Dim polystr As Polyline = Trans3.GetObject(colectie_poly_str(i).ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                                    For j = 0 To polystr.NumberOfVertices - 1
                                        Dim pTTT1 As New Point3d
                                        pTTT1 = polystr.GetPoint3dAt(j)
                                        Dim Clossest_pt As New Point3d
                                        Clossest_pt = Poly_before_offset.GetClosestPointTo(pTTT1, Vector3d.ZAxis, False)
                                        Dim Len1 As Double = pTTT1.GetVectorTo(Clossest_pt).Length
                                        If Len1 > CDbl(TextBox_max_dist.Text) Then
                                            Este_in_afara = True
                                        Else
                                            Este_in_afara = False
                                            Exit For
                                        End If
                                    Next
                                    If Este_in_afara = True Then
                                        MsgBox("Structure is outside of " & TextBox_max_dist.Text)
                                    End If
                                Next





                                If Este_in_afara = False Then
                                    BTrecord.AppendEntity(Poly_after_offset)
                                    Trans3.AddNewlyCreatedDBObject(Poly_after_offset, True)
                                    Trans3.TransactionManager.QueueForGraphicsFlush()


                                    Dim P1 As New Point3d
                                    P1 = PolyCL.GetClosestPointTo(Poly_before_offset.StartPoint, Vector3d.ZAxis, False)
                                    Dim P2 As New Point3d
                                    P2 = PolyCL.GetClosestPointTo(Poly_before_offset.EndPoint, Vector3d.ZAxis, False)
                                    Dim TextS As Double = CDbl(TextBox_text_size.Text)
                                    Dim Arrows As Double = CDbl(TextBox_arrow_size.Text)


                                    If CheckBox_sta.Checked = True Then
                                        Dim Chainage1 As String = Get_chainage_feet_from_double(PolyCL.GetDistAtPoint(P1), 0)
                                        Dim Mleader1 As New MLeader
                                        Mleader1 = Creaza_Mleader_nou_fara_UCS_transform(Poly_after_offset.StartPoint, "STA = " & Chainage1, TextS, Arrows, Arrows, 2 * TextS, 3 * TextS)
                                        Mleader1.Linetype = "CONTINUOUS"
                                        Mleader1.LineWeight = LineWeight.LineWeight000

                                        Dim Mleader2 As New MLeader
                                        Dim Chainage2 As String = Get_chainage_feet_from_double(PolyCL.GetDistAtPoint(P2), 0)
                                        Mleader2 = Creaza_Mleader_nou_fara_UCS_transform(Poly_after_offset.EndPoint, "STA = " & Chainage2, TextS, Arrows, Arrows, 2 * TextS, 3 * TextS)
                                        Mleader2.Linetype = "CONTINUOUS"
                                        Mleader2.LineWeight = LineWeight.LineWeight000
                                    End If

                                    If CheckBox_MP.Checked = True Then
                                        Dim Chainage1 As Double = PolyCL.GetDistAtPoint(P1)
                                        Dim Mleader1 As New MLeader
                                        Mleader1 = Creaza_Mleader_nou_fara_UCS_transform(Poly_after_offset.StartPoint, "MP = " & Get_String_Rounded(Chainage1 / 5280, 2), TextS, Arrows, Arrows, 2 * TextS, 5 * TextS)
                                        Mleader1.Linetype = "CONTINUOUS"
                                        Mleader1.LineWeight = LineWeight.LineWeight000
                                        Dim Mleader2 As New MLeader
                                        Dim Chainage2 As Double = PolyCL.GetDistAtPoint(P2)
                                        Mleader2 = Creaza_Mleader_nou_fara_UCS_transform(Poly_after_offset.EndPoint, "MP = " & Get_String_Rounded(Chainage2 / 5280, 2), TextS, Arrows, Arrows, 2 * TextS, 5 * TextS)
                                        Mleader2.Linetype = "CONTINUOUS"
                                        Mleader2.LineWeight = LineWeight.LineWeight000
                                    End If

                                    Dim Acmap As Autodesk.Gis.Map.Platform.AcMapMap = Autodesk.Gis.Map.Platform.AcMapMap.GetCurrentMap
                                    Dim Curent_system As String = Acmap.GetMapSRS()
                                    If CheckBox_lat_long.Checked = True Then
                                        If String.IsNullOrEmpty(Curent_system) = True Then
                                            MsgBox("Please set your coordinate system")
                                            CheckBox_lat_long.Checked = False
                                        End If
                                    End If


                                    If CheckBox_lat_long.Checked = True Then
                                        Dim String_LL84 As String = "GEOGCS[" & Chr(34) & "LL84" & Chr(34) & ",DATUM[" & Chr(34) & "WGS84" & Chr(34) & ",SPHEROID[" & Chr(34) & "WGS84" & Chr(34) & ",6378137.000,298.25722293]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.01745329251994]]"

                                        Dim Coord_factory1 As New OSGeo.MapGuide.MgCoordinateSystemFactory
                                        Dim CoordSys1 As OSGeo.MapGuide.MgCoordinateSystem = Coord_factory1.Create(Curent_system)
                                        Dim CoordSys2 As OSGeo.MapGuide.MgCoordinateSystem = Coord_factory1.Create(String_LL84)
                                        Dim Transform1 As OSGeo.MapGuide.MgCoordinateSystemTransform = Coord_factory1.GetTransform(CoordSys1, CoordSys2)



                                        Dim x1 As Double = Poly_after_offset.StartPoint.X
                                        Dim y1 As Double = Poly_after_offset.StartPoint.Y
                                        Dim Coord1 As OSGeo.MapGuide.MgCoordinate = Transform1.Transform(x1, y1)
                                        Dim Lat1 As Double = Coord1.Y
                                        Dim Long1 As Double = Coord1.X
                                        Dim Mleader1 As New MLeader
                                        Mleader1 = Creaza_Mleader_nou_fara_UCS_transform(Poly_after_offset.StartPoint, "Lon = " & Round(Long1, 4) & vbCrLf & "Lat = " & Get_String_Rounded(Lat1, 4), TextS, Arrows, Arrows, 4 * TextS, TextS)
                                        Mleader1.Linetype = "CONTINUOUS"
                                        Mleader1.LineWeight = LineWeight.LineWeight000

                                        Dim x2 As Double = Poly_after_offset.EndPoint.X
                                        Dim y2 As Double = Poly_after_offset.EndPoint.Y
                                        Dim Coord2 As OSGeo.MapGuide.MgCoordinate = Transform1.Transform(x2, y2)
                                        Dim Lat2 As Double = Coord2.Y
                                        Dim Long2 As Double = Coord2.X
                                        Dim Mleader2 As New MLeader
                                        Mleader2 = Creaza_Mleader_nou_fara_UCS_transform(Poly_after_offset.EndPoint, "Lon = " & Round(Long2, 4) & vbCrLf & "Lat = " & Get_String_Rounded(Lat2, 4), TextS, Arrows, Arrows, 4 * TextS, TextS)
                                        Mleader2.Linetype = "CONTINUOUS"
                                        Mleader2.LineWeight = LineWeight.LineWeight000

                                    End If



                                End If


                            End If

                            If colectie_linii_si_cercuri.Count > 0 Then
                                For i = 0 To colectie_linii_si_cercuri.Count - 1
                                    Dim Dbobj1 As DBObject = Trans3.GetObject(colectie_linii_si_cercuri(i).ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                                    Dbobj1.Erase()
                                Next
                            End If
                            If IsNothing(Poly_before_offset) = False Then
                                Poly_before_offset.Erase()
                            End If


                            Trans3.Commit()

                        End Using
                    End If

                End Using

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



    Private Sub Button_pick_corner_Click(sender As Object, e As EventArgs) Handles Button_pick_corner.Click
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
                        Dim Prompt_pt As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify THE VIEWPORT lower left point")

                        Prompt_pt.AllowNone = True
                        Pt_rezult = Editor1.GetPoint(Prompt_pt)

                        If Pt_rezult.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            TextBox_x.Text = Get_String_Rounded(Pt_rezult.Value.X, 2)
                            TextBox_y.Text = Get_String_Rounded(Pt_rezult.Value.Y, 2)
                        End If





                    End Using
                End Using

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
    End Sub

    Private Sub Button_load_graph_Click(sender As Object, e As EventArgs) Handles Button_load_graph.Click
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


                        Dim Rezultat_hor As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                        Dim Object_prompt_hor As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_prompt_hor.MessageForAdding = vbLf & "Select a known vertical line (STATION) and the label for it:"

                        Object_prompt_hor.SingleOnly = False
                        Rezultat_hor = Editor1.GetSelection(Object_prompt_hor)


                        Dim Rezultat_vert As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                        Dim Object_prompt_vert As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_prompt_vert.MessageForAdding = vbLf & "Select a known horizontal line (ELEVATION) and the label for it:"

                        Object_prompt_vert.SingleOnly = False
                        Rezultat_vert = Editor1.GetSelection(Object_prompt_vert)

                        Dim Rezultat_vert2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                        Dim Object_prompt_vert2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_prompt_vert2.MessageForAdding = vbLf & "Select an elevation label from the other side:"
                        Object_prompt_vert2.SingleOnly = True
                        Rezultat_vert2 = Editor1.GetSelection(Object_prompt_vert2)

                        Dim x01, y01, x02, y02 As Double
                        Dim x03, y03, x04, y04 As Double

                        If Rezultat_hor.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK And Rezultat_vert.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK And Rezultat_vert2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Dim Empty_array() As ObjectId


                            Dim mText_cunoscut As Autodesk.AutoCAD.DatabaseServices.MText
                            Dim Text_cunoscut As Autodesk.AutoCAD.DatabaseServices.DBText
                            Dim Linia_cunoscuta As Autodesk.AutoCAD.DatabaseServices.Line
                            Dim PolyLinia_cunoscuta As Autodesk.AutoCAD.DatabaseServices.Polyline



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
                            Obj4 = Rezultat_hor.Value.Item(1)
                            Ent4 = Obj4.ObjectId.GetObject(OpenMode.ForRead)

                            Dim Obj5 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj5 = Rezultat_vert2.Value.Item(0)
                            Dim Ent5 As Entity
                            Ent5 = Obj5.ObjectId.GetObject(OpenMode.ForRead)



                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                                mText_cunoscut = Ent1
                                Dim String_Mtext As String = Replace(mText_cunoscut.Text, "'", "")

                                If IsNumeric(String_Mtext) = True Then
                                    Elevatia_cunoscuta = CDbl(String_Mtext)
                                    Elevation_X2 = mText_cunoscut.Location.X
                                End If

                            End If

                            If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                                mText_cunoscut = Ent2
                                Dim String_Mtext As String = Replace(mText_cunoscut.Text, "'", "")
                                If IsNumeric(String_Mtext) = True Then
                                    Elevatia_cunoscuta = CDbl(String_Mtext)
                                    Elevation_X2 = mText_cunoscut.Location.X
                                End If

                            End If

                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                                Text_cunoscut = Ent1
                                Dim String_Text As String = Replace(Text_cunoscut.TextString, "'", "")
                                If IsNumeric(String_Text) = True Then
                                    Elevatia_cunoscuta = CDbl(String_Text)
                                    Elevation_X2 = Text_cunoscut.Position.X
                                End If

                            End If

                            If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                                Text_cunoscut = Ent2
                                Dim String_Text As String = Replace(Text_cunoscut.TextString, "'", "")
                                If IsNumeric(String_Text) = True Then
                                    Elevatia_cunoscuta = CDbl(String_Text)
                                    Elevation_X2 = Text_cunoscut.Position.X
                                End If

                            End If

                            If TypeOf Ent5 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                                mText_cunoscut = Ent5
                                Dim String_Mtext As String = Replace(mText_cunoscut.Text, "'", "")

                                If IsNumeric(String_Mtext) = True Then
                                    Elevation_X1 = mText_cunoscut.Location.X
                                End If

                            End If

                            If TypeOf Ent5 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                                Text_cunoscut = Ent5
                                Dim String_Text As String = Replace(Text_cunoscut.TextString, "'", "")
                                If IsNumeric(String_Text) = True Then
                                    Elevation_X1 = Text_cunoscut.Position.X
                                End If

                            End If
                            If Elevatia_cunoscuta = -100000 Then
                                Freeze_operations = False
                                Editor1.SetImpliedSelection(Empty_array)
                                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                Exit Sub
                            End If



                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                                Linia_cunoscuta = Ent1
                                x01 = Linia_cunoscuta.StartPoint.X
                                y01 = Linia_cunoscuta.StartPoint.Y
                                x02 = Linia_cunoscuta.EndPoint.X
                                y02 = Linia_cunoscuta.EndPoint.Y
                                If Abs(y01 - y02) > 0.001 Then
                                    Freeze_operations = False
                                    Editor1.SetImpliedSelection(Empty_array)
                                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                    Exit Sub
                                End If
                            End If


                            If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                                Linia_cunoscuta = Ent2
                                x01 = Linia_cunoscuta.StartPoint.X
                                y01 = Linia_cunoscuta.StartPoint.Y
                                x02 = Linia_cunoscuta.EndPoint.X
                                y02 = Linia_cunoscuta.EndPoint.Y
                                If Abs(y01 - y02) > 0.001 Then
                                    Freeze_operations = False
                                    Editor1.SetImpliedSelection(Empty_array)
                                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                    Exit Sub
                                End If
                            End If

                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                PolyLinia_cunoscuta = Ent1

                                x01 = PolyLinia_cunoscuta.StartPoint.X
                                y01 = PolyLinia_cunoscuta.StartPoint.Y
                                x02 = PolyLinia_cunoscuta.EndPoint.X
                                y02 = PolyLinia_cunoscuta.EndPoint.Y
                                If Abs(y01 - y02) > 0.001 Then
                                    Freeze_operations = False
                                    Editor1.SetImpliedSelection(Empty_array)
                                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                    Exit Sub
                                End If
                            End If
                            If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                PolyLinia_cunoscuta = Ent2

                                x01 = PolyLinia_cunoscuta.StartPoint.X
                                y01 = PolyLinia_cunoscuta.StartPoint.Y
                                x02 = PolyLinia_cunoscuta.EndPoint.X
                                y02 = PolyLinia_cunoscuta.EndPoint.Y
                                If Abs(y01 - y02) > 0.001 Then
                                    Freeze_operations = False
                                    Editor1.SetImpliedSelection(Empty_array)
                                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                    Exit Sub
                                End If

                            End If




                            If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                                mText_cunoscut = Ent3
                                Dim numar_fara_plus As String = Replace(mText_cunoscut.Text, "+", "")
                                If IsNumeric(numar_fara_plus) = True Then
                                    Chainage_cunoscuta = CDbl(numar_fara_plus)
                                    Chainage_Y = mText_cunoscut.Location.Y
                                End If

                            End If

                            If TypeOf Ent4 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                                mText_cunoscut = Ent4
                                Dim numar_fara_plus As String = Replace(mText_cunoscut.Text, "+", "")
                                If IsNumeric(numar_fara_plus) = True Then
                                    Chainage_cunoscuta = CDbl(numar_fara_plus)
                                    Chainage_Y = mText_cunoscut.Location.Y
                                End If

                            End If

                            If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                                Text_cunoscut = Ent3
                                Dim numar_fara_plus As String = Replace(Text_cunoscut.TextString, "+", "")
                                If IsNumeric(numar_fara_plus) = True Then
                                    Chainage_cunoscuta = CDbl(numar_fara_plus)
                                    Chainage_Y = Text_cunoscut.Position.Y
                                End If

                            End If

                            If TypeOf Ent4 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                                Text_cunoscut = Ent4
                                Dim numar_fara_plus As String = Replace(Text_cunoscut.TextString, "+", "")
                                If IsNumeric(numar_fara_plus) = True Then
                                    Chainage_cunoscuta = CDbl(numar_fara_plus)
                                    Chainage_Y = Text_cunoscut.Position.Y
                                End If

                            End If



                            If Chainage_cunoscuta = -100000 Then
                                Freeze_operations = False
                                Editor1.SetImpliedSelection(Empty_array)
                                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                Exit Sub
                            End If



                            If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                                Linia_cunoscuta = Ent3
                                x03 = Linia_cunoscuta.StartPoint.X
                                y03 = Linia_cunoscuta.StartPoint.Y
                                x04 = Linia_cunoscuta.EndPoint.X
                                y04 = Linia_cunoscuta.EndPoint.Y
                                If Abs(x03 - x04) > 0.001 Then
                                    Freeze_operations = False
                                    Editor1.SetImpliedSelection(Empty_array)
                                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                    Exit Sub
                                End If
                            End If


                            If TypeOf Ent4 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                                Linia_cunoscuta = Ent4
                                x03 = Linia_cunoscuta.StartPoint.X
                                y03 = Linia_cunoscuta.StartPoint.Y
                                x04 = Linia_cunoscuta.EndPoint.X
                                y04 = Linia_cunoscuta.EndPoint.Y
                                If Abs(x03 - x04) > 0.001 Then
                                    Freeze_operations = False
                                    Editor1.SetImpliedSelection(Empty_array)
                                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                    Exit Sub
                                End If

                            End If

                            If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                PolyLinia_cunoscuta = Ent3
                                x03 = PolyLinia_cunoscuta.StartPoint.X
                                y03 = PolyLinia_cunoscuta.StartPoint.Y
                                x04 = PolyLinia_cunoscuta.EndPoint.X
                                y04 = PolyLinia_cunoscuta.EndPoint.Y
                                If Abs(x03 - x04) > 0.001 Then
                                    Freeze_operations = False
                                    Editor1.SetImpliedSelection(Empty_array)
                                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                    Exit Sub
                                End If
                            End If

                            If TypeOf Ent4 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                PolyLinia_cunoscuta = Ent4
                                x03 = PolyLinia_cunoscuta.StartPoint.X
                                y03 = PolyLinia_cunoscuta.StartPoint.Y
                                x04 = PolyLinia_cunoscuta.EndPoint.X
                                y04 = PolyLinia_cunoscuta.EndPoint.Y
                                If Abs(x03 - x04) > 0.001 Then
                                    Freeze_operations = False
                                    Editor1.SetImpliedSelection(Empty_array)
                                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                    Exit Sub
                                End If
                            End If
                        End If

                        Dim Linie1 As New Line(New Point3d(x01, y01, 0), New Point3d(x02, y02, 0))
                        Dim Linie2 As New Line(New Point3d(x03, y03, 0), New Point3d(x04, y04, 0))

                        If Linie1.Length > 0.01 And Linie2.Length > 0.01 Then
                            Dim Colint1 As New Point3dCollection
                            Linie1.IntersectWith(Linie2, Intersect.ExtendBoth, Colint1, IntPtr.Zero, IntPtr.Zero)
                            If Colint1.Count > 0 Then
                                Dim Rezultat_Poly As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                                Dim Object_prompt_Poly As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                                Object_prompt_Poly.MessageForAdding = vbLf & "Select the ground Polyline"

                                Object_prompt_Poly.SingleOnly = True
                                Rezultat_Poly = Editor1.GetSelection(Object_prompt_Poly)
                                If Rezultat_Poly.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                    If TypeOf Rezultat_Poly.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead) Is Polyline Then
                                        Dim Poly1 As Polyline
                                        Poly1 = Rezultat_Poly.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)

                                        Data_table_poly = New System.Data.DataTable
                                        Data_table_poly.Columns.Add("X", GetType(Double))
                                        Data_table_poly.Columns.Add("Y", GetType(Double))

                                        For i = 0 To Poly1.NumberOfVertices - 1
                                            Data_table_poly.Rows.Add()
                                            Data_table_poly.Rows(i).Item("X") = Poly1.GetPoint3dAt(i).X
                                            Data_table_poly.Rows(i).Item("Y") = Poly1.GetPoint3dAt(i).Y
                                        Next

                                        Point_cunoscut = Colint1(0)
                                        Freeze_operations = False
                                    Else
                                        Freeze_operations = False
                                    End If
                                Else
                                    Freeze_operations = False
                                End If


                            Else
                                Freeze_operations = False
                            End If
                        Else
                            Freeze_operations = False
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

    Private Sub RadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton_generic.CheckedChanged, RadioButton_Road_permit.CheckedChanged, RadioButton_usage_permit.CheckedChanged
        If RadioButton_generic.Checked = True Then
            TextBox_north_arrow_Big_X.Text = "1.0598"
            TextBox_north_arrow_Big_y.Text = "10.1504"
            TextBox_main_viewport_height.Text = "5.4120"
            TextBox_main_viewport_width.Text = "15.9861"
            TextBox_main_viewport_center_X.Text = "8.5228"
            TextBox_main_viewport_center_Y.Text = "7.9788"
            TextBox_north_arrow.Text = "Spectra North Arrow"
            TextBox_blockScale.Text = "0.3562"
            TextBox_H_SCALE.Text = "1"
            TextBox_V_SCALE.Text = "1"
            TextBox_Layout_index_main.Text = "1"
            TextBox_W2.Text = "0.25"
            TextBox_W3.Text = "0.25"
            TextBox_delta_x2.Text = "0"
            TextBox_shiftY_viewport.Text = "0"
            TextBox_delta_x2.Text = "1"
            TextBox_H2.Text = "1.3"
            TextBox_deltaY_Vw2.Text = "0"
            TextBox_H1.Text = "0.27"
            TextBox_Layout_index_profile.Text = "1"

            TextBox_x.Text = "6.55"
            TextBox_y.Text = "1.33"

        ElseIf RadioButton_Road_permit.Checked = True Then
            TextBox_north_arrow_Big_X.Text = "14.9"
            TextBox_north_arrow_Big_y.Text = "9.07"
            TextBox_main_viewport_height.Text = "4.40214"
            TextBox_main_viewport_width.Text = "11.7"
            TextBox_main_viewport_center_X.Text = "6.91057"
            TextBox_main_viewport_center_Y.Text = "8.48370"
            TextBox_north_arrow.Text = "Spectra North Arrow"
            TextBox_blockScale.Text = "0.5"
            TextBox_H_SCALE.Text = "1"
            TextBox_V_SCALE.Text = "0.5"
            TextBox_Layout_index_main.Text = "1"
            TextBox_W2.Text = "0.25"
            TextBox_W3.Text = "0.25"
            TextBox_delta_x2.Text = "0"
            TextBox_shiftY_viewport.Text = "0"
            TextBox_delta_x2.Text = "1"
            TextBox_H2.Text = "3.77"
            TextBox_deltaY_Vw2.Text = "0"
            TextBox_H1.Text = "0.27"
            TextBox_Layout_index_profile.Text = "1"
            TextBox_dwt_template.Text = "C:\AutoCAD\Templates\Civil\SPECTRA\SPECTRA ROAD PERMIT.dwt"
            TextBox_x.Text = "6.91"
            TextBox_y.Text = "1.15"
            RadioButton20.Checked = True

        ElseIf RadioButton_usage_permit.Checked = True Then
            TextBox_north_arrow_Big_X.Text = "44.34"
            TextBox_north_arrow_Big_y.Text = "9.79"
            TextBox_main_viewport_height.Text = "4.9289"
            TextBox_main_viewport_width.Text = "8.5"
            TextBox_main_viewport_center_X.Text = "48.6026"
            TextBox_main_viewport_center_Y.Text = "8.0189"
            TextBox_north_arrow.Text = "Spectra North Arrow"
            TextBox_blockScale.Text = "0.5"
            TextBox_H_SCALE.Text = "1"
            TextBox_V_SCALE.Text = "1"
            TextBox_Layout_index_main.Text = "1"
            TextBox_W2.Text = "0.5"
            TextBox_W3.Text = "0.5"
            TextBox_delta_x2.Text = "0"
            TextBox_shiftY_viewport.Text = "0"
            TextBox_delta_x2.Text = "0"
            TextBox_H2.Text = "5.0162"
            TextBox_deltaY_Vw2.Text = "0"
            TextBox_H1.Text = "0.3"
            TextBox_Layout_index_profile.Text = "2"
            TextBox_dwt_template.Text = "C:\AutoCAD\Templates\Civil\SPECTRA\SPECTRA USACE PERMIT.dwt"


            TextBox_x.Text = "48.6242"
            TextBox_y.Text = "4.9510"
            RadioButton50.Checked = True
        End If
    End Sub

End Class


