
Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry


Public Class Profiler_convertor_form

    Private Sub Profiler_form_Load(sender As Object, e As EventArgs) Handles Me.Load


    End Sub

    Public Overridable Function InsertBlock_PROFILE_TAG(ByVal dblInsert As Point3d, ByVal Spatiu As BlockTableRecord, ByVal Layer1 As String, ByVal Chainage As String, ByVal Nume As String) As BlockReference
        Dim dlock As DocumentLock = Nothing
        Dim BlockTable1 As BlockTable
        Dim Block_table_record1 As BlockTableRecord = Nothing
        Dim br As BlockReference
        Dim id As ObjectId
        Dim db As Autodesk.AutoCAD.DatabaseServices.Database = HostApplicationServices.WorkingDatabase
        Using trans As Transaction = db.TransactionManager.StartTransaction
            Dim ed As Autodesk.AutoCAD.EditorInput.Editor = Application.DocumentManager.MdiActiveDocument.Editor

            'insert block and rename it
            Try
                Try
                    dlock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                Catch ex As Exception
                    Dim aex As New System.Exception("Error locking document for InsertBlock: " & "G_1_LINE_PROFILE_TAG" & ": ", ex)
                    Throw aex
                End Try
                BlockTable1 = trans.GetObject(db.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)
12345:
                If BlockTable1.Has("G_1_LINE_PROFILE_TAG") = True Then
                    Block_table_record1 = trans.GetObject(BlockTable1.Item("G_1_LINE_PROFILE_TAG"), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                    Spatiu = trans.GetObject(Spatiu.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)
                    'Set the Attribute Value
                    Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection
                    Dim ent As Entity
                    Dim Block_table_record1enum As BlockTableRecordEnumerator
                    br = New BlockReference(dblInsert, Block_table_record1.ObjectId)
                    br.Layer = Layer1
                    br.ScaleFactors = New Autodesk.AutoCAD.Geometry.Scale3d(1, 1, 1)

                    Spatiu.AppendEntity(br)
                    trans.AddNewlyCreatedDBObject(br, True)
                    attColl = br.AttributeCollection
                    Block_table_record1enum = Block_table_record1.GetEnumerator
                    While Block_table_record1enum.MoveNext
                        ent = Block_table_record1enum.Current.GetObject(OpenMode.ForWrite)
                        If TypeOf ent Is AttributeDefinition Then
                            Dim attdef As AttributeDefinition = ent
                            Dim attref As New AttributeReference
                            attref.SetAttributeFromBlock(attdef, br.BlockTransform)
                            'attref.TextString = attref.Tag

                            If attref.Tag = "CHAINAGE" Then
                                attref.TextString = Chainage
                            End If
                            If attref.Tag = "FIRST_LINE" Then
                                attref.TextString = Nume
                            End If
                            If attref.Tag = "SECOND_LINE" Then
                                attref.TextString = Nume
                            End If

                            attColl.AppendAttribute(attref)
                            trans.AddNewlyCreatedDBObject(attref, True)


                        End If
                    End While
                    trans.Commit()
                Else
                    Try
                        Dim Locatie_blocuri As String = Locatie.Locatie_blocuri
                        Dim Locatie_blocuri_alternativ As String = Locatie.Locatie_blocuri_my_docs
                        Dim Nume_fisier As String = "1_LINE_PROFILE_TAG.dwg"
                        Dim Locatie1 As String = Locatie_fisier(Locatie_blocuri_alternativ & Nume_fisier, Locatie_blocuri & Nume_fisier)

                        If Locatie1 = Locatie_blocuri & Nume_fisier Then
                            If Microsoft.VisualBasic.FileIO.FileSystem.DirectoryExists(Locatie.Locatie_blocuri_my_docs_) = False Then
                                Microsoft.VisualBasic.FileIO.FileSystem.CreateDirectory(Locatie.Locatie_blocuri_my_docs_)
                            End If
                            If Microsoft.VisualBasic.FileIO.FileSystem.FileExists(Locatie_blocuri_alternativ & Nume_fisier) = False Then
                                Microsoft.VisualBasic.FileIO.FileSystem.CopyFile(Locatie_blocuri & Nume_fisier, Locatie_blocuri_alternativ & Nume_fisier)
                            End If
                            Locatie1 = Locatie_blocuri_alternativ & Nume_fisier
                        End If



                        Dim ThisDrawing As Document = Application.DocumentManager.MdiActiveDocument


                        Dim Nume_block As String = "G_1_LINE_PROFILE_TAG"

                        Using docLock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument
                            Using Database2 As New Database(False, False)
                                'read block drawing do we need Lockdocument..?? Not Always..?
                                Database2.ReadDwgFile(Locatie1, System.IO.FileShare.Read, True, Nothing)
                                Using Trans1 As Transaction = ThisDrawing.TransactionManager.StartTransaction()

                                    Dim BlockTable2 As BlockTable = DirectCast(Trans1.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead, False), BlockTable)

                                    Dim idBTR As ObjectId = ThisDrawing.Database.Insert(Nume_block, Database2, False)
                                    Trans1.Commit()

                                End Using
                            End Using
                        End Using
                    Catch e As System.Exception

                        MsgBox(e.Message)
                    End Try
                    GoTo 12345
                End If ' ASTA E DE LA  If bt.Has("G_1_LINE_PROFILE_TAG")

            Catch ex As System.Exception
                Dim aex2 As New System.Exception("Error in inserting new block: " & "G_1_LINE_PROFILE_TAG" & ": ", ex)
                Throw aex2
            Finally
                If Not trans Is Nothing Then trans.Dispose()
                If Not dlock Is Nothing Then dlock.Dispose()
            End Try
        End Using
        Return br
    End Function




    Private Sub Create_profile_graf(ByVal Trans1 As Transaction, ByVal BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord, ByVal BTrecord_paper As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord, _
                                        ByVal Layer_grid_lines As String, ByVal color_index_grid_lines As Integer, _
                                        ByVal Layer_text As String, ByVal Text_style_id As Autodesk.AutoCAD.DatabaseServices.ObjectId, ByVal layer_polyline As String, _
                                        ByVal Layer_no_plot As String, ByVal Layer_Viewport As String, ByVal Layer_blocks As String, _
                                        ByVal punct_jos_stanga As Point3d, ByVal Lungime_graph_1_1 As Double, _
                                        ByVal Scale_inv_vwport As Double, ByVal interval_horizontal As Double, ByVal interval_vertical As Double, _
                                        ByVal Nr_chainage_labels As Integer, ByVal Nr_vertical_labels As Integer, _
                                        ByVal Lung_thick_vert As Double, ByVal Lungime_linie_mica_horiz As Double, _
                                        ByVal Elevatia_cunoscuta As Double, ByVal Elevatia_pt_pct_zero_masurata As Double, ByVal Xpoly As Double, ByVal Ypoly As Double, ByVal Yreference As Double, _
                                        ByVal V_scale As Double, ByVal H_scale As Double, ByVal Pr_scale As Double, ByVal nr_linii_jos As Integer, ByVal Colectie_puncte_graph As Point2dCollection, _
                                        ByVal V_scale_hmm As Double, ByVal H_scale_hmm As Double, ByVal Pr_scale_hmm As Double, ByVal Valoare_Chainage As DoubleCollection, ByVal viewport_ps_center As Point3d, _
                                        ByVal viewport_height As Double, ByVal viewport_width As Double, ByVal Descriptie_Chainage As Specialized.StringCollection, ByVal String_Chainage As Specialized.StringCollection, _
                                         ByVal Gap_PS_sus As Double, ByVal Gap_PS_jos As Double, ByVal Gap_PS_stanga_dreapta As Double)
        Try
            Dim X_grid As Double = punct_jos_stanga.X
            Dim Y_grid As Double = punct_jos_stanga.Y
            Dim Viewport_scale As Double = 1 / Scale_inv_vwport

            Dim Inaltime_graph As Double = Nr_vertical_labels * interval_vertical
            Dim Scale_Modelspace As Double = Pr_scale_hmm / H_scale_hmm
            Dim Vert_exag As Double = H_scale_hmm / V_scale_hmm


            Dim Linia_vert_stanga As New Autodesk.AutoCAD.DatabaseServices.Line
            With Linia_vert_stanga
                .StartPoint = New Point3d(X_grid, Y_grid, 0)
                .EndPoint = New Point3d(X_grid, Y_grid + Inaltime_graph * Scale_inv_vwport * Scale_Modelspace * Vert_exag, 0)
                .Layer = Layer_grid_lines
                .Linetype = "CONTINUOUS"
                .ColorIndex = 2
            End With

            BTrecord.AppendEntity(Linia_vert_stanga)
            Trans1.AddNewlyCreatedDBObject(Linia_vert_stanga, True)

            Dim Linia_jos_horiz As New Autodesk.AutoCAD.DatabaseServices.Line
            With Linia_jos_horiz
                .StartPoint = New Point3d(X_grid, Y_grid, 0)
                .EndPoint = New Point3d(X_grid + Lungime_graph_1_1 * Scale_inv_vwport * Scale_Modelspace, Y_grid, 0)
                .Layer = Layer_grid_lines
                .Linetype = "CONTINUOUS"
                .ColorIndex = 2
            End With

            BTrecord.AppendEntity(Linia_jos_horiz)
            Trans1.AddNewlyCreatedDBObject(Linia_jos_horiz, True)


            Dim Linia_vert_dreapta As New Autodesk.AutoCAD.DatabaseServices.Line
            With Linia_vert_dreapta
                .StartPoint = New Point3d(X_grid + Lungime_graph_1_1 * Scale_inv_vwport * Scale_Modelspace, Y_grid, 0)
                .EndPoint = New Point3d(X_grid + Lungime_graph_1_1 * Scale_inv_vwport * Scale_Modelspace, _
                                        Y_grid + Inaltime_graph * Scale_inv_vwport * Scale_Modelspace * Vert_exag, 0)
                .Layer = Layer_grid_lines
                .Linetype = "CONTINUOUS"
                .ColorIndex = 2
            End With

            BTrecord.AppendEntity(Linia_vert_dreapta)
            Trans1.AddNewlyCreatedDBObject(Linia_vert_dreapta, True)

            'aici desenez the grey lines hidden
            For i = 1 To Nr_vertical_labels
                Dim Linia_hidden As New Autodesk.AutoCAD.DatabaseServices.Line
                With Linia_hidden
                    .StartPoint = New Point3d(X_grid, Y_grid + i * interval_vertical * Scale_inv_vwport * Scale_Modelspace * Vert_exag, 0)
                    .EndPoint = New Point3d(X_grid + Lungime_graph_1_1 * Scale_inv_vwport * Scale_Modelspace, _
                                            Y_grid + i * interval_vertical * Scale_inv_vwport * Scale_Modelspace * Vert_exag, 0)
                    .Layer = Layer_grid_lines
                    .ColorIndex = color_index_grid_lines
                    .Linetype = ComboBox_linetype.Text
                End With

                BTrecord.AppendEntity(Linia_hidden)
                Trans1.AddNewlyCreatedDBObject(Linia_hidden, True)
            Next

            'aici desenez liniile mari horiz
            For i = 1 To Nr_vertical_labels
                Dim Linie_horiz_lunga_stanga As New Autodesk.AutoCAD.DatabaseServices.Line
                With Linie_horiz_lunga_stanga
                    .StartPoint = New Point3d(X_grid, _
                                              Y_grid + i * interval_vertical * Scale_inv_vwport * Scale_Modelspace * Vert_exag, 0)
                    .EndPoint = New Point3d(X_grid + 2 * Lungime_linie_mica_horiz * Scale_inv_vwport, _
                                            Y_grid + i * interval_vertical * Scale_inv_vwport * Scale_Modelspace * Vert_exag, 0)
                    .Layer = Layer_grid_lines
                    .Linetype = "CONTINUOUS"
                    .ColorIndex = 2
                End With

                BTrecord.AppendEntity(Linie_horiz_lunga_stanga)
                Trans1.AddNewlyCreatedDBObject(Linie_horiz_lunga_stanga, True)

                Dim Linie_horiz_lunga_dreapta As New Autodesk.AutoCAD.DatabaseServices.Line
                With Linie_horiz_lunga_dreapta
                    .StartPoint = New Point3d(X_grid + Lungime_graph_1_1 * Scale_inv_vwport * Scale_Modelspace - 2 * Lungime_linie_mica_horiz * Scale_inv_vwport, _
                                              Y_grid + i * interval_vertical * Scale_inv_vwport * Scale_Modelspace * Vert_exag, 0)
                    .EndPoint = New Point3d(X_grid + Lungime_graph_1_1 * Scale_inv_vwport * Scale_Modelspace, _
                                            Y_grid + i * interval_vertical * Scale_inv_vwport * Scale_Modelspace * Vert_exag, 0)
                    .Layer = Layer_grid_lines

                    .Linetype = "CONTINUOUS"
                    .ColorIndex = 2
                End With

                BTrecord.AppendEntity(Linie_horiz_lunga_dreapta)
                Trans1.AddNewlyCreatedDBObject(Linie_horiz_lunga_dreapta, True)
            Next

            'aici desenez liniile mici horizontale
            For i = 1 To Nr_vertical_labels
                Dim Linie_horiz_mica_stanga As New Autodesk.AutoCAD.DatabaseServices.Line
                With Linie_horiz_mica_stanga
                    .StartPoint = New Point3d(X_grid, _
                                              Y_grid + (i - 1) * interval_vertical * Scale_inv_vwport * Scale_Modelspace * Vert_exag + 0.5 * interval_vertical * Scale_inv_vwport * Scale_Modelspace * Vert_exag, 0)
                    .EndPoint = New Point3d(X_grid + Lungime_linie_mica_horiz * Scale_inv_vwport, _
                                            Y_grid + (i - 1) * interval_vertical * Scale_inv_vwport * Scale_Modelspace * Vert_exag + 0.5 * interval_vertical * Scale_inv_vwport * Scale_Modelspace * Vert_exag, 0)
                    .Layer = Layer_grid_lines

                    .Linetype = "CONTINUOUS"
                    .ColorIndex = 2
                End With

                BTrecord.AppendEntity(Linie_horiz_mica_stanga)
                Trans1.AddNewlyCreatedDBObject(Linie_horiz_mica_stanga, True)

                Dim Linie_horiz_mica_dreapta As New Autodesk.AutoCAD.DatabaseServices.Line
                With Linie_horiz_mica_dreapta
                    .StartPoint = New Point3d(X_grid + Lungime_graph_1_1 * Scale_inv_vwport * Scale_Modelspace - Lungime_linie_mica_horiz * Scale_inv_vwport, _
                                              Y_grid + (i - 1) * interval_vertical * Scale_inv_vwport * Scale_Modelspace * Vert_exag + 0.5 * interval_vertical * Scale_inv_vwport * Scale_Modelspace * Vert_exag, 0)
                    .EndPoint = New Point3d(X_grid + Lungime_graph_1_1 * Scale_inv_vwport * Scale_Modelspace, _
                                            Y_grid + (i - 1) * interval_vertical * Scale_inv_vwport * Scale_Modelspace * Vert_exag + 0.5 * interval_vertical * Scale_inv_vwport * Scale_Modelspace * Vert_exag, 0)
                    .Layer = Layer_grid_lines

                    .Linetype = "CONTINUOUS"
                    .ColorIndex = 2
                End With

                BTrecord.AppendEntity(Linie_horiz_mica_dreapta)
                Trans1.AddNewlyCreatedDBObject(Linie_horiz_mica_dreapta, True)
            Next

            If Nr_chainage_labels / 2 - CInt(Nr_chainage_labels / 2) = 0 Then
                Nr_chainage_labels = Nr_chainage_labels + 1
            End If

            Dim Lung_start_thick_vert As Double = 0.5 * (Lungime_graph_1_1 - interval_horizontal * (Nr_chainage_labels - 1))

            'aici desenez liniile verticale thick
            For i = 1 To Nr_chainage_labels
                Dim Linie_verticala_thick As New Autodesk.AutoCAD.DatabaseServices.Line
                With Linie_verticala_thick
                    .StartPoint = New Point3d(X_grid + Lung_start_thick_vert * Scale_inv_vwport * Scale_Modelspace + (i - 1) * interval_horizontal * Scale_inv_vwport * Scale_Modelspace, _
                                              Y_grid, 0)
                    .EndPoint = New Point3d(X_grid + Lung_start_thick_vert * Scale_inv_vwport * Scale_Modelspace + (i - 1) * interval_horizontal * Scale_inv_vwport * Scale_Modelspace, _
                                            Y_grid - Lung_thick_vert * Scale_inv_vwport, 0)
                    .Layer = Layer_grid_lines
                    .ColorIndex = 2
                    .Linetype = "CONTINUOUS"
                End With

                BTrecord.AppendEntity(Linie_verticala_thick)
                Trans1.AddNewlyCreatedDBObject(Linie_verticala_thick, True)

                Dim Chainage As Double = ((i - 1) * interval_horizontal - 0.5 * interval_horizontal * (Nr_chainage_labels - 1)) * Scale_inv_vwport


                Dim Mtext_chainage_label As New Autodesk.AutoCAD.DatabaseServices.MText
                Mtext_chainage_label.Layer = Layer_no_plot
                Mtext_chainage_label.Contents = Get_chainage_from_double(Chainage, 0)
                Mtext_chainage_label.TextStyleId = Text_style_id
                Mtext_chainage_label.TextHeight = 2.5 * Scale_inv_vwport
                Mtext_chainage_label.Location = New Autodesk.AutoCAD.Geometry.Point3d(X_grid + Lung_start_thick_vert * Scale_inv_vwport * Scale_Modelspace + (i - 1) * interval_horizontal * Scale_inv_vwport * Scale_Modelspace, _
                                                                                      Y_grid - Lung_thick_vert * Scale_inv_vwport - Lung_thick_vert * Scale_inv_vwport, 0)
                Mtext_chainage_label.Rotation = 0
                Mtext_chainage_label.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.TopCenter

                'Mtext_chainage_label.ColorIndex = 5
                BTrecord.AppendEntity(Mtext_chainage_label)
                Trans1.AddNewlyCreatedDBObject(Mtext_chainage_label, True)
            Next

            Dim Mtext_west As New Autodesk.AutoCAD.DatabaseServices.MText
            Mtext_west.Layer = Layer_no_plot
            Mtext_west.Contents = "NORTH/WEST/SOUTH/EAST"
            Mtext_west.TextStyleId = Text_style_id
            Mtext_west.TextHeight = 2.5 * 1.6 * Scale_inv_vwport
            Mtext_west.Location = New Autodesk.AutoCAD.Geometry.Point3d(X_grid, Y_grid - 2.5 * 1.6 * Scale_inv_vwport, 0)
            Mtext_west.Rotation = 0
            Mtext_west.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.TopLeft
            BTrecord.AppendEntity(Mtext_west)
            Trans1.AddNewlyCreatedDBObject(Mtext_west, True)

            Dim Mtext_east As New Autodesk.AutoCAD.DatabaseServices.MText
            Mtext_east.Layer = Layer_no_plot
            Mtext_east.Contents = "NORTH/WEST/SOUTH/EAST"
            Mtext_east.TextStyleId = Text_style_id
            Mtext_east.TextHeight = 2.5 * 1.6 * Scale_inv_vwport
            Mtext_east.Location = New Autodesk.AutoCAD.Geometry.Point3d(X_grid + Lungime_graph_1_1 * Scale_inv_vwport * Scale_Modelspace, Y_grid - 2.5 * 1.6 * Scale_inv_vwport, 0)
            Mtext_east.Rotation = 0
            Mtext_east.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.TopRight

            BTrecord.AppendEntity(Mtext_east)
            Trans1.AddNewlyCreatedDBObject(Mtext_east, True)

            Dim Mtext_titlu_rand1 As New Autodesk.AutoCAD.DatabaseServices.MText
            Mtext_titlu_rand1.Layer = Layer_text
            Mtext_titlu_rand1.Contents = "{\LDITCHLINE PROFILE ALONG CENTRELINE OF DITCH}"
            Mtext_titlu_rand1.TextStyleId = Text_style_id
            Mtext_titlu_rand1.TextHeight = 2.5 * 1.6 * Scale_inv_vwport
            Mtext_titlu_rand1.Location = New Autodesk.AutoCAD.Geometry.Point3d(X_grid + Lungime_graph_1_1 * Scale_Modelspace * Scale_inv_vwport / 2, Y_grid - 13 * Scale_inv_vwport, 0)
            Mtext_titlu_rand1.Rotation = 0
            Mtext_titlu_rand1.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.TopCenter

            BTrecord.AppendEntity(Mtext_titlu_rand1)
            Trans1.AddNewlyCreatedDBObject(Mtext_titlu_rand1, True)

            'Dim Linie_titlu As New Autodesk.AutoCAD.DatabaseServices.Line
            'Linie_titlu.Layer = Layer_text
            'Linie_titlu.StartPoint = New Autodesk.AutoCAD.Geometry.Point3d(X_grid + Lungime_graph_1_1 * Scale_Modelspace * Scale_inv_vwport / 2 - 30 * Scale_inv_vwport, Y_grid - 13 * Scale_inv_vwport - 5.5 * Scale_inv_vwport, 0)
            'Linie_titlu.EndPoint = New Autodesk.AutoCAD.Geometry.Point3d(X_grid + Lungime_graph_1_1 * Scale_Modelspace * Scale_inv_vwport / 2 + 30 * Scale_inv_vwport, Y_grid - 13 * Scale_inv_vwport - 5.5 * Scale_inv_vwport, 0)

            'BTrecord.AppendEntity(Linie_titlu)
            'Trans1.AddNewlyCreatedDBObject(Linie_titlu, True)

            Dim Mtext_titlu_rand2 As New Autodesk.AutoCAD.DatabaseServices.MText
            Mtext_titlu_rand2.Layer = Layer_text
            Mtext_titlu_rand2.Contents = "SCALES HORIZONTAL 1:" & H_scale_hmm * Scale_inv_vwport & " VERTICAL 1:" & V_scale_hmm * Scale_inv_vwport
            Mtext_titlu_rand2.TextStyleId = Text_style_id
            Mtext_titlu_rand2.TextHeight = 2.5 * Scale_inv_vwport
            Mtext_titlu_rand2.Location = New Autodesk.AutoCAD.Geometry.Point3d(X_grid + Lungime_graph_1_1 * Scale_Modelspace * Scale_inv_vwport / 2, Y_grid - 13 * Scale_inv_vwport - 5.5 * Scale_inv_vwport - 2.5 * Scale_inv_vwport / 2, 0)
            Mtext_titlu_rand2.Rotation = 0
            Mtext_titlu_rand2.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.TopCenter

            BTrecord.AppendEntity(Mtext_titlu_rand2)
            Trans1.AddNewlyCreatedDBObject(Mtext_titlu_rand2, True)


            Dim Distanta_de_la_elevatia_cunoscuta As Double = Elevatia_pt_pct_zero_masurata * (V_scale / H_scale) * (H_scale / Pr_scale)

            If Ypoly > Yreference Then
                Distanta_de_la_elevatia_cunoscuta = -Distanta_de_la_elevatia_cunoscuta
            End If

            Dim Elevatia_punctului As Double = Elevatia_cunoscuta - Distanta_de_la_elevatia_cunoscuta

            Dim Closest_elev_label As Double = Cel_mai_aproape_multiplu(Elevatia_punctului, 10 * Scale_inv_vwport)

            Dim Distanta_pana_la_linia_referinta As Double = (Elevatia_punctului - Closest_elev_label) * (Pr_scale_hmm / V_scale_hmm)

            Dim Cantitate_elevatie As Double = nr_linii_jos * interval_vertical

            For i = 1 To Nr_vertical_labels
                Dim Mtext_elev_stanga As New Autodesk.AutoCAD.DatabaseServices.MText
                Mtext_elev_stanga.Layer = Layer_text
                Mtext_elev_stanga.Contents = Closest_elev_label - Cantitate_elevatie * Scale_inv_vwport
                Mtext_elev_stanga.TextStyleId = Text_style_id
                Mtext_elev_stanga.TextHeight = 2.5 * Scale_inv_vwport
                Mtext_elev_stanga.Location = New Autodesk.AutoCAD.Geometry.Point3d(X_grid + 2.5 * Scale_inv_vwport, _
                                                                                   Y_grid + (i - 1) * interval_vertical * Scale_inv_vwport * Vert_exag * Scale_Modelspace + 0.75 * Scale_inv_vwport, 0)
                Mtext_elev_stanga.Rotation = 0
                Mtext_elev_stanga.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.BottomLeft
                Mtext_elev_stanga.ColorIndex = 2

                BTrecord.AppendEntity(Mtext_elev_stanga)
                Trans1.AddNewlyCreatedDBObject(Mtext_elev_stanga, True)

                Dim Mtext_elev_dreapta As New Autodesk.AutoCAD.DatabaseServices.MText
                Mtext_elev_dreapta.Layer = Layer_text
                Mtext_elev_dreapta.Contents = Closest_elev_label - Cantitate_elevatie * Scale_inv_vwport
                Mtext_elev_dreapta.TextStyleId = Text_style_id
                Mtext_elev_dreapta.TextHeight = 2.5 * Scale_inv_vwport
                Mtext_elev_dreapta.Location = New Autodesk.AutoCAD.Geometry.Point3d(X_grid + Lungime_graph_1_1 * Scale_inv_vwport * Scale_Modelspace - 2.5 * Scale_inv_vwport, _
                                                                                    Y_grid + (i - 1) * interval_vertical * Scale_inv_vwport * Vert_exag * Scale_Modelspace + 0.75 * Scale_inv_vwport, 0)
                Mtext_elev_dreapta.Rotation = 0
                Mtext_elev_dreapta.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.BottomRight
                Mtext_elev_dreapta.ColorIndex = 2

                BTrecord.AppendEntity(Mtext_elev_dreapta)
                Trans1.AddNewlyCreatedDBObject(Mtext_elev_dreapta, True)


                Cantitate_elevatie = Cantitate_elevatie - interval_vertical

            Next





            Dim poly_graph As New Autodesk.AutoCAD.DatabaseServices.Polyline
            poly_graph.Layer = layer_polyline
            For i = 0 To Colectie_puncte_graph.Count - 1
                Dim X1, Y1 As Double
                X1 = X_grid + (Lungime_graph_1_1 * Scale_Modelspace * Scale_inv_vwport / 2 - (Xpoly - Colectie_puncte_graph(i).X) * (H_scale / H_scale_hmm) * (Pr_scale_hmm / Pr_scale))
                Y1 = Y_grid + (nr_linii_jos * interval_vertical * Scale_inv_vwport * Scale_Modelspace * Vert_exag + Distanta_pana_la_linia_referinta - (Ypoly - Colectie_puncte_graph(i).Y) * (V_scale / V_scale_hmm) * (Pr_scale_hmm / Pr_scale))
                poly_graph.AddVertexAt(i, New Autodesk.AutoCAD.Geometry.Point2d(X1, Y1), 0, 0, 0)
            Next
            BTrecord.AppendEntity(poly_graph)
            Trans1.AddNewlyCreatedDBObject(poly_graph, True)


            Dim Viewport_paper As New Viewport
            Viewport_paper.SetDatabaseDefaults()
            Viewport_paper.CenterPoint = viewport_ps_center ' asta e pozitia viewport in paper space
            Viewport_paper.Height = viewport_height
            Viewport_paper.Width = viewport_width
            Viewport_paper.Layer = Layer_Viewport
            Viewport_paper.ViewDirection = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis
            Viewport_paper.ViewTarget = New Autodesk.AutoCAD.Geometry.Point3d(X_grid + Scale_inv_vwport * Scale_Modelspace * Lungime_graph_1_1 / 2, _
                                                                              Y_grid - Scale_inv_vwport * Gap_PS_jos + Scale_inv_vwport * (Inaltime_graph * Vert_exag * Scale_Modelspace + Gap_PS_jos + Gap_PS_sus) / 2, 0) ' asta e pozitia viewport in MODEL space
            Viewport_paper.ViewCenter = Autodesk.AutoCAD.Geometry.Point2d.Origin
            Viewport_paper.TwistAngle = 0

            BTrecord_paper.AppendEntity(Viewport_paper)
            Trans1.AddNewlyCreatedDBObject(Viewport_paper, True)
            Viewport_paper.On = True
            Viewport_paper.CustomScale = Viewport_scale
            Viewport_paper.Locked = True


            Dim Mtext_Paper_elevation_1 As New Autodesk.AutoCAD.DatabaseServices.MText
            Mtext_Paper_elevation_1.Layer = Layer_text
            Mtext_Paper_elevation_1.Contents = "ELEVATION"
            Mtext_Paper_elevation_1.TextStyleId = Text_style_id
            Mtext_Paper_elevation_1.TextHeight = 2.5 * Scale_inv_vwport
            Mtext_Paper_elevation_1.Location = New Autodesk.AutoCAD.Geometry.Point3d(X_grid - 1.25 * Scale_inv_vwport, _
                                                                                      Y_grid + Inaltime_graph * Scale_inv_vwport * Scale_Modelspace * Vert_exag / 2, 0)
            Mtext_Paper_elevation_1.Rotation = Math.PI / 2
            Mtext_Paper_elevation_1.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.BottomCenter
            BTrecord.AppendEntity(Mtext_Paper_elevation_1)
            Trans1.AddNewlyCreatedDBObject(Mtext_Paper_elevation_1, True)


            Dim Mtext_Paper_elevation_2 As New Autodesk.AutoCAD.DatabaseServices.MText
            Mtext_Paper_elevation_2.Layer = Layer_text
            Mtext_Paper_elevation_2.Contents = "ELEVATION"
            Mtext_Paper_elevation_2.TextStyleId = Text_style_id
            Mtext_Paper_elevation_2.TextHeight = 2.5 * Scale_inv_vwport
            Mtext_Paper_elevation_2.Location = New Autodesk.AutoCAD.Geometry.Point3d(X_grid + Lungime_graph_1_1 * Scale_inv_vwport * Scale_Modelspace + 1.25 * Scale_inv_vwport, _
                                                                                      Y_grid + Inaltime_graph * Scale_inv_vwport * Scale_Modelspace * Vert_exag / 2, 0)
            Mtext_Paper_elevation_2.Rotation = Math.PI / 2
            Mtext_Paper_elevation_2.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.TopCenter
            BTrecord.AppendEntity(Mtext_Paper_elevation_2)
            Trans1.AddNewlyCreatedDBObject(Mtext_Paper_elevation_2, True)


            Dim Linia_ref_survey_zero As New Autodesk.AutoCAD.DatabaseServices.Line
            Linia_ref_survey_zero.StartPoint = New Point3d(X_grid + Scale_inv_vwport * Scale_Modelspace * Lungime_graph_1_1 / 2, Y_grid, 0)
            Linia_ref_survey_zero.EndPoint = New Point3d(X_grid + Scale_inv_vwport * Scale_Modelspace * Lungime_graph_1_1 / 2,
                                                         Y_grid + Scale_inv_vwport * interval_vertical * (Nr_vertical_labels + 1) * Scale_Modelspace * Vert_exag, 0)
            Linia_ref_survey_zero.Layer = Layer_no_plot
            BTrecord.AppendEntity(Linia_ref_survey_zero)
            Trans1.AddNewlyCreatedDBObject(Linia_ref_survey_zero, True)

            If Valoare_Chainage.Count > 0 Then
                Dim IDs As New ObjectIdCollection
                For i = 0 To Valoare_Chainage.Count - 1
                    Dim Linia_Ref_survey_2 As New Autodesk.AutoCAD.DatabaseServices.Line
                    Linia_Ref_survey_2.StartPoint = New Point3d(X_grid + Scale_inv_vwport * Scale_Modelspace * Lungime_graph_1_1 / 2 + Valoare_Chainage(i), _
                                                                Y_grid, 0)
                    Linia_Ref_survey_2.EndPoint = New Point3d(X_grid + Scale_inv_vwport * Scale_Modelspace * Lungime_graph_1_1 / 2 + Valoare_Chainage(i), _
                                                              Y_grid + Scale_inv_vwport * Scale_Modelspace * Vert_exag * interval_vertical * (Nr_vertical_labels + 1), 0)
                    Linia_Ref_survey_2.Layer = Layer_no_plot
                    BTrecord.AppendEntity(Linia_Ref_survey_2)
                    Trans1.AddNewlyCreatedDBObject(Linia_Ref_survey_2, True)


                    Dim Colectie_temp001 As New Point3dCollection
                    Linia_Ref_survey_2.IntersectWith(poly_graph, Intersect.ExtendThis, Colectie_temp001, IntPtr.Zero, IntPtr.Zero)



                    Dim Valoare_x As Double = Colectie_temp001.Item(0).X - X_grid
                    Dim Valoare_y As Double = Colectie_temp001.Item(0).Y - Y_grid
                    Dim Pozitie_block_x As Double = Viewport_paper.CenterPoint.X - viewport_width / 2 + Gap_PS_stanga_dreapta + Valoare_x * Viewport_paper.CustomScale
                    Dim Pozitie_block_y As Double = Viewport_paper.CenterPoint.Y - viewport_height / 2 + Gap_PS_jos + Valoare_y * Viewport_paper.CustomScale

                    Dim Xmax, Xmin As Double
                    Xmax = Viewport_paper.CenterPoint.X + viewport_width / 2 + 0.002
                    Xmin = Viewport_paper.CenterPoint.X - viewport_width / 2 - 0.002

                    If Pozitie_block_x > Xmin And Pozitie_block_x < Xmax Then
                        If Valoare_Chainage(i) = 0 Then
                            InsertBlock_PROFILE_TAG(New Point3d(Pozitie_block_x, Pozitie_block_y, 0), BTrecord_paper, Layer_blocks, "0+000.0", "℄ ")
                        Else
                            Dim Descriptie As String = Replace(Descriptie_Chainage(i), vbCrLf, " ").ToUpper
                            If Strings.Left(Descriptie, 1) = " " Then Descriptie = Strings.Right(Descriptie, Len(Descriptie) - 1)
                            InsertBlock_PROFILE_TAG(New Point3d(Pozitie_block_x, Pozitie_block_y, 0), BTrecord_paper, Layer_blocks, String_Chainage(i), Descriptie)
                        End If

                    End If

                    Dim Mtext_ref_survey_description As New Autodesk.AutoCAD.DatabaseServices.MText
                    Mtext_ref_survey_description.Layer = Layer_no_plot
                    Mtext_ref_survey_description.Contents = String_Chainage(i) & " - " & Replace(Descriptie_Chainage(i), vbCrLf, " ").ToUpper
                    Mtext_ref_survey_description.TextStyleId = Text_style_id
                    Mtext_ref_survey_description.TextHeight = Scale_inv_vwport
                    Mtext_ref_survey_description.Location = New Autodesk.AutoCAD.Geometry.Point3d(X_grid + Scale_inv_vwport * Scale_Modelspace * Lungime_graph_1_1 / 2 + Valoare_Chainage(i), Y_grid, 0)
                    Mtext_ref_survey_description.Rotation = Math.PI / 2
                    Mtext_ref_survey_description.ColorIndex = 9
                    Mtext_ref_survey_description.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.MiddleLeft
                    BTrecord.AppendEntity(Mtext_ref_survey_description)
                    Trans1.AddNewlyCreatedDBObject(Mtext_ref_survey_description, True)
                    IDs.Add(Mtext_ref_survey_description.ObjectId)

                Next
                Dim DrawOrderTable1 As Autodesk.AutoCAD.DatabaseServices.DrawOrderTable = Trans1.GetObject(BTrecord.DrawOrderTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)
                DrawOrderTable1.MoveToTop(IDs)



            End If




        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Sub Button_Draw_Click(sender As System.Object, e As System.EventArgs) Handles Button_Draw.Click

        If isSECURE() = False Then Exit Sub
        If IsNumeric(TextBox_PROFILE_LENGTH.Text) = False Then
            Exit Sub
        End If
        If IsNumeric(TextBox_horizontal_increment.Text) = False Then
            Exit Sub
        End If
        If IsNumeric(TextBox_vertical_increment.Text) = False Then
            Exit Sub
        End If
        If IsNumeric(TextBox_nr_vert_labels.Text) = False Then
            Exit Sub
        End If


        Dim layer_viewport_name As String = "VPORT"
        Dim layer_viewport_description As String = "Layer used to create a view port in paper space"
        Dim layer_viewport_color As Integer = 7
        Dim layer_viewport_ltype As String = "Continuous"
        Dim layer_viewport_lweight As LineWeight = LineWeight.LineWeight025
        Dim layer_viewport_plot As Boolean = False
        Dim layer_BLOCKS_name As String = "TEXT"
        Dim layer_BLOCKS_description As String = "All text, notes baloons, section symbols, legends, door-room-wall numbers, text leader lines"
        Dim layer_BLOCKS_color As Integer = 2
        Dim layer_BLOCKS_ltype As String = "Continuous"
        Dim layer_BLOCKS_lweight As LineWeight = LineWeight.LineWeight025
        Dim layer_BLOCKS_plot As Boolean = True
        Dim layer_water_name As String = "PWATER"
        Dim layer_water_description As String = "Water"
        Dim layer_water_color As Integer = 5
        Dim layer_water_ltype As String = "Continuous"
        Dim layer_water_lweight As LineWeight = LineWeight.LineWeight035
        Dim layer_water_plot As Boolean = True
        Dim layer_Text_name As String = "TEXT"
        Dim layer_text_description As String = "All text, notes baloons, section symbols, legends, door-room-wall numbers, text leader lines"
        Dim layer_Text_color As Integer = 2
        Dim layer_Text_ltype As String = "Continuous"
        Dim layer_Text_lweight As LineWeight = LineWeight.LineWeight025
        Dim layer_Text_plot As Boolean = True
        Dim layer_grid_name As String = "PGRID"
        Dim layer_grid_description As String = "Profile grids"
        Dim layer_grid_color As Integer = 6
        Dim layer_grid_ltype As String = "TCHIDDEN"
        Dim layer_grid_lweight As LineWeight = LineWeight.LineWeight025
        Dim layer_grid_plot As Boolean = True
        Dim layer_polyline_name As String = "PGRADE"
        Dim layer_polyline_description As String = "Grade"
        Dim layer_polyline_color As Integer = 3
        Dim layer_polyline_ltype As String = "Continuous"
        Dim layer_polyline_lweight As LineWeight = LineWeight.LineWeight035
        Dim layer_polyline_plot As Boolean = True
        Dim layer_no_plot_name As String = "NO PLOT"
        Dim layer_no_plot_description As String = "NO PLOT"
        Dim layer_no_plot_color As Integer = 40
        Dim layer_no_plot_ltype As String = "Continuous"
        Dim layer_no_plot_lweight As LineWeight = LineWeight.LineWeight025
        Dim layer_no_plot_plot As Boolean = False

        Dim Color_index_grey_lines As Integer = 251
        Dim Grid_ltype As String = "TCDOT2"

        Dim Lungime_graph_1_1 As Double = CDbl(TextBox_PROFILE_LENGTH.Text)

        Dim Lung_liniuta_hor As Double = 5
        Dim Lung_liniuta_ver As Double = 2.5

        Dim Interval_hor = CDbl(TextBox_horizontal_increment.Text)
        Dim Interval_ver = CDbl(TextBox_vertical_increment.Text)
        Dim Nr_vertical_labels As Integer = CDbl(TextBox_nr_vert_labels.Text)

        Dim Scale_inv_vwport As Double
        Dim Nr_chainage_labels As Integer = 7

        Dim Nr_linii_jos As Integer

        Dim Gap_PS_sus As Double = 10
        Dim Gap_PS_jos As Double = 25
        Dim Gap_PS_stanga_dreapta As Double = 25

        Dim Inaltime_graph As Double = Nr_vertical_labels * Interval_ver

        Dim Viewport_Height As Double = Inaltime_graph + Gap_PS_jos + Gap_PS_sus
        Dim Viewport_Width As Double = Lungime_graph_1_1 + Gap_PS_stanga_dreapta * 2
        Dim Viewport_CenterPoint As Point3d

        Dim X_paper As Double = 1800
        Dim y_paper As Double = 185

        Dim Vw_Spacing_X_paper As Double = 45
        Dim Vw_Spacing_Y_paper As Double = 10





        If TextBox_Horizontal_scale.Text = "" Then
            MsgBox("Please specify horizontal scale")
            TextBox_Horizontal_scale.SelectAll()
            Exit Sub
        End If

        If IsNumeric(TextBox_Horizontal_scale.Text) = False Then
            MsgBox("Please specify horizontal scale")
            TextBox_Horizontal_scale.SelectAll()
            Exit Sub
        End If

        If TextBox_Vertical_scale.Text = "" Then
            MsgBox("Please specify vertical scale")
            TextBox_Vertical_scale.SelectAll()
            Exit Sub
        End If

        If IsNumeric(TextBox_Vertical_scale.Text) = False Then
            MsgBox("Please specify vertical scale")
            TextBox_Vertical_scale.SelectAll()
            Exit Sub
        End If

        If TextBox_printing_scale.Text = "" Then
            MsgBox("Please specify printing scale")
            TextBox_printing_scale.SelectAll()
            Exit Sub
        End If

        If IsNumeric(TextBox_printing_scale.Text) = False Then
            MsgBox("Please specify printing scale")
            TextBox_printing_scale.SelectAll()
            Exit Sub
        End If

        If ComboBox_Horizontal_scale_HMM.Text = "" Then
            MsgBox("Please specify horizontal scale")
            ComboBox_Horizontal_scale_HMM.SelectAll()
            Exit Sub
        End If

        If IsNumeric(ComboBox_Horizontal_scale_HMM.Text) = False Then
            MsgBox("Please specify horizontal scale")
            ComboBox_Horizontal_scale_HMM.SelectAll()
            Exit Sub
        End If

        If ComboBox_Vertical_scale_HMM.Text = "" Then
            MsgBox("Please specify vertical scale")
            Exit Sub
        End If

        If IsNumeric(TextBox_Vertical_scale.Text) = False Then
            MsgBox("Please specify vertical scale")
            TextBox_Vertical_scale.SelectAll()
            Exit Sub
        End If

        If TextBox_Printing_scale_HMM.Text = "" Then
            MsgBox("Please specify printing scale")
            TextBox_Printing_scale_HMM.SelectAll()
            Exit Sub
        End If

        If IsNumeric(TextBox_Printing_scale_HMM.Text) = False Then
            MsgBox("Please specify printing scale")
            TextBox_Printing_scale_HMM.SelectAll()
            Exit Sub
        End If






        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
        ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim Empty_array() As ObjectId
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try
            Dim Lock1 As Autodesk.AutoCAD.ApplicationServices.DocumentLock
            Lock1 = ThisDrawing.LockDocument
            Using Lock1
                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = "Select polyline:"

                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)


                If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.SetImpliedSelection(Empty_array)
                    Exit Sub
                End If

                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord

                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)

                    Dim BTrecord_PAPER As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BTrecord_PAPER = Trans1.GetObject(BlockTable_data1(BlockTableRecord.PaperSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)

                    Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                    Obj1 = Rezultat1.Value.Item(0)
                    Dim Ent1 As Entity
                    Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                    Dim Colectie_x As New DoubleCollection
                    Dim Colectie_y As New DoubleCollection
                    Dim Colectie_puncte_graph As New Point2dCollection

                    Dim Text_style_table As Autodesk.AutoCAD.DatabaseServices.TextStyleTable = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    Dim Text_style_ID As Autodesk.AutoCAD.DatabaseServices.ObjectId = Text_style_table.Item(ComboBox_text_style.Text)


                    If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then

                        'aici citesc chainage si descriptie

                        Dim Rezultat3 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                        Dim Object_prompt3 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_prompt3.MessageForAdding = "Select the existing chainage labels:" & vbCrLf

                        Object_prompt3.SingleOnly = False

                        Rezultat3 = Editor1.GetSelection(Object_prompt3)





                        Dim Poly1 As Autodesk.AutoCAD.DatabaseServices.Polyline = Ent1
                        For i = 0 To Poly1.NumberOfVertices - 1
                            Colectie_x.Add(Poly1.GetPoint2dAt(i).X)
                            Colectie_y.Add(Poly1.GetPoint2dAt(i).Y)
                            Colectie_puncte_graph.Add(Poly1.GetPoint2dAt(i))
                        Next

                        Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult

                        Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions("Specify 0+000 position on the polyline:" & vbCrLf)




                        PP1.AllowNone = False
                        Point1 = Editor1.GetPoint(PP1)
                        If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                            Exit Sub
                        End If

                        Dim X0, Y0 As Double
                        X0 = Poly1.GetClosestPointTo(New Point3d(Point1.Value.X, Point1.Value.Y, 0), Vector3d.ZAxis, False).X
                        Y0 = Poly1.GetClosestPointTo(New Point3d(Point1.Value.X, Point1.Value.Y, 0), Vector3d.ZAxis, False).Y

                        Dim distanta_de_la_inceputul_grafului As Double = X0 - Poly1.GetPoint2dAt(0).X

                        Dim Valoare_Chainage As New DoubleCollection
                        Dim Descriptie_Chainage As New Specialized.StringCollection
                        Dim String_Chainage As New Specialized.StringCollection


                        If Rezultat3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            If Rezultat3.Value.Count > 0 Then
                                For i = 0 To Rezultat3.Value.Count - 1
                                    Dim Obj33 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                    Obj33 = Rezultat3.Value.Item(i)
                                    Dim Ent33 As Entity
                                    Ent33 = Obj33.ObjectId.GetObject(OpenMode.ForRead)
                                    Dim Mtext_label1 As MText
                                    Dim Text_label1 As DBText

                                    If TypeOf Ent33 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                                        Mtext_label1 = Ent33
                                        Dim String1 As String = Mtext_label1.Text
                                        Dim Poz_plus As Integer = InStr(String1, "+")
                                        If Poz_plus > 0 Then
                                            Dim Partea1 As String
                                            Dim Partea1_for_chainage As String = Replace(String1, " ", "")
                                            Partea1 = extrage_chainage_din_text_de_la_inceputul_textului(Partea1_for_chainage)

                                            If IsNumeric(Replace(Partea1, "+", "")) = True Then
                                                Dim Numar1 As Double = CDbl(Replace(Partea1, "+", ""))
                                                If CheckBox_recalculate_chainage.Checked = True Then
                                                    Dim Nou_nr1 As Double
                                                    Dim Nou_partea1 As String
                                                    Nou_nr1 = Numar1 - distanta_de_la_inceputul_grafului
                                                    Nou_partea1 = Get_chainage_from_double(Nou_nr1, 1)
                                                    Valoare_Chainage.Add(Nou_nr1)
                                                    String_Chainage.Add(Nou_partea1)
                                                Else
                                                    Valoare_Chainage.Add(Numar1)
                                                    String_Chainage.Add(Partea1)
                                                End If


                                                Dim Descr1 As String = String1.Replace(Partea1, "")
                                                If Strings.Left(Descr1, 1) = " " Then Descr1 = Strings.Right(Descr1, Len(Descr1) - 1)
                                                If Descr1 = "" Or Descr1 = " " Or Descr1 = "  " Then
                                                    Descr1 = "xxx"
                                                End If

                                                Descriptie_Chainage.Add(Descr1)
                                            Else
                                                Dim Partea11 As String
                                                Dim Partea11_for_chainage As String = Replace(String1, " ", "")
                                                Partea11 = extrage_chainage_din_text_de_la_sfarsitul_textului(Partea11_for_chainage)
                                                If IsNumeric(Replace(Partea11, "+", "")) = True Then
                                                    Dim Numar11 As Double = CDbl(Replace(Partea11, "+", ""))
                                                    If CheckBox_recalculate_chainage.Checked = True Then
                                                        Dim Nou_nr1 As Double
                                                        Dim Nou_partea1 As String
                                                        Nou_nr1 = Numar11 - distanta_de_la_inceputul_grafului
                                                        Nou_partea1 = Get_chainage_from_double(Nou_nr1, 1)
                                                        Valoare_Chainage.Add(Nou_nr1)
                                                        String_Chainage.Add(Nou_partea1)
                                                    Else
                                                        Valoare_Chainage.Add(Numar11)
                                                        String_Chainage.Add(Partea11)
                                                    End If
                                                    Dim Descr11 As String = String1.Replace(Partea11, "")
                                                    If Strings.Left(Descr11, 1) = " " Then Descr11 = Strings.Right(Descr11, Len(Descr11) - 1)
                                                    If Descr11 = "" Or Descr11 = " " Or Descr11 = "  " Then
                                                        Descr11 = "xxx"
                                                    End If

                                                    Descriptie_Chainage.Add(Descr11)
                                                End If

                                            End If
                                        End If ' asta e de la  If InStr(String1, "+") > 0
                                    End If ' asta e de la If TypeOf Ent33 Is Autodesk.AutoCAD.DatabaseServices.MText

                                    If TypeOf Ent33 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                                        Text_label1 = Ent33
                                        Dim String1 As String = Text_label1.TextString
                                        Dim Poz_plus As Integer = InStr(String1, "+")
                                        If Poz_plus > 0 Then
                                            Dim Partea1 As String
                                            Dim Partea1_for_chainage As String = Replace(String1, " ", "")
                                            Partea1 = extrage_chainage_din_text_de_la_inceputul_textului(Partea1_for_chainage)

                                            If IsNumeric(Replace(Partea1, "+", "")) = True Then
                                                Dim Numar1 As Double = CDbl(Replace(Partea1, "+", ""))
                                                If CheckBox_recalculate_chainage.Checked = True Then
                                                    Dim Nou_nr1 As Double
                                                    Dim Nou_partea1 As String
                                                    Nou_nr1 = Numar1 - distanta_de_la_inceputul_grafului
                                                    Nou_partea1 = Get_chainage_from_double(Nou_nr1, 1)
                                                    Valoare_Chainage.Add(Nou_nr1)
                                                    String_Chainage.Add(Nou_partea1)
                                                Else
                                                    Valoare_Chainage.Add(Numar1)
                                                    String_Chainage.Add(Partea1)
                                                End If

                                                Dim Descr1 As String = String1.Replace(Partea1, "")
                                                If Strings.Left(Descr1, 1) = " " Then Descr1 = Strings.Right(Descr1, Len(Descr1) - 1)
                                                If Descr1 = "" Or Descr1 = " " Or Descr1 = "  " Then
                                                    Descr1 = "xxx"
                                                End If

                                                Descriptie_Chainage.Add(Descr1)
                                            Else
                                                Dim Partea11 As String
                                                Dim Partea11_for_chainage As String = Replace(String1, " ", "")
                                                Partea11 = extrage_chainage_din_text_de_la_sfarsitul_textului(Partea11_for_chainage)
                                                If IsNumeric(Replace(Partea11, "+", "")) = True Then
                                                    Dim Numar11 As Double = CDbl(Replace(Partea11, "+", ""))
                                                    If CheckBox_recalculate_chainage.Checked = True Then
                                                        Dim Nou_nr1 As Double
                                                        Dim Nou_partea1 As String
                                                        Nou_nr1 = Numar11 - distanta_de_la_inceputul_grafului
                                                        Nou_partea1 = Get_chainage_from_double(Nou_nr1, 1)
                                                        Valoare_Chainage.Add(Nou_nr1)
                                                        String_Chainage.Add(Nou_partea1)
                                                    Else
                                                        Valoare_Chainage.Add(Numar11)
                                                        String_Chainage.Add(Partea11)
                                                    End If
                                                    Dim Descr11 As String = String1.Replace(Partea11, "")
                                                    If Strings.Left(Descr11, 1) = " " Then Descr11 = Strings.Right(Descr11, Len(Descr11) - 1)
                                                    If Descr11 = "" Or Descr11 = " " Or Descr11 = "  " Then
                                                        Descr11 = "xxx"
                                                    End If

                                                    Descriptie_Chainage.Add(Descr11)
                                                End If


                                            End If

                                        End If ' asta e de la  If InStr(String1, "+") > 0
                                    End If ' asta e de la If TypeOf Ent33 Is Autodesk.AutoCAD.DatabaseServices.dbText

                                Next ' asta e de la  For i = 0 To Rezultat3.Value.Count - 1
                            End If ' asta e de la If Rezultat3.Value.Count = 0
                        End If ' asta e de la If Rezultat3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK








                        Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                        Dim Object_prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_prompt2.MessageForAdding = vbLf & "Select a known elevation line and the label for it:"

                        Object_prompt2.SingleOnly = False
                        Rezultat2 = Editor1.GetSelection(Object_prompt2)


                        If Rezultat2.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Editor1.SetImpliedSelection(Empty_array)
                            Exit Sub
                        End If

                        If Rezultat2.Value.Count <> 2 Then
                            Editor1.SetImpliedSelection(Empty_array)
                            Exit Sub
                        End If




                        Dim x01, y01, x02, y02, dist, a As Double

                        Dim Elevatia_cunoscuta As Double = -100000
                        Dim Elevatia_pt_pct_zero_masurata As Double = -100000
                        Dim mText_cunoscut As Autodesk.AutoCAD.DatabaseServices.MText
                        Dim Text_cunoscut As Autodesk.AutoCAD.DatabaseServices.DBText
                        Dim Linia_cunoscuta As Autodesk.AutoCAD.DatabaseServices.Line

                        Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj2 = Rezultat2.Value.Item(0)
                        Dim Ent2 As Entity
                        Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)
                        Dim Obj3 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj3 = Rezultat2.Value.Item(1)
                        Dim Ent3 As Entity
                        Ent3 = Obj3.ObjectId.GetObject(OpenMode.ForRead)

                        If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                            mText_cunoscut = Ent2
                            If IsNumeric(mText_cunoscut.Text) = True Then Elevatia_cunoscuta = CDbl(mText_cunoscut.Text)
                        End If

                        If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                            mText_cunoscut = Ent3
                            If IsNumeric(mText_cunoscut.Text) = True Then Elevatia_cunoscuta = CDbl(mText_cunoscut.Text)
                        End If

                        If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                            Text_cunoscut = Ent2
                            If IsNumeric(Text_cunoscut.TextString) = True Then Elevatia_cunoscuta = CDbl(Text_cunoscut.TextString)
                        End If

                        If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                            Text_cunoscut = Ent3
                            If IsNumeric(Text_cunoscut.TextString) = True Then Elevatia_cunoscuta = CDbl(Text_cunoscut.TextString)
                        End If

                        If Elevatia_cunoscuta = -100000 Then
                            Editor1.SetImpliedSelection(Empty_array)
                            Exit Sub
                        End If

                        If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                            Linia_cunoscuta = Ent2
                            x01 = Linia_cunoscuta.StartPoint.X
                            y01 = Linia_cunoscuta.StartPoint.Y
                            x02 = Linia_cunoscuta.EndPoint.X
                            y02 = Linia_cunoscuta.EndPoint.Y
                            If Abs(y01 - y02) > 0.001 Then
                                Editor1.SetImpliedSelection(Empty_array)
                                Exit Sub
                            End If
                            dist = ((X0 - x01) ^ 2 + (Y0 - y01) ^ 2) ^ 0.5
                            a = Abs(X0 - x01)
                            Elevatia_pt_pct_zero_masurata = (dist ^ 2 - a ^ 2) ^ 0.5

                        End If

                        If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                            Dim Poly_temp As Polyline = Ent2
                            Dim Line_temp As New Line(Poly_temp.GetPoint3dAt(0), Poly_temp.GetPoint3dAt(1))
                            Linia_cunoscuta = Line_temp
                            x01 = Linia_cunoscuta.StartPoint.X
                            y01 = Linia_cunoscuta.StartPoint.Y
                            x02 = Linia_cunoscuta.EndPoint.X
                            y02 = Linia_cunoscuta.EndPoint.Y
                            If Abs(y01 - y02) > 0.001 Then
                                Editor1.SetImpliedSelection(Empty_array)
                                Exit Sub
                            End If
                            dist = ((X0 - x01) ^ 2 + (Y0 - y01) ^ 2) ^ 0.5
                            a = Abs(X0 - x01)
                            Elevatia_pt_pct_zero_masurata = (dist ^ 2 - a ^ 2) ^ 0.5

                        End If


                        If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                            Linia_cunoscuta = Ent3
                            x01 = Linia_cunoscuta.StartPoint.X
                            y01 = Linia_cunoscuta.StartPoint.Y
                            x02 = Linia_cunoscuta.EndPoint.X
                            y02 = Linia_cunoscuta.EndPoint.Y
                            If Abs(y01 - y02) > 0.001 Then
                                Editor1.SetImpliedSelection(Empty_array)
                                Exit Sub
                            End If
                            dist = ((X0 - x01) ^ 2 + (Y0 - y01) ^ 2) ^ 0.5
                            a = Abs(X0 - x01)
                            Elevatia_pt_pct_zero_masurata = (dist ^ 2 - a ^ 2) ^ 0.5

                        End If

                        If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                            Dim Poly_temp As Polyline = Ent3
                            Dim Line_temp As New Line(Poly_temp.GetPoint3dAt(0), Poly_temp.GetPoint3dAt(1))
                            Linia_cunoscuta = Line_temp
                            x01 = Linia_cunoscuta.StartPoint.X
                            y01 = Linia_cunoscuta.StartPoint.Y
                            x02 = Linia_cunoscuta.EndPoint.X
                            y02 = Linia_cunoscuta.EndPoint.Y
                            If Abs(y01 - y02) > 0.001 Then
                                Editor1.SetImpliedSelection(Empty_array)
                                Exit Sub
                            End If
                            dist = ((X0 - x01) ^ 2 + (Y0 - y01) ^ 2) ^ 0.5
                            a = Abs(X0 - x01)
                            Elevatia_pt_pct_zero_masurata = (dist ^ 2 - a ^ 2) ^ 0.5

                        End If


                        If Elevatia_pt_pct_zero_masurata = -100000 Then
                            Editor1.SetImpliedSelection(Empty_array)
                            Exit Sub
                        End If

                        Dim H_scale As Double = CDbl(TextBox_Horizontal_scale.Text)
                        Dim V_scale As Double = CDbl(TextBox_Vertical_scale.Text)
                        Dim Pr_scale As Double = CDbl(TextBox_printing_scale.Text)
                        Dim H_scale_hmm As Double = CDbl(ComboBox_Horizontal_scale_HMM.Text)
                        Dim V_scale_hmm As Double = CDbl(ComboBox_Vertical_scale_HMM.Text)
                        Dim Pr_scale_HMM As Double = CDbl(TextBox_Printing_scale_HMM.Text)




                        Dim Point_GRID As Autodesk.AutoCAD.EditorInput.PromptPointResult

                        Dim PP_GRID As New Autodesk.AutoCAD.EditorInput.PromptPointOptions("Specify grid location:")


                        PP_GRID.AllowNone = False
                        Point_GRID = Editor1.GetPoint(PP_GRID)
                        If Point_GRID.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                            Exit Sub
                        End If

                        Dim X_grid, Y_grid As Double
                        X_grid = Point_GRID.Value.X
                        Y_grid = Point_GRID.Value.Y



                        Creaza_layer_cu_linetype_si_lineweight(layer_viewport_name, layer_viewport_color, layer_viewport_ltype, layer_viewport_lweight, layer_viewport_description, layer_viewport_plot, True)

                        Creaza_layer_cu_linetype_si_lineweight(layer_no_plot_name, layer_no_plot_color, layer_no_plot_ltype, layer_no_plot_lweight, layer_no_plot_description, layer_no_plot_plot, True)

                        Creaza_layer_cu_linetype_si_lineweight(layer_BLOCKS_name, layer_BLOCKS_color, layer_BLOCKS_ltype, layer_BLOCKS_lweight, layer_BLOCKS_description, layer_BLOCKS_plot, True)

                        Creaza_layer_cu_linetype_si_lineweight(layer_water_name, layer_water_color, layer_water_ltype, layer_water_lweight, layer_water_description, layer_water_plot, True)

                        Creaza_layer_cu_linetype_si_lineweight(layer_Text_name, layer_Text_color, layer_Text_ltype, layer_Text_lweight, layer_text_description, layer_Text_plot, True)

                        Creaza_layer_cu_linetype_si_lineweight(layer_grid_name, layer_grid_color, layer_grid_ltype, layer_grid_lweight, layer_grid_description, layer_grid_plot, True)
                        Creaza_layer_cu_linetype_si_lineweight(layer_polyline_name, layer_polyline_color, layer_polyline_ltype, layer_polyline_lweight, layer_polyline_description, layer_polyline_plot, True)


                        Dim Distanta_de_la_elevatia_cunoscuta As Double = 0
                        Dim Closest10_label As Double = 0
                        Dim Distanta_pana_la_linia_referinta As Double = 0
                        Dim Closest5_label As Double = 0
                        Dim Closest2_label As Double = 0
                        Dim Closest20_label As Double = 0
                        Dim Elevatia_punctului As Double = 0
                        Dim Closest_label As Double = 0

                        If ComboBox_Horizontal_scale_HMM.Text = 1000 And ComboBox_Vertical_scale_HMM.Text = 1000 Then
                            Dim X_zero_zero As Double = X_grid



1:
                            'DE AICI 1 LA 2000

                            Scale_inv_vwport = 2
                            Nr_linii_jos = 6

                            Viewport_CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(X_paper, y_paper, 0)
                            Create_profile_graf(Trans1, BTrecord, BTrecord_PAPER, layer_grid_name, Color_index_grey_lines, layer_Text_name, Text_style_ID, layer_polyline_name, _
                                                      layer_no_plot_name, layer_viewport_name, layer_BLOCKS_name, New Point3d(X_grid, Y_grid, 0), Lungime_graph_1_1, Scale_inv_vwport, _
                                                      Interval_hor, Interval_ver, Nr_chainage_labels, Nr_vertical_labels, Lung_liniuta_ver, Lung_liniuta_hor, _
                                                      Elevatia_cunoscuta, Elevatia_pt_pct_zero_masurata, X0, Y0, y01, V_scale, H_scale, Pr_scale, Nr_linii_jos, Colectie_puncte_graph, V_scale_hmm, H_scale_hmm, Pr_scale_HMM, _
                                                       Valoare_Chainage, Viewport_CenterPoint, Viewport_Height, Viewport_Width, Descriptie_Chainage, String_Chainage, Gap_PS_sus, Gap_PS_jos, Gap_PS_stanga_dreapta)
2:
                            ' DE AICI ESTE GRIDUL 1:1000



                            X_grid = X_grid + 0.5 * Lungime_graph_1_1 * 2 - 0.5 * Lungime_graph_1_1
                            Y_grid = Y_grid + 100 + Nr_vertical_labels * Interval_ver * Scale_inv_vwport * Pr_scale_HMM / V_scale_hmm ' scale_inV_vw*interv_vert = adevaratul interval vert - am folosit valorile de la graful de sus

                            Scale_inv_vwport = 1
                            Nr_linii_jos = 6
                            Viewport_CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(X_paper, y_paper + Viewport_Height + Vw_Spacing_Y_paper, 0)
                            Create_profile_graf(Trans1, BTrecord, BTrecord_PAPER, layer_grid_name, Color_index_grey_lines, layer_Text_name, Text_style_ID, layer_polyline_name, _
                                                                                 layer_no_plot_name, layer_viewport_name, layer_BLOCKS_name, New Point3d(X_grid, Y_grid, 0), Lungime_graph_1_1, Scale_inv_vwport, _
                                                                                 Interval_hor, Interval_ver, Nr_chainage_labels, Nr_vertical_labels, Lung_liniuta_ver, Lung_liniuta_hor, _
                                                                                 Elevatia_cunoscuta, Elevatia_pt_pct_zero_masurata, X0, Y0, y01, V_scale, H_scale, Pr_scale, Nr_linii_jos, Colectie_puncte_graph, V_scale_hmm, H_scale_hmm, Pr_scale_HMM, _
                                                                                  Valoare_Chainage, Viewport_CenterPoint, Viewport_Height, Viewport_Width, Descriptie_Chainage, String_Chainage, Gap_PS_sus, Gap_PS_jos, Gap_PS_stanga_dreapta)
3:
                            'URMEAZA 1 LA 500
                            X_grid = X_grid + 0.5 * Lungime_graph_1_1 - 0.5 * Lungime_graph_1_1 * 0.5 '186.071
                            Y_grid = Y_grid + 100 + Nr_vertical_labels * Interval_ver * Scale_inv_vwport * Pr_scale_HMM / V_scale_hmm '+ 208.227
                            Scale_inv_vwport = 0.5
                            Nr_linii_jos = 6
                            Viewport_CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(X_paper, y_paper + 2 * Viewport_Height + 2 * Vw_Spacing_Y_paper, 0)
                            Create_profile_graf(Trans1, BTrecord, BTrecord_PAPER, layer_grid_name, Color_index_grey_lines, layer_Text_name, Text_style_ID, layer_polyline_name, _
                                                       layer_no_plot_name, layer_viewport_name, layer_BLOCKS_name, New Point3d(X_grid, Y_grid, 0), Lungime_graph_1_1, Scale_inv_vwport, _
                                                       Interval_hor, Interval_ver, Nr_chainage_labels, Nr_vertical_labels, Lung_liniuta_ver, Lung_liniuta_hor, _
                                                       Elevatia_cunoscuta, Elevatia_pt_pct_zero_masurata, X0, Y0, y01, V_scale, H_scale, Pr_scale, Nr_linii_jos, Colectie_puncte_graph, V_scale_hmm, H_scale_hmm, Pr_scale_HMM, _
                                                        Valoare_Chainage, Viewport_CenterPoint, Viewport_Height, Viewport_Width, Descriptie_Chainage, String_Chainage, Gap_PS_sus, Gap_PS_jos, Gap_PS_stanga_dreapta)

4:
                            'URMEAZA 1 LA 200
                            X_grid = X_grid + 0.5 * Lungime_graph_1_1 * 0.5 - 0.5 * Lungime_graph_1_1 * 0.2 '111.643
                            Y_grid = Y_grid + 100 + Nr_vertical_labels * Interval_ver * Scale_inv_vwport * Pr_scale_HMM / V_scale_hmm '+ 133.259
                            Scale_inv_vwport = 0.2
                            Nr_linii_jos = 6
                            Viewport_CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(X_paper, y_paper + 3 * Viewport_Height + 3 * Vw_Spacing_Y_paper, 0)
                            Create_profile_graf(Trans1, BTrecord, BTrecord_PAPER, layer_grid_name, Color_index_grey_lines, layer_Text_name, Text_style_ID, layer_polyline_name, _
                                                       layer_no_plot_name, layer_viewport_name, layer_BLOCKS_name, New Point3d(X_grid, Y_grid, 0), Lungime_graph_1_1, Scale_inv_vwport, _
                                                       Interval_hor, Interval_ver, Nr_chainage_labels, Nr_vertical_labels, Lung_liniuta_ver, Lung_liniuta_hor, _
                                                       Elevatia_cunoscuta, Elevatia_pt_pct_zero_masurata, X0, Y0, y01, V_scale, H_scale, Pr_scale, Nr_linii_jos, Colectie_puncte_graph, V_scale_hmm, H_scale_hmm, Pr_scale_HMM, _
                                                        Valoare_Chainage, Viewport_CenterPoint, Viewport_Height, Viewport_Width, Descriptie_Chainage, String_Chainage, Gap_PS_sus, Gap_PS_jos, Gap_PS_stanga_dreapta)
11:
                            'DE AICI 1 LA 2000 shifted 2 LINES UP
                            X_grid = X_zero_zero + Lungime_graph_1_1 * 2 + 200
                            Y_grid = Y_grid - (100 + Nr_vertical_labels * Interval_ver * 0.5 * Pr_scale_HMM / V_scale_hmm) - (100 + Nr_vertical_labels * Interval_ver * 1 * Pr_scale_HMM / V_scale_hmm) - (100 + Nr_vertical_labels * Interval_ver * 2 * Pr_scale_HMM / V_scale_hmm) '133.259 - 208.227 - 428.514
                            Scale_inv_vwport = 2
                            Nr_linii_jos = 8
                            Viewport_CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(X_paper, y_paper + 4 * Viewport_Height + 4 * Vw_Spacing_Y_paper, 0)
                            Create_profile_graf(Trans1, BTrecord, BTrecord_PAPER, layer_grid_name, Color_index_grey_lines, layer_Text_name, Text_style_ID, layer_polyline_name, _
                                                        layer_no_plot_name, layer_viewport_name, layer_BLOCKS_name, New Point3d(X_grid, Y_grid, 0), Lungime_graph_1_1, Scale_inv_vwport, _
                                                        Interval_hor, Interval_ver, Nr_chainage_labels, Nr_vertical_labels, Lung_liniuta_ver, Lung_liniuta_hor, _
                                                        Elevatia_cunoscuta, Elevatia_pt_pct_zero_masurata, X0, Y0, y01, V_scale, H_scale, Pr_scale, Nr_linii_jos, Colectie_puncte_graph, V_scale_hmm, H_scale_hmm, Pr_scale_HMM, _
                                                         Valoare_Chainage, Viewport_CenterPoint, Viewport_Height, Viewport_Width, Descriptie_Chainage, String_Chainage, Gap_PS_sus, Gap_PS_jos, Gap_PS_stanga_dreapta)
12:
                            '1 LA 1000
                            'SHIFTED GRIDS UP BY 2 GRID LINES
                            X_grid = X_grid + 0.5 * Lungime_graph_1_1 * 2 - 0.5 * Lungime_graph_1_1 '372.142
                            Y_grid = Y_grid + 100 + Nr_vertical_labels * Interval_ver * Scale_inv_vwport * Pr_scale_HMM / V_scale_hmm
                            Scale_inv_vwport = 1
                            Nr_linii_jos = 8
                            Viewport_CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(X_paper, y_paper + 5 * Viewport_Height + 5 * Vw_Spacing_Y_paper, 0)
                            Create_profile_graf(Trans1, BTrecord, BTrecord_PAPER, layer_grid_name, Color_index_grey_lines, layer_Text_name, Text_style_ID, layer_polyline_name, _
                                                       layer_no_plot_name, layer_viewport_name, layer_BLOCKS_name, New Point3d(X_grid, Y_grid, 0), Lungime_graph_1_1, Scale_inv_vwport, _
                                                       Interval_hor, Interval_ver, Nr_chainage_labels, Nr_vertical_labels, Lung_liniuta_ver, Lung_liniuta_hor, _
                                                       Elevatia_cunoscuta, Elevatia_pt_pct_zero_masurata, X0, Y0, y01, V_scale, H_scale, Pr_scale, Nr_linii_jos, Colectie_puncte_graph, V_scale_hmm, H_scale_hmm, Pr_scale_HMM, _
                                                        Valoare_Chainage, Viewport_CenterPoint, Viewport_Height, Viewport_Width, Descriptie_Chainage, String_Chainage, Gap_PS_sus, Gap_PS_jos, Gap_PS_stanga_dreapta)

13:
                            'URMEAZA 1 LA 500 shift SHIFT 2 GRID LINES UP
                            X_grid = X_grid + 0.5 * Lungime_graph_1_1 - 0.5 * Lungime_graph_1_1 * 0.5
                            Y_grid = Y_grid + 100 + Nr_vertical_labels * Interval_ver * Scale_inv_vwport * Pr_scale_HMM / V_scale_hmm
                            Scale_inv_vwport = 0.5
                            Nr_linii_jos = 8
                            Viewport_CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(X_paper, y_paper + 6 * Viewport_Height + 6 * Vw_Spacing_Y_paper, 0)
                            Create_profile_graf(Trans1, BTrecord, BTrecord_PAPER, layer_grid_name, Color_index_grey_lines, layer_Text_name, Text_style_ID, layer_polyline_name, _
                                                      layer_no_plot_name, layer_viewport_name, layer_BLOCKS_name, New Point3d(X_grid, Y_grid, 0), Lungime_graph_1_1, Scale_inv_vwport, _
                                                      Interval_hor, Interval_ver, Nr_chainage_labels, Nr_vertical_labels, Lung_liniuta_ver, Lung_liniuta_hor, _
                                                      Elevatia_cunoscuta, Elevatia_pt_pct_zero_masurata, X0, Y0, y01, V_scale, H_scale, Pr_scale, Nr_linii_jos, Colectie_puncte_graph, V_scale_hmm, H_scale_hmm, Pr_scale_HMM, _
                                                       Valoare_Chainage, Viewport_CenterPoint, Viewport_Height, Viewport_Width, Descriptie_Chainage, String_Chainage, Gap_PS_sus, Gap_PS_jos, Gap_PS_stanga_dreapta)
14:

                            'URMEAZA 1 LA 200 shift SHIFT 2 LINES UP
                            X_grid = X_grid + 0.5 * Lungime_graph_1_1 * 0.5 - 0.5 * Lungime_graph_1_1 * 0.2
                            Y_grid = Y_grid + 100 + Nr_vertical_labels * Interval_ver * Scale_inv_vwport * Pr_scale_HMM / V_scale_hmm
                            Scale_inv_vwport = 0.2
                            Nr_linii_jos = 8
                            Viewport_CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(X_paper, y_paper + 7 * Viewport_Height + 7 * Vw_Spacing_Y_paper, 0)
                            Create_profile_graf(Trans1, BTrecord, BTrecord_PAPER, layer_grid_name, Color_index_grey_lines, layer_Text_name, Text_style_ID, layer_polyline_name, _
                                                       layer_no_plot_name, layer_viewport_name, layer_BLOCKS_name, New Point3d(X_grid, Y_grid, 0), Lungime_graph_1_1, Scale_inv_vwport, _
                                                       Interval_hor, Interval_ver, Nr_chainage_labels, Nr_vertical_labels, Lung_liniuta_ver, Lung_liniuta_hor, _
                                                       Elevatia_cunoscuta, Elevatia_pt_pct_zero_masurata, X0, Y0, y01, V_scale, H_scale, Pr_scale, Nr_linii_jos, Colectie_puncte_graph, V_scale_hmm, H_scale_hmm, Pr_scale_HMM, _
                                                        Valoare_Chainage, Viewport_CenterPoint, Viewport_Height, Viewport_Width, Descriptie_Chainage, String_Chainage, Gap_PS_sus, Gap_PS_jos, Gap_PS_stanga_dreapta)
                            ' asta e de la If Este_1_la_1 = True And CheckBox_1_1_exag_10x.Checked = False
5:
                            'URMEAZA 1 LA 2000 hor 1:200 ver
                        ElseIf ComboBox_Vertical_scale_HMM.Text = 200 And ComboBox_Horizontal_scale_HMM.Text = 2000 Then

                            Lung_liniuta_hor = 5
                            Lung_liniuta_ver = 2.5
                            Scale_inv_vwport = 1
                            Nr_chainage_labels = Floor(CDbl(TextBox_PROFILE_LENGTH.Text) / CDbl(TextBox_horizontal_increment.Text))
                            Nr_linii_jos = 3
                            Gap_PS_sus = 20
                            Gap_PS_jos = 25
                            Gap_PS_stanga_dreapta = 25
                            Viewport_Height = Nr_vertical_labels * Interval_ver * Pr_scale / V_scale_hmm + Gap_PS_jos + Gap_PS_sus
                            Viewport_Width = Lungime_graph_1_1 * Pr_scale_HMM / H_scale_hmm + Gap_PS_stanga_dreapta * 2
                            X_paper = -1400
                            y_paper = 250
                            Vw_Spacing_X_paper = 0
                            Vw_Spacing_Y_paper = 50
                            Viewport_CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(X_paper, y_paper, 0)
                            Create_profile_graf(Trans1, BTrecord, BTrecord_PAPER, layer_grid_name, Color_index_grey_lines, layer_Text_name, Text_style_ID, layer_polyline_name, _
                                                       layer_no_plot_name, layer_viewport_name, layer_BLOCKS_name, New Point3d(X_grid, Y_grid, 0), Lungime_graph_1_1, Scale_inv_vwport, _
                                                       Interval_hor, Interval_ver, Nr_chainage_labels, Nr_vertical_labels, Lung_liniuta_ver, Lung_liniuta_hor, _
                                                       Elevatia_cunoscuta, Elevatia_pt_pct_zero_masurata, X0, Y0, y01, V_scale, H_scale, Pr_scale, Nr_linii_jos, Colectie_puncte_graph, V_scale_hmm, H_scale_hmm, Pr_scale_HMM, _
                                                        Valoare_Chainage, Viewport_CenterPoint, Viewport_Height, Viewport_Width, Descriptie_Chainage, String_Chainage, Gap_PS_sus, Gap_PS_jos, Gap_PS_stanga_dreapta)


6:

                            'URMEAZA 1 LA 2000 hor 1:200 vertical alt grid shifted up
                            Y_grid = Y_grid + 160
                            Nr_linii_jos = 4
                            Viewport_CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(X_paper, y_paper + Viewport_Height + Vw_Spacing_Y_paper, 0)
                            Create_profile_graf(Trans1, BTrecord, BTrecord_PAPER, layer_grid_name, Color_index_grey_lines, layer_Text_name, Text_style_ID, layer_polyline_name, _
                                                       layer_no_plot_name, layer_viewport_name, layer_BLOCKS_name, New Point3d(X_grid, Y_grid, 0), Lungime_graph_1_1, Scale_inv_vwport, _
                                                       Interval_hor, Interval_ver, Nr_chainage_labels, Nr_vertical_labels, Lung_liniuta_ver, Lung_liniuta_hor, _
                                                       Elevatia_cunoscuta, Elevatia_pt_pct_zero_masurata, X0, Y0, y01, V_scale, H_scale, Pr_scale, Nr_linii_jos, Colectie_puncte_graph, V_scale_hmm, H_scale_hmm, Pr_scale_HMM, _
                                                        Valoare_Chainage, Viewport_CenterPoint, Viewport_Height, Viewport_Width, Descriptie_Chainage, String_Chainage, Gap_PS_sus, Gap_PS_jos, Gap_PS_stanga_dreapta)


7:

                            'URMEAZA 1 LA 2000 hor 1:200 vertical alt grid shifted up 2LINES
                            Y_grid = Y_grid + 160
                            Nr_linii_jos = 5
                            Viewport_CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(X_paper, y_paper + 2 * Viewport_Height + 2 * Vw_Spacing_Y_paper, 0)
                            Create_profile_graf(Trans1, BTrecord, BTrecord_PAPER, layer_grid_name, Color_index_grey_lines, layer_Text_name, Text_style_ID, layer_polyline_name, _
                                                       layer_no_plot_name, layer_viewport_name, layer_BLOCKS_name, New Point3d(X_grid, Y_grid, 0), Lungime_graph_1_1, Scale_inv_vwport, _
                                                       Interval_hor, Interval_ver, Nr_chainage_labels, Nr_vertical_labels, Lung_liniuta_ver, Lung_liniuta_hor, _
                                                       Elevatia_cunoscuta, Elevatia_pt_pct_zero_masurata, X0, Y0, y01, V_scale, H_scale, Pr_scale, Nr_linii_jos, Colectie_puncte_graph, V_scale_hmm, H_scale_hmm, Pr_scale_HMM, _
                                                        Valoare_Chainage, Viewport_CenterPoint, Viewport_Height, Viewport_Width, Descriptie_Chainage, String_Chainage, Gap_PS_sus, Gap_PS_jos, Gap_PS_stanga_dreapta)

8:

                            'URMEAZA 1 LA 2000 hor 1:200 ver 400 long
                            Y_grid = Y_grid + 160
                            Nr_linii_jos = 3
                            Lungime_graph_1_1 = Lungime_graph_1_1 + 350
                            Nr_chainage_labels = Floor(CDbl(TextBox_PROFILE_LENGTH.Text) / CDbl(TextBox_horizontal_increment.Text))
                            Viewport_Width = Lungime_graph_1_1 * Pr_scale_HMM / H_scale_hmm + Gap_PS_stanga_dreapta * 2
                            Viewport_CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(X_paper, y_paper + 3 * Viewport_Height + 3 * Vw_Spacing_Y_paper, 0)
                            Create_profile_graf(Trans1, BTrecord, BTrecord_PAPER, layer_grid_name, Color_index_grey_lines, layer_Text_name, Text_style_ID, layer_polyline_name, _
                                                       layer_no_plot_name, layer_viewport_name, layer_BLOCKS_name, New Point3d(X_grid, Y_grid, 0), Lungime_graph_1_1, Scale_inv_vwport, _
                                                       Interval_hor, Interval_ver, Nr_chainage_labels, Nr_vertical_labels, Lung_liniuta_ver, Lung_liniuta_hor, _
                                                       Elevatia_cunoscuta, Elevatia_pt_pct_zero_masurata, X0, Y0, y01, V_scale, H_scale, Pr_scale, Nr_linii_jos, Colectie_puncte_graph, V_scale_hmm, H_scale_hmm, Pr_scale_HMM, _
                                                        Valoare_Chainage, Viewport_CenterPoint, Viewport_Height, Viewport_Width, Descriptie_Chainage, String_Chainage, Gap_PS_sus, Gap_PS_jos, Gap_PS_stanga_dreapta)

                            'pana aici este 1 LA 2000 hor 1:200 vertical 400 long each side


9:

                            'URMEAZA 1 LA 2000 hor 1:200 ver 400 long 1 line up
                            Y_grid = Y_grid + 160
                            Nr_linii_jos = 4
                            Viewport_CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(X_paper, y_paper + 4 * Viewport_Height + 4 * Vw_Spacing_Y_paper, 0)
                            Create_profile_graf(Trans1, BTrecord, BTrecord_PAPER, layer_grid_name, Color_index_grey_lines, layer_Text_name, Text_style_ID, layer_polyline_name, _
                                                       layer_no_plot_name, layer_viewport_name, layer_BLOCKS_name, New Point3d(X_grid, Y_grid, 0), Lungime_graph_1_1, Scale_inv_vwport, _
                                                       Interval_hor, Interval_ver, Nr_chainage_labels, Nr_vertical_labels, Lung_liniuta_ver, Lung_liniuta_hor, _
                                                       Elevatia_cunoscuta, Elevatia_pt_pct_zero_masurata, X0, Y0, y01, V_scale, H_scale, Pr_scale, Nr_linii_jos, Colectie_puncte_graph, V_scale_hmm, H_scale_hmm, Pr_scale_HMM, _
                                                        Valoare_Chainage, Viewport_CenterPoint, Viewport_Height, Viewport_Width, Descriptie_Chainage, String_Chainage, Gap_PS_sus, Gap_PS_jos, Gap_PS_stanga_dreapta)
10:


                            'URMEAZA 1 LA 2000 hor 1:200 ver 400 long 2 line up
                            Y_grid = Y_grid + 160
                            Nr_linii_jos = 5
                            Viewport_CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(X_paper, y_paper + 5 * Viewport_Height + 5 * Vw_Spacing_Y_paper, 0)
                            Create_profile_graf(Trans1, BTrecord, BTrecord_PAPER, layer_grid_name, Color_index_grey_lines, layer_Text_name, Text_style_ID, layer_polyline_name, _
                                                       layer_no_plot_name, layer_viewport_name, layer_BLOCKS_name, New Point3d(X_grid, Y_grid, 0), Lungime_graph_1_1, Scale_inv_vwport, _
                                                       Interval_hor, Interval_ver, Nr_chainage_labels, Nr_vertical_labels, Lung_liniuta_ver, Lung_liniuta_hor, _
                                                       Elevatia_cunoscuta, Elevatia_pt_pct_zero_masurata, X0, Y0, y01, V_scale, H_scale, Pr_scale, Nr_linii_jos, Colectie_puncte_graph, V_scale_hmm, H_scale_hmm, Pr_scale_HMM, _
                                                        Valoare_Chainage, Viewport_CenterPoint, Viewport_Height, Viewport_Width, Descriptie_Chainage, String_Chainage, Gap_PS_sus, Gap_PS_jos, Gap_PS_stanga_dreapta)


                            ' asta e de la If Este_1_la_1 = false And CheckBox_1_1_exag_10x.Checked = False
                        ElseIf ComboBox_Vertical_scale_HMM.Text = 100 And ComboBox_Horizontal_scale_HMM.Text = 1000 Then
20:
                            ' DE AICI ESTE GRIDUL 1:1000 h 1:100 vertical water
                            Lung_liniuta_hor = 5
                            Lung_liniuta_ver = 2.5
                            Scale_inv_vwport = 1
                            Nr_chainage_labels = Floor(CDbl(TextBox_PROFILE_LENGTH.Text) / CDbl(TextBox_horizontal_increment.Text))
                            Nr_linii_jos = 3
                            Gap_PS_sus = 20
                            Gap_PS_jos = 25
                            Gap_PS_stanga_dreapta = 25
                            Viewport_Height = Nr_vertical_labels * Interval_ver * Pr_scale / V_scale_hmm + Gap_PS_jos + Gap_PS_sus
                            Viewport_Width = Lungime_graph_1_1 * Pr_scale_HMM / H_scale_hmm + Gap_PS_stanga_dreapta * 2
                            X_paper = -2000
                            y_paper = 250
                            Vw_Spacing_X_paper = 0
                            Vw_Spacing_Y_paper = 50
                            Viewport_CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(X_paper, y_paper, 0)

                            Create_profile_graf(Trans1, BTrecord, BTrecord_PAPER, layer_grid_name, Color_index_grey_lines, layer_Text_name, Text_style_ID, layer_polyline_name, _
                                                       layer_no_plot_name, layer_viewport_name, layer_BLOCKS_name, New Point3d(X_grid, Y_grid, 0), Lungime_graph_1_1, Scale_inv_vwport, _
                                                       Interval_hor, Interval_ver, Nr_chainage_labels, Nr_vertical_labels, Lung_liniuta_ver, Lung_liniuta_hor, _
                                                       Elevatia_cunoscuta, Elevatia_pt_pct_zero_masurata, X0, Y0, y01, V_scale, H_scale, Pr_scale, Nr_linii_jos, Colectie_puncte_graph, V_scale_hmm, H_scale_hmm, Pr_scale_HMM, _
                                                        Valoare_Chainage, Viewport_CenterPoint, Viewport_Height, Viewport_Width, Descriptie_Chainage, String_Chainage, Gap_PS_sus, Gap_PS_jos, Gap_PS_stanga_dreapta)

                        ElseIf ComboBox_Vertical_scale_HMM.Text = 500 And ComboBox_Horizontal_scale_HMM.Text = 1000 Then
30:

                            Lung_liniuta_hor = 5
                            Lung_liniuta_ver = 2.5
                            Scale_inv_vwport = 1
                            Nr_chainage_labels = Floor(CDbl(TextBox_PROFILE_LENGTH.Text) / CDbl(TextBox_horizontal_increment.Text))
                            Nr_linii_jos = 3
                            Gap_PS_sus = 20
                            Gap_PS_jos = 25
                            Gap_PS_stanga_dreapta = 25
                            Viewport_Height = Nr_vertical_labels * Interval_ver * Pr_scale / V_scale_hmm + Gap_PS_jos + Gap_PS_sus
                            Viewport_Width = Lungime_graph_1_1 * Pr_scale_HMM / H_scale_hmm + Gap_PS_stanga_dreapta * 2
                            X_paper = -2000
                            y_paper = 250
                            Vw_Spacing_X_paper = 0
                            Vw_Spacing_Y_paper = 50
                            Viewport_CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(X_paper, y_paper, 0)

                            Create_profile_graf(Trans1, BTrecord, BTrecord_PAPER, layer_grid_name, Color_index_grey_lines, layer_Text_name, Text_style_ID, layer_polyline_name, _
                                                       layer_no_plot_name, layer_viewport_name, layer_BLOCKS_name, New Point3d(X_grid, Y_grid, 0), Lungime_graph_1_1, Scale_inv_vwport, _
                                                       Interval_hor, Interval_ver, Nr_chainage_labels, Nr_vertical_labels, Lung_liniuta_ver, Lung_liniuta_hor, _
                                                       Elevatia_cunoscuta, Elevatia_pt_pct_zero_masurata, X0, Y0, y01, V_scale, H_scale, Pr_scale, Nr_linii_jos, Colectie_puncte_graph, V_scale_hmm, H_scale_hmm, Pr_scale_HMM, _
                                                        Valoare_Chainage, Viewport_CenterPoint, Viewport_Height, Viewport_Width, Descriptie_Chainage, String_Chainage, Gap_PS_sus, Gap_PS_jos, Gap_PS_stanga_dreapta)



                            Nr_linii_jos = 5
                            Y_grid = Y_grid + 200 + Nr_vertical_labels * Interval_ver * Scale_inv_vwport * Pr_scale_HMM / V_scale_hmm

                            Viewport_CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(X_paper, y_paper + Viewport_Height + Vw_Spacing_Y_paper, 0)

                            Create_profile_graf(Trans1, BTrecord, BTrecord_PAPER, layer_grid_name, Color_index_grey_lines, layer_Text_name, Text_style_ID, layer_polyline_name, _
                                                       layer_no_plot_name, layer_viewport_name, layer_BLOCKS_name, New Point3d(X_grid, Y_grid, 0), Lungime_graph_1_1, Scale_inv_vwport, _
                                                       Interval_hor, Interval_ver, Nr_chainage_labels, Nr_vertical_labels, Lung_liniuta_ver, Lung_liniuta_hor, _
                                                       Elevatia_cunoscuta, Elevatia_pt_pct_zero_masurata, X0, Y0, y01, V_scale, H_scale, Pr_scale, Nr_linii_jos, Colectie_puncte_graph, V_scale_hmm, H_scale_hmm, Pr_scale_HMM, _
                                                        Valoare_Chainage, Viewport_CenterPoint, Viewport_Height, Viewport_Width, Descriptie_Chainage, String_Chainage, Gap_PS_sus, Gap_PS_jos, Gap_PS_stanga_dreapta)


                            Nr_linii_jos = 7
                            Y_grid = Y_grid + 200 + Nr_vertical_labels * Interval_ver * Scale_inv_vwport * Pr_scale_HMM / V_scale_hmm

                            Viewport_CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(X_paper, y_paper + 2 * Viewport_Height + 2 * Vw_Spacing_Y_paper, 0)

                            Create_profile_graf(Trans1, BTrecord, BTrecord_PAPER, layer_grid_name, Color_index_grey_lines, layer_Text_name, Text_style_ID, layer_polyline_name, _
                                                       layer_no_plot_name, layer_viewport_name, layer_BLOCKS_name, New Point3d(X_grid, Y_grid, 0), Lungime_graph_1_1, Scale_inv_vwport, _
                                                       Interval_hor, Interval_ver, Nr_chainage_labels, Nr_vertical_labels, Lung_liniuta_ver, Lung_liniuta_hor, _
                                                       Elevatia_cunoscuta, Elevatia_pt_pct_zero_masurata, X0, Y0, y01, V_scale, H_scale, Pr_scale, Nr_linii_jos, Colectie_puncte_graph, V_scale_hmm, H_scale_hmm, Pr_scale_HMM, _
                                                        Valoare_Chainage, Viewport_CenterPoint, Viewport_Height, Viewport_Width, Descriptie_Chainage, String_Chainage, Gap_PS_sus, Gap_PS_jos, Gap_PS_stanga_dreapta)



                            Nr_linii_jos = 9
                            Y_grid = Y_grid + 200 + Nr_vertical_labels * Interval_ver * Scale_inv_vwport * Pr_scale_HMM / V_scale_hmm

                            Viewport_CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(X_paper, y_paper + 3 * Viewport_Height + 3 * Vw_Spacing_Y_paper, 0)

                            Create_profile_graf(Trans1, BTrecord, BTrecord_PAPER, layer_grid_name, Color_index_grey_lines, layer_Text_name, Text_style_ID, layer_polyline_name, _
                                                       layer_no_plot_name, layer_viewport_name, layer_BLOCKS_name, New Point3d(X_grid, Y_grid, 0), Lungime_graph_1_1, Scale_inv_vwport, _
                                                       Interval_hor, Interval_ver, Nr_chainage_labels, Nr_vertical_labels, Lung_liniuta_ver, Lung_liniuta_hor, _
                                                       Elevatia_cunoscuta, Elevatia_pt_pct_zero_masurata, X0, Y0, y01, V_scale, H_scale, Pr_scale, Nr_linii_jos, Colectie_puncte_graph, V_scale_hmm, H_scale_hmm, Pr_scale_HMM, _
                                                        Valoare_Chainage, Viewport_CenterPoint, Viewport_Height, Viewport_Width, Descriptie_Chainage, String_Chainage, Gap_PS_sus, Gap_PS_jos, Gap_PS_stanga_dreapta)


                            Nr_linii_jos = 13
                            Y_grid = Y_grid + 200 + Nr_vertical_labels * Interval_ver * Scale_inv_vwport * Pr_scale_HMM / V_scale_hmm

                            Viewport_CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(X_paper, y_paper + 4 * Viewport_Height + 4 * Vw_Spacing_Y_paper, 0)

                            Create_profile_graf(Trans1, BTrecord, BTrecord_PAPER, layer_grid_name, Color_index_grey_lines, layer_Text_name, Text_style_ID, layer_polyline_name, _
                                                       layer_no_plot_name, layer_viewport_name, layer_BLOCKS_name, New Point3d(X_grid, Y_grid, 0), Lungime_graph_1_1, Scale_inv_vwport, _
                                                       Interval_hor, Interval_ver, Nr_chainage_labels, Nr_vertical_labels, Lung_liniuta_ver, Lung_liniuta_hor, _
                                                       Elevatia_cunoscuta, Elevatia_pt_pct_zero_masurata, X0, Y0, y01, V_scale, H_scale, Pr_scale, Nr_linii_jos, Colectie_puncte_graph, V_scale_hmm, H_scale_hmm, Pr_scale_HMM, _
                                                        Valoare_Chainage, Viewport_CenterPoint, Viewport_Height, Viewport_Width, Descriptie_Chainage, String_Chainage, Gap_PS_sus, Gap_PS_jos, Gap_PS_stanga_dreapta)



                            Nr_linii_jos = 17
                            Y_grid = Y_grid + 200 + Nr_vertical_labels * Interval_ver * Scale_inv_vwport * Pr_scale_HMM / V_scale_hmm

                            Viewport_CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(X_paper, y_paper + 5 * Viewport_Height + 5 * Vw_Spacing_Y_paper, 0)

                            Create_profile_graf(Trans1, BTrecord, BTrecord_PAPER, layer_grid_name, Color_index_grey_lines, layer_Text_name, Text_style_ID, layer_polyline_name, _
                                                       layer_no_plot_name, layer_viewport_name, layer_BLOCKS_name, New Point3d(X_grid, Y_grid, 0), Lungime_graph_1_1, Scale_inv_vwport, _
                                                       Interval_hor, Interval_ver, Nr_chainage_labels, Nr_vertical_labels, Lung_liniuta_ver, Lung_liniuta_hor, _
                                                       Elevatia_cunoscuta, Elevatia_pt_pct_zero_masurata, X0, Y0, y01, V_scale, H_scale, Pr_scale, Nr_linii_jos, Colectie_puncte_graph, V_scale_hmm, H_scale_hmm, Pr_scale_HMM, _
                                                        Valoare_Chainage, Viewport_CenterPoint, Viewport_Height, Viewport_Width, Descriptie_Chainage, String_Chainage, Gap_PS_sus, Gap_PS_jos, Gap_PS_stanga_dreapta)


                        ElseIf ComboBox_Vertical_scale_HMM.Text = 1000 And ComboBox_Horizontal_scale_HMM.Text = 2000 Then
31:

                            Lung_liniuta_hor = 5
                            Lung_liniuta_ver = 2.5
                            Scale_inv_vwport = 1
                            Nr_chainage_labels = Floor(CDbl(TextBox_PROFILE_LENGTH.Text) / CDbl(TextBox_horizontal_increment.Text))
                            Nr_linii_jos = 3
                            Gap_PS_sus = 20
                            Gap_PS_jos = 25
                            Gap_PS_stanga_dreapta = 25
                            Viewport_Height = Nr_vertical_labels * Interval_ver * Pr_scale / V_scale_hmm + Gap_PS_jos + Gap_PS_sus
                            Viewport_Width = Lungime_graph_1_1 * Pr_scale_HMM / H_scale_hmm + Gap_PS_stanga_dreapta * 2
                            X_paper = -2000
                            y_paper = 250
                            Vw_Spacing_X_paper = 0
                            Vw_Spacing_Y_paper = 50
                            Viewport_CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(X_paper, y_paper, 0)

                            Create_profile_graf(Trans1, BTrecord, BTrecord_PAPER, layer_grid_name, Color_index_grey_lines, layer_Text_name, Text_style_ID, layer_polyline_name, _
                                                       layer_no_plot_name, layer_viewport_name, layer_BLOCKS_name, New Point3d(X_grid, Y_grid, 0), Lungime_graph_1_1, Scale_inv_vwport, _
                                                       Interval_hor, Interval_ver, Nr_chainage_labels, Nr_vertical_labels, Lung_liniuta_ver, Lung_liniuta_hor, _
                                                       Elevatia_cunoscuta, Elevatia_pt_pct_zero_masurata, X0, Y0, y01, V_scale, H_scale, Pr_scale, Nr_linii_jos, Colectie_puncte_graph, V_scale_hmm, H_scale_hmm, Pr_scale_HMM, _
                                                        Valoare_Chainage, Viewport_CenterPoint, Viewport_Height, Viewport_Width, Descriptie_Chainage, String_Chainage, Gap_PS_sus, Gap_PS_jos, Gap_PS_stanga_dreapta)



                            Nr_linii_jos = 5
                            Y_grid = Y_grid + 200 + Nr_vertical_labels * Interval_ver * Scale_inv_vwport * Pr_scale_HMM / V_scale_hmm

                            Viewport_CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(X_paper, y_paper + Viewport_Height + Vw_Spacing_Y_paper, 0)

                            Create_profile_graf(Trans1, BTrecord, BTrecord_PAPER, layer_grid_name, Color_index_grey_lines, layer_Text_name, Text_style_ID, layer_polyline_name, _
                                                       layer_no_plot_name, layer_viewport_name, layer_BLOCKS_name, New Point3d(X_grid, Y_grid, 0), Lungime_graph_1_1, Scale_inv_vwport, _
                                                       Interval_hor, Interval_ver, Nr_chainage_labels, Nr_vertical_labels, Lung_liniuta_ver, Lung_liniuta_hor, _
                                                       Elevatia_cunoscuta, Elevatia_pt_pct_zero_masurata, X0, Y0, y01, V_scale, H_scale, Pr_scale, Nr_linii_jos, Colectie_puncte_graph, V_scale_hmm, H_scale_hmm, Pr_scale_HMM, _
                                                        Valoare_Chainage, Viewport_CenterPoint, Viewport_Height, Viewport_Width, Descriptie_Chainage, String_Chainage, Gap_PS_sus, Gap_PS_jos, Gap_PS_stanga_dreapta)

                            Nr_linii_jos = 7
                            Y_grid = Y_grid + 200 + Nr_vertical_labels * Interval_ver * Scale_inv_vwport * Pr_scale_HMM / V_scale_hmm

                            Viewport_CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(X_paper, y_paper + 2 * Viewport_Height + 2 * Vw_Spacing_Y_paper, 0)

                            Create_profile_graf(Trans1, BTrecord, BTrecord_PAPER, layer_grid_name, Color_index_grey_lines, layer_Text_name, Text_style_ID, layer_polyline_name, _
                                                       layer_no_plot_name, layer_viewport_name, layer_BLOCKS_name, New Point3d(X_grid, Y_grid, 0), Lungime_graph_1_1, Scale_inv_vwport, _
                                                       Interval_hor, Interval_ver, Nr_chainage_labels, Nr_vertical_labels, Lung_liniuta_ver, Lung_liniuta_hor, _
                                                       Elevatia_cunoscuta, Elevatia_pt_pct_zero_masurata, X0, Y0, y01, V_scale, H_scale, Pr_scale, Nr_linii_jos, Colectie_puncte_graph, V_scale_hmm, H_scale_hmm, Pr_scale_HMM, _
                                                        Valoare_Chainage, Viewport_CenterPoint, Viewport_Height, Viewport_Width, Descriptie_Chainage, String_Chainage, Gap_PS_sus, Gap_PS_jos, Gap_PS_stanga_dreapta)


                            Nr_linii_jos = 9
                            Y_grid = Y_grid + 200 + Nr_vertical_labels * Interval_ver * Scale_inv_vwport * Pr_scale_HMM / V_scale_hmm

                            Viewport_CenterPoint = New Autodesk.AutoCAD.Geometry.Point3d(X_paper, y_paper + 3 * Viewport_Height + 3 * Vw_Spacing_Y_paper, 0)

                            Create_profile_graf(Trans1, BTrecord, BTrecord_PAPER, layer_grid_name, Color_index_grey_lines, layer_Text_name, Text_style_ID, layer_polyline_name, _
                                                       layer_no_plot_name, layer_viewport_name, layer_BLOCKS_name, New Point3d(X_grid, Y_grid, 0), Lungime_graph_1_1, Scale_inv_vwport, _
                                                       Interval_hor, Interval_ver, Nr_chainage_labels, Nr_vertical_labels, Lung_liniuta_ver, Lung_liniuta_hor, _
                                                       Elevatia_cunoscuta, Elevatia_pt_pct_zero_masurata, X0, Y0, y01, V_scale, H_scale, Pr_scale, Nr_linii_jos, Colectie_puncte_graph, V_scale_hmm, H_scale_hmm, Pr_scale_HMM, _
                                                        Valoare_Chainage, Viewport_CenterPoint, Viewport_Height, Viewport_Width, Descriptie_Chainage, String_Chainage, Gap_PS_sus, Gap_PS_jos, Gap_PS_stanga_dreapta)

                        End If ' asta e de la scales displayed
                        Trans1.Commit()
                    End If ' asta e de la If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline
                    Editor1.Regen()
                End Using
                Editor1.SetImpliedSelection(Empty_array)
            End Using ' ASTA E DE LA Using Lock1
        Catch ex As Exception

            MsgBox(ex.Message)
            Editor1.SetImpliedSelection(Empty_array)
        End Try
    End Sub










    Private Sub ComboBox_Vertical_scale_HMM_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_Vertical_scale_HMM.SelectedIndexChanged
        Try
            If ComboBox_Vertical_scale_HMM.Text = "1000" Then
                ComboBox_Horizontal_scale_HMM.Text = 1000
                TextBox_PROFILE_LENGTH.Text = 750
                TextBox_horizontal_increment.Text = 100
                TextBox_vertical_increment.Text = 10
                TextBox_nr_vert_labels.Text = 11
            End If

            If ComboBox_Vertical_scale_HMM.Text = "200" Then
                ComboBox_Horizontal_scale_HMM.Text = 2000
                TextBox_PROFILE_LENGTH.Text = 525
                TextBox_horizontal_increment.Text = 100
                TextBox_vertical_increment.Text = 2
                TextBox_nr_vert_labels.Text = 6
            End If

            If ComboBox_Vertical_scale_HMM.Text = "100" Then
                ComboBox_Horizontal_scale_HMM.Text = "1000"
                TextBox_PROFILE_LENGTH.Text = 360
                TextBox_horizontal_increment.Text = 50
                TextBox_vertical_increment.Text = 1
                TextBox_nr_vert_labels.Text = 6
            End If

            If ComboBox_Vertical_scale_HMM.Text = "500" Then
                ComboBox_Horizontal_scale_HMM.Text = 1000
                TextBox_horizontal_increment.Text = 100
                TextBox_vertical_increment.Text = 5
                TextBox_nr_vert_labels.Text = 19
                TextBox_PROFILE_LENGTH.Text = 950
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub ComboBox_hor_scale_HMM_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_Horizontal_scale_HMM.SelectedIndexChanged
        Try
            If ComboBox_Horizontal_scale_HMM.Text = "1000" Then
                TextBox_PROFILE_LENGTH.Text = 750
                TextBox_horizontal_increment.Text = 100
                TextBox_vertical_increment.Text = 10
                TextBox_nr_vert_labels.Text = 11
            End If

            If ComboBox_Horizontal_scale_HMM.Text = "2000" Then
                TextBox_PROFILE_LENGTH.Text = 1900
                TextBox_horizontal_increment.Text = 200
                TextBox_vertical_increment.Text = 20
                TextBox_nr_vert_labels.Text = 17
            End If


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

End Class