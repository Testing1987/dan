Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class Object_data_tools_form
    Dim colectie1 As Specialized.StringCollection
    Private Sub Object_data_tools_form_Load(sender As Object, e As EventArgs) Handles Me.Load

        Incarca_existing_layers_to_combobox(ComboBox_layer_linelist)
        With ComboBox_layer_linelist
            If .Items.Count > 0 Then
                If .Items.Contains("Text_PROPERTY") = True Then
                    .SelectedIndex = .Items.IndexOf("Text_PROPERTY")
                Else
                    .SelectedIndex = 0
                End If
            End If
        End With


        Incarca_existing_layers_to_combobox(ComboBox_layer_owner)
        With ComboBox_layer_owner
            If .Items.Count > 0 Then
                If .Items.Contains("Text_PROPERTY") = True Then
                    .SelectedIndex = .Items.IndexOf("Text_PROPERTY")
                Else
                    .SelectedIndex = 0
                End If
            End If
        End With

        Incarca_existing_layers_to_combobox(ComboBox_layer_street)
        With ComboBox_layer_street
            If .Items.Count > 0 Then
                If .Items.Contains("Text_Lables") = True Then
                    .SelectedIndex = .Items.IndexOf("Text_Lables")
                Else
                    .SelectedIndex = 0
                End If
            End If
        End With

        Incarca_existing_layers_to_combobox(ComboBox_layer_stream)
        With ComboBox_layer_stream
            If .Items.Count > 0 Then
                If .Items.Contains("Text_STREAM LABEL") = True Then
                    .SelectedIndex = .Items.IndexOf("Text_STREAM LABEL")
                Else
                    .SelectedIndex = 0
                End If
            End If
        End With

        Incarca_existing_layers_to_combobox(ComboBox_layer_access_road)
        With ComboBox_layer_access_road
            If .Items.Count > 0 Then
                If .Items.Contains("Text_ACCESS ROAD") = True Then
                    .SelectedIndex = .Items.IndexOf("Text_ACCESS ROAD")
                Else
                    .SelectedIndex = 0
                End If
            End If
        End With
        Incarca_existing_layers_to_combobox(ComboBox_mleader)
        With ComboBox_mleader
            If .Items.Count > 0 Then
                If .Items.Contains("E&S_blocks") = True Then
                    .SelectedIndex = .Items.IndexOf("E&S_blocks")
                Else
                    .SelectedIndex = 0
                End If
            End If
        End With
    End Sub

    Private Sub Button_refresh_click(sender As Object, e As EventArgs) Handles Button_refresh.Click

        Dim Text_ComboBox_layer_linelist As String = ComboBox_layer_linelist.Text
        Incarca_existing_layers_to_combobox(ComboBox_layer_linelist)
        With ComboBox_layer_linelist
            If .Items.Count > 0 Then
                If .Items.Contains(Text_ComboBox_layer_linelist) = True And Not Text_ComboBox_layer_linelist = "0" Then
                    .SelectedIndex = .Items.IndexOf(Text_ComboBox_layer_linelist)
                Else
                    If .Items.Contains("Text_PROPERTY") = True Then
                        .SelectedIndex = .Items.IndexOf("Text_PROPERTY")
                    Else
                        .SelectedIndex = 0
                    End If
                End If

            End If
        End With

        Dim Text_ComboBox_layer_owner As String = ComboBox_layer_owner.Text
        Incarca_existing_layers_to_combobox(ComboBox_layer_owner)
        With ComboBox_layer_owner
            If .Items.Count > 0 Then
                If .Items.Contains(Text_ComboBox_layer_owner) = True And Not Text_ComboBox_layer_owner = "0" Then
                    .SelectedIndex = .Items.IndexOf(Text_ComboBox_layer_owner)
                Else
                    If .Items.Contains("Text_PROPERTY") = True Then
                        .SelectedIndex = .Items.IndexOf("Text_PROPERTY")
                    Else
                        .SelectedIndex = 0
                    End If
                End If

            End If
        End With

        Dim Text_ComboBox_layer_street As String = ComboBox_layer_street.Text
        Incarca_existing_layers_to_combobox(ComboBox_layer_street)
        With ComboBox_layer_street
            If .Items.Count > 0 Then
                If .Items.Contains(Text_ComboBox_layer_street) = True And Not Text_ComboBox_layer_street = "0" Then
                    .SelectedIndex = .Items.IndexOf(Text_ComboBox_layer_street)
                Else
                    If .Items.Contains("Text_Lables") = True Then
                        .SelectedIndex = .Items.IndexOf("Text_Lables")
                    Else
                        .SelectedIndex = 0
                    End If
                End If

            End If
        End With

        Dim Text_ComboBox_layer_stream As String = ComboBox_layer_stream.Text
        Incarca_existing_layers_to_combobox(ComboBox_layer_stream)
        With ComboBox_layer_stream
            If .Items.Count > 0 Then
                If .Items.Contains(Text_ComboBox_layer_stream) = True And Not Text_ComboBox_layer_stream = "0" Then
                    .SelectedIndex = .Items.IndexOf(Text_ComboBox_layer_stream)
                Else
                    If .Items.Contains("Text_STREAM LABEL") = True Then
                        .SelectedIndex = .Items.IndexOf("Text_STREAM LABEL")
                    Else
                        .SelectedIndex = 0
                    End If
                End If

            End If
        End With

        Dim Text_ComboBox_layer_access_road As String = ComboBox_layer_access_road.Text
        Incarca_existing_layers_to_combobox(ComboBox_layer_access_road)
        With ComboBox_layer_access_road
            If .Items.Count > 0 Then
                If .Items.Contains(Text_ComboBox_layer_access_road) = True And Not Text_ComboBox_layer_access_road = "0" Then
                    .SelectedIndex = .Items.IndexOf(Text_ComboBox_layer_access_road)
                Else
                    If .Items.Contains("Text_ACCESS ROAD") = True Then
                        .SelectedIndex = .Items.IndexOf("Text_ACCESS ROAD")
                    Else
                        .SelectedIndex = 0
                    End If
                End If

            End If
        End With

        Dim Text_ComboBox_mleader As String = ComboBox_mleader.Text
        Incarca_existing_layers_to_combobox(ComboBox_mleader)
        With ComboBox_mleader
            If .Items.Count > 0 Then
                If .Items.Contains(Text_ComboBox_mleader) = True And Not Text_ComboBox_mleader = "0" Then
                    .SelectedIndex = .Items.IndexOf(Text_ComboBox_mleader)
                Else
                    If .Items.Contains("E&S_blocks") = True Then
                        .SelectedIndex = .Items.IndexOf("E&S_blocks")
                    Else
                        .SelectedIndex = 0
                    End If
                End If

            End If
        End With
    End Sub


    Private Sub Button_align_ucs_to_rectangle_Click(sender As Object, e As EventArgs) Handles Button_align_ucs_to_rectangle.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Try

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt.MessageForAdding = vbLf & "Select blue sky rectangle:"

            Rezultat1 = Editor1.GetSelection(Object_Prompt)
            colectie1 = New Specialized.StringCollection
            ascunde_butoanele_pentru_forms(Me, colectie1)
            If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                afiseaza_butoanele_pentru_forms(Me, colectie1)
                Exit Sub

            End If



            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then

                    Using Lock_dwg As DocumentLock = ThisDrawing.LockDocument
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim CurentUCSmatrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
                            Dim CurentUCS As CoordinateSystem3d = CurentUCSmatrix.CoordinateSystem3d
                            Dim Point0 As New Point3d
                            Dim Point1 As New Point3d
                            Dim PointM As New Point3d
                            Dim ucs_NAME As String = ""

                            If Rezultat1.Value.Count = 2 Then
                                For i = 0 To Rezultat1.Value.Count - 1
                                    Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                    Obj1 = Rezultat1.Value.Item(i)
                                    Dim Ent1 As Entity
                                    Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                                    If TypeOf Ent1 Is DBText Then
                                        Dim Text1 As DBText = Ent1
                                        ucs_NAME = Text1.TextString

                                    End If
                                    If TypeOf Ent1 Is Polyline Then
                                        Dim POLY1 As Polyline = Ent1
                                        If POLY1.NumberOfVertices = 4 Then
                                            Dim Point2 As New Point3d
                                            Point0 = POLY1.GetPoint3dAt(0)
                                            Point1 = POLY1.GetPoint3dAt(1)
                                            Point2 = POLY1.GetPoint3dAt(3)
                                            PointM = New Point3d(0.5 * Point1.X + 0.5 * Point2.X, 0.5 * Point1.Y + 0.5 * Point2.Y, 0)

                                        End If
                                    End If

                                Next
                            End If
                            If Rezultat1.Value.Count = 1 Then
                                For i = 0 To Rezultat1.Value.Count - 1
                                    Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                    Obj1 = Rezultat1.Value.Item(i)
                                    Dim Ent1 As Entity
                                    Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                                    ucs_NAME = "UCS_1"
                                    If TypeOf Ent1 Is Polyline Then
                                        Dim POLY1 As Polyline = Ent1
                                        If POLY1.NumberOfVertices = 4 Then
                                            Dim Point2 As New Point3d
                                            Point0 = POLY1.GetPoint3dAt(0)
                                            Point1 = POLY1.GetPoint3dAt(1)
                                            Point2 = POLY1.GetPoint3dAt(3)
                                            PointM = New Point3d(0.5 * Point1.X + 0.5 * Point2.X, 0.5 * Point1.Y + 0.5 * Point2.Y, 0)

                                        End If
                                    End If

                                Next
                            End If


                            Dim Ucs_table As UcsTable = Trans1.GetObject(ThisDrawing.Database.UcsTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                            Dim View_Table As ViewTable = Trans1.GetObject(ThisDrawing.Database.ViewTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                            If Not Point0.X = 0 And Not Point0.Y = 0 And Not Point1.X = 0 And Not Point1.Y = 0 Then
                                If Not ucs_NAME = "" Then

                                    Dim Ratio As Double = ThisDrawing.Editor.GetCurrentView.Width / ThisDrawing.Editor.GetCurrentView.Height
                                    Dim Latime_view As Double = 1.3 * Point0.GetVectorTo(Point1).Length
                                    Dim Ucs1 As UcsTableRecord

                                    If Ucs_table.Has(ucs_NAME) = False Then
                                        Ucs1 = New UcsTableRecord
                                        Ucs1.Name = ucs_NAME
                                        Ucs_table.UpgradeOpen()
                                        Ucs_table.Add(Ucs1)
                                        Trans1.AddNewlyCreatedDBObject(Ucs1, True)
                                        Ucs1.Origin = Point0
                                        Ucs1.XAxis = Point0.GetVectorTo(Point1)
                                        Ucs1.YAxis = Vector3d.ZAxis.CrossProduct(Point0.GetVectorTo(Point1))

                                        Dim ViewportTableRecord1 As ViewportTableRecord
                                        ViewportTableRecord1 = Trans1.GetObject(ThisDrawing.Editor.ActiveViewportId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)
                                        ViewportTableRecord1.IconAtOrigin = True
                                        ViewportTableRecord1.IconEnabled = True
                                        ViewportTableRecord1.SetUcs(Ucs1.ObjectId)
                                        ThisDrawing.Editor.UpdateTiledViewportsFromDatabase()
                                    Else
                                        Ucs1 = Ucs_table(ucs_NAME).GetObject(OpenMode.ForWrite)
                                        Ucs1.Origin = Point0
                                        Ucs1.XAxis = Point0.GetVectorTo(Point1)
                                        Ucs1.YAxis = Vector3d.ZAxis.CrossProduct(Point0.GetVectorTo(Point1))

                                        Dim ViewportTableRecord1 As ViewportTableRecord
                                        ViewportTableRecord1 = Trans1.GetObject(ThisDrawing.Editor.ActiveViewportId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)
                                        ViewportTableRecord1.IconAtOrigin = True
                                        ViewportTableRecord1.IconEnabled = True
                                        ViewportTableRecord1.SetUcs(Ucs1.ObjectId)
                                        ThisDrawing.Editor.UpdateTiledViewportsFromDatabase()

                                    End If

                                    Dim Rotatie As Double = Vector3d.XAxis.GetAngleTo(Ucs1.YAxis) + PI / 2
                                    Dim view1 As ViewTableRecord

                                    If View_Table.Has(ucs_NAME) = False Then
                                        View_Table.UpgradeOpen()
                                        view1 = New ViewTableRecord
                                        view1.CenterPoint = New Point2d(0, 0)
                                        view1.Target = PointM
                                        view1.ViewTwist = Rotatie
                                        view1.Width = Latime_view
                                        view1.Height = Latime_view / Ratio
                                        view1.Name = ucs_NAME
                                        View_Table.Add(view1)
                                        Trans1.AddNewlyCreatedDBObject(view1, True)
                                    Else
                                        view1 = View_Table(ucs_NAME).GetObject(OpenMode.ForWrite)
                                        view1.CenterPoint = New Point2d(0, 0)
                                        view1.Target = PointM
                                        view1.ViewTwist = Rotatie
                                        view1.Width = Latime_view
                                        view1.Height = Latime_view / Ratio
                                    End If


                                    ThisDrawing.Editor.SetCurrentView(view1)


                                End If
                            End If






                            Trans1.Commit()
                        End Using
                    End Using

                End If
            End If

            afiseaza_butoanele_pentru_forms(Me, colectie1)
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
            afiseaza_butoanele_pentru_forms(Me, colectie1)
        End Try
    End Sub


    Private Sub Button_line_list_symbol_Click(sender As Object, e As EventArgs) Handles Button_line_list_symbol.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Try
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            colectie1 = New Specialized.StringCollection
            ascunde_butoanele_pentru_forms(Me, colectie1)

            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptEntityResult
lbl1:
            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select entity containing object data:")

            Rezultat1 = Editor1.GetEntity(Object_Prompt)

            If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, colectie1)
                Exit Sub
            End If

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then

                    Using Lock_dwg As DocumentLock = ThisDrawing.LockDocument


                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim Ent1 As Entity = Trans1.GetObject(Rezultat1.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                            If TypeOf Ent1 Is Dimension Then
                                Dim xxx1 As Dimension = Ent1
                                'MsgBox(Poly1.GetBulgeAt(0) & vbCrLf & Poly1.GetBulgeAt(1))
                            End If
                            Dim Pick_point As Point3d = Rezultat1.PickedPoint '.TransformBy(Editor1.CurrentUserCoordinateSystem)
                            Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                            Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables
                            Dim Id1 As ObjectId = Ent1.ObjectId
                            Dim Records1 As Autodesk.Gis.Map.ObjectData.Records
                            If Not Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False).Count = 0 Then
                                Records1 = Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                Dim Record1 As Autodesk.Gis.Map.ObjectData.Record
                                Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                Dim LineList As String = ""
                                For Each Record1 In Records1
                                    Dim Tabla1 As Autodesk.Gis.Map.ObjectData.Table
                                    Tabla1 = Tables1(Record1.TableName)
                                    If Tabla1.Name.ToUpper = "E_BDY_PROPERTY" Then
                                        Dim Field_defs1 As Autodesk.Gis.Map.ObjectData.FieldDefinitions
                                        Field_defs1 = Tabla1.FieldDefinitions
                                        For i = 0 To Record1.Count - 1
                                            Dim Field_def1 As Autodesk.Gis.Map.ObjectData.FieldDefinition
                                            Field_def1 = Field_defs1(i)
                                            If Field_def1.Name.ToUpper = "LINELIST" Then
                                                Valoare_record1 = Record1(i)
                                                LineList = Valoare_record1.StrValue
                                            End If
                                        Next
                                    End If
                                Next

                                If Not LineList = "" Then
                                    Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                                    Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify insertion point:")
                                    PP1.UseBasePoint = True
                                    PP1.BasePoint = Pick_point
                                    PP1.AllowNone = False
                                    Point1 = Editor1.GetPoint(PP1)

                                    If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                        Editor1.WriteMessage(vbLf & "Command:")
                                        afiseaza_butoanele_pentru_forms(Me, colectie1)
                                        Exit Sub
                                    End If

                                    Dim CurentUCSmatrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
                                    Dim CurentUCS As CoordinateSystem3d = CurentUCSmatrix.CoordinateSystem3d

                                    Dim Planul_curent As Plane = New Plane(New Point3d(0, 0, 0), CurentUCS.Zaxis)
                                    Dim Rotatie As Double = CurentUCS.Xaxis.GetAngleTo(Vector3d.XAxis)

                                    'New Point3d(0, 0, 0).GetVectorTo(New Point3d(x2, y2, 0)).AngleOnPlane(Planul_curent)

                                    Dim X1 As Double = Point1.Value.TransformBy(CurentUCSmatrix).X
                                    Dim Y1 As Double = Point1.Value.TransformBy(CurentUCSmatrix).Y
                                    Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                    BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)

                                    Dim Mtext1 As New Autodesk.AutoCAD.DatabaseServices.MText()
                                    Mtext1.SetDatabaseDefaults()
                                    Mtext1.LineSpacingFactor = 0.875
                                    Mtext1.Contents = "{\Fromans|c0;\W0.65;" & LineList & "}"
                                    Mtext1.TextHeight = 10
                                    Mtext1.Location = New Autodesk.AutoCAD.Geometry.Point3d(X1, Y1, 0)
                                    Mtext1.Rotation = 0
                                    Mtext1.Attachment = AttachmentPoint.MiddleCenter
                                    Mtext1.Layer = ComboBox_layer_linelist.Text
                                    Mtext1.ColorIndex = 1
                                    BTrecord.AppendEntity(Mtext1)
                                    Trans1.AddNewlyCreatedDBObject(Mtext1, True)

                                    Dim Pt1 As New Point2d(X1 - 50, Y1 + 9)
                                    Dim Pt2 As New Point2d(X1 - 50, Y1 - 9)
                                    Dim Pt3 As New Point2d(X1 + 50, Y1 - 9)
                                    Dim Pt4 As New Point2d(X1 + 50, Y1 + 9)
                                    Dim Poly1 As New Polyline
                                    Poly1.AddVertexAt(0, Pt1, 0.76, 0, 0)
                                    Poly1.AddVertexAt(1, Pt2, 0, 0, 0)
                                    Poly1.AddVertexAt(2, Pt3, 0.76, 0, 0)
                                    Poly1.AddVertexAt(3, Pt4, 0, 0, 0)
                                    Poly1.TransformBy(Matrix3d.Rotation(Rotatie, Vector3d.ZAxis, Point1.Value.TransformBy(CurentUCSmatrix)))
                                    Poly1.Closed = True
                                    Poly1.Layer = ComboBox_layer_linelist.Text
                                    Poly1.ColorIndex = 1
                                    BTrecord.AppendEntity(Poly1)
                                    Trans1.AddNewlyCreatedDBObject(Poly1, True)
                                    Dim Hatch1 As New Hatch
                                    BTrecord.AppendEntity(Hatch1)
                                    Trans1.AddNewlyCreatedDBObject(Hatch1, True)
                                    Hatch1.SetHatchPattern(HatchPatternType.PreDefined, "SOLID")
                                    Dim oBJiD_COL As New ObjectIdCollection
                                    oBJiD_COL.Add(Poly1.ObjectId)
                                    Hatch1.AppendLoop(HatchLoopTypes.External, oBJiD_COL)
                                    Hatch1.Layer = Poly1.Layer
                                    Hatch1.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 255, 255)
                                    Hatch1.EvaluateHatch(True)

                                    oBJiD_COL = New ObjectIdCollection
                                    oBJiD_COL.Add(Hatch1.ObjectId)
                                    Dim DrawOrderTable1 As Autodesk.AutoCAD.DatabaseServices.DrawOrderTable = Trans1.GetObject(BTrecord.DrawOrderTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)
                                    DrawOrderTable1.MoveToBottom(oBJiD_COL)

                                End If



                                Trans1.Commit()

                            End If



                        End Using
                    End Using

                End If
            End If

            GoTo lbl1
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, colectie1)
    End Sub




    Private Sub Button_owner_Click(sender As Object, e As EventArgs) Handles Button_owner.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Try

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            colectie1 = New Specialized.StringCollection
            ascunde_butoanele_pentru_forms(Me, colectie1)

            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptEntityResult
lbl1:
            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select entity containing object data:")

            Rezultat1 = Editor1.GetEntity(Object_Prompt)

            If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, colectie1)
                Exit Sub

            End If



            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then

                    Using Lock_dwg As DocumentLock = ThisDrawing.LockDocument


                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim Ent1 As Entity = Trans1.GetObject(Rezultat1.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                            If TypeOf Ent1 Is Dimension Then
                                Dim xxx1 As Dimension = Ent1
                                'MsgBox(Poly1.GetBulgeAt(0) & vbCrLf & Poly1.GetBulgeAt(1))
                            End If
                            Dim Pick_point As Point3d = Rezultat1.PickedPoint '.TransformBy(Editor1.CurrentUserCoordinateSystem)
                            Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                            Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables
                            Dim Id1 As ObjectId = Ent1.ObjectId
                            Dim Records1 As Autodesk.Gis.Map.ObjectData.Records
                            If Not Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False).Count = 0 Then
                                Records1 = Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                Dim Record1 As Autodesk.Gis.Map.ObjectData.Record
                                Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                Dim OwnerName As String = ""
                                For Each Record1 In Records1
                                    Dim Tabla1 As Autodesk.Gis.Map.ObjectData.Table
                                    Tabla1 = Tables1(Record1.TableName)
                                    If Tabla1.Name.ToUpper = "E_BDY_PROPERTY" Then
                                        Dim Field_defs1 As Autodesk.Gis.Map.ObjectData.FieldDefinitions
                                        Field_defs1 = Tabla1.FieldDefinitions
                                        For i = 0 To Record1.Count - 1
                                            Dim Field_def1 As Autodesk.Gis.Map.ObjectData.FieldDefinition
                                            Field_def1 = Field_defs1(i)
                                            If Field_def1.Name.ToUpper = "OWNERNAME" Then
                                                Valoare_record1 = Record1(i)
                                                OwnerName = Valoare_record1.StrValue
                                            End If
                                        Next
                                    End If
                                Next

                                If Not OwnerName = "" Then
                                    Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                                    Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify insertion point:")
                                    PP1.UseBasePoint = True
                                    PP1.BasePoint = Pick_point
                                    PP1.AllowNone = False
                                    Point1 = Editor1.GetPoint(PP1)

                                    If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                        Editor1.WriteMessage(vbLf & "Command:")
                                        afiseaza_butoanele_pentru_forms(Me, colectie1)
                                        Exit Sub
                                    End If

                                    Dim CurentUCSmatrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
                                    Dim CurentUCS As CoordinateSystem3d = CurentUCSmatrix.CoordinateSystem3d

                                    Dim Planul_curent As Plane = New Plane(New Point3d(0, 0, 0), CurentUCS.Zaxis)
                                    Dim Rotatie As Double = CurentUCS.Xaxis.GetAngleTo(Vector3d.XAxis)

                                    'New Point3d(0, 0, 0).GetVectorTo(New Point3d(x2, y2, 0)).AngleOnPlane(Planul_curent)

                                    Dim X1 As Double = Point1.Value.TransformBy(CurentUCSmatrix).X
                                    Dim Y1 As Double = Point1.Value.TransformBy(CurentUCSmatrix).Y

                                    Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                    BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)


                                    Dim Mtext1 As New Autodesk.AutoCAD.DatabaseServices.MText()
                                    Mtext1.SetDatabaseDefaults()
                                    Mtext1.LineSpacingFactor = 0.875
                                    Mtext1.Contents = "{\Fromans|c0;" & OwnerName & "}"
                                    Mtext1.TextHeight = 10
                                    Mtext1.Location = New Autodesk.AutoCAD.Geometry.Point3d(X1, Y1, 0)
                                    Mtext1.Rotation = 0
                                    Mtext1.Attachment = AttachmentPoint.TopCenter
                                    Mtext1.Layer = ComboBox_layer_owner.Text
                                    Mtext1.ColorIndex = 3
                                    BTrecord.AppendEntity(Mtext1)
                                    Trans1.AddNewlyCreatedDBObject(Mtext1, True)


                                End If



                                Trans1.Commit()

                            End If



                        End Using
                    End Using

                End If
            End If

            GoTo lbl1
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, colectie1)
    End Sub

    Private Sub Button_street_Click(sender As Object, e As EventArgs) Handles Button_street.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Try

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            colectie1 = New Specialized.StringCollection
            ascunde_butoanele_pentru_forms(Me, colectie1)

            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptEntityResult
lbl1:
            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select entity containing object data:")

            Rezultat1 = Editor1.GetEntity(Object_Prompt)

            If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, colectie1)
                Exit Sub

            End If



            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then

                    Using Lock_dwg As DocumentLock = ThisDrawing.LockDocument


                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim Ent1 As Entity = Trans1.GetObject(Rezultat1.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                            If TypeOf Ent1 Is Dimension Then
                                Dim xxx1 As Dimension = Ent1
                                'MsgBox(Poly1.GetBulgeAt(0) & vbCrLf & Poly1.GetBulgeAt(1))
                            End If
                            Dim Pick_point As Point3d = Rezultat1.PickedPoint '.TransformBy(Editor1.CurrentUserCoordinateSystem)
                            Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                            Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables
                            Dim Id1 As ObjectId = Ent1.ObjectId
                            Dim Records1 As Autodesk.Gis.Map.ObjectData.Records
                            If Not Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False).Count = 0 Then
                                Records1 = Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                Dim Record1 As Autodesk.Gis.Map.ObjectData.Record
                                Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                Dim RoadName As String = ""
                                For Each Record1 In Records1
                                    Dim Tabla1 As Autodesk.Gis.Map.ObjectData.Table
                                    Tabla1 = Tables1(Record1.TableName)
                                    If Tabla1.Name.ToUpper = "E_FEA_ROADCL" Then
                                        Dim Field_defs1 As Autodesk.Gis.Map.ObjectData.FieldDefinitions
                                        Field_defs1 = Tabla1.FieldDefinitions
                                        For i = 0 To Record1.Count - 1
                                            Dim Field_def1 As Autodesk.Gis.Map.ObjectData.FieldDefinition
                                            Field_def1 = Field_defs1(i)
                                            If Field_def1.Name.ToUpper = "NAME" Then
                                                Valoare_record1 = Record1(i)
                                                RoadName = Valoare_record1.StrValue
                                            End If
                                        Next
                                    End If
                                Next

                                If Not RoadName = "" Then
                                    Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                                    Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify insertion point:")
                                    PP1.UseBasePoint = True
                                    PP1.BasePoint = Pick_point
                                    PP1.AllowNone = False
                                    Point1 = Editor1.GetPoint(PP1)

                                    If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                        Editor1.WriteMessage(vbLf & "Command:")
                                        afiseaza_butoanele_pentru_forms(Me, colectie1)
                                        Exit Sub
                                    End If

                                    Dim Point2 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                                    Dim PP2 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify label rotation:")
                                    PP2.UseBasePoint = True
                                    PP2.BasePoint = Point1.Value
                                    PP2.AllowNone = False
                                    Point2 = Editor1.GetPoint(PP2)

                                    If Point2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                        Editor1.WriteMessage(vbLf & "Command:")
                                        afiseaza_butoanele_pentru_forms(Me, colectie1)
                                        Exit Sub
                                    End If

                                    Dim CurentUCSmatrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
                                    Dim CurentUCS As CoordinateSystem3d = CurentUCSmatrix.CoordinateSystem3d

                                    Dim Planul_curent As Plane = New Plane(New Point3d(0, 0, 0), CurentUCS.Zaxis)
                                    Dim Rotatie As Double = CurentUCS.Xaxis.GetAngleTo(Vector3d.XAxis)

                                    'New Point3d(0, 0, 0).GetVectorTo(New Point3d(x2, y2, 0)).AngleOnPlane(Planul_curent)

                                    Dim X1 As Double = Point1.Value.TransformBy(CurentUCSmatrix).X
                                    Dim Y1 As Double = Point1.Value.TransformBy(CurentUCSmatrix).Y

                                    Dim X01 As Double = Point1.Value.X
                                    Dim Y01 As Double = Point1.Value.Y
                                    Dim X2 As Double = Point2.Value.X
                                    Dim Y2 As Double = Point2.Value.Y
                                    Dim rotatie1 As Double = GET_Bearing_rad(X01, Y01, X2, Y2)

                                    Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                    BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)


                                    Dim Mtext1 As New Autodesk.AutoCAD.DatabaseServices.MText()
                                    Mtext1.SetDatabaseDefaults()
                                    Mtext1.LineSpacingFactor = 0.875
                                    Mtext1.Contents = "{\Fromans|c0;" & RoadName & "}"
                                    Mtext1.TextHeight = 8
                                    Mtext1.Location = New Autodesk.AutoCAD.Geometry.Point3d(X1, Y1, 0)
                                    Mtext1.Rotation = rotatie1
                                    Mtext1.Attachment = AttachmentPoint.MiddleLeft
                                    Mtext1.Layer = ComboBox_layer_street.Text
                                    Mtext1.ColorIndex = 256
                                    BTrecord.AppendEntity(Mtext1)
                                    Trans1.AddNewlyCreatedDBObject(Mtext1, True)


                                End If



                                Trans1.Commit()

                            End If



                        End Using
                    End Using

                End If
            End If

            GoTo lbl1
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, colectie1)
    End Sub

    Private Sub Button_stream_Click(sender As Object, e As EventArgs) Handles Button_stream.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Try

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            colectie1 = New Specialized.StringCollection
            ascunde_butoanele_pentru_forms(Me, colectie1)

            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptEntityResult
lbl1:
            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select entity containing object data:")

            Rezultat1 = Editor1.GetEntity(Object_Prompt)

            If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, colectie1)
                Exit Sub

            End If



            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then

                    Using Lock_dwg As DocumentLock = ThisDrawing.LockDocument


                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim Ent1 As Entity = Trans1.GetObject(Rezultat1.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                            If TypeOf Ent1 Is Dimension Then
                                Dim xxx1 As Dimension = Ent1
                                'MsgBox(Poly1.GetBulgeAt(0) & vbCrLf & Poly1.GetBulgeAt(1))
                            End If
                            Dim Pick_point As Point3d = Rezultat1.PickedPoint '.TransformBy(Editor1.CurrentUserCoordinateSystem)
                            Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                            Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables
                            Dim Id1 As ObjectId = Ent1.ObjectId
                            Dim Records1 As Autodesk.Gis.Map.ObjectData.Records
                            If Not Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False).Count = 0 Then
                                Records1 = Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                Dim Record1 As Autodesk.Gis.Map.ObjectData.Record
                                Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                Dim StreamName As String = ""
                                For Each Record1 In Records1
                                    Dim Tabla1 As Autodesk.Gis.Map.ObjectData.Table
                                    Tabla1 = Tables1(Record1.TableName)
                                    If Tabla1.Name.ToUpper = "E_ENV_STREAMCL" Then
                                        Dim Field_defs1 As Autodesk.Gis.Map.ObjectData.FieldDefinitions
                                        Field_defs1 = Tabla1.FieldDefinitions
                                        For i = 0 To Record1.Count - 1
                                            Dim Field_def1 As Autodesk.Gis.Map.ObjectData.FieldDefinition
                                            Field_def1 = Field_defs1(i)
                                            If Field_def1.Name.ToUpper = "NAME" Then
                                                Valoare_record1 = Record1(i)
                                                StreamName = Valoare_record1.StrValue
                                            End If
                                        Next
                                    End If
                                Next

                                If Not StreamName = "" Then
                                    Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                                    Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify second point:")
                                    PP1.UseBasePoint = True
                                    PP1.BasePoint = Pick_point
                                    PP1.AllowNone = False
                                    Point1 = Editor1.GetPoint(PP1)

                                    If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                        Editor1.WriteMessage(vbLf & "Command:")
                                        afiseaza_butoanele_pentru_forms(Me, colectie1)
                                        Exit Sub
                                    End If



                                    Dim CurentUCSmatrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
                                    Dim CurentUCS As CoordinateSystem3d = CurentUCSmatrix.CoordinateSystem3d

                                    Dim Planul_curent As Plane = New Plane(New Point3d(0, 0, 0), CurentUCS.Zaxis)
                                    Dim Rotatie As Double = CurentUCS.Xaxis.GetAngleTo(Vector3d.XAxis)

                                    'New Point3d(0, 0, 0).GetVectorTo(New Point3d(x2, y2, 0)).AngleOnPlane(Planul_curent)

                                    Dim X1 As Double = Point1.Value.TransformBy(CurentUCSmatrix).X
                                    Dim Y1 As Double = Point1.Value.TransformBy(CurentUCSmatrix).Y

                                    Dim X01 As Double = Point1.Value.X
                                    Dim Y01 As Double = Point1.Value.Y

                                    Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                    BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)

                                    CREAZA_LEADER(Pick_point, "{\Fromans|c0;" & StreamName.ToUpper & "}", 8, 0, ComboBox_layer_stream.Text, 18, 12, 12, 256, 256, Point1.Value)
                                    Dim Pt1 As New Point2d(X1 + 12, Y1 + 7.5)
                                    Dim Pt2 As New Point2d(X1 + 12, Y1 - 7.5)
                                    Dim Pt3 As New Point2d(X1 + 80 + 12, Y1 - 7.5)
                                    Dim Pt4 As New Point2d(X1 + 80 + 12, Y1 + 7.5)
                                    Dim Poly1 As New Polyline
                                    Poly1.AddVertexAt(0, Pt1, 0, 0, 0)
                                    Poly1.AddVertexAt(1, Pt2, 0, 0, 0)
                                    Poly1.AddVertexAt(2, Pt3, 0, 0, 0)
                                    Poly1.AddVertexAt(3, Pt4, 0, 0, 0)
                                    Poly1.TransformBy(Matrix3d.Rotation(Rotatie, Vector3d.ZAxis, Point1.Value.TransformBy(CurentUCSmatrix)))
                                    Poly1.Closed = True
                                    Poly1.Layer = ComboBox_layer_stream.Text
                                    Poly1.ColorIndex = 256
                                    BTrecord.AppendEntity(Poly1)
                                    Trans1.AddNewlyCreatedDBObject(Poly1, True)
                                    Dim Hatch1 As New Hatch
                                    BTrecord.AppendEntity(Hatch1)
                                    Trans1.AddNewlyCreatedDBObject(Hatch1, True)
                                    Hatch1.SetHatchPattern(HatchPatternType.PreDefined, "SOLID")
                                    Dim oBJiD_COL As New ObjectIdCollection
                                    oBJiD_COL.Add(Poly1.ObjectId)
                                    Hatch1.AppendLoop(HatchLoopTypes.External, oBJiD_COL)
                                    Hatch1.Layer = Poly1.Layer
                                    Hatch1.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 255, 255)
                                    Hatch1.EvaluateHatch(True)

                                    oBJiD_COL = New ObjectIdCollection
                                    oBJiD_COL.Add(Hatch1.ObjectId)
                                    Dim DrawOrderTable1 As Autodesk.AutoCAD.DatabaseServices.DrawOrderTable = Trans1.GetObject(BTrecord.DrawOrderTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)
                                    DrawOrderTable1.MoveToBottom(oBJiD_COL)

                                End If



                                Trans1.Commit()

                            End If



                        End Using
                    End Using

                End If
            End If

            GoTo lbl1
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, colectie1)
    End Sub

    Private Sub Button_access_ROAD_Click(sender As Object, e As EventArgs) Handles Button_access_ROAD.Click


        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Try

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            colectie1 = New Specialized.StringCollection
            ascunde_butoanele_pentru_forms(Me, colectie1)

            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptEntityResult
lbl1:
            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select entity containing object data:")

            Rezultat1 = Editor1.GetEntity(Object_Prompt)

            If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, colectie1)
                Exit Sub

            End If



            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then

                    Using Lock_dwg As DocumentLock = ThisDrawing.LockDocument


                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim Ent1 As Entity = Trans1.GetObject(Rezultat1.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                            If TypeOf Ent1 Is Dimension Then
                                Dim xxx1 As Dimension = Ent1
                                'MsgBox(Poly1.GetBulgeAt(0) & vbCrLf & Poly1.GetBulgeAt(1))
                            End If
                            Dim Pick_point As Point3d = Rezultat1.PickedPoint '.TransformBy(Editor1.CurrentUserCoordinateSystem)
                            Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                            Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables
                            Dim Id1 As ObjectId = Ent1.ObjectId
                            Dim Records1 As Autodesk.Gis.Map.ObjectData.Records
                            If Not Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False).Count = 0 Then
                                Records1 = Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                Dim Record1 As Autodesk.Gis.Map.ObjectData.Record
                                Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue
                                Dim accessroadName As String = ""
                                For Each Record1 In Records1
                                    Dim Tabla1 As Autodesk.Gis.Map.ObjectData.Table
                                    Tabla1 = Tables1(Record1.TableName)
                                    If Tabla1.Name.ToUpper = "P_FEA_ACCESSRDCL" Then
                                        Dim Field_defs1 As Autodesk.Gis.Map.ObjectData.FieldDefinitions
                                        Field_defs1 = Tabla1.FieldDefinitions
                                        For i = 0 To Record1.Count - 1
                                            Dim Field_def1 As Autodesk.Gis.Map.ObjectData.FieldDefinition
                                            Field_def1 = Field_defs1(i)
                                            If Field_def1.Name.ToUpper = "NAME" Then
                                                Valoare_record1 = Record1(i)
                                                accessroadName = Valoare_record1.StrValue
                                            End If
                                        Next
                                    End If
                                Next

                                If Not accessroadName = "" Then
                                    Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                                    Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify second point:")
                                    PP1.UseBasePoint = True
                                    PP1.BasePoint = Pick_point
                                    PP1.AllowNone = False
                                    Point1 = Editor1.GetPoint(PP1)

                                    If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                        Editor1.WriteMessage(vbLf & "Command:")
                                        afiseaza_butoanele_pentru_forms(Me, colectie1)
                                        Exit Sub
                                    End If



                                    Dim CurentUCSmatrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
                                    Dim CurentUCS As CoordinateSystem3d = CurentUCSmatrix.CoordinateSystem3d

                                    Dim Planul_curent As Plane = New Plane(New Point3d(0, 0, 0), CurentUCS.Zaxis)
                                    Dim Rotatie As Double = CurentUCS.Xaxis.GetAngleTo(Vector3d.XAxis)

                                    'New Point3d(0, 0, 0).GetVectorTo(New Point3d(x2, y2, 0)).AngleOnPlane(Planul_curent)

                                    Dim X1 As Double = Point1.Value.TransformBy(CurentUCSmatrix).X
                                    Dim Y1 As Double = Point1.Value.TransformBy(CurentUCSmatrix).Y

                                    Dim X01 As Double = Point1.Value.X
                                    Dim Y01 As Double = Point1.Value.Y

                                    Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                    BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)

                                    CREAZA_LEADER(Pick_point, "{\Fromans|c0;" & accessroadName.ToUpper & "}", 8, 0, ComboBox_layer_access_road.Text, 18, 12, 12, 24, 7, Point1.Value)
                                    Dim Pt1 As New Point2d(X1 + 12, Y1 + 7.5)
                                    Dim Pt2 As New Point2d(X1 + 12, Y1 - 7.5)
                                    Dim Pt3 As New Point2d(X1 + 65 + 12 + 5.625, Y1 - 7.5)
                                    Dim Pt4 As New Point2d(X1 + 65 + 12 + 5.625, Y1 + 7.5)
                                    Dim Poly1 As New Polyline
                                    Poly1.AddVertexAt(0, Pt1, 0.75, 0, 0)
                                    Poly1.AddVertexAt(1, Pt2, 0, 0, 0)
                                    Poly1.AddVertexAt(2, Pt3, 0.75, 0, 0)
                                    Poly1.AddVertexAt(3, Pt4, 0, 0, 0)
                                    Poly1.TransformBy(Matrix3d.Rotation(Rotatie, Vector3d.ZAxis, Point1.Value.TransformBy(CurentUCSmatrix)))
                                    Poly1.Closed = True
                                    Poly1.Layer = ComboBox_layer_access_road.Text
                                    Poly1.ColorIndex = 24
                                    BTrecord.AppendEntity(Poly1)
                                    Trans1.AddNewlyCreatedDBObject(Poly1, True)
                                    Dim Hatch1 As New Hatch
                                    BTrecord.AppendEntity(Hatch1)
                                    Trans1.AddNewlyCreatedDBObject(Hatch1, True)
                                    Hatch1.SetHatchPattern(HatchPatternType.PreDefined, "SOLID")
                                    Dim oBJiD_COL As New ObjectIdCollection
                                    oBJiD_COL.Add(Poly1.ObjectId)
                                    Hatch1.AppendLoop(HatchLoopTypes.External, oBJiD_COL)
                                    Hatch1.Layer = Poly1.Layer
                                    Hatch1.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 255, 255)
                                    Hatch1.EvaluateHatch(True)

                                    oBJiD_COL = New ObjectIdCollection
                                    oBJiD_COL.Add(Hatch1.ObjectId)
                                    Dim DrawOrderTable1 As Autodesk.AutoCAD.DatabaseServices.DrawOrderTable = Trans1.GetObject(BTrecord.DrawOrderTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)
                                    DrawOrderTable1.MoveToBottom(oBJiD_COL)

                                End If



                                Trans1.Commit()

                            End If



                        End Using
                    End Using

                End If
            End If

            GoTo lbl1
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, colectie1)
    End Sub

    Private Sub Button_Mleader_Click(sender As Object, e As EventArgs) Handles Button_Mleader.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Try

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            colectie1 = New Specialized.StringCollection
            ascunde_butoanele_pentru_forms(Me, colectie1)


lbl1:

            Using Lock_dwg As DocumentLock = ThisDrawing.LockDocument


                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                    Dim Point0 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                    Dim PP0 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify first point:")
                    PP0.AllowNone = False
                    Point0 = Editor1.GetPoint(PP0)

                    If Point0.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                        Editor1.WriteMessage(vbLf & "Command:")
                        afiseaza_butoanele_pentru_forms(Me, colectie1)
                        Exit Sub
                    End If


                    Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                    Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify second point:")
                    PP1.UseBasePoint = True
                    PP1.BasePoint = Point0.Value
                    PP1.AllowNone = False
                    Point1 = Editor1.GetPoint(PP1)

                    If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                        Editor1.WriteMessage(vbLf & "Command:")
                        afiseaza_butoanele_pentru_forms(Me, colectie1)
                        Exit Sub
                    End If



                    Dim CurentUCSmatrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
                    Dim CurentUCS As CoordinateSystem3d = CurentUCSmatrix.CoordinateSystem3d

                    Dim Planul_curent As Plane = New Plane(New Point3d(0, 0, 0), CurentUCS.Zaxis)
                    Dim Rotatie As Double = CurentUCS.Xaxis.GetAngleTo(Vector3d.XAxis)

                    'New Point3d(0, 0, 0).GetVectorTo(New Point3d(x2, y2, 0)).AngleOnPlane(Planul_curent)

                    Dim X1 As Double = Point1.Value.TransformBy(CurentUCSmatrix).X
                    Dim Y1 As Double = Point1.Value.TransformBy(CurentUCSmatrix).Y

                    Dim X01 As Double = Point1.Value.X
                    Dim Y01 As Double = Point1.Value.Y

                    Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)
                    Dim content As String = TextBox_mleader.Text



                    CREAZA_LEADER(Point0.Value, "{\Fromans|c0;" & content & "}", 8, 0, ComboBox_layer_access_road.Text, 18, 3.6, 4, 3, 3, Point1.Value)
                    Dim Pt1 As New Point2d(X1 + 3.6, Y1 + 7.5)
                    Dim Pt2 As New Point2d(X1 + 3.6, Y1 - 7.5)
                    Dim Pt3 As New Point2d(X1 + 32 + 3.6, Y1 - 7.5)
                    Dim Pt4 As New Point2d(X1 + 32 + 3.6, Y1 + 7.5)
                    Dim Poly1 As New Polyline
                    Poly1.AddVertexAt(0, Pt1, 0, 0, 0)
                    Poly1.AddVertexAt(1, Pt2, 0, 0, 0)
                    Poly1.AddVertexAt(2, Pt3, 0, 0, 0)
                    Poly1.AddVertexAt(3, Pt4, 0, 0, 0)
                    Poly1.TransformBy(Matrix3d.Rotation(Rotatie, Vector3d.ZAxis, Point1.Value.TransformBy(CurentUCSmatrix)))
                    Poly1.Closed = True
                    Poly1.Layer = ComboBox_mleader.Text
                    Poly1.ColorIndex = 3
                    BTrecord.AppendEntity(Poly1)
                    Trans1.AddNewlyCreatedDBObject(Poly1, True)
                    Dim Hatch1 As New Hatch
                    BTrecord.AppendEntity(Hatch1)
                    Trans1.AddNewlyCreatedDBObject(Hatch1, True)
                    Hatch1.SetHatchPattern(HatchPatternType.PreDefined, "SOLID")
                    Dim oBJiD_COL As New ObjectIdCollection
                    oBJiD_COL.Add(Poly1.ObjectId)
                    Hatch1.AppendLoop(HatchLoopTypes.External, oBJiD_COL)
                    Hatch1.Layer = Poly1.Layer
                    Hatch1.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 255, 255)
                    Hatch1.EvaluateHatch(True)

                    oBJiD_COL = New ObjectIdCollection
                    oBJiD_COL.Add(Hatch1.ObjectId)
                    Dim DrawOrderTable1 As Autodesk.AutoCAD.DatabaseServices.DrawOrderTable = Trans1.GetObject(BTrecord.DrawOrderTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)
                    DrawOrderTable1.MoveToBottom(oBJiD_COL)





                    Trans1.Commit()





                End Using
            End Using



            GoTo lbl1
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, colectie1)
    End Sub
End Class


