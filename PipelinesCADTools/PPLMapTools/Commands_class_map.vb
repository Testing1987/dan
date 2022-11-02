Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class Commands_class_map

    <CommandMethod("CPU1")> _
    Public Sub gaseste_nr_serial_al_disc_C()
        Dim disk As New Management.ManagementObject("Win32_LogicalDisk.DeviceID=""C:""")
        Dim diskPropertyB As Management.PropertyData = disk.Properties("VolumeSerialNumber")
        MsgBox(diskPropertyB.Value.ToString())
    End Sub
    <CommandMethod("PPLOBJECTDATA")> _
    Public Sub Show_OBJECT_DATA_TOOLS_FORM()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Object_data_tools_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Object_data_tools_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    <CommandMethod("PPL1")>
    Public Sub Select_doua_poly_then_transfer_to_EXCEL()
        If isSECURE() = False Then Exit Sub
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim CurentUCSmatrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        Try
            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt.MessageForAdding = vbLf & "Select CL polyline:"

            Object_Prompt.SingleOnly = True

            Rezultat1 = Editor1.GetSelection(Object_Prompt)


            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If

            Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

            Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt2.MessageForAdding = vbLf & "Select A PROPERTY polyline:"

            Object_Prompt2.SingleOnly = True

            Rezultat2 = Editor1.GetSelection(Object_Prompt2)


            If Rezultat2.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If

            Dim Poly1 As Polyline
            Dim Point_on_poly As New Point3d
            Dim Layer_property As String



            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then

                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                        Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj2 = Rezultat2.Value.Item(0)
                        Dim Ent2 As Entity
                        Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)

                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                            Poly1 = Ent1
                            Layer_property = Ent2.Layer
                            Trans1.Commit()
                        Else
                            Editor1.WriteMessage("No Polyline")
                            Editor1.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If
                    End Using
                End If
            End If

            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)
                Dim X1, Y1 As Double
                Dim LayerTable1 As Autodesk.AutoCAD.DatabaseServices.LayerTable = Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                If LayerTable1.Has(Layer_property) = True Then
                    Dim Layer1 As Autodesk.AutoCAD.DatabaseServices.LayerTableRecord = Trans1.GetObject(LayerTable1(Layer_property), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    Dim Filtru1(0) As Autodesk.AutoCAD.DatabaseServices.TypedValue
                    Filtru1(0) = New Autodesk.AutoCAD.DatabaseServices.TypedValue(Autodesk.AutoCAD.DatabaseServices.DxfCode.LayerName, Layer1.Name)
                    Dim Selection_Filter1 As New Autodesk.AutoCAD.EditorInput.SelectionFilter(Filtru1)
                    Dim Selection_result1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Selection_result1 = Editor1.SelectAll(Selection_Filter1)
                    Dim Selset1 As Autodesk.AutoCAD.EditorInput.SelectionSet
                    Selset1 = Selection_result1.Value
                    If Not Selset1 Is Nothing Then
                        If Selset1.Count > 0 Then

                            Dim Table_data1 As New System.Data.DataTable
                            Table_data1.Columns.Add("OBJECT_ID", GetType(ObjectId))
                            Table_data1.Columns.Add("START_POINT", GetType(Point3d))
                            Table_data1.Columns.Add("START_CHAINAGE", GetType(Double))
                            Table_data1.Columns.Add("END_POINT", GetType(Point3d))
                            Table_data1.Columns.Add("END_CHAINAGE", GetType(Double))
                            Table_data1.Columns.Add("OWNER1", GetType(String))
                            Table_data1.Columns.Add("OWNER2", GetType(String))
                            Table_data1.Columns.Add("PARCEL_CENTER_POINT", GetType(Point3d))
                            Table_data1.Columns.Add("MID_POINT", GetType(Point3d))

                            Dim Index_table As Double = 0
                            Table_data1.Rows.Add()
                            Table_data1.Rows(Index_table).Item("START_POINT") = Poly1.StartPoint
                            Table_data1.Rows(Index_table).Item("START_CHAINAGE") = 0

                            For i = 0 To Selset1.Count - 1
                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Selset1.Item(i)
                                Dim Ent1 As Autodesk.AutoCAD.DatabaseServices.Entity
                                Ent1 = Trans1.GetObject(Obj1.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                If TypeOf Ent1 Is Polyline Then
                                    Dim Poly2 As Polyline = Ent1

                                    Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                                    Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables
                                    Dim Id1 As ObjectId = Ent1.ObjectId
                                    Dim Records1 As Autodesk.Gis.Map.ObjectData.Records
                                    If Not Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False).Count = 0 Then
                                        Records1 = Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                        Dim Record1 As Autodesk.Gis.Map.ObjectData.Record
                                        Dim Valoare_record1 As Autodesk.Gis.Map.Utilities.MapValue

                                        Dim Col_int As New Point3dCollection

                                        Poly1.IntersectWith(Poly2, Intersect.OnBothOperands, Col_int, IntPtr.Zero, IntPtr.Zero)
                                        If Col_int.Count > 0 Then
                                            Table_data1.Rows.Add()
                                            Index_table = Index_table + 1
                                            Table_data1.Rows(Index_table).Item("OBJECT_ID") = Ent1.ObjectId
                                            Dim DBCol1 As New DBObjectCollection
                                            DBCol1.Add(Poly2)
                                            Dim DBCol2 As New DBObjectCollection
                                            DBCol2 = Region.CreateFromCurves(DBCol1)
                                            Dim Region1 As Autodesk.AutoCAD.DatabaseServices.Region
                                            Region1 = DBCol2(0)
                                            Dim Solid3D As New Solid3d
                                            Solid3D.Extrude(Region1, 1, 0)

                                            Table_data1.Rows(Index_table).Item("PARCEL_CENTER_POINT") = New Point3d(Solid3D.MassProperties.Centroid.X, Solid3D.MassProperties.Centroid.Y, Poly2.Elevation)


                                            Dim No_int As Double = Col_int.Count

                                            Select Case No_int
                                                Case 2
                                                    For j = 0 To Col_int.Count - 1
                                                        If j = 0 Then
                                                            Table_data1.Rows(Index_table).Item("START_POINT") = Col_int(j)
                                                            Table_data1.Rows(Index_table).Item("START_CHAINAGE") = Poly1.GetDistAtPoint(Col_int(j))
                                                        End If
                                                        If j = 1 Then
                                                            Table_data1.Rows(Index_table).Item("END_POINT") = Col_int(j)
                                                            Table_data1.Rows(Index_table).Item("END_CHAINAGE") = Poly1.GetDistAtPoint(Col_int(j))
                                                        End If


                                                        For Each Record1 In Records1
                                                            Dim Tabla1 As Autodesk.Gis.Map.ObjectData.Table
                                                            Tabla1 = Tables1(Record1.TableName)
                                                            If Tabla1.Name.ToUpper = "E_BDY_PROPERTY" Then
                                                                Dim Field_defs1 As Autodesk.Gis.Map.ObjectData.FieldDefinitions
                                                                Field_defs1 = Tabla1.FieldDefinitions
                                                                For k = 0 To Record1.Count - 1
                                                                    Dim Field_def1 As Autodesk.Gis.Map.ObjectData.FieldDefinition
                                                                    Field_def1 = Field_defs1(k)
                                                                    If Field_def1.Name.ToUpper = "OWNNAME1" Then
                                                                        Valoare_record1 = Record1(k)
                                                                        If Not Replace(Valoare_record1.StrValue, " ", "") = "" Then
                                                                            Table_data1.Rows(Index_table).Item("OWNER1") = Valoare_record1.StrValue
                                                                        End If
                                                                    End If

                                                                    If Field_def1.Name.ToUpper = "OWNNAME2" Then
                                                                        Valoare_record1 = Record1(k)
                                                                        If Not Replace(Valoare_record1.StrValue, " ", "") = "" Then
                                                                            Table_data1.Rows(Index_table).Item("OWNER2") = Valoare_record1.StrValue
                                                                        End If
                                                                    End If
                                                                Next
                                                            End If
                                                        Next
                                                    Next

                                                Case 1
                                                    For j = 0 To Col_int.Count - 1
                                                        If IsDBNull(Table_data1.Rows(Index_table).Item("START_POINT")) = True Then
                                                            Table_data1.Rows(Index_table).Item("START_POINT") = Col_int(j)
                                                            If IsDBNull(Table_data1.Rows(Index_table).Item("START_CHAINAGE")) = True Then Table_data1.Rows(Index_table).Item("START_CHAINAGE") = Poly1.GetDistAtPoint(Col_int(j))
                                                        Else
                                                            If IsDBNull(Table_data1.Rows(Index_table).Item("END_POINT")) = True Then Table_data1.Rows(Index_table).Item("END_POINT") = Col_int(j)
                                                            If IsDBNull(Table_data1.Rows(Index_table).Item("END_CHAINAGE")) = True Then Table_data1.Rows(Index_table).Item("END_CHAINAGE") = Poly1.GetDistAtPoint(Col_int(j))
                                                        End If

                                                        For Each Record1 In Records1
                                                            Dim Tabla1 As Autodesk.Gis.Map.ObjectData.Table
                                                            Tabla1 = Tables1(Record1.TableName)
                                                            If Tabla1.Name.ToUpper = "E_BDY_PROPERTY" Then
                                                                Dim Field_defs1 As Autodesk.Gis.Map.ObjectData.FieldDefinitions
                                                                Field_defs1 = Tabla1.FieldDefinitions
                                                                For k = 0 To Record1.Count - 1
                                                                    Dim Field_def1 As Autodesk.Gis.Map.ObjectData.FieldDefinition
                                                                    Field_def1 = Field_defs1(k)
                                                                    If Field_def1.Name.ToUpper = "OWNNAME1" Then
                                                                        Valoare_record1 = Record1(k)
                                                                        If Not Replace(Valoare_record1.StrValue, " ", "") = "" Then
                                                                            Table_data1.Rows(Index_table).Item("OWNER1") = Valoare_record1.StrValue
                                                                        End If
                                                                    End If

                                                                    If Field_def1.Name.ToUpper = "OWNNAME2" Then
                                                                        Valoare_record1 = Record1(k)
                                                                        If Not Replace(Valoare_record1.StrValue, " ", "") = "" Then
                                                                            Table_data1.Rows(Index_table).Item("OWNER2") = Valoare_record1.StrValue
                                                                        End If
                                                                    End If
                                                                Next
                                                            End If
                                                        Next
                                                    Next

                                                Case Else
                                                    For j = 0 To Col_int.Count - 1
                                                        If j = 0 Then
                                                            Table_data1.Rows(Index_table).Item("START_POINT") = Col_int(j)
                                                            Table_data1.Rows(Index_table).Item("START_CHAINAGE") = Poly1.GetDistAtPoint(Col_int(j))
                                                        End If
                                                        If j = 1 Then
                                                            Table_data1.Rows(Index_table).Item("END_POINT") = Col_int(j)
                                                            Table_data1.Rows(Index_table).Item("END_CHAINAGE") = Poly1.GetDistAtPoint(Col_int(j))
                                                        End If
                                                        If j = 3 Then
                                                            Table_data1.Rows.Add()
                                                            Index_table = Index_table + 1
                                                            Table_data1.Rows(Index_table).Item("OBJECT_ID") = Ent1.ObjectId
                                                            Table_data1.Rows(Index_table).Item("START_POINT") = Col_int(j)
                                                            Table_data1.Rows(Index_table).Item("START_CHAINAGE") = Poly1.GetDistAtPoint(Col_int(j))
                                                            Table_data1.Rows(Index_table).Item("PARCEL_CENTER_POINT") = New Point3d(Solid3D.MassProperties.Centroid.X, Solid3D.MassProperties.Centroid.Y, Poly2.Elevation)
                                                        End If
                                                        If j = 4 Then
                                                            Table_data1.Rows(Index_table).Item("END_POINT") = Col_int(j)
                                                            Table_data1.Rows(Index_table).Item("END_POINT") = Poly1.GetDistAtPoint(Col_int(j))
                                                        End If
                                                        If j = 5 Then
                                                            Table_data1.Rows.Add()
                                                            Index_table = Index_table + 1
                                                            Table_data1.Rows(Index_table).Item("OBJECT_ID") = Ent1.ObjectId
                                                            Table_data1.Rows(Index_table).Item("START_POINT") = Col_int(j)
                                                            Table_data1.Rows(Index_table).Item("START_CHAINAGE") = Poly1.GetDistAtPoint(Col_int(j))
                                                            Table_data1.Rows(Index_table).Item("PARCEL_CENTER_POINT") = New Point3d(Solid3D.MassProperties.Centroid.X, Solid3D.MassProperties.Centroid.Y, Poly2.Elevation)
                                                        End If
                                                        If j = 6 Then
                                                            Table_data1.Rows(Index_table).Item("END_POINT") = Col_int(j)
                                                            Table_data1.Rows(Index_table).Item("END_POINT") = Poly1.GetDistAtPoint(Col_int(j))
                                                        End If
                                                        If j = 7 Then
                                                            Table_data1.Rows.Add()
                                                            Index_table = Index_table + 1
                                                            Table_data1.Rows(Index_table).Item("OBJECT_ID") = Ent1.ObjectId
                                                            Table_data1.Rows(Index_table).Item("START_POINT") = Col_int(j)
                                                            Table_data1.Rows(Index_table).Item("START_CHAINAGE") = Poly1.GetDistAtPoint(Col_int(j))
                                                            Table_data1.Rows(Index_table).Item("PARCEL_CENTER_POINT") = New Point3d(Solid3D.MassProperties.Centroid.X, Solid3D.MassProperties.Centroid.Y, Poly2.Elevation)
                                                        End If
                                                        If j = 8 Then
                                                            Table_data1.Rows(Index_table).Item("END_POINT") = Col_int(j)
                                                            Table_data1.Rows(Index_table).Item("END_POINT") = Poly1.GetDistAtPoint(Col_int(j))
                                                        End If

                                                        For Each Record1 In Records1
                                                            Dim Tabla1 As Autodesk.Gis.Map.ObjectData.Table
                                                            Tabla1 = Tables1(Record1.TableName)
                                                            If Tabla1.Name.ToUpper = "E_BDY_PROPERTY" Then
                                                                Dim Field_defs1 As Autodesk.Gis.Map.ObjectData.FieldDefinitions
                                                                Field_defs1 = Tabla1.FieldDefinitions
                                                                For k = 0 To Record1.Count - 1
                                                                    Dim Field_def1 As Autodesk.Gis.Map.ObjectData.FieldDefinition
                                                                    Field_def1 = Field_defs1(k)
                                                                    If Field_def1.Name.ToUpper = "OWNNAME1" Then
                                                                        Valoare_record1 = Record1(k)
                                                                        If Not Replace(Valoare_record1.StrValue, " ", "") = "" Then
                                                                            Table_data1.Rows(Index_table).Item("OWNER1") = Valoare_record1.StrValue
                                                                        End If
                                                                    End If

                                                                    If Field_def1.Name.ToUpper = "OWNNAME2" Then
                                                                        Valoare_record1 = Record1(k)
                                                                        If Not Replace(Valoare_record1.StrValue, " ", "") = "" Then
                                                                            Table_data1.Rows(Index_table).Item("OWNER2") = Valoare_record1.StrValue
                                                                        End If
                                                                    End If
                                                                Next
                                                            End If
                                                        Next
                                                    Next



                                            End Select

                                            Table_data1.Rows(Index_table).Item("END_POINT") = Poly1.EndPoint
                                            Table_data1.Rows(Index_table).Item("END_CHAINAGE") = Poly1.Length

                                        End If
                                    End If
                                End If
                            Next


                            If Table_data1.Rows.Count > 0 Then
                                Dim DataView1 As New DataView(Table_data1)
                                DataView1.Sort = "START_CHAINAGE"
                                Dim Rotatie_text As Double = PI / 4

                                For i = 0 To DataView1.Count - 1

                                    If IsDBNull(DataView1.Item(i)("OWNER1")) = False Then
                                        If Not DataView1.Item(i)("OWNER1") = "" Then

                                            Dim Text1 As New DBText
                                            If IsDBNull(DataView1.Item(i)("PARCEL_CENTER_POINT")) = False Then
                                                Text1.Position = DataView1.Item(i)("PARCEL_CENTER_POINT")
                                                Text1.TextString = DataView1.Item(i)("OWNER1")
                                                Text1.Height = 2
                                                Text1.Rotation = Rotatie_text

                                                BTrecord.AppendEntity(Text1)
                                                Trans1.AddNewlyCreatedDBObject(Text1, True)
                                            End If

                                        End If
                                    End If


                                Next

                            End If




                        End If
                    End If
                End If


                Trans1.Commit()
            End Using


            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try



    End Sub

    <CommandMethod("ucs_align")> _
    Public Sub ucs_align_to_polyline()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Try


            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt.MessageForAdding = vbLf & "Select blue sky rectangle:"

            Rezultat1 = Editor1.GetSelection(Object_Prompt)

            If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                Exit Sub

            End If



            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        If Rezultat1.Value.Count = 2 Then

                            Dim CurentUCSmatrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
                            Dim CurentUCS As CoordinateSystem3d = CurentUCSmatrix.CoordinateSystem3d
                            Dim Point0 As New Point3d
                            Dim Point1 As New Point3d
                            Dim PointM As New Point3d
                            Dim ucs_NAME As String = ""

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




                            Dim Ucs_table As UcsTable = Trans1.GetObject(ThisDrawing.Database.UcsTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                            Dim View_Table As ViewTable = Trans1.GetObject(ThisDrawing.Database.ViewTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                            If Not Point0.X = 0 And Not Point0.Y = 0 And Not Point1.X = 0 And Not Point1.Y = 0 And Not PointM.X = 0 And Not PointM.Y = 0 Then
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




                        End If

                        Trans1.Commit()
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
    End Sub

    <CommandMethod("PPL_RECHAINAGE_EXCEL")> _
    Public Sub Show_POINT_RECHAINAGE_EXCEL_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is New_chainage_calc_from_excel Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New New_chainage_calc_from_excel
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    <CommandMethod("clintersector")> _
    Public Sub Show_intersect_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Intersection_Form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Intersection_Form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    <CommandMethod("PPL_SEC_FENCE")> _
    Public Sub Show_protection_fence_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Protection_fence_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Protection_fence_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    <CommandMethod("GE")> _
    Public Sub Add_to_clipboard_lat_longs_AUTOCAD_ENGINE()
        If isSECURE() = False Then Exit Sub
        Try
            'Gis.Map.Platform.AcMapMap
            Dim Acmap As Autodesk.Gis.Map.Platform.AcMapMap = Autodesk.Gis.Map.Platform.AcMapMap.GetCurrentMap
            Dim Curent_system As String = Acmap.GetMapSRS()
            If String.IsNullOrEmpty(Curent_system) = True Then
                MsgBox("Please set your coordinate system", , "Dan says...")
                Exit Sub
            End If
            Dim String_UTM83_12 As String = "PROJCS[" & Chr(34) & "UTM83-12" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL83" & Chr(34) & ",DATUM[" & Chr(34) & "NAD83" & Chr(34) & ",SPHEROID[" & Chr(34) & "GRS1980" & Chr(34) & ",6378137.000,298.25722210]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",500000.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-111.00000000000000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.9996],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
            Dim String_UTM83_11 As String = "PROJCS[" & Chr(34) & "UTM83-11" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL83" & Chr(34) & ",DATUM[" & Chr(34) & "NAD83" & Chr(34) & ",SPHEROID[" & Chr(34) & "GRS1980" & Chr(34) & ",6378137.000,298.25722210]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",500000.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-117.00000000000000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.9996],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
            Dim String_CANA83_10TM115 As String = "PROJCS[" & Chr(34) & "CANA83-10TM115" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL83" & Chr(34) & ",DATUM[" & Chr(34) & "NAD83" & Chr(34) & ",SPHEROID[" & Chr(34) & "GRS1980" & Chr(34) & ",6378137.000,298.25722210]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.999200000000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-115.00000000000000],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.00000000000000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
            Dim String_CANA83_3TM114 As String = "PROJCS[" & Chr(34) & "CANA83-3TM114" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL83" & Chr(34) & ",DATUM[" & Chr(34) & "NAD83" & Chr(34) & ",SPHEROID[" & Chr(34) & "GRS1980" & Chr(34) & ",6378137.000,298.25722210]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.999900000000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-114.00000000000000],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.00000000000000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
            Dim String_CANA83_3TM111 As String = "PROJCS[" & Chr(34) & "CANA83-3TM111" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL83" & Chr(34) & ",DATUM[" & Chr(34) & "NAD83" & Chr(34) & ",SPHEROID[" & Chr(34) & "GRS1980" & Chr(34) & ",6378137.000,298.25722210]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.999900000000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-111.00000000000000],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.00000000000000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
            Dim String_CANA83_3TM117 As String = "PROJCS[" & Chr(34) & "CANA83-3TM117" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL83" & Chr(34) & ",DATUM[" & Chr(34) & "NAD83" & Chr(34) & ",SPHEROID[" & Chr(34) & "GRS1980" & Chr(34) & ",6378137.000,298.25722210]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.999900000000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-117.00000000000000],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.00000000000000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
            Dim String_CANA83_3TM120 As String = "PROJCS[" & Chr(34) & "CANA83-3TM120" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL83" & Chr(34) & ",DATUM[" & Chr(34) & "NAD83" & Chr(34) & ",SPHEROID[" & Chr(34) & "GRS1980" & Chr(34) & ",6378137.000,298.25722210]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.999900000000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-120.00000000000000],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.00000000000000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
            Dim String_UTM27_11 As String = "PROJCS[" & Chr(34) & "UTM27-11" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL27" & Chr(34) & ",DATUM[" & Chr(34) & "NAD27" & Chr(34) & ",SPHEROID[" & Chr(34) & "CLRK66" & Chr(34) & ",6378206.400,294.97869821]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",500000.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-117.00000000000000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.9996],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
            Dim String_UTM27_12 As String = "PROJCS[" & Chr(34) & "UTM27-12" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL27" & Chr(34) & ",DATUM[" & Chr(34) & "NAD27" & Chr(34) & ",SPHEROID[" & Chr(34) & "CLRK66" & Chr(34) & ",6378206.400,294.97869821]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",500000.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-111.00000000000000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.9996],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
            Dim String_LL84 As String = "GEOGCS[" & Chr(34) & "LL84" & Chr(34) & ",DATUM[" & Chr(34) & "WGS84" & Chr(34) & ",SPHEROID[" & Chr(34) & "WGS84" & Chr(34) & ",6378137.000,298.25722293]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.01745329251994]]"
            Dim String_LL83 As String = "GEOGCS[" & Chr(34) & "LL83" & Chr(34) & ",DATUM[" & Chr(34) & "NAD83" & Chr(34) & ",SPHEROID[" & Chr(34) & "GRS1980" & Chr(34) & ",6378137.000,298.25722210]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.01745329251994]]"

            Dim Coord_factory1 As New OSGeo.MapGuide.MgCoordinateSystemFactory
            Dim CoordSys1 As OSGeo.MapGuide.MgCoordinateSystem = Coord_factory1.Create(Curent_system)
            Dim CoordSys2 As OSGeo.MapGuide.MgCoordinateSystem = Coord_factory1.Create(String_LL84)
            Dim Transform1 As OSGeo.MapGuide.MgCoordinateSystemTransform = Coord_factory1.GetTransform(CoordSys1, CoordSys2)



            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
            Dim Database1 As Autodesk.AutoCAD.DatabaseServices.Database
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor


            ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Database1 = ThisDrawing.Database
            Editor1 = ThisDrawing.Editor

            Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult

            Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions("Specify point:")

            Point1 = Editor1.GetPoint(PP1)
            If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                Exit Sub
            End If

            Dim x1 As Double = Point1.Value.X
            Dim y1 As Double = Point1.Value.Y
            Dim Coord1 As OSGeo.MapGuide.MgCoordinate = Transform1.Transform(x1, y1)
            Dim Lat1 As Double = Coord1.Y
            Dim Long1 As Double = Coord1.X




            Windows.Forms.Clipboard.SetDataObject(Lat1 & Chr(176) & "," & Long1 & Chr(176))

        Catch ex As Exception

            MsgBox(ex.Message)

        End Try



    End Sub
    <CommandMethod("PPL_COORD_CONV")> _
    Public Sub Show_coord_sys_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Coordinate_systems_in_Excel Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Coordinate_systems_in_Excel
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    <CommandMethod("PPL_highlight")> _
    Public Sub highlight()
        Try

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
            Editor1 = ThisDrawing.Editor
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

            Using lock As DocumentLock = ThisDrawing.LockDocument

                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select objects containing object data:"
                Object_Prompt.SingleOnly = False
                Rezultat1 = Editor1.GetSelection(Object_Prompt)



                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    If IsNothing(Rezultat1) = False Then
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                            Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables

                            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                            BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)

                            For i = 0 To Rezultat1.Value.Count - 1
                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Rezultat1.Value.Item(i)
                                Dim Ent1 As Entity
                                Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                                Dim Id1 As ObjectId = Ent1.ObjectId
                                Dim Records1 As Autodesk.Gis.Map.ObjectData.Records
                                If Not Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False).Count = 0 Then

                                    Records1 = Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                    Dim Record1 As Autodesk.Gis.Map.ObjectData.Record
                                    For Each Record1 In Records1
                                        Dim Tabla1 As Autodesk.Gis.Map.ObjectData.Table
                                        Tabla1 = Tables1(Record1.TableName)

                                        Dim Field_defs1 As Autodesk.Gis.Map.ObjectData.FieldDefinitions
                                        Field_defs1 = Tabla1.FieldDefinitions
                                        For ii = 0 To Record1.Count - 1
                                            Dim Field_def1 As Autodesk.Gis.Map.ObjectData.FieldDefinition
                                            Field_def1 = Field_defs1(ii)
                                            Dim Nume_field As String = Field_def1.Name.ToUpper
                                            Dim Valoare_field As String = Record1(ii).StrValue
                                            If Nume_field.ToUpper.Contains("DESC") = True Then
                                                If Valoare_field.ToUpper.Contains("O") = True Or Valoare_field.ToUpper.Contains("FA") = True Then
                                                    Ent1.UpgradeOpen()
                                                    Ent1.ColorIndex = 1
                                                    Exit For
                                                End If
                                            End If



                                        Next

                                    Next
                                End If
                            Next







                            Editor1.Regen()
                            Trans1.Commit()
                        End Using
                    End If
                End If
            End Using


            MsgBox("Done")
            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")


        Catch ex As Exception

            MsgBox(ex.Message)
        End Try
    End Sub


    <CommandMethod("PPL_od2block")> _
    Public Sub Show_object_data_to_block_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Object_data_to_block_Form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next

        Try
            Dim forma1 As New Object_data_to_block_Form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    <CommandMethod("PPL_CROSSBAND")> _
    Public Sub SHOW_Crossing_Band_Form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Crossing_Band_Form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Crossing_Band_Form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub



End Class
